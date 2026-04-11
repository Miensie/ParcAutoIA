"""
FleetInsight AI — Smart Workbook Parser v2.2
=============================================
Corrections :
  - Conversion robuste des dates Excel sérielles (bug 1970 corrigé)
  - Support : serial Excel, ISO string, datetime Python, timestamp JS ms
  - Inclusion Gpe électrogène
"""

from __future__ import annotations
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from typing import Any
import warnings
warnings.filterwarnings("ignore")

EXCEL_EPOCH = datetime(1899, 12, 31)

# ─────────────────────────────────────────────────────────────────────────────
# CONVERSION DE DATES — ROBUSTE (bug 1970 corrigé)
# ─────────────────────────────────────────────────────────────────────────────

def _parse_date(v: Any) -> datetime | None:
    """
    Convertit n'importe quelle valeur en datetime.
    Cas traités :
      1. datetime Python natif
      2. Numéro sériel Excel  (int/float entre 1 et 109 574)
         ex: 46044 -> 02/01/2026
         Correction : EXCEL_EPOCH = 1899-12-31 (pas 12-30)
         Formule : EPOCH + timedelta(days=serial)
      3. Timestamp JavaScript en ms  (>= 1_000_000_000_000)
      4. Chaîne de date (ISO, française DD/MM/YYYY, etc.)
    """
    if v is None:
        return None

    # 1. Déjà un datetime Python
    if isinstance(v, datetime):
        return v

    # 2 & 3. Numérique
    if isinstance(v, (int, float)):
        try:
            if pd.isna(v):
                return None
        except Exception:
            pass
        val = float(v)

        # Timestamp JavaScript en millisecondes (> 10^12)
        if val >= 1_000_000_000_000:
            try:
                return datetime.utcfromtimestamp(val / 1000)
            except Exception:
                return None

        # Numéro sériel Excel : plage 1..109574 (1900-2199)
        # Serial 1 = 1900-01-01, Serial 2 = 1900-01-02, etc.
        if 1 <= val <= 109_574:
            try:
                days = int(val) - 1  # Correction : Excel utilise offset 1, pas 0
                # EPOCH(1899-12-31) + (days-1) donne la bonne date
                return EXCEL_EPOCH + timedelta(days=days)
            except Exception:
                return None

        # Timestamp Unix en secondes (plage moderne ~2000-2100)
        # 946_684_800 = 2000-01-01,  4_102_444_800 = 2100-01-01
        if 946_684_800 <= val <= 4_102_444_800:
            try:
                return datetime.utcfromtimestamp(val)
            except Exception:
                return None

        return None

    # 4. Chaîne
    if isinstance(v, str):
        s = v.strip()
        if not s or s.lower() in ("none", "null", "nan", ""):
            return None
        for fmt in (
            "%Y-%m-%dT%H:%M:%S.%f",
            "%Y-%m-%dT%H:%M:%S",
            "%Y-%m-%dT%H:%M",
            "%Y-%m-%d",
            "%d/%m/%Y",
            "%d-%m-%Y",
            "%m/%d/%Y",
            "%d/%m/%y",
            "%d %B %Y",
        ):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
        try:
            return pd.to_datetime(s, dayfirst=True).to_pydatetime()
        except Exception:
            return None

    return None


def _fmt_date(v: Any) -> str:
    d = _parse_date(v)
    return d.strftime("%d/%m/%Y") if d else "—"


# ─────────────────────────────────────────────────────────────────────────────
# MAPPING DES FEUILLES
# ─────────────────────────────────────────────────────────────────────────────

SHEET_ALIASES: dict[str, list[str]] = {
    "vehicules":  ["LISTE DE VEH","LISTE VEH","VEHICULES","PARC","INVENTAIRE"],
    "carburant":  ["SUIVI_CARBURANT","CARBURANT","FUEL","SUIVI CARBURANT"],
    "entretien":  ["ENTRETIEN ET REPARATIONS","ENTRETIEN","MAINTENANCE",
                   "ENTRETIEN & REPARATIONS","REPARATIONS"],
    "sorties":    ["SORTIES VEH","SORTIES","MISSIONS","SORTIES VEHICULES"],
    "vt":         ["VT","VISITE TECHNIQUE","VISITES TECHNIQUES","CONTROLE TECHNIQUE"],
    "generateur": ["GPE ÉLECTROGÈNE","GPE ELECTROGENE","GENERATEUR",
                   "GROUPE ELECTROGENE","Gpe électrogène","GE",
                   "GROUPE ELECTROG","ELECTROGENE"],
}

JUNK = ["total général","total general","étiquettes de lignes","étiquettes de colonnes",
        "(vide)","somme de","nombre de","grand total","sous-total"]


def _is_junk(row):
    for v in row:
        if v is None: continue
        s = str(v).lower().strip()
        if any(k in s for k in JUNK): return True
    return False


def _non_null(row):
    return sum(1 for v in row if v is not None and str(v).strip())


def _header_idx(rows):
    for i, row in enumerate(rows):
        if _is_junk(row): continue
        nn = [v for v in row if v is not None and str(v).strip()]
        if len(nn) >= 2 and sum(1 for v in nn if isinstance(v, str)) >= len(nn)*0.5:
            return i
    return 0


def _build_df(raw_rows, h_idx):
    hrow = raw_rows[h_idx]
    headers, seen = [], {}
    for v in hrow:
        n = str(v).strip() if (v is not None and str(v).strip()) else "_"
        seen.setdefault(n, 0)
        headers.append(n if seen[n] == 0 else f"{n}_{seen[n]}")
        seen[n] += 1

    data = []
    for row in raw_rows[h_idx+1:]:
        if _is_junk(row) or _non_null(row) == 0: continue
        padded = list(row) + [None]*max(0, len(headers)-len(row))
        data.append(dict(zip(headers, padded[:len(headers)])))

    df = pd.DataFrame(data).dropna(how="all")
    # Conversion numérique prudente (évite de convertir les sériels de dates)
    for col in df.columns:
        if col.startswith("_"): continue
        try:
            c = pd.to_numeric(df[col], errors="coerce")
            # Pas de conversion si ressemble à sériels de dates (40k-80k)
            if c.notna().sum() >= len(df)*0.5 and not c.dropna().between(30000,80000).all():
                df[col] = c
        except Exception:
            pass
    return df


# ─────────────────────────────────────────────────────────────────────────────
# CLASSE PRINCIPALE
# ─────────────────────────────────────────────────────────────────────────────

class FleetParser:
    def __init__(self, raw_sheets: dict[str, list[dict]]):
        self.raw = raw_sheets
        self.dfs: dict[str, pd.DataFrame] = {}
        self.sheet_map: dict[str, str] = {}
        self._parse()

    def _resolve(self, key):
        aliases = [a.upper() for a in SHEET_ALIASES.get(key, [])]
        for name in self.raw:
            c = name.strip().upper()
            if any(a == c or a in c or c in a for a in aliases):
                return name
        return None

    def _to_rows(self, rows):
        result = []
        for row in rows:
            try:
                mi = max(int(k) for k in row.keys())
                vals = [row.get(str(i)) for i in range(mi+1)]
            except Exception:
                vals = list(row.values())
            result.append(vals)
        return result

    def _parse_sheet(self, rows):
        if not rows: return pd.DataFrame()
        raw = self._to_rows(rows)
        return _build_df(raw, _header_idx(raw))

    def _parse(self):
        for key in SHEET_ALIASES:
            name = self._resolve(key)
            if name:
                self.sheet_map[key] = name
                self.dfs[key] = self._parse_sheet(self.raw[name])
        recognized = set(self.sheet_map.values())
        for name, rows in self.raw.items():
            if name not in recognized:
                self.dfs[f"__extra__{name}"] = self._parse_sheet(rows)

    def get(self, key) -> pd.DataFrame:
        return self.dfs.get(key, pd.DataFrame())

    def summary(self) -> dict:
        return {
            "feuilles_reconnues": list(self.sheet_map.keys()),
            "feuilles_total":     len(self.raw),
            "mapping":            dict(self.sheet_map),
        }

    @staticmethod
    def find_col(df, *kws):
        for kw in kws:
            for col in df.columns:
                if kw.upper() in col.upper():
                    return col
        return None

    @staticmethod
    def parse_date_col(df, col):
        """Applique _parse_date sur toute la colonne."""
        return df[col].apply(_parse_date)

    @staticmethod
    def fmt_date(v):
        return _fmt_date(v)