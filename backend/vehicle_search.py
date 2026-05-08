"""
FleetInsight IA — Vehicle Search Module
=========================================
Recherche toutes les informations d'un véhicule
par immatriculation dans toutes les feuilles du classeur.
"""

from __future__ import annotations
import pandas as pd
import numpy as np
from datetime import datetime
from fleet_parser import FleetParser, _fmt_date, _parse_date
import warnings
warnings.filterwarnings("ignore")

TODAY = datetime.now()


def _num(s): return pd.to_numeric(s, errors="coerce")
def _str(v): return str(v).strip() if v is not None and str(v).strip() not in ("None","nan","") else "—"
def _fmt(n):
    if n is None or (isinstance(n, float) and np.isnan(n)): return "—"
    n = float(n)
    if n >= 1_000_000: return f"{n/1_000_000:.1f}M FCFA"
    if n >= 1_000:     return f"{n/1_000:.0f}K FCFA"
    return f"{n:.0f} FCFA"


def _match(val, immat: str) -> bool:
    """Correspondance flexible : insensible à la casse et aux espaces multiples."""
    if val is None: return False
    v = " ".join(str(val).upper().split())
    q = " ".join(immat.upper().split())
    return v == q or q in v or v in q


class VehicleSearchEngine:
    """
    Moteur de recherche d'un véhicule par immatriculation.
    Croise toutes les feuilles du classeur et retourne
    un profil complet du véhicule.
    """

    def __init__(self, parser: FleetParser):
        self.p = parser

    def search(self, immat: str) -> dict:
        """
        Recherche un véhicule par immatriculation.
        Retourne None si non trouvé, sinon le profil complet.
        """
        immat = " ".join(immat.strip().upper().split())
        if not immat:
            return {"found": False, "error": "Immatriculation vide."}

        # 1. Fiche véhicule (LISTE DE VEH)
        fiche = self._get_fiche(immat)

        # 2. Carburant
        carburant = self._get_carburant(immat)

        # 3. Entretien & Réparations
        entretien = self._get_entretien(immat)

        # 4. VT
        vt = self._get_vt(immat)

        # 5. Sorties / Missions
        sorties = self._get_sorties(immat)

        # Si aucune donnée trouvée dans aucune feuille → véhicule inconnu
        if (not fiche and not carburant["transactions"]
                and not entretien["interventions"]
                and vt is None and not sorties["missions"]):
            # Chercher les immatriculations proches (suggestions)
            suggestions = self._suggest(immat)
            return {
                "found": False,
                "immat_recherchee": immat,
                "error": f"Aucun véhicule trouvé pour '{immat}'.",
                "suggestions": suggestions,
            }

        # Score santé simplifié basé sur les données disponibles
        score_info = self._compute_score(carburant, entretien, vt)

        return {
            "found":     True,
            "immat":     immat,
            "fiche":     fiche,
            "carburant": carburant,
            "entretien": entretien,
            "vt":        vt,
            "sorties":   sorties,
            "score":     score_info,
        }

    # ── Fiche véhicule ────────────────────────────────────────────────────

    def _get_fiche(self, immat: str) -> dict | None:
        df = self.p.get("vehicules")
        if df.empty: return None
        fc = self.p.find_col
        immat_col  = fc(df, "IMMAT")
        marque_col = fc(df, "MARQUE")
        type_col   = fc(df, "TYPE")
        no_col     = fc(df, "N°", "NO", "NUM")

        if not immat_col: return None

        mask = df[immat_col].apply(lambda v: _match(v, immat))
        row  = df[mask]
        if row.empty: return None
        r = row.iloc[0]

        return {
            "immatriculation": _str(r.get(immat_col)),
            "marque":          _str(r.get(marque_col)) if marque_col else "—",
            "type":            _str(r.get(type_col))   if type_col   else "—",
            "numero":          _str(r.get(no_col))      if no_col    else "—",
        }

    # ── Carburant ─────────────────────────────────────────────────────────

    def _get_carburant(self, immat: str) -> dict:
        df = self.p.get("carburant")
        empty = {"total_depense": 0, "total_fmt": "—",
                 "nb_transactions": 0, "transactions": []}
        if df.empty: return empty
        fc = self.p.find_col
        v_col  = fc(df, "VEHICULE")
        m_col  = fc(df, "MONTANT")
        d_col  = fc(df, "DATE")
        p_col  = fc(df, "PRODUIT")
        mo_col = fc(df, "MOIS")
        c_col  = fc(df, "CARTE")
        if not v_col: return empty

        mask = df[v_col].apply(lambda v: _match(v, immat))
        sub  = df[mask].copy()
        if sub.empty: return empty

        # Montants
        montants = _num(sub[m_col]) if m_col else pd.Series(dtype=float)
        total    = float(montants.sum()) if not montants.empty else 0.0

        # Transactions triées par date
        transactions = []
        for _, r in sub.iterrows():
            d = _parse_date(r.get(d_col)) if d_col else None
            transactions.append({
                "mois":    _str(r.get(mo_col)) if mo_col else "—",
                "date":    _fmt_date(r.get(d_col)) if d_col else "—",
                "produit": _str(r.get(p_col))  if p_col  else "—",
                "carte":   _str(r.get(c_col))  if c_col  else "—",
                "montant": float(_num(r.get(m_col))) if m_col and not pd.isna(_num(r.get(m_col))) else 0,
                "montant_fmt": _fmt(float(_num(r.get(m_col)))) if m_col else "—",
            })

        # Tri par date décroissante
        transactions.sort(key=lambda x: x["date"], reverse=True)

        # Par produit
        par_produit = {}
        if p_col and m_col:
            for _, r in sub.groupby(p_col)[m_col]:
                pass
            gp = sub.groupby(p_col)[m_col].apply(lambda x: _num(x).sum()) if p_col and m_col else {}
            par_produit = {str(k): float(v) for k, v in gp.items()} if hasattr(gp, 'items') else {}

        return {
            "total_depense":    round(total, 0),
            "total_fmt":        _fmt(total),
            "nb_transactions":  len(sub),
            "transactions":     transactions[:30],   # 30 dernières
            "par_produit":      par_produit,
        }

    # ── Entretien ─────────────────────────────────────────────────────────

    def _get_entretien(self, immat: str) -> dict:
        df = self.p.get("entretien")
        empty = {"total_depense": 0, "total_fmt": "—",
                 "nb_interventions": 0, "interventions": [],
                 "en_atelier": False, "duree_atelier": None}
        if df.empty: return empty
        fc = self.p.find_col
        v_col  = fc(df, "VEHICULE")
        d_col  = fc(df, "DEPOT")
        r_col  = fc(df, "RETOUR")
        t_col  = fc(df, "TYPE")
        m_col  = fc(df, "MONTANT")
        g_col  = fc(df, "GARAGE")
        if not v_col: return empty

        mask = df[v_col].apply(lambda v: _match(v, immat))
        sub  = df[mask].copy()
        if sub.empty: return empty

        montants = _num(sub[m_col]) if m_col else pd.Series(dtype=float)
        total    = float(montants.dropna().sum()) if not montants.empty else 0.0

        interventions = []
        en_atelier    = False
        duree_atelier = None

        for _, r in sub.iterrows():
            depot  = _parse_date(r.get(d_col)) if d_col else None
            retour = _parse_date(r.get(r_col)) if r_col else None
            immo   = depot is not None and retour is None
            if immo:
                en_atelier = True
                if depot:
                    duree_atelier = (TODAY - depot).days

            montant = float(_num(r.get(m_col))) if m_col and not pd.isna(_num(r.get(m_col))) else None
            interventions.append({
                "date_depot":  _fmt_date(r.get(d_col)) if d_col else "—",
                "date_retour": _fmt_date(r.get(r_col)) if r_col else "—",
                "type":        _str(r.get(t_col))  if t_col else "—",
                "montant":     montant,
                "montant_fmt": _fmt(montant) if montant else "Non renseigné",
                "garage":      _str(r.get(g_col))  if g_col else "—",
                "en_cours":    immo,
            })

        interventions.sort(key=lambda x: x["date_depot"], reverse=True)

        return {
            "total_depense":     round(total, 0),
            "total_fmt":         _fmt(total),
            "nb_interventions":  len(sub),
            "interventions":     interventions,
            "en_atelier":        en_atelier,
            "duree_atelier":     duree_atelier,
        }

    # ── VT ────────────────────────────────────────────────────────────────

    def _get_vt(self, immat: str) -> dict | None:
        df = self.p.get("vt")
        if df.empty: return None
        fc = self.p.find_col
        v_col   = fc(df, "VEHICULE")
        s_col   = fc(df, "STATUT", "Statuts", "STATUS")
        d_col   = fc(df, "EXPIR", "DATE")
        aff_col = fc(df, "AFFECTATION")
        if not v_col: return None

        mask = df[v_col].apply(lambda v: _match(v, immat))
        row  = df[mask]
        if row.empty: return None
        r = row.iloc[0]

        statut     = _str(r.get(s_col)).title() if s_col else "—"
        expiration = _fmt_date(r.get(d_col)) if d_col else "—"
        aff        = _str(r.get(aff_col)) if aff_col else "—"

        # Calcul jours restants
        jours_restants = None
        date_exp = _parse_date(r.get(d_col)) if d_col else None
        if date_exp:
            jours_restants = (date_exp - TODAY).days

        return {
            "statut":          statut,
            "expiration":      expiration,
            "affectation":     aff,
            "jours_restants":  jours_restants,
            "alerte": (
                "EXPIRÉE"   if jours_restants is not None and jours_restants < 0
                else "URGENTE"  if jours_restants is not None and jours_restants <= 30
                else "VIGILANCE" if jours_restants is not None and jours_restants <= 60
                else "OK"
            ) if statut.lower() == "oui" else (
                "NON_CONFORME" if statut.lower() == "non"
                else "EN_ATTENTE"
            ),
        }

    # ── Sorties ───────────────────────────────────────────────────────────

    def _get_sorties(self, immat: str) -> dict:
        df = self.p.get("sorties")
        empty = {"km_total": 0, "nb_missions": 0, "missions": []}
        if df.empty: return empty
        fc = self.p.find_col
        v_col    = fc(df, "VEHICULE")
        d_col    = fc(df, "DATE")
        kd_col   = fc(df, "KM DEPART", "KM_DEPART", "DEPART")
        kr_col   = fc(df, "KM RETOUR", "KM_RETOUR", "RETOUR")
        km_col   = fc(df, "KM PARCOURU", "KM PARCOURU")
        dest_col = fc(df, "DESTINATION", "DEST")
        cond_col = fc(df, "NOMS", "PRENOMS", "CONDUCTEUR", "CHAUFFEUR")
        mo_col   = fc(df, "MOIS")
        if not v_col: return empty

        mask = df[v_col].apply(lambda v: _match(v, immat))
        sub  = df[mask].copy()
        if sub.empty: return empty

        # KM parcouru : calculer si colonne absente ou invalide
        if kd_col and kr_col:
            kms = (_num(sub[kr_col]) - _num(sub[kd_col])).clip(lower=0)
        elif km_col:
            kms = _num(sub[km_col]).clip(lower=0)
        else:
            kms = pd.Series([0]*len(sub))

        km_total = float(kms.sum())

        # KM retour max = dernier kilométrage connu
        km_dernier = None
        if kr_col:
            vals = _num(sub[kr_col]).dropna()
            if not vals.empty: km_dernier = float(vals.max())

        missions = []
        for i, (_, r) in enumerate(sub.iterrows()):
            km_p = float(kms.iloc[i]) if i < len(kms) else 0
            missions.append({
                "mois":      _str(r.get(mo_col))   if mo_col   else "—",
                "date":      _fmt_date(r.get(d_col)) if d_col  else "—",
                "km_depart": int(_num(r.get(kd_col))) if kd_col and not pd.isna(_num(r.get(kd_col))) else None,
                "km_retour": int(_num(r.get(kr_col))) if kr_col and not pd.isna(_num(r.get(kr_col))) else None,
                "km_parcouru": int(km_p) if km_p > 0 else 0,
                "destination": _str(r.get(dest_col)) if dest_col else "—",
                "conducteur":  _str(r.get(cond_col)) if cond_col else "—",
            })

        missions.sort(key=lambda x: x["date"], reverse=True)

        return {
            "km_total":     int(km_total),
            "km_dernier":   km_dernier,
            "nb_missions":  len(sub),
            "missions":     missions[:30],
        }

    # ── Score santé simplifié ─────────────────────────────────────────────

    def _compute_score(self, carb: dict, entr: dict, vt: dict | None) -> dict:
        """Score santé simplifié pour la fiche véhicule."""
        score = 100
        issues = []

        # VT
        if vt:
            if vt["alerte"] == "NON_CONFORME":
                score -= 30
                issues.append("VT non effectuée (−30 pts)")
            elif vt["alerte"] == "EXPIRÉE":
                score -= 20
                issues.append("VT expirée (−20 pts)")
            elif vt["alerte"] == "URGENTE":
                score -= 10
                issues.append("VT expire bientôt (−10 pts)")

        # En atelier
        if entr["en_atelier"]:
            score -= 20
            d = entr["duree_atelier"]
            issues.append(f"En atelier depuis {d}j (−20 pts)" if d else "En atelier (−20 pts)")

        # Pannes curatives
        prev_kw = ["vidange","révision","revision","contrôle","controle","filtre"]
        nb_pannes = sum(
            1 for i in entr["interventions"]
            if not any(k in i["type"].lower() for k in prev_kw)
            and i["type"] != "—"
        )
        if nb_pannes >= 3:
            score -= 25
            issues.append(f"{nb_pannes} pannes curatives (−25 pts)")
        elif nb_pannes >= 1:
            score -= 10
            issues.append(f"{nb_pannes} panne(s) curative(s) (−10 pts)")

        score = max(0, score)
        if score >= 80:   statut = "OPTIMAL"
        elif score >= 60: statut = "BON"
        elif score >= 40: statut = "SURVEILLER"
        elif score >= 20: statut = "CRITIQUE"
        else:             statut = "REMPLACER"

        return {"score": score, "statut": statut, "details": issues}

    # ── Suggestions ───────────────────────────────────────────────────────

    def _suggest(self, immat: str) -> list[str]:
        """Retourne les immatriculations proches si aucune correspondance exacte."""
        df = self.p.get("vehicules")
        if df.empty: return []
        fc  = self.p.find_col
        col = fc(df, "IMMAT")
        if not col: return []

        q = immat.replace(" ", "").upper()
        suggestions = []
        for v in df[col].dropna():
            v_clean = str(v).replace(" ","").upper()
            if q[:4] in v_clean or v_clean[:4] in q:
                suggestions.append(str(v).strip())
        return suggestions[:5]


def search_vehicle(raw_sheets: dict, immat: str) -> dict:
    """Point d'entrée : parse le classeur et cherche l'immatriculation."""
    from fleet_parser import FleetParser
    import json

    parser = FleetParser(raw_sheets)
    engine = VehicleSearchEngine(parser)
    result = engine.search(immat)

    # Sérialisation JSON-safe
    def safe(o):
        if isinstance(o, (int,)) and not isinstance(o, bool): return int(o)
        if isinstance(o, float):
            if o != o or o == float('inf') or o == float('-inf'): return None
            return round(o, 2)
        if isinstance(o, dict):  return {k: safe(v) for k, v in o.items()}
        if isinstance(o, list):  return [safe(i) for i in o]
        return o

    return safe(result)
