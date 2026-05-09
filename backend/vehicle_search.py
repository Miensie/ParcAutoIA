"""
FleetInsight IA — Vehicle Search Module (CORRIGÉ)
==================================================
Corrections appliquées :
  1. Sort carburant/missions : utilise datetime ISO (pas string formaté)
     → les "—" ne remontent plus en tête de liste
  2. Filtrage lignes vides : exclut les lignes matchées sans aucune donnée utile
     (sous-headers, totaux, lignes orphelines du multi-bloc SUIVI_CARBURANT)
  3. nb_transactions / nb_missions : reflète le vrai compte après filtrage
"""

from __future__ import annotations
import pandas as pd
import numpy as np
from datetime import datetime
from fleet_parser import FleetParser, _fmt_date, _parse_date
import warnings
warnings.filterwarnings("ignore")

TODAY = datetime.now()

# ─── Helpers ──────────────────────────────────────────────────────────────────

def _num(s):
    return pd.to_numeric(s, errors="coerce")

def _str(v):
    return str(v).strip() if v is not None and str(v).strip() not in ("None", "nan", "") else "—"

def _fmt(n):
    if n is None or (isinstance(n, float) and np.isnan(n)):
        return "—"
    n = float(n)
    if n >= 1_000_000: return f"{n/1_000_000:.1f}M FCFA"
    if n >= 1_000:     return f"{n/1_000:.0f}K FCFA"
    return f"{n:.0f} FCFA"

def _match(val, immat: str) -> bool:
    """Correspondance flexible : insensible à la casse et aux espaces multiples."""
    if val is None:
        return False
    v = " ".join(str(val).upper().split())
    q = " ".join(immat.upper().split())
    return v == q or q in v or v in q

def _date_sort_key(date_obj) -> str:
    """
    Clé de tri ISO pour les dates.
    CORRECTION BUG #1 : on ne trie plus sur la string formatée "DD/MM/YYYY"
    car "—" (U+2014) a une valeur Unicode supérieure à tous les chiffres,
    ce qui fait remonter toutes les lignes sans date en haut de la liste
    avec reverse=True.
    Une string ISO "" (vide) trie toujours après toutes les vraies dates.
    """
    if date_obj is None:
        return ""
    try:
        return date_obj.isoformat()
    except Exception:
        return ""


# ─── Moteur de recherche ──────────────────────────────────────────────────────

class VehicleSearchEngine:
    """
    Moteur de recherche d'un véhicule par immatriculation.
    Croise toutes les feuilles du classeur et retourne
    un profil complet du véhicule.
    """

    def __init__(self, parser: FleetParser):
        self.p = parser

    # ── Point d'entrée ────────────────────────────────────────────────────

    def search(self, immat: str) -> dict:
        immat = " ".join(immat.strip().upper().split())
        if not immat:
            return {"found": False, "error": "Immatriculation vide."}

        fiche     = self._get_fiche(immat)
        carburant = self._get_carburant(immat)
        entretien = self._get_entretien(immat)
        vt        = self._get_vt(immat)
        sorties   = self._get_sorties(immat)

        if (not fiche
                and not carburant["transactions"]
                and not entretien["interventions"]
                and vt is None
                and not sorties["missions"]):
            return {
                "found": False,
                "immat_recherchee": immat,
                "error": f"Aucun véhicule trouvé pour '{immat}'.",
                "suggestions": self._suggest(immat),
            }

        return {
            "found":     True,
            "immat":     immat,
            "fiche":     fiche,
            "carburant": carburant,
            "entretien": entretien,
            "vt":        vt,
            "sorties":   sorties,
            "score":     self._compute_score(carburant, entretien, vt),
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
    # CORRECTION BUG #1 + #2

    def _get_carburant(self, immat: str) -> dict:
        df = self.p.get("carburant")
        empty = {"total_depense": 0, "total_fmt": "—",
                 "nb_transactions": 0, "transactions": []}
        if df.empty: return empty

        fc = self.p.find_col
        v_col  = fc(df, "VEHICULE", "IMMAT", "VEH")
        m_col  = fc(df, "MONTANT")
        d_col  = fc(df, "DATE")
        p_col  = fc(df, "PRODUIT")
        mo_col = fc(df, "MOIS")
        c_col  = fc(df, "CARTE", "N° CARTE", "NUMERO CARTE")
        if not v_col: return empty

        mask = df[v_col].apply(lambda v: _match(v, immat))
        sub  = df[mask].copy()
        if sub.empty: return empty

        # Montant total (sur toutes les lignes matchées, même incomplètes)
        montants = _num(sub[m_col]) if m_col else pd.Series(dtype=float)
        total    = float(montants.sum()) if not montants.empty else 0.0

        transactions = []
        for _, r in sub.iterrows():
            date_obj  = _parse_date(r.get(d_col)) if d_col else None
            date_fmt  = _fmt_date(r.get(d_col))   if d_col else "—"
            montant_v = _num(r.get(m_col)) if m_col else float("nan")
            montant_f = float(montant_v) if not pd.isna(montant_v) else 0.0
            produit   = _str(r.get(p_col)) if p_col else "—"
            carte     = _str(r.get(c_col)) if c_col else "—"

            # ── BUG #2 : filtrer les lignes vides ──────────────────────
            # Une ligne est considérée vide si elle n'a ni date, ni montant,
            # ni produit renseignés. Ces lignes proviennent généralement des
            # sous-en-têtes de blocs dans SUIVI_CARBURANT ou de lignes de
            # séparation/totaux que _match a capturées par erreur.
            has_data = (
                date_fmt != "—"
                or montant_f > 0
                or produit   != "—"
            )
            if not has_data:
                continue

            transactions.append({
                # ── BUG #1 : clé de tri séparée en ISO ─────────────────
                "_date_sort": _date_sort_key(date_obj),
                "mois":       _str(r.get(mo_col)) if mo_col else "—",
                "date":       date_fmt,
                "produit":    produit,
                "carte":      carte,
                "montant":    montant_f,
                "montant_fmt": _fmt(montant_f),
            })

        # Tri décroissant sur la clé ISO (les vides restent à la fin)
        transactions.sort(key=lambda x: x["_date_sort"], reverse=True)

        # On ne garde que les 30 plus récentes
        transactions = transactions[:30]

        # Nettoyage de la clé interne avant sérialisation
        for t in transactions:
            t.pop("_date_sort", None)

        # Par produit
        par_produit = {}
        if p_col and m_col:
            gp = sub.groupby(p_col)[m_col].apply(lambda x: _num(x).sum())
            par_produit = {str(k): float(v) for k, v in gp.items()} if hasattr(gp, "items") else {}

        return {
            "total_depense":   round(total, 0),
            "total_fmt":       _fmt(total),
            "nb_transactions": len(transactions),   # ← vrai compte après filtrage
            "transactions":    transactions,
            "par_produit":     par_produit,
        }

    # ── Entretien ─────────────────────────────────────────────────────────

    def _get_entretien(self, immat: str) -> dict:
        df = self.p.get("entretien")
        empty = {"total_depense": 0, "total_fmt": "—",
                 "nb_interventions": 0, "interventions": [],
                 "en_atelier": False, "duree_atelier": None}
        if df.empty: return empty

        fc = self.p.find_col
        v_col  = fc(df, "VEHICULE", "IMMAT", "VEH")
        m_col  = fc(df, "MONTANT", "COUT", "COÛT")
        dd_col = fc(df, "DEPOT", "DATE DEPOT", "DATE ENTRÉE", "ENTREE")
        dr_col = fc(df, "RETOUR", "DATE RETOUR", "DATE SORTIE", "SORTIE")
        t_col  = fc(df, "TYPE", "NATURE", "LIBELLE", "OBJET", "TRAVAUX")
        if not v_col: return empty

        mask = df[v_col].apply(lambda v: _match(v, immat))
        sub  = df[mask].copy()
        if sub.empty: return empty

        montants   = _num(sub[m_col]) if m_col else pd.Series(dtype=float)
        total      = float(montants.sum()) if not montants.empty else 0.0
        en_atelier = False
        duree      = None

        interventions = []
        for _, r in sub.iterrows():
            dd_obj = _parse_date(r.get(dd_col)) if dd_col else None
            dr_obj = _parse_date(r.get(dr_col)) if dr_col else None
            dd_fmt = _fmt_date(r.get(dd_col))   if dd_col else "—"
            dr_fmt = _fmt_date(r.get(dr_col))   if dr_col else "—"
            en_cours = (dr_obj is None)
            if en_cours:
                en_atelier = True
                if dd_obj:
                    duree = (TODAY - dd_obj).days
            montant_v  = _num(r.get(m_col)) if m_col else float("nan")
            montant_f  = float(montant_v)   if not pd.isna(montant_v) else 0.0
            type_i     = _str(r.get(t_col)) if t_col else "—"

            # Filtre lignes vides
            if dd_fmt == "—" and montant_f == 0 and type_i == "—":
                continue

            interventions.append({
                "_date_sort":  _date_sort_key(dd_obj),
                "date_depot":  dd_fmt,
                "date_retour": dr_fmt,
                "en_cours":    en_cours,
                "type":        type_i,
                "montant":     montant_f,
                "montant_fmt": _fmt(montant_f),
            })

        interventions.sort(key=lambda x: x["_date_sort"], reverse=True)
        interventions = interventions[:30]
        for i in interventions:
            i.pop("_date_sort", None)

        return {
            "total_depense":    round(total, 0),
            "total_fmt":        _fmt(total),
            "nb_interventions": len(interventions),
            "interventions":    interventions,
            "en_atelier":       en_atelier,
            "duree_atelier":    duree,
        }

    # ── VT ────────────────────────────────────────────────────────────────

    def _get_vt(self, immat: str) -> dict | None:
        df = self.p.get("vt")
        if df.empty: return None

        fc = self.p.find_col
        v_col   = fc(df, "VEHICULE", "IMMAT", "VEH")
        s_col   = fc(df, "STATUT", "Statuts", "STATUS")
        d_col   = fc(df, "EXPIR", "DATE")
        aff_col = fc(df, "AFFECTATION")
        if not v_col: return None

        mask = df[v_col].apply(lambda v: _match(v, immat))
        row  = df[mask]
        if row.empty: return None
        r = row.iloc[0]

        statut        = _str(r.get(s_col)).title() if s_col else "—"
        expiration    = _fmt_date(r.get(d_col)) if d_col else "—"
        aff           = _str(r.get(aff_col))    if aff_col else "—"
        jours_restants = None
        date_exp      = _parse_date(r.get(d_col)) if d_col else None
        if date_exp:
            jours_restants = (date_exp - TODAY).days

        return {
            "statut":         statut,
            "expiration":     expiration,
            "affectation":    aff,
            "jours_restants": jours_restants,
            "alerte": (
                "EXPIRÉE"    if jours_restants is not None and jours_restants < 0
                else "URGENTE"   if jours_restants is not None and jours_restants <= 30
                else "VIGILANCE" if jours_restants is not None and jours_restants <= 60
                else "OK"
            ) if statut.lower() == "oui" else (
                "NON_CONFORME" if statut.lower() == "non"
                else "EN_ATTENTE"
            ),
        }

    # ── Sorties / Missions ────────────────────────────────────────────────
    # CORRECTION BUG #1 + #2

    def _get_sorties(self, immat: str) -> dict:
        df = self.p.get("sorties")
        empty = {"km_total": 0, "km_dernier": None, "nb_missions": 0, "missions": []}
        if df.empty: return empty

        fc = self.p.find_col
        v_col    = fc(df, "VEHICULE", "IMMAT", "VEH")
        # DATE : on essaie plusieurs variantes
        d_col    = fc(df, "DATE DE SORTIE", "DATE SORTIE", "DATE DEPART",
                       "DATE DE DÉPART", "DATE", "JOUR")
        kd_col   = fc(df, "KM DEPART", "KM_DEPART", "KM DE DÉPART",
                       "KM DÉPART", "DÉPART KM", "DEPART")
        kr_col   = fc(df, "KM RETOUR", "KM_RETOUR", "KM DE RETOUR",
                       "RETOUR KM", "RETOUR")
        km_col   = fc(df, "KM PARCOURU", "KMS PARCOURUS", "DISTANCE",
                       "KM EFFECTUE", "KM EFFECTUÉ")
        # DESTINATION : plusieurs variantes
        dest_col = fc(df, "DESTINATION", "LIEU DESTINATION", "LIEU",
                       "MOTIF", "OBJET MISSION", "DEST")
        # CONDUCTEUR : plusieurs variantes
        cond_col = fc(df, "CONDUCTEUR", "CHAUFFEUR", "NOM CONDUCTEUR",
                       "NOMS ET PRENOMS", "NOMS PRENOMS", "NOM PRENOM",
                       "NOMS", "PRENOMS", "AGENT")
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
            kms = pd.Series([np.nan] * len(sub), index=sub.index)

        km_total = float(kms.sum())  # sum() ignore les NaN

        # KM retour max = dernier kilométrage connu
        km_dernier = None
        if kr_col:
            vals = _num(sub[kr_col]).dropna()
            if not vals.empty:
                km_dernier = float(vals.max())

        missions = []
        for i, (_, r) in enumerate(sub.iterrows()):
            raw_km = kms.iloc[i] if i < len(kms) else np.nan
            km_p   = float(raw_km) if not pd.isna(raw_km) else 0.0

            date_obj = _parse_date(r.get(d_col)) if d_col else None
            date_fmt = _fmt_date(r.get(d_col))   if d_col else "—"
            dest     = _str(r.get(dest_col))      if dest_col else "—"
            cond     = _str(r.get(cond_col))      if cond_col else "—"

            # ── BUG #2 : exclure les lignes sans aucune info utile ──────
            has_data = (
                date_fmt != "—"
                or km_p   >  0
                or dest   != "—"
                or cond   != "—"
            )
            if not has_data:
                continue

            missions.append({
                # ── BUG #1 : clé de tri ISO ─────────────────────────────
                "_date_sort":  _date_sort_key(date_obj),
                "mois":        _str(r.get(mo_col)) if mo_col else "—",
                "date":        date_fmt,
                "km_depart":   int(_num(r.get(kd_col)))
                               if kd_col and not pd.isna(_num(r.get(kd_col))) else None,
                "km_retour":   int(_num(r.get(kr_col)))
                               if kr_col and not pd.isna(_num(r.get(kr_col))) else None,
                "km_parcouru": int(km_p) if km_p > 0 else 0,
                "destination": dest,
                "conducteur":  cond,
            })

        # Tri par date décroissante (ISO : vides restent en bas)
        missions.sort(key=lambda x: x["_date_sort"], reverse=True)
        missions = missions[:30]
        for m in missions:
            m.pop("_date_sort", None)

        return {
            "km_total":    int(km_total),
            "km_dernier":  km_dernier,
            "nb_missions": len(missions),   # ← compte réel après filtrage
            "missions":    missions,
        }

    # ── Score santé simplifié ─────────────────────────────────────────────

    def _compute_score(self, carb: dict, entr: dict, vt: dict | None) -> dict:
        score = 100

        # VT
        if vt is None:
            score -= 20
        else:
            a = vt.get("alerte", "EN_ATTENTE")
            if a == "EXPIRÉE":       score -= 30
            elif a == "NON_CONFORME": score -= 25
            elif a == "URGENTE":      score -= 15
            elif a == "VIGILANCE":    score -= 10
            elif a == "EN_ATTENTE":   score -= 5

        # Atelier
        if entr.get("en_atelier"):
            d = entr.get("duree_atelier") or 0
            score -= min(20, 5 + d // 3)

        # Pannes curatives
        nb = entr.get("nb_interventions", 0)
        score -= min(10, nb // 3)

        score = max(0, min(100, score))
        statut = (
            "OPTIMAL"   if score >= 85
            else "BON"       if score >= 70
            else "SURVEILLER" if score >= 50
            else "CRITIQUE"  if score >= 30
            else "REMPLACER"
        )
        return {
            "score":        score,
            "statut":       statut,
            "detail_vt":    vt.get("alerte") if vt else "INCONNU",
            "en_atelier":   entr.get("en_atelier", False),
            "nb_pannes":    entr.get("nb_interventions", 0),
        }

    # ── Suggestions ───────────────────────────────────────────────────────

    def _suggest(self, immat: str) -> list[str]:
        df = self.p.get("vehicules")
        if df.empty: return []
        col = self.p.find_col(df, "IMMAT")
        if not col: return []
        q = immat.replace(" ", "").upper()
        suggestions = []
        for v in df[col].dropna():
            v_clean = str(v).replace(" ", "").upper()
            if q[:4] in v_clean or v_clean[:4] in q:
                suggestions.append(str(v).strip())
        return suggestions[:5]


# ─── Point d'entrée ───────────────────────────────────────────────────────────

def search_vehicle(raw_sheets: dict, immat: str) -> dict:
    """Parse le classeur et cherche l'immatriculation."""
    parser = FleetParser(raw_sheets)
    engine = VehicleSearchEngine(parser)
    result = engine.search(immat)

    def safe(o):
        if isinstance(o, bool):   return o
        if isinstance(o, int):    return int(o)
        if isinstance(o, float):
            if o != o or o == float("inf") or o == float("-inf"): return None
            return round(o, 2)
        if isinstance(o, dict):   return {k: safe(v) for k, v in o.items()}
        if isinstance(o, list):   return [safe(i) for i in o]
        return o

    return safe(result)
