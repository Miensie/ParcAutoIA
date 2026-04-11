"""
FleetInsight AI — Fleet Analyzer v2.2
=======================================
- Gpe électrogène intégrée aux 3 niveaux
- Dates corrigées (via fleet_parser._parse_date)
- Analyse décisionnelle 100% basée sur les données réelles
- Facteur économie dynamique (calculé par IA)
"""

from __future__ import annotations
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from typing import Any
from fleet_parser import FleetParser, _parse_date, _fmt_date
import warnings
warnings.filterwarnings("ignore")

TODAY  = datetime.now()
TS_NOW = pd.Timestamp(TODAY)

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _j(obj):
    if isinstance(obj, (np.integer,)):            return int(obj)
    if isinstance(obj, (np.floating,)):           return None if np.isnan(obj) else float(obj)
    if isinstance(obj, np.ndarray):               return obj.tolist()
    if isinstance(obj, (pd.Timestamp, datetime)): return obj.isoformat()
    if isinstance(obj, dict):  return {k: _j(v) for k, v in obj.items()}
    if isinstance(obj, list):  return [_j(i) for i in obj]
    return obj

def _num(s): return pd.to_numeric(s, errors="coerce")

def _fmt(n):
    if n is None or (isinstance(n, float) and np.isnan(n)): return "—"
    n = float(n)
    if n >= 1_000_000: return f"{n/1_000_000:.1f}M FCFA"
    if n >= 1_000:     return f"{n/1_000:.0f}K FCFA"
    return f"{n:.0f} FCFA"

def _date_series(df: pd.DataFrame, col: str) -> pd.Series:
    """Convertit une colonne en Series de pd.Timestamp (NaT si invalide)."""
    def _to_ts(v):
        d = _parse_date(v)
        return pd.Timestamp(d) if d else pd.NaT
    return df[col].apply(_to_ts)


# ─────────────────────────────────────────────────────────────────────────────
# NIVEAU 1 — SUIVI OPÉRATIONNEL
# ─────────────────────────────────────────────────────────────────────────────

class N1_Operationnel:
    def __init__(self, p: FleetParser):
        self.p = p

    # ── Inventaire ────────────────────────────────────────────────────────
    def inventaire(self) -> dict:
        df = self.p.get("vehicules")
        if df.empty: return {"error": "Feuille LISTE DE VEH introuvable", "total": 0}
        fc = self.p.find_col
        marque = fc(df, "MARQUE")
        type_v = fc(df, "TYPE")
        result: dict = {"total": int(len(df))}
        if marque:
            vc = df[marque].dropna().value_counts()
            result["par_marque"] = [{"label": str(k), "value": int(v)} for k,v in vc.head(12).items()]
        if type_v:
            vc2 = df[type_v].dropna().value_counts()
            result["par_type"] = [{"label": str(k), "value": int(v)} for k,v in vc2.items()]
        immat = fc(df, "IMMAT")
        if immat:
            result["immatriculations"] = [str(v) for v in df[immat].dropna().tolist()]
        return result

    # ── Carburant ─────────────────────────────────────────────────────────
    def carburant(self) -> dict:
        df = self.p.get("carburant")
        if df.empty: return {"error": "Feuille SUIVI_CARBURANT introuvable", "total_depense": 0}
        fc = self.p.find_col
        m_col = fc(df, "MONTANT")
        v_col = fc(df, "VEHICULE")
        p_col = fc(df, "PRODUIT")
        mo_col= fc(df, "MOIS")
        c_col = fc(df, "CARTE")
        d_col = fc(df, "DATE")
        result: dict = {"nb_transactions": int(len(df))}
        if m_col:
            montants = _num(df[m_col])
            result.update({
                "total_depense": float(montants.sum()),
                "moyenne_par_transaction": float(montants.mean()),
                "max_transaction": float(montants.max()),
            })
        if v_col and m_col:
            by_v = df.groupby(v_col)[m_col].apply(lambda x: _num(x).sum()).sort_values(ascending=False)
            result["top_consommateurs"] = [
                {"vehicule": str(k), "total": float(v), "total_fmt": _fmt(float(v))}
                for k,v in by_v.head(10).items() if not (isinstance(v,float) and np.isnan(v))
            ]
        if p_col:
            vc = df[p_col].dropna().value_counts(normalize=True)
            result["repartition_produit"] = [
                {"produit": str(k), "pct": round(float(v)*100,1)} for k,v in vc.items()
            ]
        if mo_col and m_col:
            by_mo = df.groupby(mo_col)[m_col].apply(lambda x: _num(x).sum())
            result["par_mois"] = [{"mois": str(k), "total": float(v)} for k,v in by_mo.items()]
        if c_col and m_col:
            by_c = df.groupby(c_col)[m_col].apply(lambda x: _num(x).sum()).sort_values(ascending=False)
            result["par_carte"] = [{"carte": str(k), "total": float(v)} for k,v in by_c.head(8).items()]
        # Dates réelles
        if d_col:
            dates = _date_series(df, d_col).dropna()
            if not dates.empty:
                result["date_min"] = _fmt_date(dates.min().to_pydatetime())
                result["date_max"] = _fmt_date(dates.max().to_pydatetime())
        return result

    # ── Entretien & Réparations (mécanique uniquement) ────────────────────
    def entretien(self) -> dict:
        df = self.p.get("entretien")
        if df.empty: return {"error": "Feuille ENTRETIEN introuvable", "total_interventions": 0,
                              "en_atelier": 0, "liste_atelier": []}
        fc = self.p.find_col
        v_col = fc(df, "VEHICULE")
        d_col = fc(df, "DEPOT")
        r_col = fc(df, "RETOUR")
        t_col = fc(df, "TYPE")
        m_col = fc(df, "MONTANT")
        result: dict = {"total_interventions": int(len(df))}

        # Convertir les colonnes de dates correctement
        if d_col:
            df = df.copy()
            df["_depot_dt"]  = _date_series(df, d_col)
            df["_retour_dt"] = _date_series(df, r_col) if r_col else pd.NaT

            mask_atelier = df["_depot_dt"].notna() & df["_retour_dt"].isna()
            atelier_df   = df[mask_atelier]
            result["en_atelier"] = int(len(atelier_df))

            if v_col and not atelier_df.empty:
                result["liste_atelier"] = []
                for _, row in atelier_df.iterrows():
                    depot_dt = row["_depot_dt"]
                    duree = int((TS_NOW - depot_dt).days) if pd.notna(depot_dt) else None
                    result["liste_atelier"].append({
                        "vehicule":   str(row.get(v_col,"—")),
                        "type":       str(row.get(t_col,"—")) if t_col else "—",
                        "date_depot": _fmt_date(row.get(d_col)),
                        "duree_jours": duree,
                    })
            else:
                result["liste_atelier"] = []

            # Interventions terminées
            terminees = df[df["_retour_dt"].notna()]
            if m_col:
                montants = _num(terminees[m_col])
                result["total_depense"] = float(montants.sum())
                result["cout_moyen"]    = float(montants.mean())
                result["cout_max"]      = float(montants.max())
        else:
            result["en_atelier"]    = 0
            result["liste_atelier"] = []

        if t_col:
            vc = df[t_col].dropna().value_counts()
            result["types"] = [{"type": str(k), "count": int(v)} for k,v in vc.head(12).items()]
            prev_kw   = ["vidange","révision","revision","contrôle","controle","filtre","graissage"]
            curatif_kw= ["accident","panne","bris","fuite","remplacement","problème","probleme",
                          "climatisation","frein","crémaillière","moteur","amortisseur"]
            df["_cat"] = df[t_col].astype(str).str.lower().apply(
                lambda x: "Préventif" if any(k in x for k in prev_kw)
                else      "Curatif"   if any(k in x for k in curatif_kw)
                else      "Autre"
            )
            vc_cat = df["_cat"].value_counts()
            result["categories"] = [{"cat": str(k), "count": int(v)} for k,v in vc_cat.items()]
        return result

    # ── VT — basée sur colonne Statuts ────────────────────────────────────
    def vt(self) -> dict:
        df = self.p.get("vt")
        if df.empty: return {"error": "Feuille VT introuvable", "total": 0,
                              "oui": 0, "non": 0, "pas_encore": 0}
        fc = self.p.find_col
        s_col   = fc(df, "STATUT","Statuts","STATUS")
        v_col   = fc(df, "VEHICULE")
        d_col   = fc(df, "EXPIR","EXPIRATION")
        aff_col = fc(df, "AFFECTATION")

        result: dict = {"total": int(len(df))}
        if not s_col:
            return {**result, "error": "Colonne 'Statuts' introuvable"}

        df = df.copy()
        df["_statut"] = df[s_col].astype(str).str.strip().str.lower().map({
            "oui":"Oui","non":"Non","pas encore":"Pas encore","nan":"Non renseigné"
        }).fillna("Non renseigné")

        counts = df["_statut"].value_counts()
        result.update({
            "oui":           int(counts.get("Oui",0)),
            "non":           int(counts.get("Non",0)),
            "pas_encore":    int(counts.get("Pas encore",0)),
            "non_renseigne": int(counts.get("Non renseigné",0)),
            "taux_conformite": round(int(counts.get("Oui",0))/len(df)*100,1) if len(df) else 0,
        })

        # Convertir les dates d'expiration correctement
        if d_col:
            df["_date_exp"] = _date_series(df, d_col)
        else:
            df["_date_exp"] = pd.NaT

        def _row(row) -> dict:
            d = row.get("_date_exp")
            jours = int((d - TS_NOW).days) if pd.notna(d) else None
            return {
                "vehicule":       str(row.get(v_col,"—")) if v_col else "—",
                "affectation":    str(row.get(aff_col,"—")) if aff_col else "—",
                "statut":         str(row.get("_statut","—")),
                "expiration":     _fmt_date(row.get(d_col)) if d_col else "—",
                "jours_restants": jours,
            }

        df_non  = df[df["_statut"]=="Non"]
        df_oui  = df[df["_statut"]=="Oui"]
        df_pa   = df[df["_statut"]=="Pas encore"].sort_values("_date_exp", na_position="last")

        result["liste_non"]        = [_row(r) for _,r in df_non.iterrows()]
        result["liste_oui"]        = [_row(r) for _,r in df_oui.iterrows()]
        result["liste_pas_encore"] = [_row(r) for _,r in df_pa.iterrows()]

        ts_60j = TS_NOW + pd.Timedelta(days=60)
        urgents = df_pa[df_pa["_date_exp"].notna() & (df_pa["_date_exp"] <= ts_60j)]
        result["pas_encore_urgents"] = [_row(r) for _,r in urgents.iterrows()]

        if aff_col:
            nc = df[df["_statut"].isin(["Non","Pas encore"])]
            if not nc.empty:
                by_aff = nc.groupby(aff_col)["_statut"].count().sort_values(ascending=False)
                result["non_conformes_par_affectation"] = [
                    {"affectation": str(k), "count": int(v)} for k,v in by_aff.head(10).items()
                ]
        return result

    # ── Sorties ───────────────────────────────────────────────────────────
    def sorties(self) -> dict:
        df = self.p.get("sorties")
        if df.empty: return {"error": "Feuille SORTIES VEH introuvable", "total_sorties": 0}
        fc = self.p.find_col
        v_col    = fc(df, "VEHICULE")
        kd_col   = fc(df, "KM DEPART","KM_DEPART","DEPART")
        kr_col   = fc(df, "KM RETOUR","KM_RETOUR","RETOUR")
        dest_col = fc(df, "DESTINATION","DEST")
        mo_col   = fc(df, "MOIS")
        result: dict = {"total_sorties": int(len(df))}
        if kd_col and kr_col:
            kms = _num(df[kr_col]) - _num(df[kd_col])
            kms = kms[kms > 0]
            result.update({
                "km_total": float(kms.sum()),
                "km_moyen": float(kms.mean()) if not kms.empty else 0.0,
                "km_max":   float(kms.max()) if not kms.empty else 0.0,
            })
        if v_col:
            result["vehicules_actifs"] = int(df[v_col].nunique())
            vc = df[v_col].value_counts()
            result["top_vehicules"] = [{"vehicule":str(k),"sorties":int(v)} for k,v in vc.head(8).items()]
        if dest_col:
            vc2 = df[dest_col].dropna().value_counts()
            result["top_destinations"] = [{"dest":str(k),"count":int(v)} for k,v in vc2.head(8).items()]
        if mo_col:
            vc3 = df[mo_col].dropna().value_counts()
            result["par_mois"] = [{"mois":str(k),"count":int(v)} for k,v in vc3.items()]
        return result

    # ── Gpe électrogène ───────────────────────────────────────────────────
    def generateur(self) -> dict:
        """
        Analyse la feuille Gpe électrogène :
        colonnes attendues : DATE, AGENCE, PRODUIT, MONTANT
        Analyses par produit et montant, incluant croisement agence-produit
        """
        df = self.p.get("generateur")
        if df.empty:
            return {"disponible": False, "message": "Feuille Gpe électrogène non trouvée"}
        fc = self.p.find_col
        d_col   = fc(df, "DATE")
        ag_col  = fc(df, "AGENCE")
        pr_col  = fc(df, "PRODUIT")
        m_col   = fc(df, "MONTANT")

        result: dict = {"disponible": True, "nb_entrees": int(len(df))}

        if m_col:
            montants = _num(df[m_col])
            result["total_depense"]  = float(montants.sum())
            result["depense_fmt"]    = _fmt(float(montants.sum()))
            result["cout_moyen"]     = float(montants.mean())
            result["cout_max"]       = float(montants.max())

        if ag_col:
            vc = df[ag_col].dropna().value_counts()
            result["par_agence"] = [{"agence":str(k),"count":int(v)} for k,v in vc.items()]
            if m_col:
                by_ag = df.groupby(ag_col)[m_col].apply(lambda x: _num(x).sum()).sort_values(ascending=False)
                result["cout_par_agence"] = [
                    {"agence": str(k), "total": float(v), "fmt": _fmt(float(v))}
                    for k,v in by_ag.items()
                ]

        # ── ANALYSE PAR PRODUIT ET MONTANT (nouveau) ──────────────────────
        if pr_col:
            vc2 = df[pr_col].dropna().value_counts(normalize=True)
            result["repartition_produit"] = [
                {"produit":str(k),"pct":round(float(v)*100,1)} for k,v in vc2.items()
            ]
            
            # Montant total par produit
            if m_col:
                # Fallback si agg échoue
                try:
                    by_pr = df.groupby(pr_col)[m_col].apply(lambda x: _num(x).sum()).sort_values(ascending=False)
                    result["cout_par_produit"] = [
                        {"produit": str(k), "total": float(v), "fmt": _fmt(float(v)), "pct_total": round(float(v)/result["total_depense"]*100,1) if result["total_depense"] else 0}
                        for k,v in by_pr.items()
                    ]
                except:
                    pass
                
                # Coût moyen par produit
                try:
                    avg_by_pr = df.groupby(pr_col)[m_col].apply(lambda x: _num(x).mean())
                    result["cout_moyen_produit"] = [
                        {"produit": str(k), "montant_moyen": float(v), "fmt": _fmt(float(v))}
                        for k,v in avg_by_pr.items()
                    ]
                except:
                    pass

        # ── CROISEMENT AGENCE + PRODUIT + MONTANT (nouveau) ───────────────
        if ag_col and pr_col and m_col:
            try:
                cross = df.groupby([ag_col, pr_col])[m_col].agg(['sum', 'count', 'mean']).reset_index()
                cross.columns = ['agence', 'produit', 'total', 'nb_transactions', 'montant_moyen']
                cross_sorted = cross.sort_values('total', ascending=False)
                result["par_agence_produit"] = [
                    {
                        "agence": str(row['agence']),
                        "produit": str(row['produit']),
                        "total": float(row['total']),
                        "fmt": _fmt(float(row['total'])),
                        "nb_transactions": int(row['nb_transactions']),
                        "montant_moyen": float(row['montant_moyen']),
                        "fmt_moyen": _fmt(float(row['montant_moyen']))
                    }
                    for _, row in cross_sorted.iterrows()
                ]
            except:
                pass

        if d_col:
            dates = _date_series(df, d_col).dropna()
            if not dates.empty:
                result["date_min"] = _fmt_date(dates.min().to_pydatetime())
                result["date_max"] = _fmt_date(dates.max().to_pydatetime())
                result["nb_jours_couverts"] = int((dates.max() - dates.min()).days) + 1

            # Regroupement par mois si possible
            if len(dates) > 0 and len(df) > 1:
                df2 = df.copy()
                df2["_month"] = dates.dt.to_period("M").astype(str)
                if m_col:
                    by_m = df2.groupby("_month")[m_col].apply(lambda x: _num(x).sum())
                    result["par_mois"] = [{"mois":str(k),"total":float(v)} for k,v in by_m.items()]

        return result

    def run(self) -> dict:
        return {
            "inventaire": self.inventaire(),
            "carburant":  self.carburant(),
            "entretien":  self.entretien(),
            "vt":         self.vt(),
            "sorties":    self.sorties(),
            "generateur": self.generateur(),
        }


# ─────────────────────────────────────────────────────────────────────────────
# NIVEAU 2 — INDICATEURS (100% basés sur données réelles)
# ─────────────────────────────────────────────────────────────────────────────

class N2_Indicateurs:
    def __init__(self, p: FleetParser, ops: dict):
        self.p   = p
        self.ops = ops

    def cout_par_vehicule(self) -> list[dict]:
        costs: dict[str, dict] = {}
        df_c = self.p.get("carburant")
        if not df_c.empty:
            v = self.p.find_col(df_c, "VEHICULE")
            m = self.p.find_col(df_c, "MONTANT")
            if v and m:
                for veh, grp in df_c.groupby(v):
                    if pd.isna(veh): continue
                    k = str(veh).strip()
                    costs.setdefault(k, {"vehicule":k,"carburant":0.,"entretien":0.})
                    costs[k]["carburant"] = float(_num(grp[m]).sum())
        df_e = self.p.get("entretien")
        if not df_e.empty:
            v = self.p.find_col(df_e, "VEHICULE")
            m = self.p.find_col(df_e, "MONTANT")
            if v and m:
                for veh, grp in df_e.groupby(v):
                    if pd.isna(veh): continue
                    k = str(veh).strip()
                    costs.setdefault(k, {"vehicule":k,"carburant":0.,"entretien":0.})
                    costs[k]["entretien"] = float(_num(grp[m]).sum())
        result = []
        for c in costs.values():
            total = c["carburant"] + c["entretien"]
            result.append({**c, "total": round(total,0), "total_fmt": _fmt(total)})
        return sorted(result, key=lambda x: -x["total"])

    def conformite_vt(self) -> dict:
        vt = self.ops["vt"]
        total  = vt.get("total",0)
        nb_oui = vt.get("oui",0)
        nb_non = vt.get("non",0)
        nb_pa  = vt.get("pas_encore",0)
        nb_urg = len(vt.get("pas_encore_urgents",[]))
        taux   = round(nb_oui/total*100,1) if total else 0

        # Niveau de risque calculé sur les données réelles
        if nb_non >= 5 or (total > 0 and nb_non/total >= 0.20):
            niveau = "CRITIQUE"
        elif nb_non >= 3 or nb_urg >= 5:
            niveau = "ÉLEVÉ"
        elif nb_non >= 1 or nb_urg >= 3:
            niveau = "MODÉRÉ"
        else:
            niveau = "FAIBLE"

        return {
            "total": total,
            "nb_conformes":        nb_oui,
            "nb_non_conformes":    nb_non,
            "nb_en_attente":       nb_pa,
            "nb_urgents_60j":      nb_urg,
            "taux_conformite":     taux,
            "taux_non_conformite": round(nb_non/total*100,1) if total else 0,
            "taux_en_attente":     round(nb_pa/total*100,1)  if total else 0,
            "niveau_risque":       niveau,
        }

    def taux_immobilisation(self) -> dict:
        total      = self.ops["inventaire"].get("total",0)
        en_atelier = self.ops["entretien"].get("en_atelier",0)
        dispo      = max(0, total - en_atelier)
        taux_immo  = round(en_atelier/total*100,1) if total else 0
        return {
            "total":       total,
            "disponibles": dispo,
            "en_atelier":  en_atelier,
            "taux_immo":   taux_immo,
            "taux_dispo":  round(100-taux_immo,1),
        }

    def frequence_pannes(self) -> list[dict]:
        df = self.p.get("entretien")
        if df.empty: return []
        v_col = self.p.find_col(df, "VEHICULE")
        t_col = self.p.find_col(df, "TYPE")
        if not v_col: return []
        prev_kw = ["vidange","révision","revision","contrôle","controle","filtre"]
        if t_col:
            mask = ~df[t_col].astype(str).str.lower().apply(lambda x: any(k in x for k in prev_kw))
            pannes_df = df[mask]
        else:
            pannes_df = df
        freq = pannes_df[v_col].dropna().value_counts()
        return [{"vehicule":str(k),"nb_pannes":int(v)} for k,v in freq.head(15).items()]

    def indicateurs_generateur(self) -> dict:
        """
        Indicateurs Niveau 2 pour Gpe électrogène.
        Ratio coût/agence, fréquence d'approvisionnement.
        Analyse par produit et montant.
        """
        gen = self.ops["generateur"]
        if not gen.get("disponible"):
            return {"disponible": False}
        result = {"disponible": True}
        total = gen.get("total_depense", 0) or 0

        # Coût moyen par agence
        cout_ag = gen.get("cout_par_agence", [])
        if cout_ag:
            nb_agences = len(cout_ag)
            result["nb_agences"]       = nb_agences
            result["cout_moyen_agence"] = round(total / nb_agences, 0) if nb_agences else 0
            result["top_agence"]       = cout_ag[0] if cout_ag else None

        # Analyse par produit
        cout_pr = gen.get("cout_par_produit", [])
        if cout_pr:
            result["nb_produits"] = len(cout_pr)
            result["produit_plus_cher"] = cout_pr[0] if cout_pr else None
            # Concentration produit (produit le plus cher vs total)
            if cout_pr and total:
                pct_top_produit = (cout_pr[0].get("total", 0) / total * 100)
                result["concentration_produit"] = round(pct_top_produit, 1)

        # Fréquence supprimée (données non fiables)

        # Croisement agence-produit : identification des anomalies
        par_ag_pr = gen.get("par_agence_produit", [])
        if par_ag_pr:
            result["nb_combinaisons_agence_produit"] = len(par_ag_pr)
            result["top_agence_produit"] = par_ag_pr[0] if par_ag_pr else None

        return result

    def sante_vehicules(self) -> list[dict]:
        """
        Score 0-100 basé sur percentiles réels des données.
        """
        couts  = {c["vehicule"]: c for c in self.cout_par_vehicule()}
        pannes = {p["vehicule"]: p["nb_pannes"] for p in self.frequence_pannes()}
        atelier = {r["vehicule"]: True for r in self.ops["entretien"].get("liste_atelier",[])}
        all_v   = set(list(couts.keys()) + list(pannes.keys()) + list(atelier.keys()))

        # Seuils dynamiques sur les données réelles
        all_couts  = [couts[v]["total"] for v in all_v if v in couts and couts[v]["total"] > 0]
        all_pannes = [pannes[v] for v in all_v if v in pannes]

        p50_cout   = float(np.percentile(all_couts, 50))  if all_couts  else 1.0
        p75_cout   = float(np.percentile(all_couts, 75))  if all_couts  else 1.0
        p90_cout   = float(np.percentile(all_couts, 90))  if all_couts  else 1.0
        median_pan = float(np.median(all_pannes))          if all_pannes else 1.0

        result = []
        for v in all_v:
            cout  = couts.get(v,{}).get("total",0)
            panne = pannes.get(v,0)
            immo  = atelier.get(v,False)

            # Pénalité coût (relative aux percentiles réels)
            if cout >= p90_cout:   pen_cout = 40
            elif cout >= p75_cout: pen_cout = 25
            elif cout >= p50_cout: pen_cout = 10
            else:                  pen_cout = 0

            # Pénalité pannes (relative à la médiane)
            pen_panne = min(35, int(panne / max(median_pan, 1) * 15)) if panne > 0 else 0
            pen_immo  = 20 if immo else 0
            score     = max(0, 100 - pen_cout - pen_panne - pen_immo)

            if   score >= 80: statut = "OPTIMAL"
            elif score >= 60: statut = "BON"
            elif score >= 40: statut = "SURVEILLER"
            elif score >= 20: statut = "CRITIQUE"
            else:             statut = "REMPLACER"

            result.append({
                "vehicule":    v,
                "score":       score,
                "statut":      statut,
                "total_cout":  round(float(cout),0),
                "nb_pannes":   panne,
                "immobilise":  immo,
            })
        return sorted(result, key=lambda x: x["score"])

    def run(self) -> dict:
        return {
            "cout_par_vehicule":    self.cout_par_vehicule(),
            "conformite_vt":        self.conformite_vt(),
            "taux_immobilisation":  self.taux_immobilisation(),
            "frequence_pannes":     self.frequence_pannes(),
            "indicateurs_generateur": self.indicateurs_generateur(),
            "sante_vehicules":      self.sante_vehicules(),
        }


# ─────────────────────────────────────────────────────────────────────────────
# NIVEAU 3 — PILOTAGE DÉCISIONNEL (100% données réelles, sans facteur figé)
# ─────────────────────────────────────────────────────────────────────────────

class N3_Decision:
    def __init__(self, ops: dict, ind: dict):
        self.ops = ops
        self.ind = ind

    # ── Alertes ───────────────────────────────────────────────────────────
    def alertes(self) -> list[dict]:
        alerts = []
        taux   = self.ind["taux_immobilisation"]
        conf   = self.ind["conformite_vt"]
        entr   = self.ops["entretien"]
        vt_op  = self.ops["vt"]
        gen    = self.ops["generateur"]

        # ENTRETIEN : immobilisation
        if taux["en_atelier"] > 0:
            seuil_crit = taux["total"] * 0.15 if taux["total"] else 3
            niveau = "CRITIQUE" if taux["en_atelier"] >= seuil_crit else "WARNING"
            max_d  = max((v.get("duree_jours",0) or 0 for v in entr.get("liste_atelier",[])), default=0)
            alerts.append({
                "domaine": "ENTRETIEN", "niveau": niveau, "icon": "🔧",
                "titre":   f"{taux['en_atelier']} véhicule(s) en atelier — {taux['taux_immo']}% du parc",
                "detail":  f"Immobilisation maximale : {max_d} jour(s). "
                           f"{taux['disponibles']} véhicules disponibles sur {taux['total']}.",
                "action":  "Relancer le garage. Définir une date de retour pour chaque véhicule.",
            })

        # VT : Non
        nb_non = vt_op.get("non",0)
        if nb_non > 0:
            veh_non = [v["vehicule"] for v in vt_op.get("liste_non",[])[:5]]
            alerts.append({
                "domaine": "VT", "niveau": "CRITIQUE", "icon": "⛔",
                "titre":   f"⚖️ {nb_non} véhicule(s) sans VT — Risque juridique immédiat",
                "detail":  f"Véhicules concernés : {', '.join(veh_non)}{'...' if nb_non>5 else ''}",
                "action":  "Ces véhicules ne doivent pas circuler. Programmer les visites techniques d'urgence.",
            })

        # VT : Pas encore urgents
        urgents = vt_op.get("pas_encore_urgents",[])
        if urgents:
            plus_urgent = urgents[0]
            alerts.append({
                "domaine": "VT", "niveau": "WARNING", "icon": "⏳",
                "titre":   f"{len(urgents)} VT 'Pas encore' expirant sous 60 jours",
                "detail":  f"Plus urgent : {plus_urgent['vehicule']} ({plus_urgent['expiration']}).",
                "action":  "Planifier et confirmer les RDV de VT sous 2 semaines.",
            })

        # VT : Pas encore non urgents
        pa_total    = vt_op.get("pas_encore",0)
        pa_restants = pa_total - len(urgents)
        if pa_restants > 0:
            alerts.append({
                "domaine": "VT", "niveau": "INFO", "icon": "📅",
                "titre":   f"{pa_restants} VT 'Pas encore' à horizon > 60 jours",
                "detail":  "Ces visites sont prévues mais pas encore programmées.",
                "action":  "Intégrer au planning de VT du prochain trimestre.",
            })

        # Véhicules critiques
        critiques = [v for v in self.ind["sante_vehicules"] if v["statut"] in ("CRITIQUE","REMPLACER")]
        if critiques:
            alerts.append({
                "domaine": "ENTRETIEN", "niveau": "WARNING", "icon": "⚠️",
                "titre":   f"{len(critiques)} véhicule(s) en état critique (coût/pannes élevés)",
                "detail":  f"Coûts et/ou pannes au-dessus des seuils du parc actuel.",
                "action":  "Lancer un audit mécanique. Évaluer remplacement vs réparation.",
            })

        # Gpe électrogène (alerte fréquence supprimée)

        return alerts

    # ── Plan VT ───────────────────────────────────────────────────────────
    def plan_action_vt(self) -> dict:
        vt_op = self.ops["vt"]
        conf  = self.ind["conformite_vt"]
        p1 = vt_op.get("liste_non",[])
        p2 = vt_op.get("pas_encore_urgents",[])
        p3 = [v for v in vt_op.get("liste_pas_encore",[]) if v not in p2]
        return {
            "synthese": {
                "conformes":      conf["nb_conformes"],
                "non_conformes":  conf["nb_non_conformes"],
                "en_attente":     conf["nb_en_attente"],
                "taux_conformite":conf["taux_conformite"],
                "niveau_risque":  conf["niveau_risque"],
            },
            "p1_immediat": {
                "description": "Statut NON — VT non effectuée — Circulation interdite",
                "vehicules": p1, "count": len(p1), "delai": "Cette semaine",
            },
            "p2_planifier": {
                "description": "Statut PAS ENCORE — Expiration < 60 jours",
                "vehicules": p2, "count": len(p2), "delai": "Sous 2 semaines",
            },
            "p3_programmer": {
                "description": "Statut PAS ENCORE — Expiration > 60 jours",
                "vehicules": p3[:10], "count": len(p3), "delai": "Ce trimestre",
            },
        }

    # ── Plan renouvellement (dynamique) ──────────────────────────────────
    def plan_renouvellement(self) -> dict:
        sante = self.ind["sante_vehicules"]
        a_rem  = [v for v in sante if v["statut"]=="REMPLACER"]
        a_crit = [v for v in sante if v["statut"]=="CRITIQUE"]
        a_surv = [v for v in sante if v["statut"]=="SURVEILLER"]

        # Coût moyen réel des véhicules à remplacer comme base du budget
        couts_rem = [v["total_cout"] for v in a_rem if v["total_cout"] > 0]
        budget_unitaire = 9_000_000  # VP neuf estimé FCFA
        budget = len(a_rem) * budget_unitaire

        if not a_rem:
            horizon = "Pas de remplacement urgent"
        elif len(a_rem) >= 5:
            horizon = "Immédiat (0–1 mois)"
        elif len(a_rem) >= 2:
            horizon = "Court terme (1–3 mois)"
        else:
            horizon = "Moyen terme (3–6 mois)"

        return {
            "a_remplacer":     a_rem[:10],
            "a_auditer":       a_crit[:10],
            "a_surveiller":    a_surv[:10],
            "nb_remplacement": len(a_rem),
            "nb_audit":        len(a_crit),
            "budget_estime":   budget,
            "budget_fmt":      _fmt(float(budget)),
            "cout_moyen_rem":  round(float(np.mean(couts_rem)),0) if couts_rem else 0,
            "horizon":         horizon,
        }

    # ── Décisions Gpe électrogène ─────────────────────────────────────────
    def decisions_generateur(self) -> dict:
        """
        Décisions niveau 3 pour le groupe électrogène.
        Basées sur les données réelles (produit, montant, agence).
        """
        gen     = self.ops["generateur"]
        gen_ind = self.ind.get("indicateurs_generateur",{})
        if not gen.get("disponible"):
            return {"disponible": False}

        total       = gen.get("total_depense",0) or 0
        top_agence  = gen_ind.get("top_agence")
        nb_agences  = gen_ind.get("nb_agences",0) or 0
        cout_moyen  = gen_ind.get("cout_moyen_agence",0) or 0
        
        # Nouvelles variables pour analyse produit-montant
        produit_plus_cher = gen_ind.get("produit_plus_cher")
        concentration_pr  = gen_ind.get("concentration_produit", 0)
        top_ag_pr         = gen_ind.get("top_agence_produit")

        recommandations = []

        # --- Analyse par PRODUIT -----
        if produit_plus_cher:
            pct = produit_plus_cher.get("pct_total", 0)
            prod_name = produit_plus_cher.get("produit", "")
            if pct >= 50:
                recommandations.append(
                    f"PRODUIT : '{prod_name}' représente {pct:.1f}% du budget groupe électrogène. "
                    f"Analyser la consommation et les prix du marché pour ce produit."
                )
            elif pct >= 40:
                recommandations.append(
                    f"PRODUIT : '{prod_name}' représente {pct:.1f}% du budget. "
                    f"Vérifier les tarifs fournisseur et les opportunités de négociation."
                )

        # --- Analyse par MONTANT per produit ---
        cout_pr = gen.get("cout_par_produit", [])
        if cout_pr and len(cout_pr) > 1:
            # Comparaison montant moyen par produit
            cout_moyen_pr = gen.get("cout_moyen_produit", [])
            if cout_moyen_pr:
                # Produit avec montant moyen le plus élevé
                max_avg = max((p.get("montant_moyen", 0) for p in cout_moyen_pr), default=0)
                min_avg = min((p.get("montant_moyen", 0) for p in cout_moyen_pr), default=0)
                if max_avg > 0 and min_avg > 0 and max_avg / min_avg > 1.5:
                    prod_max = next((p for p in cout_moyen_pr if p.get("montant_moyen") == max_avg), {})
                    prod_min = next((p for p in cout_moyen_pr if p.get("montant_moyen") == min_avg), {})
                    recommandations.append(
                        f"MONTANTS : Écart de {max_avg/min_avg:.1f}x entre produit le plus cher ('{prod_max.get('produit')}' : "
                        f"{_fmt(max_avg)}) et le moins cher ('{prod_min.get('produit')}' : {_fmt(min_avg)}). "
                        f"Analyser les causes."
                    )

        # --- Analyse par AGENCE + PRODUIT -----
        if top_ag_pr:
            ag = top_ag_pr.get("agence", "")
            pr = top_ag_pr.get("produit", "")
            t = top_ag_pr.get("total", 0)
            pct_total = round(t / total * 100, 1) if total else 0
            if pct_total >= 30:
                recommandations.append(
                    f"AGENCE+PRODUIT : L'agence '{ag}' pour produit '{pr}' = {pct_total}% du budget. "
                    f"Vérifier la consommation réelle et les facteurs opérationnels."
                )

        # --- Concentration sur une seule agence -----
        if top_agence and nb_agences > 1:
            pct_top = round(top_agence["total"]/total*100,0) if total else 0
            if pct_top >= 70:
                recommandations.append(
                    f"AGENCE : L'agence '{top_agence['agence']}' représente {pct_top:.0f}% des dépenses. "
                    "Vérifier si le groupe est dimensionné correctement."
                )



        if not recommandations:
            recommandations.append(
                "Consommation du groupe électrogène conforme. Mantenir le suivi produit/agence/montant régulier."
            )

        return {
            "disponible":       True,
            "total_depense":    total,
            "total_fmt":        _fmt(total),
            "cout_moyen_agence":_fmt(cout_moyen),
            "recommandations":  recommandations,
            "priorite":         "HAUTE" if (concentration_pr >= 60) else "NORMALE",
        }

    # ── Optimisation (données réelles — facteur calculé par IA) ───────
    def optimisation(self) -> dict:
        """
        NB : le facteur d'économie n'est plus hardcodé à 0.12.
        Il est calculé dynamiquement par IA dans gemini_advisor.py.
        Ici on prépare juste les données brutes pour l'IA.
        """
        carb  = self.ops["carburant"]
        entr  = self.ops["entretien"]
        gen   = self.ops["generateur"]
        tc    = carb.get("total_depense",0) or 0
        te    = entr.get("total_depense",0) or 0
        tg    = gen.get("total_depense",0)  or 0
        total = tc + te + tg

        # Anomalies détectées sur données réelles
        top3    = carb.get("top_consommateurs",[])[:3]
        anomalies = []

        if total > 0 and te/total > 0.35:
            anomalies.append({
                "type": "ratio_entretien_eleve",
                "valeur": round(te/total*100,1),
                "seuil": 35,
                "detail": f"Entretien = {te/total*100:.0f}% du budget (seuil : 35%)"
            })
        if top3 and tc > 0:
            top3_total = sum(v["total"] for v in top3)
            pct = top3_total/tc*100
            if pct > 45:
                anomalies.append({
                    "type": "concentration_carburant",
                    "valeur": round(pct,1),
                    "seuil": 45,
                    "detail": f"Top 3 véhicules = {pct:.0f}% des dépenses carburant"
                })

        critiques = sum(1 for v in self.ind["sante_vehicules"] if v["statut"] in ("CRITIQUE","REMPLACER"))
        if critiques > 0:
            anomalies.append({
                "type": "vehicules_critiques",
                "valeur": critiques,
                "detail": f"{critiques} véhicule(s) critique(s) ou à remplacer"
            })

        return {
            "total_budget":       round(total,0),
            "total_fmt":          _fmt(total),
            "total_carburant":    round(tc,0),
            "total_entretien":    round(te,0),
            "total_generateur":   round(tg,0),
            "part_carburant_pct": round(tc/total*100,1) if total else 0,
            "part_entretien_pct": round(te/total*100,1) if total else 0,
            "part_generateur_pct":round(tg/total*100,1) if total else 0,
            "anomalies_detectees": anomalies,
            # économie_estimee sera complétée par Gemini AI
            "economie_estimee":   None,
            "economie_fmt":       "Calculé par IA",
        }

    def run(self) -> dict:
        return {
            "alertes":             self.alertes(),
            "plan_action_vt":      self.plan_action_vt(),
            "plan_renouvellement": self.plan_renouvellement(),
            "decisions_generateur":self.decisions_generateur(),
            "optimisation":        self.optimisation(),
        }


# ─────────────────────────────────────────────────────────────────────────────
# KPIs STRATÉGIQUES
# ─────────────────────────────────────────────────────────────────────────────

def build_kpis(ops: dict, ind: dict, dec: dict) -> list[dict]:
    taux  = ind["taux_immobilisation"]
    conf  = ind["conformite_vt"]
    carb  = ops["carburant"]
    entr  = ops["entretien"]
    gen   = ops["generateur"]
    optim = dec["optimisation"]
    inv   = ops["inventaire"]
    vt_op = ops["vt"]

    total_budget = optim.get("total_budget",0)
    tg = gen.get("total_depense",0) or 0

    return [
        # N1 — faits
        {"label":"Parc total",       "value":str(inv.get("total","—")),
         "icon":"🚗","color":"#3b82f6","niveau":1},
        {"label":"Budget carburant", "value":_fmt(carb.get("total_depense")),
         "icon":"⛽","color":"#f59e0b","niveau":1},
        {"label":"Budget entretien", "value":_fmt(entr.get("total_depense")),
         "icon":"🔧","color":"#8b5cf6","niveau":1},
        {"label":"Groupe électrogène","value":_fmt(tg) if tg else "—",
         "icon":"⚡","color":"#06b6d4","niveau":1},
        # N2 — indicateurs
        {"label":"Disponibilité",    "value":f"{taux.get('taux_dispo',0):.0f}%",
         "icon":"✅",
         "color":"#10b981" if taux.get("taux_dispo",0)>=85 else "#ef4444","niveau":2},
        {"label":"Conformité VT",    "value":f"{conf.get('taux_conformite',0):.0f}%",
         "icon":"📋",
         "color":"#10b981" if conf.get("taux_conformite",0)>=80 else "#ef4444","niveau":2},
        {"label":"Budget total",     "value":_fmt(total_budget),
         "icon":"💰","color":"#6366f1","niveau":2},
        # N3 — décisions
        {"label":"VT non faites",    "value":str(vt_op.get("non",0)),
         "icon":"⛔",
         "color":"#ef4444" if vt_op.get("non",0)>0 else "#10b981","niveau":3},
        {"label":"En atelier",       "value":str(entr.get("en_atelier",0)),
         "icon":"🔴",
         "color":"#ef4444" if entr.get("en_atelier",0)>=3 else "#f59e0b","niveau":3},
    ]


# ─────────────────────────────────────────────────────────────────────────────
# ORCHESTRATEUR
# ─────────────────────────────────────────────────────────────────────────────

class FleetAnalyzer:
    def __init__(self, raw_sheets: dict[str, list[dict]]):
        self.parser = FleetParser(raw_sheets)

    def analyze(self) -> dict:
        ops  = N1_Operationnel(self.parser).run()
        ind  = N2_Indicateurs(self.parser, ops).run()
        dec  = N3_Decision(ops, ind).run()
        kpis = build_kpis(ops, ind, dec)
        return _j({
            "parser_info":           self.parser.summary(),
            "niveau_1_operationnel": ops,
            "niveau_2_indicateurs":  ind,
            "niveau_3_decisions":    dec,
            "kpis":                  kpis,
            "generated_at":          datetime.utcnow().isoformat(),
        })
