"""
FleetInsight AI — IA Advisor
=====================================
Utilise Google IA (via AI Studio API) pour :
  1. Interpréter automatiquement les résultats d'analyse
  2. Calculer dynamiquement le potentiel d'économie
  3. Générer des recommandations stratégiques en langage naturel
  4. Analyser le groupe électrogène
  5. Produire un résumé exécutif

API : Google AI Studio  →  https://generativelanguage.googleapis.com
Modèle : gemini-2.5-flash-lite  (rapide, économique)
"""

from __future__ import annotations
import json
import re
import os
import httpx
from typing import Any

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GEMINI_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    "gemini-2.5-flash-lite:generateContent"
)


# ─────────────────────────────────────────────────────────────────────────────
# APPEL API IA
# ─────────────────────────────────────────────────────────────────────────────

async def _gemini(prompt: str, json_mode: bool = True) -> str:
    """Envoie un prompt à Gemini et retourne le texte de réponse."""
    if not GEMINI_API_KEY:
        return json.dumps({"error": "GEMINI_API_KEY non configurée"})

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.3,
            "maxOutputTokens": 2048,
        },
    }

    try:
        async with httpx.AsyncClient(timeout=30) as client:
            res = await client.post(
                f"{GEMINI_URL}?key={GEMINI_API_KEY}",
                json=payload,
            )
            res.raise_for_status()
            data = res.json()
            text = data["candidates"][0]["content"]["parts"][0]["text"]
            # Nettoyer les balises markdown ```json ... ```
            text = re.sub(r"```(?:json)?\s*", "", text).strip().rstrip("`").strip()
            return text
    except Exception as e:
        return json.dumps({"error": str(e)})


def _parse_json(text: str) -> dict:
    """Parse le JSON retourné par Gemini, avec fallback."""
    try:
        return json.loads(text)
    except Exception:
        return {"raw_response": text, "parse_error": True}


# ─────────────────────────────────────────────────────────────────────────────
# FONCTIONS D'ANALYSE IA
# ─────────────────────────────────────────────────────────────────────────────

async def analyze_budget_and_savings(analysis: dict) -> dict:
    """
    Gemini calcule le potentiel d'économie RÉEL basé sur les données.
    Remplace l'ancien facteur figé de 0.12.
    """
    ops   = analysis.get("niveau_1_operationnel", {})
    ind   = analysis.get("niveau_2_indicateurs", {})
    optim = analysis.get("niveau_3_decisions", {}).get("optimisation", {})

    prompt = f"""Tu es un expert en gestion de flotte automobile en Côte d'Ivoire.
Analyse ce budget de parc automobile et calcule le potentiel d'économie RÉEL.

DONNÉES BUDGET :
- Budget carburant total : {ops.get('carburant',{}).get('total_depense', 0):,.0f} FCFA
- Budget entretien total : {ops.get('entretien',{}).get('total_depense', 0):,.0f} FCFA
- Budget groupe électrogène : {ops.get('generateur',{}).get('total_depense', 0):,.0f} FCFA
- Budget total parc : {optim.get('total_budget', 0):,.0f} FCFA
- Part carburant : {optim.get('part_carburant_pct', 0)}%
- Part entretien : {optim.get('part_entretien_pct', 0)}%

ANOMALIES DÉTECTÉES :
{json.dumps(optim.get('anomalies_detectees', []), ensure_ascii=False, indent=2)}

SANTÉ VÉHICULES :
- Véhicules à remplacer : {len([v for v in ind.get('sante_vehicules',[]) if v['statut']=='REMPLACER'])}
- Véhicules critiques : {len([v for v in ind.get('sante_vehicules',[]) if v['statut']=='CRITIQUE'])}
- Taux disponibilité : {ind.get('taux_immobilisation',{}).get('taux_dispo', 0)}%
- Taux conformité VT : {ind.get('conformite_vt',{}).get('taux_conformite', 0)}%

TOP 3 consommateurs carburant :
{json.dumps(ops.get('carburant',{}).get('top_consommateurs',[])[:3], ensure_ascii=False)}

Réponds UNIQUEMENT en JSON valide avec cette structure exacte :
{{
  "economie_estimee": <nombre FCFA>,
  "economie_pct": <pourcentage entre 5 et 40>,
  "economie_fmt": "<string formaté ex: 50K FCFA>",
  "leviers": [
    {{"levier": "<action>", "economie_potentielle": <FCFA>, "priorite": "<HAUTE|MOYENNE|FAIBLE>", "detail": "<explication>"}}
  ],
  "score_gestion": <0 à 100>,
  "niveau_maturite": "<DÉBUTANT|EN DÉVELOPPEMENT|AVANCÉ|EXPERT>",
  "synthese": "<2-3 phrases d'analyse en français>"
}}"""

    text   = await _gemini(prompt)
    result = _parse_json(text)

    # Fallback si Gemini indisponible
    if "error" in result or "parse_error" in result:
        total = optim.get("total_budget", 0) or 0
        # Calcul fallback basé sur les anomalies réelles
        pct = 0.0
        anomalies = optim.get("anomalies_detectees", [])
        for a in anomalies:
            if a["type"] == "ratio_entretien_eleve": pct += 0.08
            if a["type"] == "concentration_carburant": pct += 0.10
            if a["type"] == "vehicules_critiques":     pct += 0.05
        pct = max(0.05, min(pct, 0.30))
        eco = round(total * pct, 0)
        result = {
            "economie_estimee": eco,
            "economie_pct": round(pct * 100, 0),
            "economie_fmt": f"{eco/1000:.0f}K FCFA" if eco >= 1000 else f"{eco:.0f} FCFA",
            "leviers": [],
            "score_gestion": 60,
            "niveau_maturite": "EN DÉVELOPPEMENT",
            "synthese": "Analyse IA indisponible. Estimation basée sur les anomalies détectées.",
            "gemini_unavailable": True,
        }

    return result


async def interpret_vt_compliance(vt_data: dict, conformite: dict) -> dict:
    """Gemini interprète la conformité VT et génère des recommandations légales."""
    prompt = f"""Tu es un expert en réglementation automobile en Côte d'Ivoire.

DONNÉES VISITES TECHNIQUES (VT) :
- Total véhicules : {vt_data.get('total', 0)}
- Statut OUI (VT faite) : {vt_data.get('oui', 0)} — {conformite.get('taux_conformite', 0)}%
- Statut NON (VT non faite) : {vt_data.get('non', 0)}
- Statut PAS ENCORE : {vt_data.get('pas_encore', 0)}
- Niveau de risque : {conformite.get('niveau_risque', '—')}

Véhicules sans VT (statut NON) :
{json.dumps([v['vehicule'] for v in vt_data.get('liste_non', [])], ensure_ascii=False)}

Réponds UNIQUEMENT en JSON valide :
{{
  "risque_juridique": "<description du risque légal en CI>",
  "sanctions_possibles": "<amendes, retraits de permis, etc.>",
  "recommandation_urgente": "<action immédiate à prendre>",
  "plan_mise_en_conformite": "<étapes concrètes>",
  "impact_assurance": "<conséquences sur la couverture assurance>",
  "message_direction": "<message court pour la direction, en 1 phrase percutante>"
}}"""

    text = await _gemini(prompt)
    return _parse_json(text)


async def analyze_generator(gen_data: dict, gen_ind: dict) -> dict:
    """Gemini analyse les dépenses du groupe électrogène par produit et montant."""
    if not gen_data.get("disponible"):
        return {"analyse": "Données groupe électrogène non disponibles"}

    # Préparation des données détaillées produit/montant
    cout_produit_str = json.dumps(gen_data.get('cout_par_produit', [])[:10], ensure_ascii=False)
    cout_moyen_produit_str = json.dumps(gen_data.get('cout_moyen_produit', [])[:10], ensure_ascii=False)
    par_agence_produit_str = json.dumps(gen_data.get('par_agence_produit', [])[:15], ensure_ascii=False)

    prompt = f"""Tu es un expert en gestion d'énergie et d'équipements industriels en Afrique de l'Ouest.

DONNÉES GROUPE ÉLECTROGÈNE (analyse PRODUIT + MONTANT + AGENCE) :
- Nombre d'entrées : {gen_data.get('nb_entrees', 0)}
- Dépense totale : {gen_data.get('total_depense', 0):,.0f} FCFA
- Coût moyen par transaction : {gen_data.get('cout_moyen', 0):,.0f} FCFA
- Coût maximum observé : {gen_data.get('cout_max', 0):,.0f} FCFA
- Période couverte : {gen_data.get('date_min', '—')} → {gen_data.get('date_max', '—')}
- Nombre de jours : {gen_data.get('nb_jours_couverts', 0)}
- Nombre de produits utilisés : {gen_ind.get('nb_produits', 0)}
- Nombre d'agences : {gen_ind.get('nb_agences', 0)}

ANALYSE PAR PRODUIT :
- Répartition : {json.dumps(gen_data.get('repartition_produit', []), ensure_ascii=False)}
- Coût par produit (TOP 10) : {cout_produit_str}
- Coût moyen par produit : {cout_moyen_produit_str}

CROISEMENT AGENCE × PRODUIT × MONTANT (TOP 15) :
{par_agence_produit_str}

DÉPENSES PAR AGENCE :
{json.dumps(gen_data.get('cout_par_agence', []), ensure_ascii=False)}

Analyse ces données et réponds UNIQUEMENT en JSON valide :
{{
  "diagnostic": "<analyse de la consommation par produit/montant/agence>",
  "niveau_consommation": "<FAIBLE|NORMAL|ÉLEVÉ|CRITIQUE>",
  "cout_journalier_moyen": <FCFA>,
  "produit_principal": {{"nom": "<nom>", "pct": <pct:float>, "recommandation": "<action>"}},
  "anomalies_detectees": ["<anomalie1>", "<anomalie2>"],
  "recommendations_produit": ["<rec1>", "<rec2>"],
  "recommendations_agence": ["<rec1>", "<rec2>"],
  "optimisation_possible": "<description économies potentielles par produit/montant>",
  "alerte": "<null ou alerte CRITIQUE>"
}}"""

    text = await _gemini(prompt)
    return _parse_json(text)


async def generate_executive_summary(analysis: dict, savings: dict, vt_insight: dict) -> dict:
    """
    IA génère un résumé exécutif complet du parc automobile,
    prêt à être présenté à la direction.
    """
    ops  = analysis.get("niveau_1_operationnel", {})
    ind  = analysis.get("niveau_2_indicateurs", {})
    dec  = analysis.get("niveau_3_decisions", {})
    inv  = ops.get("inventaire", {})
    taux = ind.get("taux_immobilisation", {})
    conf = ind.get("conformite_vt", {})

    prompt = f"""Tu es le Directeur Logistique d'une grande entreprise en Côte d'Ivoire.
Rédige un résumé exécutif de la situation du parc automobile, destiné à la Direction Générale.

CHIFFRES CLÉS :
- Parc total : {inv.get('total', 0)} véhicules
- Disponibilité : {taux.get('taux_dispo', 0)}% ({taux.get('en_atelier', 0)} en atelier)
- Conformité VT : {conf.get('taux_conformite', 0)}% — Risque {conf.get('niveau_risque', '—')}
- VT non faites : {ops.get('vt',{}).get('non', 0)} véhicules (risque juridique)
- Budget carburant : {ops.get('carburant',{}).get('total_depense', 0):,.0f} FCFA
- Budget entretien : {ops.get('entretien',{}).get('total_depense', 0):,.0f} FCFA
- Budget groupe électrogène : {ops.get('generateur',{}).get('total_depense', 0):,.0f} FCFA
- Score de gestion : {savings.get('score_gestion', '—')}/100
- Économie potentielle : {savings.get('economie_fmt', '—')}
- Véhicules à remplacer : {dec.get('plan_renouvellement',{}).get('nb_remplacement', 0)}

ALERTES ACTIVES : {len(dec.get('alertes', []))} alertes
- Critiques : {sum(1 for a in dec.get('alertes',[]) if a['niveau']=='CRITIQUE')}
- Warnings  : {sum(1 for a in dec.get('alertes',[]) if a['niveau']=='WARNING')}

Réponds UNIQUEMENT en JSON valide :
{{
  "titre": "<titre du rapport>",
  "date_rapport": "<date du jour>",
  "statut_global": "<CRITIQUE|PRÉOCCUPANT|ACCEPTABLE|BON|EXCELLENT>",
  "message_cle": "<1 phrase percutante sur l'état du parc>",
  "points_forts": ["<force 1>", "<force 2>"],
  "points_attention": ["<point 1>", "<point 2>", "<point 3>"],
  "decisions_urgentes": ["<decision 1>", "<decision 2>"],
  "objectif_30_jours": "<objectif principal à atteindre>",
  "resume_narratif": "<paragraphe de 4-5 phrases pour la direction>"
}}"""

    text = await _gemini(prompt)
    return _parse_json(text)


# ─────────────────────────────────────────────────────────────────────────────
# ORCHESTRATEUR IA
# ─────────────────────────────────────────────────────────────────────────────

async def run_gemini_analysis(analysis: dict) -> dict:
    """
    Lance toutes les analyses IA en parallèle et retourne les résultats.
    """
    import asyncio

    ops = analysis.get("niveau_1_operationnel", {})
    ind = analysis.get("niveau_2_indicateurs", {})

    # Lancer en parallèle
    savings_task   = analyze_budget_and_savings(analysis)
    vt_task        = interpret_vt_compliance(
        ops.get("vt", {}),
        ind.get("conformite_vt", {})
    )
    gen_task       = analyze_generator(
        ops.get("generateur", {}),
        ind.get("indicateurs_generateur", {})
    )

    savings, vt_insight, gen_insight = await asyncio.gather(
        savings_task, vt_task, gen_task
    )

    # Résumé exécutif (nécessite les résultats précédents)
    executive = await generate_executive_summary(analysis, savings, vt_insight)

    return {
        "budget_et_economies":  savings,
        "conformite_vt":        vt_insight,
        "groupe_electrogene":   gen_insight,
        "resume_executif":      executive,
        "gemini_model":         "gemini-2.0-flash",
        "powered_by":           "Google AI Studio",
    }


async def answer_question(analysis: dict, question: str) -> dict:
    """Répond à une question basée sur le contexte d'analyse du parc."""
    if not question or not question.strip():
        return {"error": "Question vide"}

    ops   = analysis.get("niveau_1_operationnel", {})
    ind   = analysis.get("niveau_2_indicateurs", {})
    dec   = analysis.get("niveau_3_decisions", {})

    prompt = f"""Tu es un assistant expert en gestion de flotte automobile en Côte d'Ivoire.
Tu disposes des résultats d'analyse du parc suivants :
- Opérationnel : {json.dumps(ops, ensure_ascii=False, indent=2)}
- Indicateurs : {json.dumps(ind, ensure_ascii=False, indent=2)}
- Décisions : {json.dumps(dec, ensure_ascii=False, indent=2)}

Question : {question}

Réponds en français de façon claire et structurée, avec des actions recommandées.
"""

    text = await _gemini(prompt)
    if not text:
        return {"error": "Pas de réponse de l'IA"}

    return {"answer": text}
