import os
from datetime import date
import re
import pandas as pd
from jinja2 import Environment, FileSystemLoader

def generate_suivi_html(
    fichier_excel: str,
    sortie_html: str = None,
    nom_feuille: str = "Suivi Actions",
    ordre_voulu=None,
    ingenieurs_en_conge=None,
    colonnes=None,
    tri_priorite_ascendant: bool = True,
    page_length: int = 25,
    trier_autres_alpha: bool = False,
    nom_col_priorite_affiche: str = "Priorité"
) -> str:
    """
    Génère un HTML unique listant les actions par ingénieur.
    - `ingenieurs_en_conge`: liste d'ingénieurs à marquer "En congé" (section sans tableau).
    - `ordre_voulu`: ordre imposé d’ingénieurs avant d’ajouter les autres.
    - `trier_autres_alpha`: True => les "autres" sont ajoutés triés par ordre alphabétique.
    - `tri_priorite_ascendant`: True => 1->10 ; False => 10->1
    - `page_length`: nombre de lignes affichées par page (DataTables).
    - `nom_col_priorite_affiche`: nom de la colonne "Priorité" à afficher dans le tableau (si votre Excel varie).
    
    Retourne le chemin du fichier HTML généré.
    """
    if ordre_voulu is None:
        ordre_voulu = []
    if ingenieurs_en_conge is None:
        ingenieurs_en_conge = []
    if colonnes is None:
        colonnes = [
            "Bâtiments",
            "Intitulé de l'action",
            "Date souhaitée         (demandeur)",
            nom_col_priorite_affiche,
            "Avancement de l'action (décision, commentaire,…)",
            "Type (Machine/Humain/Deux)",
            "Etat"
        ]

    # --- I/O & date ---
    today = date.today().strftime("%Y-%m-%d")
    if sortie_html is None:
        base_dir = os.path.dirname(fichier_excel) or "."
        sortie_html = os.path.join(base_dir, f"Tableaux_Suivi_{today}_tous.html")
    os.makedirs(os.path.dirname(sortie_html) or ".", exist_ok=True)

    # --- Lecture & préparation ---
    df = pd.read_excel(fichier_excel, sheet_name=nom_feuille, skiprows=10)
    df = df[df["Etat"].isin(["En cours", "Non démarrée"])].copy()
    df = df.fillna("")

    # colonne de priorité (numérique) pour tri
    # (on s'appuie sur la colonne d'affichage demandée)
    if nom_col_priorite_affiche not in df.columns:
        raise ValueError(f"La colonne '{nom_col_priorite_affiche}' est introuvable dans la feuille '{nom_feuille}'.")
    df["__prio_num"] = pd.to_numeric(df[nom_col_priorite_affiche], errors="coerce").fillna(0).astype(int)

    # liste des ingénieurs
    present = list(df["Prise en charge par"].dropna().unique())
    autres = [n for n in present if n not in ordre_voulu]
    if trier_autres_alpha:
        autres = sorted(autres)
    ingenieurs = [n for n in ordre_voulu if n in present] + autres

    # --- Préparation des données pour le template ---
    order_dir = 'asc' if tri_priorite_ascendant else 'desc'
    ICON = {"Machine": "⚙️", "Humain": "👤", "Deux": "🤝"}
    sections = []
    for resp in ingenieurs:
        section = {"nom": resp, "id": resp.replace(" ", "_")}
        if resp in ingenieurs_en_conge:
            section["conge"] = True
            sections.append(section)
            continue
        grp = df[df["Prise en charge par"] == resp].copy()
        if grp.empty:
            section["rows"] = []
            sections.append(section)
            continue
        grp = grp.sort_values("__prio_num", ascending=tri_priorite_ascendant)
        rows = []
        for _, row in grp.iterrows():
            row_vals = []
            for col in colonnes:
                if col == nom_col_priorite_affiche:
                    val = int(row["__prio_num"])
                elif col == "Type (Machine/Humain/Deux)":
                    val = f"{ICON.get(row[col], '')} {row[col]}"
                else:
                    val = row[col]
                row_vals.append(val)
            rows.append(row_vals)
        section["rows"] = rows
        sections.append(section)

    templates_dir = os.path.join(os.path.dirname(__file__), "templates")
    env = Environment(loader=FileSystemLoader(templates_dir))
    template = env.get_template("suivi.html.j2")
    html = template.render(
        today=today,
        ingenieurs=ingenieurs,
        colonnes=colonnes,
        sections=sections,
        page_length=page_length,
        order_dir=order_dir,
    )

    with open(sortie_html, "w", encoding="utf-8") as f:
        f.write(html)

    return sortie_html
# ======================
# Exemple d'utilisation :
# ======================

if __name__ == "__main__":
    fichier_excel  = r"C:\Users\i.bitar\OneDrive - EGIS Group\Documents\Professionnel\SHARED FOLDERS\SUIVI_PRODUCTION\Suivi_de_production.xlsm"
    sortie_html    = fr"C:\Users\i.bitar\OneDrive - EGIS Group\Documents\Professionnel\SHARED FOLDERS\SUIVI_PRODUCTION\Tableaux_Suivi_{date.today().strftime('%Y-%m-%d')}_tous.html"

    # ordre imposé
    ordre = [ "Viet", "Matthieu",  "Vinh","Maxime", "Ibrahim"]

    # personnes en congé (optionnel) — ex. ["Viet", "Nora"]
    en_conge = ["Samih", "Benjamin","Guillaume"]

    path = generate_suivi_html(
        fichier_excel=fichier_excel,
        sortie_html=sortie_html,
        nom_feuille="Suivi Actions",
        ordre_voulu=ordre,
        ingenieurs_en_conge=en_conge,
        tri_priorite_ascendant=True,   # 1 -> 10
        page_length=25,
        trier_autres_alpha=False,      # mets True si tu veux les autres triés A→Z
        nom_col_priorite_affiche="Priorité"
    )
    print(f"✅ Fichier généré : {path}")
