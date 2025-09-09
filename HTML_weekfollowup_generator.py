import os
from datetime import date
import re
import html
import pandas as pd

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
    nom_col_priorite_affiche: str = "Priorité",
    skiprows: int = 10,
    etats_conserves=None,
) -> str:
    """
    Génère un HTML unique listant les actions par ingénieur.
    - `ingenieurs_en_conge`: liste d'ingénieurs à marquer "En congé" (section sans tableau).
    - `ordre_voulu`: ordre imposé d’ingénieurs avant d’ajouter les autres.
    - `trier_autres_alpha`: True => les "autres" sont ajoutés triés par ordre alphabétique.
    - `tri_priorite_ascendant`: True => 1->10 ; False => 10->1
    - `page_length`: nombre de lignes affichées par page (DataTables).
    - `nom_col_priorite_affiche`: nom de la colonne "Priorité" à afficher dans le tableau (si votre Excel varie).
    - `skiprows`: nombre de lignes ignorées au début de la feuille Excel.
    - `etats_conserves`: liste des états conservés dans la colonne "Etat".

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
    if etats_conserves is None:
        etats_conserves = ["En cours", "Non démarrée"]
    df = pd.read_excel(fichier_excel, sheet_name=nom_feuille, skiprows=skiprows)
    df = df[df["Etat"].isin(etats_conserves)].copy()
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

    # --- HTML head ---
    order_dir = 'asc' if tri_priorite_ascendant else 'desc'
    html_output = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<title>Suivi des actions – {today}</title>
<link rel="stylesheet" href="https://cdn.datatables.net/1.13.4/css/jquery.dataTables.min.css"/>
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>
<style>
  body {{ font-family: Arial, sans-serif; margin:20px; }}
  h1, h2 {{ font-family: Arial, sans-serif; }}
  .toc ul {{
    display: flex; flex-wrap: wrap; gap: 8px;
    padding: 0; margin: 0 0 30px 0;
  }}
  .toc li {{ list-style: none; }}
  .toc a {{
    display: block;
    padding: 6px 12px;
    background: #f2f2f2;
    border-radius: 20px;
    text-decoration: none;
    color: #0066cc;
    transition: background 0.2s;
  }}
  .toc a:hover {{ background: #ddeeff; }}
  table {{ table-layout: fixed; width:100%; border-collapse: collapse; margin-bottom:40px; }}
  th, td {{ padding:6px; border:1px solid #ddd; word-wrap: break-word; }}
  th {{ background:#e6e6e6; font-size:14px; text-align:left; }}
  td {{ font-size:12px; text-align:left; white-space: pre-wrap; }}
  .en-conge {{ color:#a00; font-style: italic; margin: 6px 0 16px 0; }}
</style>
<script>
$(document).ready(function() {{
    $('table.display').each(function() {{
      $(this).DataTable({{
        paging:      true,
        pageLength:  {page_length},
        ordering:    true,
        order:       [[3, '{order_dir}']],      // tri sur la colonne Priorité
        columnDefs:  [{{ targets: 3, type: 'num' }}],
        fixedHeader: true,
        scrollX:     true,
        language:    {{ url: 'https://cdn.datatables.net/plug-ins/1.13.4/i18n/fr-FR.json' }}
      }});
    }});
}});
</script>
</head>
<body>
<h1>Suivi des actions – {today}</h1>

<nav class="toc">
  <ul>
"""

    # --- Sommaire
    for nom in ingenieurs:
        anchor = nom.replace(" ", "_")
        html_output += f"    <li><a href='#{anchor}'>{nom}</a></li>\n"

    html_output += """  </ul>
</nav>

<p><strong>Conventions de lecture du tableau :</strong></p>
<ul>
  <li><strong>Actions “Machine”</strong> : Toute action de calcul machine dont les données sont prêtes doit être priorisée, car il s’agit de temps machine et non de temps humain.</li>
</ul>
"""

    ICON = {"Machine": "⚙️", "Humain": "👤", "Deux": "🤝"}

    # --- Sections par ingénieur
    for resp in ingenieurs:
        # Section header
        html_output += "<hr style='margin:40px 0; border:none; border-top:1px solid #ccc;'/>\n"
        html_output += f"<h2 id='{resp.replace(' ','_')}'>Actions de {resp}</h2>\n"

        # Si l'ingénieur est en congé -> note et on passe à la suite
        if resp in ingenieurs_en_conge:
            html_output += "<p class='en-conge'>En congé — pas d'actions listées pour cette période.</p>\n"
            continue

        grp = df[df["Prise en charge par"] == resp].copy()
        if grp.empty:
            html_output += "<p class='en-conge'>Aucune action à afficher.</p>\n"
            continue

        # Tri Python (complémentaire au tri DataTables côté client)
        grp = grp.sort_values("__prio_num", ascending=tri_priorite_ascendant)

        # Construction du tableau
        html_output += "<table class='display'><thead><tr>\n"
        for col in colonnes:
            html_output += f"  <th>{col}</th>\n"
        html_output += "</tr></thead><tbody>\n"

        for _, row in grp.iterrows():
            html_output += "<tr>"
            for col in colonnes:
                if col == nom_col_priorite_affiche:
                    val = int(row["__prio_num"])
                elif col == "Type (Machine/Humain/Deux)":
                    val = f"{ICON.get(row[col],'')} {row[col]}"
                else:
                    val = row[col]
                val = html.escape(str(val))
                html_output += f"<td>{val}</td>"
            html_output += "</tr>\n"

        html_output += "</tbody></table>\n"

    html_output += "</body>\n</html>"

    # --- Écriture du fichier
    with open(sortie_html, "w", encoding="utf-8") as f:
        f.write(html_output)

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
        nom_col_priorite_affiche="Priorité",
        skiprows=10,
        etats_conserves=["En cours", "Non démarrée"],
    )
    print(f"✅ Fichier généré : {path}")
