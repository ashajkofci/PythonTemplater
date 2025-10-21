# -*- coding: utf-8 -*-
"""
attestation.py (formerly generate_attestations.py)

Script de génération d'attestations Word à partir d'un modèle .docx
et d'un fichier CSV de donataires.

- Remplace les champs {NOM}, {MONTANT}, {DATE} du modèle Word fourni.
- Ne modifie jamais le template : un nouveau .docx est créé par donataire.
- Détecte au mieux les colonnes (prénom, nom, organisation, montant, civilité).
- Permet de forcer certaines colonnes via des arguments CLI.
- Gère des montants tels que "55 + 100" en les additionnant.
- Sortie : un .docx par donataire dans un dossier, et un ZIP optionnel.

Exemple d'utilisation :
    python attestation.py \
        --csv "MJ-FAM -Contacts-MAJ-25-AOUT(Donateurs 2022-2023-2024).csv" \
        --template "DONATION-AMIS-JENISCH.docx" \
        --date "Lausanne, le 16 octobre 2025" \
        --outdir "attestations_word" \
        --zip

Dépendances : pandas, python-docx
    pip install pandas python-docx
"""
import argparse
import os
import sys

# Import core functionality
from templater_core import read_csv_any, find_columns, find_amount_in_row, build_display_name, slugify
import pandas as pd
from docx import Document



import zipfile
from templater_core import replace_placeholders


def slugify(value: str) -> str:
    import re
    from templater_core import normalize_spaces
    value = normalize_spaces(value)
    value = re.sub(r"[^\w\s-]", "", value, flags=re.UNICODE)
    value = value.strip().replace(" ", "_")
    return value[:120]


def replace_placeholders_compat(doc: Document, mapping):
    """Wrapper for backward compatibility."""
    replace_placeholders(doc, mapping)


def run(csv_path, template_path, date_text, outdir, make_zip, first_col, last_col, org_col, civ_col, amount_col):
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"CSV introuvable: {csv_path}")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Modèle .docx introuvable: {template_path}")

    df = read_csv_any(csv_path)

    # Colonnes auto si non fournies
    auto_first, auto_last, auto_org, auto_civ, auto_amt = find_columns(df)
    first_col = first_col or auto_first
    last_col = last_col or auto_last
    org_col = org_col or auto_org
    civ_col = civ_col or auto_civ
    amount_col = amount_col or auto_amt

    os.makedirs(outdir, exist_ok=True)

    generated_files = []
    for _, row in df.iterrows():
        display_name = build_display_name(row, first_col, last_col, org_col, civ_col)
        amount = find_amount_in_row(row, amount_col)
        if amount <= 0:
            continue

        # Ouvre un nouveau Document basé sur le template pour CHAQUE donataire
        doc = Document(template_path)
        mapping = {
            "{NOM}": display_name,
            "{MONTANT}": f"{int(amount) if float(amount).is_integer() else amount:.0f}",
            "{DATE}": date_text,
        }
        replace_placeholders(doc, mapping)

        fname = f"Attestation_{slugify(display_name)}.docx"
        out_path = os.path.join(outdir, fname)
        doc.save(out_path)
        generated_files.append(out_path)

    zip_path = None
    if make_zip and generated_files:
        zip_path = os.path.join(outdir, "Attestations_Dons_2025_Word.zip")
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for fp in generated_files:
                zf.write(fp, arcname=os.path.basename(fp))

    return generated_files, zip_path


def main():
    parser = argparse.ArgumentParser(description="Générer des attestations Word depuis un CSV et un template .docx")
    parser.add_argument("--csv", default="MJ-FAM -Contacts-MAJ-25-AOUT(Donateurs 2022-2023-2024).csv", help="Chemin du CSV")
    parser.add_argument("--template", default="DONATION-AMIS-JENISCH.docx", help="Chemin du modèle .docx avec {NOM}, {MONTANT}, {DATE}")
    parser.add_argument("--date", default="Lausanne, le 16 octobre 2025", help="Texte pour {DATE}")
    parser.add_argument("--outdir", default="attestations_word", help="Dossier de sortie pour les .docx générés")
    parser.add_argument("--zip", action="store_true", help="Créer aussi une archive ZIP de sortie")
    # Colonnes optionnelles (si l'auto-détection ne convient pas)
    parser.add_argument("--first-col", dest="first_col", default=None, help="Nom exact de la colonne 'Prénom'")
    parser.add_argument("--last-col", dest="last_col", default=None, help="Nom exact de la colonne 'Nom'")
    parser.add_argument("--org-col", dest="org_col", default=None, help="Nom exact de la colonne 'Organisation'")
    parser.add_argument("--civ-col", dest="civ_col", default=None, help="Nom exact de la colonne 'Civilité'")
    parser.add_argument("--amount-col", dest="amount_col", default=None, help="Nom exact de la colonne 'Montant'")
    args = parser.parse_args()

    generated_files, zip_path = run(
        csv_path=args.csv,
        template_path=args.template,
        date_text=args.date,
        outdir=args.outdir,
        make_zip=args.zip,
        first_col=args.first_col,
        last_col=args.last_col,
        org_col=args.org_col,
        civ_col=args.civ_col,
        amount_col=args.amount_col,
    )

    print(f"Générés: {len(generated_files)} fichiers dans '{args.outdir}'")
    if zip_path:
        print(f"Archive: {zip_path}")


if __name__ == "__main__":
    main()
