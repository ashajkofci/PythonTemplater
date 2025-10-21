# -*- coding: utf-8 -*-
"""
templater_core.py

Core functionality for CSV to DOCX template generation.
Extracted from attestation.py to be reusable by both CLI and GUI.
"""
import os
import re
import sys
import zipfile
import pandas as pd
from docx import Document

# Encodage sûr pour Windows / Mac
DEFAULT_FS_ENCODING = sys.getfilesystemencoding() or "utf-8"

# Détection robuste du séparateur et encodage CSV
ORG_KEYWORDS = (
    "SA", "SÀRL", "SARL", "Sarl", "Association", "Fondation", "Société", "GmbH", "AG", "Ltd", "Inc"
)
FEMALE_HINTS = (
    "anne", "anna", "elle", "ette", "ine", "ene", "ène",
    "a", "ia", "ya", "na", "ina", "liane", "iane", "line", "rine",
    "otte", "ille", "ise", "yse", "cie", "lie", "rie", "nie", "xie",
    "zia", "cia", "tia", "ria", "aude", "onde", "hilde", "rude", "ude", "iette"
)
FEMALE_NAMES = {
    "virginia","gudrun","marianne","cécile","catherine","lise","nicole","sylvie",
    "eliane","blandine","monique","geneviève","laurence","liliane","ursula","herta",
    "paulette","françoise","elisabeth","elisa","christiane","cynthia","efinizia",
    "jacqueline","annemarie","myriam","liliana","anne","ramona","béatrice",
    "vivianne","thérèse","heidi","edith","monika","julia","iris","hélène","pauline",
    "marie","paola","beatrice","francoise","therese","monica","isabelle"
}


def read_csv_any(path: str) -> pd.DataFrame:
    """Lit un CSV en devinant encodage et séparateur."""
    encodings = ["utf-8-sig", "utf-8", "cp1252", "latin1"]
    last_exc = None
    for enc in encodings:
        try:
            df = pd.read_csv(path, sep=None, engine="python", encoding=enc, dtype=str)
            return df.fillna("")
        except Exception as e:
            last_exc = e
    raise RuntimeError(f"Impossible de lire le CSV: {path}\nDernière erreur: {last_exc}")


def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()


def find_columns(df: pd.DataFrame):
    """Devine les colonnes de prénom / nom / organisation / montant / civilité."""
    cols = {c: c.lower() for c in df.columns}
    def col_like(*keys):
        for k in cols:
            lk = cols[k]
            if any(kw in lk for kw in keys):
                return k
        return None

    first = col_like("prénom", "prenom", "first", "vorname")
    last = col_like("nom", "lastname", "last", "name")
    org = col_like("organisation", "société", "societe", "raison", "entreprise", "compagnie", "company", "institution")
    civ = col_like("civilit", "titre", "title", "civility")
    amt = col_like("montant", "amount", "don", "contribution")

    return first, last, org, civ, amt


def parse_amount(val: str) -> float:
    """
    Extrait un montant d'une chaîne, gère "55 + 100" -> 155.
    Ignore les codes postaux probables (>= 1000 et <= 9999) si mélangés à d'autres nombres.
    """
    if not val:
        return 0.0
    s = val.replace("\xa0", " ")
    # Si on a un pattern type "55 + 100 + 40"
    if "+" in s:
        parts = [p.strip() for p in re.split(r"\s*\+\s*", s) if p.strip()]
        total = 0.0
        for p in parts:
            pc = p.replace(" ", "").replace("\u00A0", "").replace(".-", "").replace(".–", "")
            pc = pc.replace(".", "").replace(",", ".")
            nums = re.findall(r"\d+(?:\.\d+)?", pc)
            if not nums:
                continue
            valf = float(nums[0])
            if 1000 <= valf <= 9999 and len(nums) > 1:
                continue
            total += valf
        return total

    # Sinon, on prend le dernier nombre « pertinent »
    nums = re.findall(r"\d{1,5}(?:[.,]\d+)?", s)
    if not nums:
        return 0.0
    # Retire les CP probables (4 chiffres) si d'autres nombres existent
    cleaned = []
    for n in nums:
        n2 = n.replace(".", "").replace(",", ".")
        try:
            f = float(n2)
        except Exception:
            continue
        if 1000 <= f <= 9999 and len(nums) > 1:
            continue
        cleaned.append(f)
    return cleaned[-1] if cleaned else 0.0


def find_amount_in_row(row, amount_col: str | None) -> float:
    if amount_col and amount_col in row:
        return parse_amount(str(row[amount_col]))
    # Essaye colonnes contenant 'montant'…
    for c in row.index:
        lc = c.lower()
        if any(k in lc for k in ["montant", "amount", "don"]):
            v = parse_amount(str(row[c]))
            if v > 0:
                return v
    # Sinon balaye de droite à gauche (probables colonnes fin de tableau)
    for c in reversed(row.index.tolist()):
        v = parse_amount(str(row[c]))
        if v > 0 and v < 50000:
            return v
    return 0.0


def is_organization(name_fields) -> bool:
    blob = " ".join([str(x) for x in name_fields]).strip()
    return any(k in blob for k in ORG_KEYWORDS)


def infer_civility(firstname: str, fallback="Monsieur/Madame") -> str:
    fn = normalize_spaces(firstname).lower()
    # Cas couples "X et Y"
    if " et " in f" {fn} " or " & " in f" {fn} ":
        return "Monsieur et Madame"
    # Heuristique "féminin" par dictionnaire / terminaison
    base = fn.split()[0] if fn else ""
    if base in FEMALE_NAMES:
        return "Madame"
    if any(base.endswith(suf) for suf in FEMALE_HINTS):
        return "Madame"
    return "Monsieur" if base else fallback


def build_display_name(row, first_col, last_col, org_col, civ_col):
    """
    Construit la chaîne {NOM} à injecter, avec civilité.
    Si 'organisation' détectée, on garde tel quel (sans civilité).
    """
    firstname = row.get(first_col, "") if first_col else ""
    lastname = row.get(last_col, "") if last_col else ""
    orgname = row.get(org_col, "") if org_col else ""

    # Nom « brut » potentiel depuis des colonnes combinées
    if not (firstname or lastname or orgname):
        for c in row.index:
            lc = c.lower()
            if any(k in lc for k in ["nom complet", "donataire", "beneficiaire", "raison sociale", "nom"]):
                candidate = normalize_spaces(str(row[c]))
                if candidate:
                    orgname = candidate
                    break

    # Entreprise / organisation
    if orgname and is_organization([orgname]):
        return normalize_spaces(orgname)

    # Couples du type "Hélène et Fabio" avec nom unique
    if firstname and (" et " in firstname or " & " in firstname):
        civ = "Monsieur et Madame"
        disp = normalize_spaces(f"{civ} {firstname} {lastname}".strip())
        return disp

    # Cas général personne physique
    civ = ""
    if civ_col and row.get(civ_col, ""):
        civ = normalize_spaces(str(row[civ_col]))
    else:
        civ = infer_civility(firstname)

    full = normalize_spaces(f"{civ} {firstname} {lastname}").strip()
    if not firstname and lastname:
        full = normalize_spaces(f"{civ} {lastname}")
    if not full:
        full = normalize_spaces(orgname or lastname or firstname or "Donataire")
    return full


def get_placeholders_from_template(template_path: str) -> list:
    """Extract placeholders from a DOCX template (e.g., {NOM}, {MONTANT})."""
    doc = Document(template_path)
    placeholders = set()
    
    # Check paragraphs
    for p in doc.paragraphs:
        matches = re.findall(r'\{[^}]+\}', p.text)
        placeholders.update(matches)
    
    # Check tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    matches = re.findall(r'\{[^}]+\}', p.text)
                    placeholders.update(matches)
    
    return sorted(list(placeholders))


def replace_placeholders(doc: Document, mapping):
    """Remplace les placeholders dans paragraphes ET tableaux."""
    def _replace_in_paragraph(p):
        for key, val in mapping.items():
            if key in p.text:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, val)

    # Paragraphes
    for p in doc.paragraphs:
        _replace_in_paragraph(p)

    # Tableaux
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_in_paragraph(p)


def slugify(value: str) -> str:
    value = normalize_spaces(value)
    value = re.sub(r"[^\w\s-]", "", value, flags=re.UNICODE)
    value = value.strip().replace(" ", "_")
    return value[:120]


def generate_documents(csv_path, template_path, outdir, field_mapping, 
                      filename_field=None, filename_prefix="", filename_suffix="",
                      make_zip=False, progress_callback=None):
    """
    Generate DOCX documents from CSV and template.
    
    Args:
        csv_path: Path to CSV file
        template_path: Path to DOCX template
        outdir: Output directory
        field_mapping: Dict mapping placeholder names to CSV column names
        filename_field: CSV column to use for filename (optional)
        filename_prefix: Prefix for output filenames
        filename_suffix: Suffix for output filenames
        make_zip: Whether to create a ZIP archive
        progress_callback: Callback function(current, total, message)
    
    Returns:
        (generated_files, zip_path)
    """
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"CSV introuvable: {csv_path}")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Modèle .docx introuvable: {template_path}")

    df = read_csv_any(csv_path)
    os.makedirs(outdir, exist_ok=True)

    generated_files = []
    total_rows = len(df)
    
    for idx, row in df.iterrows():
        if progress_callback:
            progress_callback(idx, total_rows, f"Processing row {idx+1}/{total_rows}")
        
        # Build mapping from template placeholders to values
        mapping = {}
        skip_row = False
        
        for placeholder, csv_column in field_mapping.items():
            if csv_column and csv_column in row:
                value = str(row[csv_column]).strip()
                mapping[placeholder] = value
            else:
                mapping[placeholder] = ""
        
        # Skip empty rows or rows without critical data
        if not any(mapping.values()):
            continue
        
        # Open a new Document based on the template for EACH row
        doc = Document(template_path)
        replace_placeholders(doc, mapping)
        
        # Determine filename
        if filename_field and filename_field in row:
            base_name = slugify(str(row[filename_field]))
        else:
            # Use first non-empty value as fallback
            base_name = slugify(next((v for v in mapping.values() if v), f"document_{idx}"))
        
        fname = f"{filename_prefix}{base_name}{filename_suffix}.docx"
        out_path = os.path.join(outdir, fname)
        
        # Handle duplicate filenames
        counter = 1
        while os.path.exists(out_path):
            fname = f"{filename_prefix}{base_name}_{counter}{filename_suffix}.docx"
            out_path = os.path.join(outdir, fname)
            counter += 1
        
        doc.save(out_path)
        generated_files.append(out_path)

    zip_path = None
    if make_zip and generated_files:
        zip_path = os.path.join(outdir, "generated_documents.zip")
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for fp in generated_files:
                zf.write(fp, arcname=os.path.basename(fp))
    
    if progress_callback:
        progress_callback(total_rows, total_rows, f"Complete! Generated {len(generated_files)} files")

    return generated_files, zip_path
