import pandas as pd
import re
import os
import string

# ----- Étape 1 : Extraction -----
def extraire_siren_depuis_excel(fichier_excel, feuille='User', dossier_sortie='OUT/GENERIX'):
    os.makedirs(dossier_sortie, exist_ok=True)
    nom_base = os.path.splitext(os.path.basename(fichier_excel))[0]
    fichier_sortie = os.path.join(dossier_sortie, f"{nom_base}_extrait.xlsx")

    df = pd.read_excel(fichier_excel, sheet_name=feuille, engine='openpyxl')
    colonnes = ['identification', 'login', 'role']
    df_extrait = df[colonnes].copy()

    # 🔁 Dédoublonnage intelligent
    lignes_avant = len(df_extrait)
    df_extrait['has_role'] = df_extrait['role'].notna() & df_extrait['role'].astype(str).str.strip().ne('')
    df_extrait.sort_values(by='has_role', ascending=False, inplace=True)
    df_extrait = df_extrait.drop_duplicates(subset='identification', keep='first')
    df_extrait.drop(columns='has_role', inplace=True)
    lignes_apres = len(df_extrait)
    doublons_supprimes = lignes_avant - lignes_apres

    # 🧪 Extraction du SIREN
    def extraire_siren(identifiant):
        match = re.match(r'^FR\d{2}(\d{9})$', str(identifiant))
        return match.group(1) if match else 'étrangère ou inconnue'

    df_extrait['SIREN'] = df_extrait['identification'].apply(extraire_siren)
    df_extrait.to_excel(fichier_sortie, index=False, engine='openpyxl')

    print(f"✅ Fichier XLSX généré : {fichier_sortie}")
    print(f"ℹ️ {doublons_supprimes} doublon(s) supprimé(s) sur la colonne 'identification'")
    return fichier_sortie

# ----- Fonction utilitaire -----
def lettre_colonne_vers_index(lettre):
    lettre = lettre.strip().upper()
    return sum([(string.ascii_uppercase.index(c) + 1) * (26 ** i) 
                for i, c in enumerate(reversed(lettre))]) - 1

# ----- Étape 2 : Enrichissement -----
def enrichir_referentiel(fichier_sortie, dossier_injection='IN/INJECTION-DES-EDI'):
    fichiers_injection = [f for f in os.listdir(dossier_injection) if f.lower().endswith('.xlsx')]
    if not fichiers_injection:
        print("❌ Aucun fichier trouvé dans IN/INJECTION-DES-EDI.")
        return
    elif len(fichiers_injection) > 1:
        print("❌ Plusieurs fichiers trouvés dans IN/INJECTION-DES-EDI.")
        return
    chemin_injection = os.path.join(dossier_injection, fichiers_injection[0])

    df_injection = pd.read_excel(chemin_injection, engine='openpyxl')
    df_injection_extrait = df_injection[['identification', 'registration']].copy()
    df_injection_extrait.rename(columns={'registration': 'SIREN'}, inplace=True)
    df_injection_extrait['role'] = 'GNX-ADMINISTRATEURLECTURESEULE'

    df_sortie = pd.read_excel(fichier_sortie, engine='openpyxl')
    df_enrichi = pd.concat([df_sortie, df_injection_extrait], ignore_index=True)
    df_enrichi.to_excel(fichier_sortie, index=False, engine='openpyxl')

    print(f"✅ Référentiel enrichi avec les données d'injection du fichier IN/INJECTION-DES-EDI : {fichier_sortie}")

# ----- Étape 3 : Comparaison -----
def comparer_avec_export(fichier_out='OUT/GENERIX', fichier_export='IN/EXPORT-FILIALE-A-COMPARER'):
    fichiers_export = [f for f in os.listdir(fichier_export) if f.lower().endswith('.xlsx')]
    if not fichiers_export:
        print("❌ Aucun fichier trouvé dans IN/EXPORT-FILIALE-A-COMPARER.")
        return
    elif len(fichiers_export) > 1:
        print("❌ Plusieurs fichiers trouvés dans IN/EXPORT-FILIALE-A-COMPARER.")
        return
    chemin_export = os.path.join(fichier_export, fichiers_export[0])

    fichiers_out = [f for f in os.listdir(fichier_out) if f.lower().endswith('.xlsx')]
    if not fichiers_out:
        print("❌ Aucun fichier généré trouvé dans OUT.")
        return
    chemin_out = os.path.join(fichier_out, fichiers_out[0])

    df_out = pd.read_excel(chemin_out, engine='openpyxl')
    df_export = pd.read_excel(chemin_export, engine='openpyxl')

    print(f"\n🔍 Fichier à comparer : {fichiers_export[0]}")
    print("Ex : colonne A = TVA, colonne C = SIREN")

    tva_col = input("🤔 Lettre de la colonne contenant le Numéro TVA : ").strip()
    siren_col = input("🤔 Lettre de la colonne contenant le SIREN : ").strip()

    try:
        col_tva = lettre_colonne_vers_index(tva_col)
        col_siren = lettre_colonne_vers_index(siren_col)
    except Exception:
        print("❌ Erreur dans les lettres de colonnes fournies.")
        return

    tva_values = df_export.iloc[:, col_tva].astype(str).str.strip().str.upper()
    siren_values = df_export.iloc[:, col_siren].astype(str).str.strip()

    df_out['Correspondance'] = ""
    df_out['login_upper'] = df_out['login'].astype(str).str.strip().str.upper()
    df_out['SIREN_str'] = df_out['SIREN'].astype(str).str.strip()

    tva_values_clean = tva_values.replace(['NAN', 'NONE', ''], None)
    siren_values_clean = siren_values.replace(['NAN', 'NONE', '', '000000NAN'], None)

    # Recherche des correspondances (sans affichage de détail)
    for idx, (tva, siren) in enumerate(zip(tva_values_clean, siren_values_clean)):
        if pd.isna(tva) and pd.isna(siren):
            continue

        match_tva = df_out['login_upper'] == tva if not pd.isna(tva) else False
        match_siren = df_out['SIREN_str'] == siren if not pd.isna(siren) else False
        masque = match_tva | match_siren

        if masque.any():
            df_out.loc[masque, 'Correspondance'] = "found"

    df_out.drop(columns=['login_upper', 'SIREN_str'], inplace=True)

    # ----- Mapping rôle → canal -----
    role_to_canal = {
        "GNX-ADMINISTRATEUREPDF2C": "EPDF",
        "GNX-UTILISATEUREPDF2C": "EPDF",
        "GNX-ADMINISTRATEURPORTAIL2CEPDFOA": "EPDF",
        "GNX-LECTURESEULE": "EDI",
        "GNX-ADMINISTRATEUREPDFOA": "OA",
        "GNX-UTILISATEUREPDFOA": "OA",
        "GNX-ADMINISTRATEURPORTAIL2C": "OCR",
        "GNX-UTILISATEURPORTAIL2C": "OCR",
        "GNX-ADMINISTRATEURLECTURESEULE": "EDI",
    }

    def mapping_role(r):
        if pd.isna(r) or str(r).strip() == '':
            return "AUTRE"
        return role_to_canal.get(str(r).strip(), "AUTRE")

    df_out['canal'] = df_out['role'].apply(mapping_role)

    # Comptage exact après mapping
    correspondances = df_out[df_out['Correspondance'] == 'found']
    found_count = len(correspondances)
    canaux_groupes = correspondances.groupby('canal').size()

    df_out.to_excel(chemin_out, index=False, engine='openpyxl')
    print(f"\n✅ Correspondances ajoutées dans : {chemin_out}")
    print(f"ℹ️ {found_count} correspondance(s) trouvée(s) sur {len(df_export)} lignes analysées dans le fichier provenant de la filiale.")

    print("\nℹ️ Nombre de fournisseurs trouvés par canal :")
    print(canaux_groupes)

    total_par_canal = canaux_groupes.sum()
    difference = found_count - total_par_canal
    if difference != 0:
        incoherents = correspondances[correspondances['canal'] == 'AUTRE']
        print(f"\n❓ Incohérence détectée : {difference} correspondance(s) 'found' sans canal catégorisé.")
        print("🔎 Détail des lignes concernées :")
        print(incoherents[['identification', 'login', 'role']].to_string(index=False))

# ----- Script principal -----
if __name__ == "__main__":
    dossier_in = 'IN/GENERIX'
    fichiers = [f for f in os.listdir(dossier_in) if f.lower().endswith('.xlsx')]

    if not fichiers:
        print("❌ Aucun fichier .xlsx trouvé dans le dossier IN/GENERIX.")
    elif len(fichiers) > 1:
        print("❌ Plusieurs fichiers trouvés dans IN/GENERIX. Ce script attend un seul fichier.")
    else:
        chemin_fichier = os.path.join(dossier_in, fichiers[0])
        fichier_sortie = extraire_siren_depuis_excel(chemin_fichier)
        enrichir_referentiel(fichier_sortie)
        comparer_avec_export()
