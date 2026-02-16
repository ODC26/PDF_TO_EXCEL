import pandas as pd
import os

# Chemin du fichier
fichier = "precomptes_mupol.xlsx"

# Vérifier si le fichier existe
if not os.path.exists(fichier):
    print(f"Le fichier {fichier} n'existe pas!")
    exit()

# Lire le fichier Excel
print(f"Lecture du fichier {fichier}...")
df = pd.read_excel(fichier)

print(f"\nNombre total de lignes : {len(df)}")
print(f"\nColonnes disponibles : {list(df.columns)}")

# Colonnes à vérifier
colonnes_toutes = ['NO_MATRICULE', 'RECUP_NOM_AGENT(NO_MATRICULE)', 'MT_MENSUALITE', 'NOMBRE_AYANTS_DROIT']
colonnes_partielles = ['NO_MATRICULE', 'MT_MENSUALITE', 'NOMBRE_AYANTS_DROIT']

# Vérifier que les colonnes existent
colonnes_manquantes = [col for col in colonnes_toutes if col not in df.columns]
if colonnes_manquantes:
    print(f"\nATTENTION : Colonnes manquantes : {colonnes_manquantes}")
    exit()

print("\n" + "="*80)
print("1. DOUBLONS COMPLETS (4 colonnes identiques)")
print("="*80)

# Trouver les doublons complets (les 4 colonnes)
doublons_complets = df[df.duplicated(subset=colonnes_toutes, keep=False)]

if len(doublons_complets) > 0:
    print(f"\n{len(doublons_complets)} lignes avec doublons complets trouvées :")
    print("\nDoublons complets (triés par NO_MATRICULE) :")
    doublons_complets_sorted = doublons_complets.sort_values(by=colonnes_toutes)
    print(doublons_complets_sorted[colonnes_toutes].to_string(index=True))
    
    # Sauvegarder dans un fichier Excel
    doublons_complets_sorted.to_excel("doublons_complets.xlsx", index=False)
    print(f"\n✓ Doublons complets sauvegardés dans 'doublons_complets.xlsx'")
else:
    print("\nAucun doublon complet trouvé.")

print("\n" + "="*80)
print("2. DOUBLONS PARTIELS (3 colonnes identiques, RECUP_NOM_AGENT différent)")
print("="*80)

# Trouver les doublons partiels (les 3 colonnes sans RECUP_NOM_AGENT)
doublons_partiels = df[df.duplicated(subset=colonnes_partielles, keep=False)]

if len(doublons_partiels) > 0:
    print(f"\n{len(doublons_partiels)} lignes avec doublons partiels trouvées :")
    
    # Exclure les doublons complets pour ne montrer que les partiels
    # (ceux où NO_MATRICULE, MT_MENSUALITE, NOMBRE_AYANTS_DROIT sont identiques
    # mais RECUP_NOM_AGENT est différent)
    doublons_partiels_only = doublons_partiels[~doublons_partiels.duplicated(subset=colonnes_toutes, keep=False)]
    
    if len(doublons_partiels_only) > 0:
        print(f"\nDoublons partiels uniquement (sans les doublons complets) : {len(doublons_partiels_only)} lignes")
        doublons_partiels_sorted = doublons_partiels_only.sort_values(by=colonnes_partielles)
        print(doublons_partiels_sorted[colonnes_toutes].to_string(index=True))
        
        # Sauvegarder dans un fichier Excel
        doublons_partiels_sorted.to_excel("doublons_partiels.xlsx", index=False)
        print(f"\n✓ Doublons partiels sauvegardés dans 'doublons_partiels.xlsx'")
    else:
        print("\nAucun doublon partiel trouvé (en dehors des doublons complets).")
    
    # Afficher tous les doublons partiels (y compris complets)
    print(f"\n\nTOUS les doublons partiels (y compris complets) : {len(doublons_partiels)} lignes")
    doublons_partiels_all_sorted = doublons_partiels.sort_values(by=colonnes_partielles)
    print(doublons_partiels_all_sorted[colonnes_toutes].to_string(index=True))
    
    # Sauvegarder dans un fichier Excel
    doublons_partiels_all_sorted.to_excel("tous_doublons_partiels.xlsx", index=False)
    print(f"\n✓ Tous les doublons partiels sauvegardés dans 'tous_doublons_partiels.xlsx'")
else:
    print("\nAucun doublon partiel trouvé.")

print("\n" + "="*80)
print("RÉSUMÉ")
print("="*80)
print(f"Total lignes : {len(df)}")
print(f"Doublons complets (4 colonnes) : {len(doublons_complets) if len(doublons_complets) > 0 else 0}")
print(f"Doublons partiels (3 colonnes) : {len(doublons_partiels) if len(doublons_partiels) > 0 else 0}")
print(f"Lignes uniques : {len(df) - len(doublons_partiels) if len(doublons_partiels) > 0 else len(df)}")
