import pandas as pd
import sys

def verifier_doublons_adherents(fichier_excel='ADHERENTS_JAN2026.xlsx', nom_colonne='MATRICULE'):
    """
    Vérifie s'il y a des doublons dans la colonne MATRICULE d'un fichier Excel ADHERENTS.
    
    Args:
        fichier_excel (str): Chemin vers le fichier Excel
        nom_colonne (str): Nom de la colonne à vérifier (par défaut 'MATRICULE')
    """
    try:
        # Lire le fichier Excel
        print(f"Lecture du fichier: {fichier_excel}")
        df = pd.read_excel(fichier_excel)
        
        # Vérifier si la colonne existe
        if nom_colonne not in df.columns:
            print(f"❌ Erreur: La colonne '{nom_colonne}' n'existe pas dans le fichier.")
            print(f"Colonnes disponibles: {', '.join(df.columns)}")
            return False
        
        # Afficher les informations générales
        print(f"\nInformations générales:")
        print(f"  - Total de lignes: {len(df)}")
        print(f"  - Valeurs non nulles dans '{nom_colonne}': {df[nom_colonne].notna().sum()}")
        print(f"  - Valeurs nulles dans '{nom_colonne}': {df[nom_colonne].isna().sum()}")
        
        # Vérifier les doublons (en excluant les valeurs nulles)
        df_non_null = df[df[nom_colonne].notna()]
        doublons = df_non_null[df_non_null[nom_colonne].duplicated(keep=False)]
        
        if doublons.empty:
            print(f"\n✅ Aucun doublon trouvé dans la colonne '{nom_colonne}'!")
            return True
        else:
            nb_doublons = len(doublons)
            valeurs_doublons = df_non_null[nom_colonne][df_non_null[nom_colonne].duplicated(keep=False)].unique()
            
            print(f"\n❌ {nb_doublons} doublons trouvés dans la colonne '{nom_colonne}'!")
            print(f"\nNombre de valeurs distinctes en doublon: {len(valeurs_doublons)}")
            print(f"\nValeurs en doublon:")
            
            # Trier les valeurs pour un affichage plus clair
            for valeur in sorted(valeurs_doublons):
                count = (df[nom_colonne] == valeur).sum()
                indices = df[df[nom_colonne] == valeur].index.tolist()
                lignes_excel = [idx + 2 for idx in indices]  # +2 car ligne 1 = en-tête, index commence à 0
                print(f"  - '{valeur}': {count} occurrences (lignes Excel: {lignes_excel})")
            
            print(f"\nDétail complet des lignes avec doublons:")
            # Afficher toutes les colonnes importantes pour les doublons
            colonnes_afficher = [col for col in df.columns if col in [nom_colonne, 'NOM', 'PRENOM', 'DATE_NAISSANCE', 'nom', 'prenom']]
            if colonnes_afficher:
                print(doublons[colonnes_afficher].to_string())
            else:
                print(doublons.to_string())
            
            # Préparer le DataFrame des doublons avec les colonnes supplémentaires
            doublons_export = doublons.copy()
            
            # Ajouter la colonne "Ligne_Excel" (numéro de ligne dans le fichier Excel)
            doublons_export['Ligne_Excel'] = doublons_export.index + 2  # +2 car ligne 1 = en-tête
            
            # Ajouter la colonne "Occurrence" (numéro d'occurrence pour chaque matricule)
            doublons_export['Occurrence'] = doublons_export.groupby(nom_colonne).cumcount() + 1
            
            # Réorganiser les colonnes pour mettre les nouvelles colonnes au début après le matricule
            cols = list(doublons_export.columns)
            # Trouver l'index de la colonne matricule
            if nom_colonne in cols:
                idx_matricule = cols.index(nom_colonne)
                # Retirer les colonnes ajoutées de leur position actuelle
                cols.remove('Ligne_Excel')
                cols.remove('Occurrence')
                # Les insérer juste après le matricule
                cols.insert(idx_matricule + 1, 'Occurrence')
                cols.insert(idx_matricule + 2, 'Ligne_Excel')
                doublons_export = doublons_export[cols]
            
            # Trier par matricule puis par occurrence pour un affichage plus clair
            doublons_export = doublons_export.sort_values(by=[nom_colonne, 'Occurrence'])
            
            # Sauvegarder les doublons dans un fichier séparé
            fichier_sortie = fichier_excel.replace('.xlsx', '_doublons.xlsx')
            doublons_export.to_excel(fichier_sortie, index=False)
            print(f"\n📄 Les doublons ont été exportés vers: {fichier_sortie}")
            print(f"   Colonnes ajoutées: 'Occurrence' (numéro de l'occurrence), 'Ligne_Excel' (numéro de ligne dans le fichier)")
            
            return False
            
    except FileNotFoundError:
        print(f"❌ Erreur: Le fichier '{fichier_excel}' n'existe pas.")
        return False
    except Exception as e:
        print(f"❌ Erreur lors de la lecture du fichier: {str(e)}")
        return False

if __name__ == "__main__":
    # Chemin par défaut
    fichier = "ADHERENTS_JAN2026.xlsx"
    colonne = "MATRICULE"
    
    # Permet de passer le fichier et la colonne en arguments
    if len(sys.argv) > 1:
        fichier = sys.argv[1]
    if len(sys.argv) > 2:
        colonne = sys.argv[2]
    
    print("=" * 70)
    print("VÉRIFICATION DES DOUBLONS - FICHIER ADHERENTS")
    print("=" * 70)
    
    verifier_doublons_adherents(fichier, colonne)
