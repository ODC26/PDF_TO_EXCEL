import pandas as pd
import sys

def verifier_doublons(fichier_excel, nom_colonne='col'):
    """
    Vérifie s'il y a des doublons dans une colonne spécifique d'un fichier Excel.
    
    Args:
        fichier_excel (str): Chemin vers le fichier Excel
        nom_colonne (str): Nom de la colonne à vérifier (par défaut 'col')
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
        
        # Vérifier les doublons
        doublons = df[df[nom_colonne].duplicated(keep=False)]
        
        if doublons.empty:
            print(f"✅ Aucun doublon trouvé dans la colonne '{nom_colonne}'!")
            print(f"Total de lignes: {len(df)}")
            return True
        else:
            nb_doublons = len(doublons)
            valeurs_doublons = df[nom_colonne][df[nom_colonne].duplicated(keep=False)].unique()
            
            print(f"❌ {nb_doublons} doublons trouvés dans la colonne '{nom_colonne}'!")
            print(f"\nValeurs en doublon:")
            for valeur in valeurs_doublons:
                count = (df[nom_colonne] == valeur).sum()
                print(f"  - '{valeur}': {count} occurrences")
            
            print(f"\nDétail des lignes avec doublons:")
            print(doublons[[nom_colonne]].to_string())
            
            # Sauvegarder les doublons dans un fichier séparé
            fichier_sortie = fichier_excel.replace('.xlsx', '_doublons.xlsx')
            doublons.to_excel(fichier_sortie, index=False)
            print(f"\n📄 Les doublons ont été exportés vers: {fichier_sortie}")
            
            return False
            
    except FileNotFoundError:
        print(f"❌ Erreur: Le fichier '{fichier_excel}' n'existe pas.")
        return False
    except Exception as e:
        print(f"❌ Erreur lors de la lecture du fichier: {str(e)}")
        return False

if __name__ == "__main__":
    # Chemin par défaut
    fichier = "resultat_jan_2026.xlsx"
    colonne = "col"
    
    # Permet de passer le fichier et la colonne en arguments
    if len(sys.argv) > 1:
        fichier = sys.argv[1]
    if len(sys.argv) > 2:
        colonne = sys.argv[2]
    
    print("=" * 60)
    print("VÉRIFICATION DES DOUBLONS")
    print("=" * 60)
    
    verifier_doublons(fichier, colonne)
