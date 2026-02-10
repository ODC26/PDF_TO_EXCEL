import pandas as pd
import sys

def supprimer_doublons_adherents(fichier_excel='ADHERENTS_JAN2026.xlsx', nom_colonne='MATRICULE'):
    """
    Supprime les doublons exacts dans un fichier Excel ADHERENTS.
    Un doublon est considéré comme exact si le MATRICULE et TOUTES les autres colonnes sont identiques.
    Crée une nouvelle colonne avec les numéros ORDRE des lignes supprimées.
    
    Args:
        fichier_excel (str): Chemin vers le fichier Excel
        nom_colonne (str): Nom de la colonne matricule à vérifier (par défaut 'MATRICULE')
    """
    try:
        # Lire le fichier Excel
        print(f"Lecture du fichier: {fichier_excel}")
        df = pd.read_excel(fichier_excel)
        
        # Convertir les colonnes numériques en nombres
        colonnes_numeriques = ['AYANTS_DROIT', 'MENSUALITE', 'MENSUALITEE', 'AYANT_DROIT']
        for col in colonnes_numeriques:
            if col in df.columns:
                print(f"  - Conversion de '{col}' en nombre...")
                # Remplacer les espaces et autres caractères non numériques
                df[col] = df[col].astype(str).str.replace(' ', '').str.replace(',', '.')
                # Convertir en numérique (les erreurs deviennent NaN)
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Vérifier si la colonne MATRICULE existe
        if nom_colonne not in df.columns:
            print(f"❌ Erreur: La colonne '{nom_colonne}' n'existe pas dans le fichier.")
            print(f"Colonnes disponibles: {', '.join(df.columns)}")
            return False
        
        # Vérifier si la colonne ORDRE existe, sinon la créer
        if 'ORDRE' not in df.columns:
            print(f"⚠️  La colonne 'ORDRE' n'existe pas. Création automatique...")
            df.insert(0, 'ORDRE', range(1, len(df) + 1))
        
        # Afficher les informations générales
        print(f"\nInformations initiales:")
        print(f"  - Total de lignes: {len(df)}")
        print(f"  - Valeurs non nulles dans '{nom_colonne}': {df[nom_colonne].notna().sum()}")
        print(f"  - Valeurs nulles dans '{nom_colonne}': {df[nom_colonne].isna().sum()}")
        
        # Ajouter la colonne pour stocker les ORDRE supprimés
        df['DOUBLONS_SUPPRIMES'] = ''
        
        # Identifier les doublons basés sur le MATRICULE uniquement
        df_non_null = df[df[nom_colonne].notna()].copy()
        matricules_doublons = df_non_null[df_non_null[nom_colonne].duplicated(keep=False)][nom_colonne].unique()
        
        print(f"\n🔍 Analyse des doublons potentiels...")
        print(f"  - Nombre de matricules ayant plusieurs occurrences: {len(matricules_doublons)}")
        
        # Liste pour stocker les indices des lignes à supprimer
        indices_a_supprimer = []
        
        # Statistiques
        nb_vrais_doublons = 0
        nb_faux_doublons = 0
        details_doublons = []
        
        # Pour chaque matricule en doublon
        for matricule in matricules_doublons:
            # Obtenir toutes les lignes avec ce matricule
            lignes_matricule = df[df[nom_colonne] == matricule].copy()
            
            if len(lignes_matricule) < 2:
                continue
            
            # Comparer toutes les colonnes SAUF 'DOUBLONS_SUPPRIMES' et 'ORDRE'
            colonnes_a_comparer = [col for col in df.columns if col not in ['DOUBLONS_SUPPRIMES', 'ORDRE']]
            
            # Grouper les lignes identiques
            # On utilise toutes les colonnes pour détecter si c'est un vrai doublon
            lignes_matricule['groupe_hash'] = lignes_matricule[colonnes_a_comparer].apply(
                lambda row: hash(tuple(str(x) for x in row)), axis=1
            )
            
            # Pour chaque groupe de lignes identiques
            groupes = lignes_matricule.groupby('groupe_hash')
            
            for groupe_hash, groupe_lignes in groupes:
                if len(groupe_lignes) > 1:
                    # C'est un vrai doublon (toutes les colonnes sont identiques)
                    nb_vrais_doublons += len(groupe_lignes) - 1
                    
                    # Garder la première ligne, supprimer les autres
                    indices_groupe = groupe_lignes.index.tolist()
                    premiere_ligne_idx = indices_groupe[0]
                    lignes_a_supprimer = indices_groupe[1:]
                    
                    # Récupérer les numéros ORDRE des lignes à supprimer
                    ordres_supprimes = [str(int(df.loc[idx, 'ORDRE'])) for idx in lignes_a_supprimer]
                    ordres_str = ', '.join(ordres_supprimes)
                    
                    # Mettre à jour la colonne DOUBLONS_SUPPRIMES de la ligne conservée
                    df.at[premiere_ligne_idx, 'DOUBLONS_SUPPRIMES'] = ordres_str
                    
                    # Ajouter les indices à la liste de suppression
                    indices_a_supprimer.extend(lignes_a_supprimer)
                    
                    # Détail pour l'affichage
                    details_doublons.append({
                        'matricule': matricule,
                        'nb_occurrences': len(groupe_lignes),
                        'ligne_conservee': int(df.loc[premiere_ligne_idx, 'ORDRE']),
                        'lignes_supprimees': ordres_supprimes
                    })
                else:
                    # Même matricule mais données différentes (faux doublon)
                    nb_faux_doublons += 1
        
        # Supprimer les lignes en doublon
        if indices_a_supprimer:
            print(f"\n📊 Résultats de l'analyse:")
            print(f"  - Vrais doublons trouvés (données identiques): {nb_vrais_doublons}")
            print(f"  - Faux doublons (même matricule, données différentes): {nb_faux_doublons}")
            print(f"  - Lignes à supprimer: {len(indices_a_supprimer)}")
            
            print(f"\n📋 Détail des doublons supprimés:")
            for detail in details_doublons:
                print(f"  - Matricule '{detail['matricule']}':")
                print(f"    • {detail['nb_occurrences']} occurrences trouvées")
                print(f"    • Ligne conservée: ORDRE {detail['ligne_conservee']}")
                print(f"    • Lignes supprimées: ORDRE {', '.join(detail['lignes_supprimees'])}")
            
            # Supprimer les lignes
            df_nettoye = df.drop(indices_a_supprimer)
            
            print(f"\n✅ Nettoyage effectué:")
            print(f"  - Lignes avant: {len(df)}")
            print(f"  - Lignes après: {len(df_nettoye)}")
            print(f"  - Lignes supprimées: {len(indices_a_supprimer)}")
            
            # Sauvegarder le fichier nettoyé
            fichier_sortie = fichier_excel.replace('.xlsx', '_nettoye.xlsx')
            df_nettoye.to_excel(fichier_sortie, index=False)
            print(f"\n💾 Fichier nettoyé sauvegardé: {fichier_sortie}")
            print(f"   La colonne 'DOUBLONS_SUPPRIMES' contient les numéros ORDRE des lignes supprimées")
            
            # Créer un rapport détaillé des suppressions
            if details_doublons:
                rapport_df = pd.DataFrame(details_doublons)
                rapport_df['lignes_supprimees'] = rapport_df['lignes_supprimees'].apply(lambda x: ', '.join(x))
                fichier_rapport = fichier_excel.replace('.xlsx', '_rapport_suppressions.xlsx')
                rapport_df.to_excel(fichier_rapport, index=False)
                print(f"   Rapport détaillé: {fichier_rapport}")
            
            return True
        else:
            print(f"\n✅ Aucun vrai doublon trouvé!")
            print(f"   (Vrais doublons = même matricule ET toutes les colonnes identiques)")
            if nb_faux_doublons > 0:
                print(f"\n⚠️  {nb_faux_doublons} matricules en doublon avec des données différentes ont été trouvés.")
                print(f"   Ces lignes ne sont PAS supprimées car les données diffèrent.")
            
            # Sauvegarder quand même avec la colonne DOUBLONS_SUPPRIMES (vide)
            fichier_sortie = fichier_excel.replace('.xlsx', '_nettoye.xlsx')
            df.to_excel(fichier_sortie, index=False)
            print(f"\n💾 Fichier sauvegardé avec colonne 'DOUBLONS_SUPPRIMES': {fichier_sortie}")
            
            return True
            
    except FileNotFoundError:
        print(f"❌ Erreur: Le fichier '{fichier_excel}' n'existe pas.")
        return False
    except Exception as e:
        print(f"❌ Erreur lors du traitement: {str(e)}")
        import traceback
        traceback.print_exc()
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
    print("SUPPRESSION DES DOUBLONS - FICHIER ADHERENTS")
    print("=" * 70)
    print("Ce script supprime uniquement les vrais doublons:")
    print("- Même MATRICULE")
    print("- ET toutes les autres colonnes identiques")
    print("=" * 70)
    
    supprimer_doublons_adherents(fichier, colonne)
