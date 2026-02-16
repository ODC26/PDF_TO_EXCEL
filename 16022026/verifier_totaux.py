import pandas as pd

# Lire les fichiers
df_orig = pd.read_excel('precomptes_mupol.xlsx')
df_traite = pd.read_excel('precomptes_mupol_traite_20260216_144708.xlsx', sheet_name='SANS_DOUBLONS')

print('='*80)
print('FICHIER ORIGINAL (precomptes_mupol.xlsx):')
print('='*80)
print(f'Lignes totales: {len(df_orig)}')
print(f'Total MT_MENSUALITE: {df_orig["MT_MENSUALITE"].sum():.0f}')
print(f'Total NOMBRE_AYANTS_DROIT (numérique): {pd.to_numeric(df_orig["NOMBRE_AYANTS_DROIT"], errors="coerce").sum():.0f}')

print('\n' + '='*80)
print('FICHIER SANS_DOUBLONS:')
print('='*80)
print(f'Lignes: {len(df_traite)}')
print(f'Total MT_MENSUALITE: {df_traite["MT_MENSUALITE"].sum():.0f}')
print(f'Total NOMBRE_AYANTS_DROIT (numérique): {pd.to_numeric(df_traite["NOMBRE_AYANTS_DROIT"], errors="coerce").sum():.0f}')

print('\n' + '='*80)
print('DIFFÉRENCE:')
print('='*80)
print(f'Lignes supprimées: {len(df_orig) - len(df_traite)}')
print(f'MT_MENSUALITE perdue: {df_orig["MT_MENSUALITE"].sum() - df_traite["MT_MENSUALITE"].sum():.0f}')
print(f'NOMBRE_AYANTS_DROIT perdu: {pd.to_numeric(df_orig["NOMBRE_AYANTS_DROIT"], errors="coerce").sum() - pd.to_numeric(df_traite["NOMBRE_AYANTS_DROIT"], errors="coerce").sum():.0f}')

print('\n' + '='*80)
print('DERNIÈRES LIGNES DU FICHIER SANS_DOUBLONS:')
print('='*80)
print(df_traite.tail(5)[['NO_MATRICULE', 'RECUP_NOM_AGENT(NO_MATRICULE)', 'MT_MENSUALITE', 'NOMBRE_AYANTS_DROIT']])
