from datetime import datetime
import pandas as pd

date = datetime.now()
file_date = str(date)[11:19].replace(':', '-')

df_to_convert = pd.read_excel('../wb/DI_DRIM_IdF.xlsx')
working_df = pd.DataFrame()


class WorkBook:
    def __init__(self, excel_input_file, data_frame):
        self.data_frame = pd.read_excel(excel_input_file)

    def export_data(self):

        working_df['IdDemande'] = df_to_convert['IdDemande']
        working_df['Rubrique'] = df_to_convert['Rubrique']
        working_df['Date de création'] = df_to_convert['Date de création']
        working_df['Titre'] = df_to_convert['Titre']
        working_df['DRIM'] = df_to_convert['DRIM']
        working_df['UE'] = df_to_convert['UE']
        working_df['Ville'] = df_to_convert['Ville']
        working_df['Intervenant'] = df_to_convert['Intervenant']
        working_df['Statut de la demande'] = df_to_convert['Statut de la demande']
        working_df['Etape Exécution'] = df_to_convert['Etape Exécution']

        working_df.fillna(value='N/A', inplace=True)

        working_df.to_excel(f'../wb/export_{file_date}.xlsx')
