from datetime import datetime
import pandas as pd

date = datetime.now()
file_date = str(date)[11:19].replace(':', '-')

df_to_convert = pd.read_excel('../wb/DI_DRIM_IdF.xlsx')
working_df = pd.DataFrame()


class WorkBook:
    def __init__(self, excel_input_file, data_frame):
        self.excel_input_file = pd.read_excel(excel_input_file)
        self.data_frame = data_frame

    def export_data(self):
        working_df['IdDemande'] = self.excel_input_file['IdDemande']
        working_df['Rubrique'] = self.excel_input_file['Rubrique']
        working_df['Date de création'] = self.excel_input_file['Date de création']
        working_df['Titre'] = self.excel_input_file['Titre']
        working_df['DRIM'] = self.excel_input_file['DRIM']
        working_df['UE'] = self.excel_input_file['UE']
        working_df['Ville'] = self.excel_input_file['Ville']
        working_df['Intervenant'] = self.excel_input_file['Intervenant']
        working_df['Statut de la demande'] = self.excel_input_file['Statut de la demande']
        working_df['Etape Exécution'] = self.excel_input_file['Etape Exécution']

        working_df.fillna(value='N/A', inplace=True)

        self.data_frame = working_df
        working_df.to_excel(f'../wb/export_{file_date}.xlsx')

    def get_last_ten(self):
        requests_by_date = self.excel_input_file.sort_values(by='Date de création', ascending=False)
        last_ten_requests = requests_by_date[0:9]

        last_ten_requests.to_excel(f'../wb/DI-par-date_{file_date}.xlsx')
        print(last_ten_requests[0:9])
