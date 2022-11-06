from datetime import datetime
import pandas as pd

df_to_convert = pd.read_excel('../wb/DI_DRIM_IdF.xlsx')
working_df = pd.DataFrame()


class WorkBook:
    def __init__(self, excel_input_file, data_frame):
        self.excel_input_file = excel_input_file
        self.data_frame = data_frame

    def read_file(self, sheet):
        excel_file = pd.read_excel(self.excel_input_file, sheet_name=sheet)

        return excel_file

    def export_data(self):
        date = datetime.now()
        file_date = str(date)[11:19].replace(':', '-')

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

    def filter_by(self, criteria):
        filtered_dataframe = self.excel_input_file[criteria]

        print(filtered_dataframe)

    def filter_by_multiple(self, criterias):
        new_dataframe = pd.DataFrame()

        for criteria in criterias:
            new_dataframe[criteria] = self.excel_input_file[criteria]

        print(new_dataframe)

    def get_last_ten(self):
        requests_by_date = self.excel_input_file.sort_values(by='Date de création', ascending=False)
        last_ten_requests = requests_by_date[0:9]

        last_ten_requests.to_excel(f'../wb/DI-par-date_{file_date}.xlsx')
        print(last_ten_requests[0:9])

    def get_count(self, column_to_count):
        count = len(self.excel_input_file[column_to_count])

        print(count)

    def get_weekly_ongoing_requests(self):
        transpo = self.read_file('Transposition')

        columns = transpo['Année Créa']

        print(columns)
