import pandas as pd
from utils.excel_converter import WorkBook

#                   Init

new_dataframe = pd.DataFrame()

excel_export = WorkBook(
    '../wb/DI_DRIM_IdF.xlsx',
    new_dataframe
)

df_to_convert = pd.read_excel('../wb/DI_DRIM_IdF.xlsx')

#                   Methods

# excel_export.export_data()
excel_export.get_last_ten()
