import pandas as pd
from utils.excel_converter import WorkBook

#                   Init

new_dataframe = pd.DataFrame()

excel_export = WorkBook(
    '../wb/DI_DRIM_IdF.xlsx',
    new_dataframe
)

#                   Methods

# excel_export.export_data()
excel_export.filter_by_multiple(['Rue', 'UE', 'Ville'])
