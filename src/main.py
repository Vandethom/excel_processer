import pandas as pd
from utils.excel_converter import WorkBook

#                   Init

new_dataframe = pd.DataFrame()

df = WorkBook(
    '../wb/analysis.xlsx',
    new_dataframe
)

excel_export = WorkBook(
    '../wb/DI_DRIM_IdF.xlsx',
    new_dataframe
)

#                   Methods

# excel_export.export_data()
# excel_export.get_count('Rue')

df.print_excel_weekly_ongoing_requests()
