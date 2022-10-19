import pandas as pd

def export_file(manager_df: pd.DataFrame, output_path: str):
    manager_df.to_csv(output_path, index=False)

    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    manager_df.to_excel(writer, sheet_name='Sheet1', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    cell_format_header = workbook.add_format({'bold': True, 'font_color': 'white', 'font_size': 14,
                                              'bg_color': '#3b5998', 'align': 'center', 'valign': 'vcenter'})

    for col_num, value in enumerate(manager_df.columns.values):
        worksheet.write(0, col_num, value, cell_format_header)

    writer.close()