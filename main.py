import pandas as pd
import os
import support as sp

class SegregateFiles:
    managers_lst = []

    def process_masterfile(filepath: str, output_path: str):
        """:param filepath: Path of input file
        :param output_path: output directory"""

        file_name = os.path.basename(filepath)
        print(file_name)
        master_df = pd.read_excel(filepath, index_col=False)

        managers_lst = master_df.Manager.unique().tolist()

        for manager in managers_lst:
            temp_df = master_df[master_df['Manager'] == manager]
            op_file_name = manager + "_" + file_name
            sp.export_file(temp_df, output_path + "/" + op_file_name)




