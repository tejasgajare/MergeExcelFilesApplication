from tkinter import filedialog
import pandas as pd
import numpy as np


class Merger:
    def __init__(self):
        self.measurer = np.vectorize(len)
        self.selected_files = list()
        self.selected_files_set = set()
        self.data = None
        self.writer = None

    @staticmethod
    def combine_data(paths, include_params=False):
        if not include_params:
            data_dict = {}
            for idx, path in enumerate(paths):
                temp_df = pd.read_excel(path)
                index = temp_df.columns.values[0]
                columns = temp_df.columns.values
                data_dict['WALK_' + str(idx + 1)] = pd.DataFrame(columns=columns, data=temp_df).set_index(index)

            return pd.concat(data_dict, axis=1)
        else:
            data_frames = []
            for path in paths:
                temp_df = pd.read_excel(path)
                data_frames.append(temp_df)
            return pd.concat(data_frames, axis=1)

    def write_to_file(self, path, final_dataframe, formatting=False):
        writer = pd.ExcelWriter(path, engine='xlsxwriter')
        final_dataframe.to_excel(writer, sheet_name='Sheet1')

        if formatting:
            worksheet = writer.sheets['Sheet1']
            max_lengths = [self.measurer(final_dataframe.index.values.astype(str)).max(axis=0)]
            max_lengths.extend(self.measurer(final_dataframe.values.astype(str)).max(axis=0))
            for idx, column_len in enumerate(max_lengths):
                worksheet.set_column(idx, idx, column_len + 2)

        return writer

    def merge_selected_files(self, include_params):
        print(f"include_params: {include_params}")
        self.data = self.combine_data(self.selected_files, include_params)

    def file_dialog_select_files(self, listbox):
        selected = filedialog.askopenfilenames(
            initialdir='/',
            initialfile='',
            filetypes=[("Excel", "*.xlsx"), ("Excel", "*.xls")]
        )
        print(f"Selected Files : {selected}")
        for file in selected:
            if file in self.selected_files_set:
                continue
            else:
                self.selected_files.append(file)
                self.selected_files_set.add(file)
                name = f"{len(self.selected_files)} | {file.split('/')[-1]}"
                listbox.insert("end", name)

        return len(self.selected_files)

    @staticmethod
    def file_dialog_save_file():
        rep = filedialog.asksaveasfilename(
            confirmoverwrite=True,
            defaultextension=".xlsx",
            initialdir='/',
            initialfile='merged.xlsx',
            filetypes=[("Excel", "*.xlsx"), ("Excel", "*.xls")]
        )
        return rep

    def clear_all(self, listbox):
        self.selected_files.clear()
        self.selected_files_set.clear()
        listbox.delete(0, 'end')

    def save_merged_files(self, path, formatting):
        print(f"formatting: {formatting}")
        self.writer = self.write_to_file(path, self.data, formatting)
        self.writer.save()
        print(f"file saved at location {path}")
