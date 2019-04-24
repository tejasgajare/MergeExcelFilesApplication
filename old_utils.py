import xlrd
import xlwt
from tkinter import filedialog

new_book = xlwt.Workbook()
new_sheet = new_book.add_sheet("Sheet 1")
selected_files = list()
selected_files_set = set()


def write_column_to_sheet(values, column_idx):
    # This function writes a single column to new file
    global new_sheet
    for i in range(len(values)):
        new_sheet.row(i + 1).write(column_idx, values[i])


def write_coloured_row(sheet, idx, value, total_columns):
    # This functions writes a colored row over the newly appended data
    # to help differentiate the data visually.
    if idx % 2 == 0:
        style = xlwt.Style.easyxf('pattern: pattern solid, fore_colour rose;'
                                  'font: colour black, bold True;')

    else:
        style = xlwt.Style.easyxf('pattern: pattern solid, fore_colour sky_blue;'
                                  'font: colour black, bold True;')

    col = idx * total_columns
    # HARD-CODED #######
    row = 0
    for i in range(total_columns):
        if i == total_columns // 2:
            sheet.write(row, col + i, value, style)
        else:
            sheet.write(row, col + i, '', style)


def append_data_into_file(old_sheet, append_from_column_idx):
    # This Functions iterates over columns of new file
    for col in range(old_sheet.ncols):
        column_values = old_sheet.col_values(col)
        write_column_to_sheet(column_values, append_from_column_idx)
        append_from_column_idx += 1


def merge_selected_files():
    global new_book
    global new_sheet

    new_book = xlwt.Workbook()
    new_sheet = new_book.add_sheet("Sheet 1")

    append_from_column_idx = 0
    for idx, PATH in enumerate(selected_files):
        old_book = xlrd.open_workbook(PATH)
        old_sheet = old_book.sheet_by_index(0)
        value = "WALK " + str(idx + 1)

        write_coloured_row(new_sheet, idx, value, old_sheet.ncols)
        append_data_into_file(old_sheet, append_from_column_idx)
        append_from_column_idx += old_sheet.ncols


# Action Frame
def sort_listbox(listbox):
    listbox.delete(0, 'end')
    selected_files.sort()
    print(selected_files)
    for idx, file in enumerate(selected_files):
        name = f"{idx + 1} | {file.split('/')[-1]}"
        listbox.insert("end", name)


def file_dialog_select_files(listbox):
    selected = filedialog.askopenfilenames(
        initialdir='/',
        initialfile='',
        filetypes=[("Excel", "*.xlsx"), ("Excel", "*.xls")]
    )
    print(f"Selected Files : {selected}")
    for file in selected:
        if file in selected_files_set:
            continue
        else:
            selected_files.append(file)
            selected_files_set.add(file)
            name = f"{len(selected_files)} | {file.split('/')[-1]}"
            listbox.insert("end", name)

    return len(selected_files)


def file_dialog_save_file():
    rep = filedialog.asksaveasfilename(
        confirmoverwrite=True,
        defaultextension=".xls",
        initialdir='/',
        initialfile='merged.xls',
        filetypes=[("Excel", "*.x lsx"), ("Excel", "*.xls")]
    )
    return rep


def clear_all(listbox):
    selected_files.clear()
    selected_files_set.clear()
    listbox.delete(0, 'end')


def save_merged_files(path):
    new_book.save(path)
    print(f"file saved at location {path}")
