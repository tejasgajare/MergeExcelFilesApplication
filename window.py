from tkinter import *
import utils


class MainWindow(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
        self.master.title("Merge Files")

        self.information_frame = InformationFrame(self)
        self.information_frame.grid(row=0, column=0, padx=(10, 10), pady=(10, 10))

        self.action_frame = ActionFrame(self)
        self.action_frame.grid(row=0, column=1, padx=(10, 10), pady=(10, 10))

    def clear_all(self):
        utils.clear_all(self.information_frame.selectedFilesListbox)
        self.action_frame.selected_files_count = 0
        self.action_frame.mergeFilesButton.config(state=DISABLED)
        self.action_frame.saveButton.config(state=DISABLED)
        self.action_frame.statusLabel.config(bg="red")
        self.action_frame.statusLabel.config(text="Please Select Files")

        self.information_frame.sort_checkbox.config(state=NORMAL)
        self.information_frame.filesCountLabel.config(text="0 Files Selected")


# Information Frame
class InformationFrame(LabelFrame):
    def __init__(self, master):
        LabelFrame.__init__(self, master, text="Information")

        self.selectedFilesListbox = Listbox(self, height=10, width=30)
        self.selectedFilesListbox.grid(row=0, column=0, padx=(10, 10), pady=(10, 0))
        scrollbar = Scrollbar(self, orient="vertical")
        scrollbar.config(command=self.selectedFilesListbox.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.selectedFilesListbox.config(yscrollcommand=scrollbar.set)

        self.filesCountLabel = Label(self, text="0 Files Selected")
        self.filesCountLabel.grid(row=1, column=0, padx=(10, 10), pady=(5, 0), sticky='w')

        # Modify Frame
        self.modify_frame = LabelFrame(self, text="Options")
        self.modify_frame.grid(row=2, column=0, padx=(10, 10), pady=(10, 10))
        self.clearAllButton = Button(self.modify_frame, text='Clear All', width=10, state=DISABLED,
                                     command=master.clear_all)
        self.clearAllButton.grid(row=0, column=1, padx=(10, 10), pady=(10, 10))

        self.sort_enable = BooleanVar()
        self.sort_checkbox = Checkbutton(self.modify_frame, text="Sort", variable=self.sort_enable,
                                         onvalue=True, offvalue=False, command=self.sort)
        self.sort_checkbox.grid(row=0, column=0, padx=(10, 30), pady=(5, 5))

    def sort(self):
        #  TODO: Functionality not needed
        self.selectedFilesListbox.delete(0, 'end')
        utils.sort_listbox(self.selectedFilesListbox)
        print("Checkbox toggled")


class ActionFrame(LabelFrame):
    def __init__(self, master):
        LabelFrame.__init__(self, master, text="Controls")
        self.master = master
        self.information_frame = master.information_frame
        self.selected_files_count = 0

        self.selectFilesButton = Button(self, text='Select Files', width=25, command=self.select_files)
        self.selectFilesButton.grid(row=1, column=0, padx=(25, 25), pady=(20, 15))

        self.mergeFilesButton = Button(self, text='Merge Files', width=25, state=DISABLED,
                                       command=self.merge_files)
        self.mergeFilesButton.grid(row=2, column=0, padx=(25, 25), pady=(15, 15))

        self.saveButton = Button(self, text='Save Files', width=25, state=DISABLED, command=self.save)
        self.saveButton.grid(row=3, column=0, padx=(25, 25), pady=(15, 25))

        self.message_frame = LabelFrame(self, text="Message", width=500)
        self.message_frame.grid(row=4, column=0, sticky="nsew", padx=(10, 10), pady=(10, 10))
        self.statusLabel = Label(self.message_frame,
                                 text="Please Select Files",
                                 bg="red",)
        self.statusLabel.grid(row=0, column=0, padx=(20, 20), pady=(20, 20))

    def select_files(self):
        # Allows to select multiple files and returns their absolute path.

        self.selected_files_count = utils.file_dialog_select_files(self.information_frame.selectedFilesListbox)
        self.information_frame.filesCountLabel["text"] = str(self.selected_files_count) + " Files Selected"

        if self.selected_files_count > 0:
            self.mergeFilesButton.config(state=NORMAL)
            self.information_frame.sort_checkbox.config(state=NORMAL)
            self.information_frame.clearAllButton.config(state=NORMAL)
            self.statusLabel.config(bg="green")
            self.statusLabel.config(text="You can now merge files")

    def merge_files(self):
        utils.merge_selected_files()

        self.mergeFilesButton.config(state=DISABLED)
        self.information_frame.sort_checkbox.config(state=DISABLED)
        self.saveButton.config(state=NORMAL)
        self.statusLabel.config(bg="yellow")
        self.statusLabel.config(text="Please Save File Before Exit")

    def save(self):
        save_file_name = utils.file_dialog_save_file()
        try:
            utils.save_merged_files(save_file_name)
            self.statusLabel.config(bg="cyan")
            self.statusLabel.config(text="Merged File Saved To Location: \n" + save_file_name)
        except PermissionError:
            self.statusLabel.config(bg="red")
            self.statusLabel.config(text="Please Close all Excel Files and Try Again")
