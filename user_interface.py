import customtkinter
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import convertion_logic
global save_path
save_path=""

class ToplevelWindow(customtkinter.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("400x300")

        self.label = customtkinter.CTkLabel(self, text="ToplevelWindow")
        self.label.pack(padx=20, pady=20)


class App(customtkinter.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("500x400")
        self.button_1 = customtkinter.CTkButton(self, text="Open File", command=self.select_file)
        self.button_1.pack(side="top", padx=20, pady=20)
        self.button_2 = customtkinter.CTkButton(self, text="Convert Unicode to ANSI",command=self.convert_uni_to_ansi)
        self.button_2.pack(side="top",padx=10,pady=2)
        self.toplevel_window = None
        self.file_path = None

    def select_file(self):
        path = filedialog.askopenfilename(title="Select A DocX file to open", filetypes=[("DocX Files",".docx")])
        if path:
            self.file_path = path
        print(self.file_path)

    def convert_uni_to_ansi(self):
        print(self.file_path)
        save_path = filedialog.asksaveasfilename(title="Where to save the converted file?",defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
        try:
            convertion_logic.replace_and_highlight(doc_path=self.file_path,save_path=save_path)
            self.destroy()
        except:
            messagebox.showerror(title="Error", message="Can't edit file. out of permission ?")
        

    def open_toplevel(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = ToplevelWindow(self)  # create window if its None or destroyed
            self.toplevel_window.focus()
        ##else:
            ##self.toplevel_window.focus()  # if window exists focus it


app = App()
app.mainloop()