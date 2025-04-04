import customtkinter

class testbench(customtkinter.CTk):
    def __init__(self):
        global file
        super().__init__()
        self.title("WordBuddy")
        self.geometry("700x630")
        self.resizable(False, False)

        self.file_path = customtkinter.CTkEntry(self, placeholder_text="File Path")
        self.file_path.pack(pady=20, padx=20)
        self.button = customtkinter.CTkButton(self, text="SHOW", command=testbench.button_press)
        self.button.pack(pady=20, padx=20)

        file = self.file_path

    def button_press():
        print (file)
    

app = testbench()
app.mainloop()