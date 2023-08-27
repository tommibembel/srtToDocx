import os

import customtkinter

import functions

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # configure window
        self.title("srt to docx converter")
        self.geometry(f"{600}x{300}")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2, 3), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(1, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Options", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        openFileSwitchVar = customtkinter.StringVar(value="on")
        self.openAfterConvertLabel_switch = customtkinter.CTkSwitch(self.sidebar_frame, text="Open converted file",
                                                                    variable=openFileSwitchVar,
                                                                    onvalue="on", offvalue="off")
        self.openAfterConvertLabel_switch.grid(row=4, column=0, padx=20, pady=(20,0))

        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "100%", "125%", "150%", "200%"],
                                                               command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))

        # create input file entry and button
        self.inputFileEntry = customtkinter.CTkEntry(self, placeholder_text="Input File")
        self.inputFileEntry.grid(row=0, column=1, columnspan=1, padx=(20, 0), pady=(20, 20), sticky="new")

        self.inputFileButton = customtkinter.CTkButton(master=self, text="Open File", fg_color="transparent",
                                                       border_width=2, text_color=("gray10", "#DCE4EE"),
                                                       command=self.open_file_dialog_event)
        self.inputFileButton.grid(row=0, column=3, padx=(20, 20), pady=(20, 20), sticky="new")

        # create output file entry and button
        self.outputFileEntry = customtkinter.CTkEntry(self, placeholder_text="output File")
        self.outputFileEntry.grid(row=1, column=1, columnspan=1, padx=(20, 0), pady=(20, 20), sticky="new")

        self.outputFileButton = customtkinter.CTkButton(master=self, text="Output File", fg_color="transparent",
                                                       border_width=2, text_color=("gray10", "#DCE4EE"),
                                                       command=self.save_file_dialog_event)
        self.outputFileButton.grid(row=1, column=3, padx=(20, 20), pady=(20, 20), sticky="new")

        self.convertButton = customtkinter.CTkButton(master=self, text="CONVERTITTIE!",
                                                       command=self.convert_event)
        self.convertButton.grid(row=2, column= 1, columnspan=3, padx=(20, 20), pady=(20, 20), sticky="new")


        # set default values
        self.appearance_mode_optionemenu.set("Dark")
        self.scaling_optionemenu.set("100%")

    def open_file_dialog_event(self):
        file = customtkinter.filedialog.askopenfilename(initialdir=os.getcwd(),
                                                        title="Select file",
                                                        filetypes=(("srt files", "*.srt"), ))
        if file.endswith('.srt') and file is not None:
            self.inputFileEntry.delete(0, len(self.inputFileEntry.get()))
            self.inputFileEntry.insert(0, file)
            self.outputFileEntry.delete(0, len(self.outputFileEntry.get()))
            self.outputFileEntry.insert(0, file.replace(".srt", ".docx"))
        else:
            self.inputFileEntry.delete(0, len(self.inputFileEntry.get()))

    def save_file_dialog_event(self):
        file = customtkinter.filedialog.asksaveasfilename(initialdir=os.getcwd(),
                                                          initialfile=os.path.basename(self.inputFileEntry.get()).replace(".srt",".docx"),
                                                          filetypes=(("MS Word", "*.docx"),))
        if file.endswith('.docx') and file is not None:
            self.outputFileEntry.delete(0, len(self.outputFileEntry.get()))
            self.outputFileEntry.insert(0, file)
        else:
            self.outputFileEntry.delete(0, len(self.outputFileEntry.get()))

    def convert_event(self):
        functions.newconvert(self.inputFileEntry.get(), self.outputFileEntry.get())
        if self.openAfterConvertLabel_switch.get() == "on":
            os.system("open \"%s\"" %self.outputFileEntry.get())
        self.inputFileEntry.delete(0, len(self.inputFileEntry.get()))
        self.outputFileEntry.delete(0, len(self.outputFileEntry.get()))

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

    def sidebar_button_event(self):
        print("sidebar_button click")


if __name__ == "__main__":
    app = App()
    app.mainloop()