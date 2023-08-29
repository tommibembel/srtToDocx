import os
import sys
import customtkinter
import functions
import configparser
import webbrowser
from PIL import Image



def check_configfile_exists(cf):
    if os.path.isfile(cf):
        config = configparser.RawConfigParser()
        config.read(cf)
        return config
    else:
        config = create_cofig_file(cf)
        return config


def create_cofig_file(cf):
    config = configparser.RawConfigParser()
    config['Main'] = {"UIScale": "100%", "Appearance": "System", "Color": "blue"}
    with open(cf, "w") as configfile:
        config.write(configfile)
    return config


def write_config(cf, conf):
    with open(cf, "w") as configfile:
        conf.write(configfile)


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


cf = "srt_to_docx_config.ini"
config = check_configfile_exists(cf)
customtkinter.set_appearance_mode(config["Main"]["Appearance"])  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme(config["Main"]["Color"])  # Themes: "blue" (standard), "green", "dark-blue"
customtkinter.set_widget_scaling(int(config["Main"]["UIScale"].replace("%", "")) / 100)
customtkinter.set_window_scaling(int(config["Main"]["UIScale"].replace("%", "")) / 100)

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # configure window
        self.title("srt to docx converter")
        self.geometry(f"{800}x{500}")
        self.minsize(800,500)
        self.resizable(True, True)
        self.iconbitmap(resource_path("assets/icon.ico"))

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

        self.gitlogo = customtkinter.CTkImage(light_image=Image.open(resource_path("assets/github-mark.png")),
                                             dark_image=Image.open(resource_path("assets/github-mark-white.png")),
                                             size=(30,30))
        self.gitlabel = customtkinter.CTkLabel(self.sidebar_frame, image=self.gitlogo, text="")
        self.gitlabel.grid(row=9, column=0, padx=20, pady=(20,20), sticky="w")
        self.gitlabel.bind(sequence="<Button-1>", command=self.open_github)

        openFileSwitchVar = customtkinter.StringVar(value="on")
        self.openAfterConvertLabel_switch = customtkinter.CTkSwitch(self.sidebar_frame, text="Open converted file",
                                                                    variable=openFileSwitchVar,
                                                                    onvalue="on", offvalue="off")
        self.openAfterConvertLabel_switch.grid(row=2, column=0, padx=(20, 20), pady=(20,20))

        self.color_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Color Mode:", anchor="w")
        self.color_mode_label.grid(row=3, column=0, padx=(20, 20), pady=(10, 0))
        self.color_mode_optionmenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["blue", "green", "dark-blue"],
                                                                 command=self.change_color_mode_event)
        self.color_mode_optionmenu.grid(row=4, column=0, padx=(20, 20), pady=(10, 20))
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=(20, 20), pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "100%", "125%", "150%", "200%"],
                                                               command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))

        # create input file entry and button
        self.inputFileEntry = customtkinter.CTkEntry(self, placeholder_text="input file path")
        self.inputFileEntry.grid(row=0, column=1, columnspan=1, padx=(20, 0), pady=(20, 20), sticky="new")

        self.inputFileButton = customtkinter.CTkButton(master=self, text="select input file", fg_color="transparent",
                                                       border_width=2, text_color=("gray10", "#DCE4EE"),
                                                       command=self.open_file_dialog_event)
        self.inputFileButton.grid(row=0, column=3, padx=(20, 20), pady=(20, 20), sticky="new")

        # create output file entry and button
        self.outputFileEntry = customtkinter.CTkEntry(self, placeholder_text="output file path")
        self.outputFileEntry.grid(row=1, column=1, columnspan=1, padx=(20, 0), pady=(20, 20), sticky="new")

        self.outputFileButton = customtkinter.CTkButton(master=self, text="select output file", fg_color="transparent",
                                                       border_width=2, text_color=("gray10", "#DCE4EE"),
                                                       command=self.save_file_dialog_event)
        self.outputFileButton.grid(row=1, column=3, padx=(20, 20), pady=(20, 20), sticky="new")

        self.convertButton = customtkinter.CTkButton(master=self, text="Convert",
                                                       command=self.convert_event)
        self.convertButton.grid(row=2, column= 1, columnspan=3, padx=(20, 20), pady=(20, 20), sticky="new")

        self.statusLabel = customtkinter.CTkLabel(self, anchor="center", text="", bg_color="transparent")
        self.statusLabel.grid(row=3, column=1, columnspan=3, padx=(20,20), pady=(20,20), sticky="sew")

        # set default values
        self.appearance_mode_optionemenu.set(config["Main"]["Appearance"])
        self.scaling_optionemenu.set(config["Main"]["UIScale"])
        self.color_mode_optionmenu.set(config["Main"]["Color"])

        self.popupWindow = None


    def open_github(self, *args):
        webbrowser.open_new_tab("https://github.com/tommibembel/srtToDocx")

    def open_file_dialog_event(self):
        self.set_status_label(2)
        file = customtkinter.filedialog.askopenfilename(initialdir=os.path.expanduser('~/Documents'),
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
        self.set_status_label(2)
        file = customtkinter.filedialog.asksaveasfilename(initialdir=os.path.expanduser('~/Documents'),
                                                          initialfile=os.path.basename(self.inputFileEntry.get()).replace(".srt",".docx"),
                                                          filetypes=(("MS Word", "*.docx"),))
        if file.endswith('.docx') and file is not None:
            self.outputFileEntry.delete(0, len(self.outputFileEntry.get()))
            self.outputFileEntry.insert(0, file)
        else:
            self.outputFileEntry.delete(0, len(self.outputFileEntry.get()))

    def convert_event(self):
        self.set_status_label(2)
        if self.inputFileEntry.get() == "" or self.outputFileEntry.get() == "":
            self.set_status_label(1, "Please set input and output file")
        else:
            functions.convert(self.inputFileEntry.get(), self.outputFileEntry.get())
            if self.openAfterConvertLabel_switch.get() == "on":
                if sys.platform == "darwin":
                    os.system("open \"%s\"" %self.outputFileEntry.get())
                #elif sys.platform == "linux":      #Don't know how linux opens a docx file ;) so this should be adapted
                #    os.startfile("%s" % self.outputFileEntry.get())
                else:
                    os.startfile("%s" % self.outputFileEntry.get())

            self.inputFileEntry.delete(0, len(self.inputFileEntry.get()))
            self.outputFileEntry.delete(0, len(self.outputFileEntry.get()))
            self.set_status_label(0)

    def set_status_label(self, status, message="Conversion successfully"):
        if status == 0:
            self.statusLabel.configure(text=message, bg_color="green")
        elif status == 1:
            self.statusLabel.configure(text=message, bg_color="red")
        else:
            self.statusLabel.configure(text="", bg_color="transparent")

    def change_appearance_mode_event(self, new_appearance_mode: str):
        self.set_status_label(2)
        customtkinter.set_appearance_mode(new_appearance_mode)
        config["Main"]["Appearance"] = new_appearance_mode
        write_config(cf, config)

    def change_color_mode_event(self, new_color_mode: str):
        self.set_status_label(2)
        config["Main"]["Color"] = new_color_mode
        write_config(cf, config)
        self.set_status_label(0, "New color set, program restart required!")

    def change_scaling_event(self, new_scaling: str):
        self.set_status_label(2)
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)
        customtkinter.set_window_scaling(new_scaling_float)
        config["Main"]["UIScale"] = new_scaling
        write_config(cf, config)


if __name__ == "__main__":
    app = App()
    app.mainloop()
