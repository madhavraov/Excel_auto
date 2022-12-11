from tkinter import messagebox, filedialog as fd
import customtkinter
from open_items import OpenItems
from tb import Excel

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):
    WIDTH = 780
    HEIGHT = 520
    source_file = ''
    source_folder = ''
    destination_folder = ''

    def __init__(self):
        super().__init__()

        self.title("Excel Automation")
        self.icon = self.iconbitmap('Images/python.ico')
        self.geometry(f"{App.WIDTH}x{App.HEIGHT}")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)  # call .on_closing() when app gets closed

        # ================ Functions ==================== #
        def select_file():
            filetypes = (('excel files', '*.xlsx'), ('All files', '*.*'))
            filename = fd.askopenfilename(title='Open a file', initialdir='/', filetypes=filetypes)
            self.entry_1.insert(0, filename)
            App.source_file = filename

        def select_folder():
            folder_name = fd.askdirectory()
            self.entry_2.insert(0, folder_name)
            App.source_folder = folder_name

        def select_folder_1():
            folder_name = fd.askdirectory()
            self.entry_3.insert(0, folder_name)
            App.destination_folder = folder_name

        def run_program():
            if self.option_menu_1.get() == 'Trail Balance':
                try:
                    tb = Excel(file=App.source_file, source_folder=App.source_folder, target_folder=App.destination_folder)
                    tb.load_file()
                    tb.get_destination_folder()
                    tb.save_data()
                    messagebox.showinfo(title='Confirmation',
                                        message='The process has been completed successfully')
                except FileNotFoundError:
                    messagebox.showerror(title='Error',
                                         message='It seems the you selected incorrect file recheck the file again!')
                except PermissionError:
                    messagebox.showerror(title='Error',
                                         message='It seems the file is open. Please close the excel file and re-run the program')
            elif self.option_menu_1.get() == 'GL Balance':
                try:
                    gl = OpenItems(file_path=App.source_file, bsr_path=App.destination_folder,output_path=App.destination_folder)
                    gl.get_source_file()
                    gl.get_destination_folder()
                    gl.data_to_destination()
                    messagebox.showinfo(title='Confirmation',
                                        message='The process has been completed successfully')
                except FileNotFoundError:
                    messagebox.showerror(title='Error',
                                         message='There are no files in this path.Please check again.')
            else:
                messagebox.showerror(title='Error',
                                     message='Please select a task to run the program.')

        # ============ create two frames ============

        # configure grid layout (2x1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.frame_left = customtkinter.CTkFrame(master=self,
                                                 width=180,
                                                 corner_radius=0)
        self.frame_left.grid(row=0, column=0, sticky="nswe", padx=20, pady=20)

        self.frame_right = customtkinter.CTkFrame(master=self,
                                                  corner_radius=0)
        self.frame_right.grid(row=0, column=1, sticky="nswe", padx=20, pady=20)

        # ============ frame_left ============

        # configure grid layout (1x11)

        self.frame_left.grid_rowconfigure(0, minsize=10)  # empty row with minsize as spacing
        self.frame_left.grid_rowconfigure(5, weight=1)  # empty row as spacing
        self.frame_left.grid_rowconfigure(8, minsize=20)  # empty row with minsize as spacing
        self.frame_left.grid_rowconfigure(11, minsize=10)  # empty row with minsize as spacing

        self.label_mode = customtkinter.CTkLabel(master=self.frame_left, text="Task:")
        self.label_mode.grid(row=1, column=0, pady=10, padx=10, sticky="w")

        self.option_menu_1 = customtkinter.CTkOptionMenu(master=self.frame_left,
                                                         values=['Select', 'Trail Balance', 'GL Balance'])
        self.option_menu_1.grid(row=3, column=0, pady=10, padx=20, sticky="w")

        self.label_mode2 = customtkinter.CTkLabel(master=self.frame_left, text="Appearance Mode:")
        self.label_mode2.grid(row=9, column=0, pady=0, padx=20, sticky="w")

        self.option_menu_3 = customtkinter.CTkOptionMenu(master=self.frame_left,
                                                         values=["System", "Light", "Dark"],
                                                         command=self.change_appearance_mode)
        self.option_menu_3.grid(row=10, column=0, pady=10, padx=20, sticky="w")

        # ============ frame_right ============

        # configure grid layout (3x7)

        self.frame_right.rowconfigure((0, 1, 2, 3), weight=1)
        self.frame_right.rowconfigure(7, weight=10)
        self.frame_right.columnconfigure((0, 1), weight=1)
        self.frame_right.columnconfigure(2, weight=0)

        # ============ frame_right ============

        self.entry_1 = customtkinter.CTkEntry(master=self.frame_right,
                                              width=120,
                                              placeholder_text="Source File")
        self.entry_1.grid(row=0, column=0, columnspan=2, pady=10, padx=20, sticky="we")

        self.button_1 = customtkinter.CTkButton(master=self.frame_right,
                                                text="Browse File",
                                                border_width=2,  # <- custom border_width
                                                fg_color=None,  # <- no fg_color
                                                command=select_file)
        self.button_1.grid(row=0, column=2, columnspan=1, pady=20, padx=20, sticky="we")

        self.entry_2 = customtkinter.CTkEntry(master=self.frame_right,
                                              width=120,
                                              placeholder_text="Source Folder")
        self.entry_2.grid(row=1, column=0, columnspan=2, pady=10, padx=20, sticky="we")

        self.button_2 = customtkinter.CTkButton(master=self.frame_right,
                                                text="Browse Folder",
                                                border_width=2,  # <- custom border_width
                                                fg_color=None,  # <- no fg_color
                                                command=select_folder)
        self.button_2.grid(row=1, column=2, columnspan=1, pady=20, padx=20, sticky="we")

        self.entry_3 = customtkinter.CTkEntry(master=self.frame_right,
                                              width=120,
                                              placeholder_text="Destination Folder")
        self.entry_3.grid(row=2, column=0, columnspan=2, pady=10, padx=20, sticky="we")

        self.button_3 = customtkinter.CTkButton(master=self.frame_right,
                                                text="Browse Folder",
                                                border_width=2,  # <- custom border_width
                                                fg_color=None,  # <- no fg_color
                                                command=select_folder_1)
        self.button_3.grid(row=2, column=2, columnspan=1, pady=20, padx=20, sticky="we")

        self.run_button = customtkinter.CTkButton(master=self.frame_right,
                                                  text="Run",
                                                  border_width=2,
                                                  fg_color=None,
                                                  command=run_program)
        self.run_button.grid(row=3, column=0, columnspan=2, pady=20, padx=20, sticky="we")

        # ================ Functions ==================== #

    def change_appearance_mode(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def on_closing(self, event=0):
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
