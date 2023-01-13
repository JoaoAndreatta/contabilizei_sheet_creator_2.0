import tkinter
import customtkinter
from PIL import Image

class App(customtkinter.CTk):
    
    def __init__(self):
        super().__init__()
        
        # configure window
        self.title("Contabilizei Planílhas")
        self.geometry(f"{1100}x{580}")
        self.iconbitmap('media\icon.ico')

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo = customtkinter.CTkImage(light_image=Image.open("media\logo_contabilizei.png"),
                                  dark_image=Image.open("media\logo_contabilizei.png"),
                                  size=(153, 23))
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, image=self.logo, text = None)
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, text='Depreciação')
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, text='Empréstimo')
        self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)

        # create sheet preview frame
        self.sheet_frame = customtkinter.CTkFrame(self, corner_radius=0)
        self.sheet_frame.grid(row=2, column=1, columnspan=2, padx=(20, 20), pady=(20, 20), sticky="nsew")

if __name__ == "__main__":

    app = App()
    app.mainloop()

