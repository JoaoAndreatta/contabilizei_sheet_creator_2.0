import tkinter
import customtkinter
from PIL import Image
from emprestimo import Emprestimo
from depreciacao import Depreciacao

class App(customtkinter.CTk):
    
    def __init__(self):
        super().__init__()
        
        # configure window
        self.title("Contabilizei Planílhas")
        self.geometry(f"{895}x{580}")
        self.resizable(width=False, height=False)
        self.iconbitmap('media\icon.ico')

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=6, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo = customtkinter.CTkImage(light_image=Image.open("media\logo_contabilizei.png"),
                                  dark_image=Image.open("media\logo_contabilizei.png"),
                                  size=(153, 23))
        self.logo_button = customtkinter.CTkButton(self.sidebar_frame, image=self.logo, text =None, 
                                                    fg_color='transparent', hover=False, command= lambda: [self.clear_window(), self.recreate_window()])
        self.logo_button.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, text='Depreciação', command=lambda: [self.clear_window(), Depreciacao.depreciacao_button(self)])
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, text='Empréstimo', command= lambda:[self.clear_window(),Emprestimo.loan_button(self)])
        self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)

        # main frame
        self.main_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color='transparent')
        self.main_frame.grid(row=0, column=1, rowspan=3 , columnspan=4, padx=(0, 0), pady=(0, 0), sticky="swn")

        # how to use widget
        self.banner = customtkinter.CTkImage(light_image=Image.open(r"media\art.png"),
                                  dark_image=Image.open(r"media\art.png"), size=(645,500))
        self.welcome_widget = customtkinter.CTkButton(self.main_frame, text=None, image=self.banner,
                                                    fg_color='transparent', hover=False)
        self.welcome_widget.grid(row=0,column=0, padx=(30,20), pady = 20, sticky='nswe')

    def recreate_window(self):
        # how to use widget
        self.welcome_widget = customtkinter.CTkButton(self.main_frame, text=None, image=self.banner,
                                                    fg_color='transparent', hover=False)
        self.welcome_widget.grid(row=0,column=0, padx=(30,20), pady = 20, sticky='nswe')

    # clear window function
    def clear_window(self):
        for widgets in self.main_frame.winfo_children():
            widgets.destroy()

if __name__ == "__main__":

    app = App()
    app.mainloop()

