import pandas as pd
from datetime import datetime as dt
from tkinter import *
from tkinter import ttk
import tkinter.font as font
import os

class NotificariMasini():

    def __init__(self):
        self.acte_mapping = {'DATA EXP RCA': 'asigurarea',
                        'DATA EXP ROV': 'rovinieta',
                        'DATA EXP ITP': 'ITP-ul'}

    def open_excel(self):
        try:
            os.system(os.environ['NotificariMasini'] + '\ITP.xlsx')
        except:
            print("Cannot open file")

    def initiate_pop_up_screen(self, text):
        pop_up = Tk()
        myfont = font.Font(size=15)
        pop_up.title('Notificare masini')
        pop_up.geometry("800x500")

        # Create Frame
        main_frame = Frame(pop_up)
        main_frame.pack(fill=BOTH, expand=1)

        # Create Canvar
        my_canvas = Canvas(main_frame)
        my_canvas.pack(side=LEFT, fill=BOTH, expand=1)

        # Add Scrollbar to the canvas
        scroll = ttk.Scrollbar(main_frame, orient=VERTICAL, command=my_canvas.yview)
        scroll.pack(side=RIGHT, fill=Y)

        # Configure the canvas
        my_canvas.configure(yscrollcommand=scroll.set)
        my_canvas.bind('<Configure>', lambda e: my_canvas.configure(scrollregion=my_canvas.bbox("all")))

        # Another frame inside canvas
        second_frame = Frame(my_canvas)

        # Add that new frame to a windows inside the canvas
        my_canvas.create_window((0, 0), window=second_frame, anchor="nw")

        # Create label for text
        pop_up_label = Label(second_frame, text=text, fg="black")
        pop_up_label['font'] = myfont
        pop_up_label.pack(pady=50)

        button = Button(second_frame, text="Apasă aici pentru a deschide tabela!", command=self.open_excel)
        button['font'] = myfont
        button.pack(pady=50)

        pop_up.mainloop()

    def load_data(self, excel_file):

        baza_de_date = pd.read_excel(excel_file)
        data_frame = pd.DataFrame(baza_de_date, columns=['NR. INMATRICULARE', 'DATA EXP RCA', 'DATA EXP ROV', 'DATA EXP ITP'])
        return data_frame

    def main(self):

        baza_de_date = self.load_data(os.environ['NotificariMasini'] + '\ITP.xlsx')
        text = ""
        for index, masina in enumerate(baza_de_date['NR. INMATRICULARE'].values):
            for coloana in baza_de_date:
                if coloana == 'NR. INMATRICULARE':
                    continue
                else:
                    data = (baza_de_date[coloana][index]).to_pydatetime()
                    data_acum = dt.now()
                    delta = (data - data_acum).days  + 1
                    if delta <= 7 and delta >=0:
                        text += "Mașina {0} îi expira {1} în {2} zile, pe data de {3}\n".format(masina, self.acte_mapping[coloana], delta, data.strftime('%d.%m.%Y'))
                    elif delta < 0:
                        text += "Mașina {0} are {1} expirat/ă pe data de {2}\n".format(masina, self.acte_mapping[coloana], data.strftime('%d.%m.%Y'))
        self.check_pop_up_requirements(text)

    def check_pop_up_requirements(self, text):
        if text != "":
            self.initiate_pop_up_screen(text)

if __name__ == "__main__":

    notificari = NotificariMasini()
    notificari.main()














