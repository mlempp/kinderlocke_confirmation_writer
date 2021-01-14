
import os
import tkinter as tk
from tkinter import filedialog
from PIL import ImageTk, Image
from pandastable import Table
import pandas as pd
#from docx import Document
#from docx.shared import Inches
#from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docxtpl import DocxTemplate
from datetime import date


class MainApp(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.parent.title('BestellDrucker 0.1')
        self.parent.geometry('1000x600')
        self.template = 'Rechnung_Vorlage_V_0.2.docx'
        ###self.parent.iconbitmap(os.getcwd() + '/highlight1_logo.ico')
        self.parent.configure(background='white')


        img = ImageTk.PhotoImage(Image.open('highlight1_logo.png').resize((110, 72)))

        pik = tk.Label(self.parent, width=110, height=72, image = img, background='white')
        pik.image = img
        pik.place(width=110,height=72, x=445, y=10)
        # panel = tk.Label(self.parent, image=img)
        # panel.place(relwidth = 0.5, relheight = 0.5, relx = 0, rely = 0)

        self.Table = tk.Frame(self.parent)
        self.Table.place(width=940,height=350, x=30, y=150)

        Load_Button = tk.Button(self.parent, text='Wähle die Auftragsliste', command = self.Load_Button_Click, background='white')
        Load_Button.place(height=25, y=90, x=30)

        Label2 =tk.Label(self.parent, text='Bestellnummer: ', background='white')
        Label2.place(height=25, y= 550, x= 30)

        self.ORDER_ID = tk.Entry(self.parent, width=70, background='white')
        self.ORDER_ID.place(height=25, y= 550, x= 220)

        Print_Button = tk.Button(self.parent, text='Erstelle Bestätigung', command = self.Print_button_Click, background='white')
        Print_Button.place(height=25, y=550, x=750)


    def Load_Button_Click(self):
        self.full_filepath = None
        self.full_filepath= filedialog.askopenfilename(initialdir=os.getcwd(), title="Wähle die Autragsliste",
                                   filetypes=(("ods files", "*.ods"), ("all files", "*.*")))
        if self.full_filepath != None:
            self.filename = self.full_filepath.split('/')[-1]
            Show_List = tk.Label(self.parent, text='geladene Bestellliste: '+ self.filename, background='white')
            Show_List.place(height=25, y=90, x=220)
            self.df_data = pd.read_excel(self.full_filepath, sheet_name = 'Auftraege', skiprows = 3, engine = 'odf')

            self.table = pt = Table(self.Table, dataframe=self.df_data,
                                    showtoolbar=False, showstatusbar=True)
            pt.show()

    def Print_button_Click(self):

        Order_ID = self.ORDER_ID.get().replace(' ','')
        self.ORDER_ID.delete(0, 'end')
        df = self.df_data
        Order_df = df[df.Bestellnummer == Order_ID].copy()
        Response = tk.Label(self.parent, text='', background='white')
        Response.place(height=25, y=550, x=880, width = 120)

        if Order_df.shape[0]> 0:
            path_output = filedialog.askdirectory(initialdir=os.getcwd(), title="Wähle Speicherort", )
            template_tmp = DocxTemplate(self.template)
            context = self.get_context(Order_df)
            template_tmp.render(context)
            template_tmp.save(path_output+ '/' + Order_ID + '.docx')
            Response = tk.Label(self.parent, text='gedruckt', background='white')
            Response.place(height=25, y=550, x=880)
        else:
            Response = tk.Label(self.parent, text='nicht gefunden', background='white')
            Response.place(height=25, y=550, x=880)


    def get_context(self, tmp_df):
        tmp_dict = {
            'name': tmp_df.Besteller.unique()[0],
            'adress': tmp_df['Adresse, Straße'].unique()[0],
            'city': tmp_df['Adresse, Stadt'].unique()[0],
            'date': date.today().strftime("%d.%m.%Y"),
            'order_ID': tmp_df.Bestellnummer.unique()[0],
            'i_order_final': ["%.2f €" % tmp_df['Gesamtpreis Bestellung'].mean()][0].replace('.', ','),
            'i_ship': ["%.2f €" % tmp_df['Versandkosten'].mean()][0].replace('.', ','),
            'i_order_total': ["%.2f €" % (tmp_df['Gesamtpreis Bestellung'].mean() - tmp_df['Versandkosten'].mean())][
                0].replace('.', ','),
            'row_contents': []
        }

        for ix, (i, row) in enumerate(tmp_df.iterrows()):
            tmp_dict['row_contents'].append({
                'pos': str(ix + 1),
                'description': row.Bestellung,
                'a': str(1),
                'rate': ["%.2f €" % row.Preis][0].replace('.', ','),
                'total': ["%.2f €" % row.Preis][0].replace('.', ','),
            })

        return tmp_dict



