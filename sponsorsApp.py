from customtkinter import *
import customtkinter
import tkinter.ttk as ttk
import CTkMessagebox
from CTkMessagebox import CTkMessagebox
from lxml import html
import pandas as pd
import os.path
import os
import numpy as np
from pandastable import Table, TableModel
from tkinter import END
import tkinter as tk
import os.path
import requests
from customtkinter import filedialog 

class Tabs(customtkinter.CTkTabview):
    def __init__(self, master):
        super().__init__(master)

        # create tabs
        self.add("Dashboard")
        self.add("Sponsors")

        # add widgets on tabs

class CredentialsFrame(customtkinter.CTkFrame):
    
    token = ""
    
    def __init__(self, master):
        
        super().__init__(master)
        
        self.inputToken = StringVar()
        self.inputToken.set(CredentialsFrame.token)

        # ENTRY
        self.entryToken = customtkinter.CTkEntry(self, width=100, textvariable=self.inputToken)
        self.entryToken.grid(column=1, row=9, sticky="EW")

        # LABELS
        customtkinter.CTkLabel(self, text="Token: ").grid(column=0, row=9, sticky="EW")
        # BUTTON
        self.buttonLogin = customtkinter.CTkButton(self, text="Save sponsor data to file", command=self.saveSponsors)
        self.buttonLogin.grid(column=1, row=11, sticky="EW")
        for widget in self.winfo_children():
                widget.grid(padx=5, pady=2.5)

    def saveSponsors(self):
            
            groups = ["Elite", "Master - 1", "Master - 2", "Master - 3", "Master - 4", "Master - 5", "Pro - 1", "Pro - 2", "Pro - 3", "Pro - 4", "Pro - 5", "Pro - 6", "Pro - 7", "Pro - 8", "Pro - 9", "Pro - 10", "Pro - 11", "Pro - 12", "Pro - 13", "Pro - 14", "Pro - 15", "Pro - 16", "Pro - 17", "Pro - 18", "Pro - 19", "Pro - 20", "Pro - 21", "Pro - 22", "Pro - 23", "Pro - 24", "Pro - 25", "Amateur - 1", "Amateur - 2", "Amateur - 3", "Amateur - 4", "Amateur - 5", "Amateur - 6", "Amateur - 7","Amateur - 8", "Amateur - 9", "Amateur - 10", "Amateur - 11", "Amateur - 12", "Amateur - 13", "Amateur - 14", "Amateur - 15", "Amateur - 16", "Amateur - 17", "Amateur - 18", "Amateur - 19", "Amateur - 20", "Amateur - 21", "Amateur - 22", "Amateur - 23", "Amateur - 24", "Amateur - 25", "Amateur - 26", "Amateur - 27", "Amateur - 28", "Amateur - 29", "Amateur - 30", "Amateur - 31", "Amateur - 32", "Amateur - 33", "Amateur - 34", "Amateur - 35", "Amateur - 36", "Amateur - 37", "Amateur - 38", "Amateur - 39", "Amateur - 40", "Amateur - 41", "Amateur - 42", "Amateur - 43", "Amateur - 44", "Amateur - 45", "Amateur - 46", "Amateur - 47", "Amateur - 48", "Amateur - 49", "Amateur - 50", "Amateur - 51", "Amateur - 52", "Amateur - 53", "Amateur - 54", "Amateur - 55", "Amateur - 56", "Amateur - 57", "Amateur - 58", "Amateur - 59", "Amateur - 60", "Amateur - 61", "Amateur - 62", "Amateur - 63", "Amateur - 64", "Amateur - 65", "Amateur - 66", "Amateur - 67", "Amateur - 68", "Amateur - 69", "Amateur - 70", "Amateur - 71", "Amateur - 72", "Amateur - 73", "Amateur - 74", "Amateur - 75", "Amateur - 76", "Amateur - 77", "Amateur - 78", "Amateur - 79", "Amateur - 80"]
            try:
                CredentialsFrame.token = self.entryToken.get()
                
                headers =  {"Content-Type":"application/json", "Authorization": f"Bearer {CredentialsFrame.token}"}
                url = "https://gpro.net/gb/backend/api/v2/ManSponsors"
                names = []
                sponsorList = []
                sponsorName = []
                sponsorGroup = []
                sponsorFinances = []
                sponsorExpectations = []
                sponsorPatience = []
                sponsorReputation = []
                sponsorImage = []
                sponsorNegotiation = []
                sponsorDuration = []
                for sufix in groups:

                    params = f"Group={sufix}"
                    r = requests.get(url, headers=headers, params=params)
                    data = r.json()
                    for sponsor in data["data"]:
                        names.append(sponsor["manName"])
                        sponsorName.append(sponsor["sponsorName"])
                        sponsorGroup.append(sponsor["sponsorGroup"])
                        sponsorFinances.append(sponsor["finances"]+1)
                        sponsorExpectations.append(sponsor["expectations"]+1)
                        sponsorPatience.append(sponsor["patience"]+1)
                        sponsorReputation.append(sponsor["reputation"]+1)
                        sponsorImage.append(sponsor["image"]+1)
                        sponsorNegotiation.append(sponsor["negotiation"]+1)
                        sponsorDuration.append(sponsor["duration"])

                iterables = [names, sponsorName, sponsorGroup, sponsorFinances, sponsorExpectations, sponsorPatience, sponsorReputation, sponsorImage, sponsorNegotiation, sponsorDuration]
                sponsorListPartial = list(zip(*iterables))
                sponsorList.extend(sponsorListPartial)
                keys = ["Name", "sponsorName", "Group", "Finances", "Expectations", "Patience", "Reputation", "Image", "Negotiation", "Duration"]
                sponsorListDict = [dict(zip(keys, values)) for values in sponsorList]
        
            
                df_sponsors = pd.DataFrame(sponsorListDict)              

                file_path = filedialog.asksaveasfilename(filetypes=[("Excel files", "*.xlsx")],initialdir = os.getcwd())
                with pd.ExcelWriter(f'{file_path}.xlsx') as writer:
                    df_sponsors.to_excel(writer, sheet_name='Sheet1', index=False)
                CTkMessagebox(title="Info", message="Sponsors saved as " + file_path + ".xlsx",
                    icon="check", option_1="Thanks")
                        
            except Exception:
                CTkMessagebox(title="Error", message="Invalid token!")
       


class SponsorFrame(customtkinter.CTkFrame):
    def __init__(self, master, title):
        super().__init__(master)

        #self.grid_rowconfigure(0, weight=1)
        self.title = title

        self.title = customtkinter.CTkLabel(self, text=self.title, fg_color="gray30", corner_radius=6)
        self.title.grid(row=0, column=0, sticky="ew", columnspan=6)

        SponsorFrame.group = StringVar()
        SponsorFrame.group.set("Elite")
    


        self.group_combobox = customtkinter.CTkComboBox(self, values=["Elite", "Master - 1", "Master - 2", "Master - 3", "Master - 4", "Master - 5", "Pro - 1", "Pro - 2", "Pro - 3", "Pro - 4", "Pro - 5", "Pro - 6", "Pro - 7", "Pro - 8", "Pro - 9", "Pro - 10", "Pro - 11", "Pro - 12", "Pro - 13", "Pro - 14", "Pro - 15", "Pro - 16", "Pro - 17", "Pro - 18", "Pro - 19", "Pro - 20", "Pro - 21", "Pro - 22", "Pro - 23", "Pro - 24", "Pro - 25", "Amateur - 1", "Amateur - 2", "Amateur - 3", "Amateur - 4", "Amateur - 5", "Amateur - 6", "Amateur - 7","Amateur - 8", "Amateur - 9", "Amateur - 10", "Amateur - 11", "Amateur - 12", "Amateur - 13", "Amateur - 14", "Amateur - 15", "Amateur - 16", "Amateur - 17", "Amateur - 18", "Amateur - 19", "Amateur - 20", "Amateur - 21", "Amateur - 22", "Amateur - 23", "Amateur - 24", "Amateur - 25", "Amateur - 26", "Amateur - 27", "Amateur - 28", "Amateur - 29", "Amateur - 30", "Amateur - 31", "Amateur - 32", "Amateur - 33", "Amateur - 34", "Amateur - 35", "Amateur - 36", "Amateur - 37", "Amateur - 38", "Amateur - 39", "Amateur - 40", "Amateur - 41", "Amateur - 42", "Amateur - 43", "Amateur - 44", "Amateur - 45", "Amateur - 46", "Amateur - 47", "Amateur - 48", "Amateur - 49", "Amateur - 50", "Amateur - 51", "Amateur - 52", "Amateur - 53", "Amateur - 54", "Amateur - 55", "Amateur - 56", "Amateur - 57", "Amateur - 58", "Amateur - 59", "Amateur - 60", "Amateur - 61", "Amateur - 62", "Amateur - 63", "Amateur - 64", "Amateur - 65", "Amateur - 66", "Amateur - 67", "Amateur - 68", "Amateur - 69", "Amateur - 70", "Amateur - 71", "Amateur - 72", "Amateur - 73", "Amateur - 74", "Amateur - 75", "Amateur - 76", "Amateur - 77", "Amateur - 78", "Amateur - 79", "Amateur - 80"], variable = SponsorFrame.group)
        self.group_combobox.grid(row=1, column=1)

        customtkinter.CTkLabel(self, text="Choose group").grid(column=0, row=1, sticky="W")

        # Button calculate
        self.calculate_button = customtkinter.CTkButton(self, text="Show Sponsors", command=self.sponsorsBtn)
        self.calculate_button.grid(row=2, column=0, columnspan=1, sticky="ew")
        self.sponsorsBtn = None
        

        for widget in self.winfo_children():
            widget.grid(padx=10, pady=5)

    def sponsorsBtn(self):
        if self.sponsorsBtn is None or not self.sponsorsBtn.winfo_exists():
            self.sponsorsBtn = SponsorWindow(self)
            self.sponsorsBtn.focus()  # create window if its None or destroyed
        else:
            self.sponsorsBtn.focus()  # if window exists focus it

class SponsorWindow(customtkinter.CTkToplevel):
    def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.geometry("800x600")
            try:
                df_sponsors = pd.read_excel(filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx")],initialdir = os.getcwd()))           
                df_to_show = df_sponsors[df_sponsors.Group == str(SponsorFrame.group.get())].sort_values(by='Duration', ascending=True)
                
                self.table_FRAME = customtkinter.CTkFrame(self)
                self.table = Table(self.table_FRAME, dataframe=df_to_show,
                                        showtoolbar=True, showstatusbar=True)
                self.table.grid(row=0, column=0, sticky="NEWS", columnspan=10)
                self.table_FRAME.pack(fill=BOTH,expand=1)
                self.table.show()
            except Exception:
                CTkMessagebox(title="Error", message="Could not retrieve sponsor data")




class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        customtkinter.set_appearance_mode("system")
        self.title("Sponsors App")
        self.tab_view = Tabs(master=self)
        self.tab_view.grid(row=0, column=0, padx=5, pady=5)
        App.warningLabel = StringVar()
        
       
        self.credentials_frame = CredentialsFrame(self.tab_view.tab("Dashboard"))
        self.credentials_frame.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="nsew")
        self.credentials_frame.configure(fg_color="transparent")

        self.setup_frame = SponsorFrame(self.tab_view.tab("Sponsors"), "Sponsors")
        self.setup_frame.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="nsew")
        self.setup_frame.configure(fg_color="transparent")


        
app = App()
app.mainloop()
