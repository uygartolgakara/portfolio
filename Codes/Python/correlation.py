# %% Import Modules

import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog

# Set pandas to suppress notifications
pd.options.mode.chained_assignment = None

# %% Hardcoded Variables

# wb_path = r"C:\Users\KUY3IB\Desktop\data.xlsx"

# %% Class Initialize

class CorrelationTool():

    def __init__(self, excel_path: str):

        self.df = pd.read_excel(excel_path)
        self.filter_dropdowns = {}

        H = list(self.df.columns)

        self.L = [h.split("value ")[1] for h in H if h.startswith("value ")]
        self.F = [h for h in H if not h.startswith("value")]
        self.F = [f for f in self.F if "date" not in f.lower()]
        O = [sorted(list(set(list(self.df[f])))) for f in self.F]

        self.FO = {self.F[i]: O[i] for i in range(len(self.F))}

        self.root = tk.Tk()
        self.root.title("Filter Selection")

        i = 0

        for F, O in self.FO.items():

            label = ttk.Label(self.root, text=F, width=20)
            label.grid(row=i, column=0, sticky='w')

            var = tk.StringVar(value=O)
            dropdown = ttk.Combobox(self.root, values=O, textvariable=var,
                                    state="readonly", width=40)
            dropdown.set("")
            dropdown.grid(row=i, column=1)

            self.filter_dropdowns[F] = var

            i += 1

        reset_button = tk.Button(
            self.root,
            text="Reset",
            command=self.reset,
            width=20,
            background="red",
            foreground="white"
            )
        reset_button.grid(row=i, column=0)

        calculate_button = tk.Button(
            self.root,
            text="Calculate",
            command=self.calculate,
            width=36,
            background="green",
            foreground="white"
            )
        calculate_button.grid(row=i, column=1)

        self.root.mainloop()

    def reset(self):

        for var in self.filter_dropdowns.values():
            var.set(())

    def calculate(self):

        global cdf, cdf1, cdf2, avai

        text = ""
        print()

        cdf = self.df.copy()

        for F, var in self.filter_dropdowns.items():
            option = var.get()

            text += f"{F} = {option}\n"
            print(f"{F} = {option}")

            if option != "":
                cdf = cdf[cdf[F] == option]

        text += "\n"
        print()

        cdf1 = cdf[["value " + L for L in self.L]]
        cdf2 = cdf[self.F]

        for L in self.L:

            text += f"{L} Analysis:\n"
            print(f"{L} Analysis:")

            avai = list(cdf1["value " + L])
            avai = [ava for ava in avai if not pd.isna(ava)]
            AVAI = avai.copy()
            avai = list(set(avai))

            for ava in avai:
                per = AVAI.count(ava) / len(AVAI) * 100
                text += f"{per:.2f} % - Value: {ava}\n"
                print(f"{per:.2f} % - Value: {ava}")

            text += "\n"
            print()

        messagebox.showinfo("Information", text)

# co = CorrelationTool(r"C:\Users\KUY3IB\Desktop\data.xlsx")

excel_path = filedialog.askopenfilename(filetypes=[("Excel File", "*.xlsx")])

if excel_path != "":
    co = CorrelationTool(excel_path)











