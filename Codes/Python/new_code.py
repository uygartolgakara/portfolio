"""Author: Uygar Tolga Kara - Date: Fri Sep 22 19:26:18 2023."""
# -*- coding: utf-8 -*-

# Import Statements
from pandas import isna, options
from PyPDF2 import PdfReader
from tabula import read_pdf
from re import findall, search, S
from json import dump

# Set Options
options.mode.chained_assignment = None


class PDFAnalyzer:
    """Provide 2 functions for analyzing pdf data."""

    def __init__(self, pdf_path: str):
        """Initialize the class."""
        self.pdf_path = pdf_path
        self.pdf_reader = PdfReader(self.pdf_path)
        self.pdf_pagenum = len(self.pdf_reader.pages)
        self.pdf_pagelist = self.gather_pages()
        self.mode_list = self.gather_modes()
        self.mode_dict = self.gather_mode_dict()
        self.dataframes = read_pdf(self.pdf_path, pages="all")
        self.pid_dataframes = self.filter_dataframes()
        self.rule_dict = self.gather_rule_dict()
        self.make_output_files()

    def gather_pages(self):
        """Return a list of pdf page strings."""
        output = [self.pdf_reader.pages[i].extract_text()
                  for i in range(self.pdf_pagenum)]
        return output

    def gather_modes(self):
        """Return a list of modes inside pdf."""
        content_page = [page for page in self.pdf_pagelist
                        if "Inhaltsverzeichnis" in page][0]
        output = findall(r"(Mode\$\d\d).*?(\d{1,2})", content_page, S)
        return output

    def gather_mode_dict(self):
        """Return dictionary from modes."""
        mode_dict = {}
        for index, mode_tuple in enumerate(self.mode_list):
            mode_name = mode_tuple[0]
            page_start = int(mode_tuple[1]) - 1
            page_end = int(self.mode_list[index+1][1]) - 1 \
                if index != len(self.mode_list) - 1 else self.pdf_pagenum

            pid_list = []
            for idx in range(page_start, page_end):
                page = self.pdf_pagelist[idx]
                search = findall(r"(PID\$(?:\d|\w){2})", page, S)
                pid_list.append(search)
            pid_list = [pid for pid_group in pid_list for pid in pid_group]
            pid_list = sorted(set(pid_list))

            mode_dict[mode_name] = pid_list
        return mode_dict

    def filter_dataframes(self):
        """Gather tables inside pdf file using tabula."""
        output = [dataframe for dataframe in self.dataframes
                  if list(dataframe.columns)[0].startswith("PID$")]
        output = [data for data in output if "Bit:" in str(data)]
        return output

    def gather_rule_dict(self):
        """Return dictionary from rules."""
        rule_dict = {}
        for index, pid_dataframe in enumerate(self.pid_dataframes):

            title = pid_dataframe.columns[0]
            pattern = r"^PID\$[0-9A-Z]{2}$"
            pid_name = title if search(pattern, title) else title.split()[0]
            rule_dict[pid_name] = {}

            bit_data = []
            for idx, row in pid_dataframe.iterrows():
                related = False
                for column in row:
                    if (search(r"Bit:\s?\d\s?=", str(column))
                            or search("^[A-Z]:", str(column))):
                        related = True
                        break
                bit_data.append(related)
            pid_dataframe = pid_dataframe[bit_data]
            pid_dataframe.reset_index(drop=True, inplace=True)

            for idx, row in reversed(list(pid_dataframe.iterrows())):
                if not search(r"Bit:\s?\d\s?=", row.str.cat(sep=''), S):
                    pid_dataframe.drop(idx, inplace=True)
            pid_dataframe.reset_index(drop=True, inplace=True)
            pid_dataframe.iloc[:, 0].fillna(method='ffill', inplace=True)

            memory = ""
            for idx, row in pid_dataframe.iterrows():
                row_str = row.str.cat(sep='')
                pattern = r"(?:([A-Z])\:\s?)?.*?Bit: (\d)\s?=\s?([0-1X])"
                dataset = findall(pattern, row_str, S)
                label, bit_num, bit_val = dataset[0]
                memory = label if label != "" else memory
                label = memory if label == "" else label
                if label not in list(rule_dict[pid_name].keys()):
                    rule_dict[pid_name][label] = {}
                rule_dict[pid_name][label]["Bit" + bit_num] = bit_val

        return rule_dict

    def make_output_files(self):
        """Make rule and mode json files."""
        with open("mode.json", "w") as json_file:
            dump(self.mode_dict, json_file)
        with open("mode.json", "w") as json_file:
            dump(self.mode_dict, json_file)

    def check_pid(self, pid_name: str) -> list:
        """Return mode names that pid is related in."""
        found_list = []
        for mode, pid_list in self.mode_dict.items():
            if pid_name in pid_list:
                found_list.append(mode)

        found_str = "\n".join(found_list)
        print(f"\n{pid_name} is inside:\n{found_str}\n")
        return found_list

    def check_number(self, number: int, pid_name: str) -> int:
        """Turn given integer to binary, check rules in PID name."""
        bit_length = 1
        while 2 ** bit_length <= number:
            bit_length += 1
        if bit_length % 8 != 0:
            bit_length = (bit_length // 8 + 1) * 8

        item_bit = bin(number)[2:].zfill(bit_length)

        if pid_name not in list(self.rule_dict.keys()):
            print(f"\n{pid_name} not available. Available options:")
            print("\n".join(list(self.rule_dict.keys())) + "\n")
            return "Not Available"

        rule_dict = self.rule_dict[pid_name]
        if len(rule_dict)*8 != bit_length:
            print(f"\nBit length for given number: {bit_length}")
            print(f"Bit length needed forat pid: {len(rule_dict)*8}")
            print("Make sure that bit numbers are equal.\n")
            return "Bit Lengths Not Compatible"

        rule_bit = ""
        for rule_letter, bit_dict in rule_dict.items():
            for i in range(7, -1, -1):
                rule_bit += bit_dict[f"Bit{i}"]

        print(f"\nGiven number: {item_bit}")
        print(f"Rules number: {rule_bit}\n")

        for i in range(len(item_bit)):
            char1 = item_bit[i]
            char2 = rule_bit[i]
            if char1 == char2 or char2 == "X":
                continue
            bit_number = len(item_bit)-(i+1)
            print(f"Given number bit {str(bit_number).zfill(2)} is a: {char1}")
            print(f"But this bit needs to be: {char2}")
            rule_letter = list(rule_dict.keys())[(bit_number // 8)]
            print("PID name: " + pid_name + " - Rule : " + rule_letter + "\n")
            item_bit = item_bit[:i] + char2 + item_bit[i + 1:]

        return int(item_bit, 2)


if __name__ == "__main__":

    # Make class object with pdf path as user input
    pdf_analyzer = PDFAnalyzer(input("PDF Path: "))