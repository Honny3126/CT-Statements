#!/usr/bin/env python3
"""
@author: Max Ramm
@date: 05/04/22
@version: 2.0.0

The following program was developed for Metcash Food and Grocery

Automation to summarise Excel invoice sheets into pdf files
"""

import os
import subprocess
import csv
import pandas as pd
import pickle
from pathlib import Path
from tkinter import Tk, Button, Frame, Label, W, E, StringVar, messagebox
from tkinter.messagebox import askyesno
from tkinter.ttk import Style, Progressbar
from tkinter.filedialog import askopenfilename, askdirectory
import pdfkit
path_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
path_to_notepad = 'C:\\Windows\\System32\\notepad.exe'

# window initialise
root = Tk()
root.title('Statements Program')
root.geometry("900x200")
root.minsize(900, 200)

# root variables
file = StringVar()
directory = StringVar()
abn = StringVar()


# Custom Exceptions
class MissingAbnFile(Exception):
    pass


def main():

    # retrieving previous directory and filename locations
    try:
        pickle_file = open("file.pkl", "rb")
        file.set(pickle.load(pickle_file))
        pickle_file.close()
    except (IOError, OSError, EOFError) as e:
        file.set("No file selected")

    try:
        pickle_file = open("dir.pkl", "rb")
        directory.set(pickle.load(pickle_file))
        pickle_file.close()
    except (IOError, OSError, EOFError) as e:
        directory.set("No directory selected")

    frame = Frame(root)
    gui(frame, file, directory, abn)
    root.mainloop()


def file_select_csv():
    file_win = Tk()
    file_win.withdraw()
    file_path = askopenfilename(filetypes=[("CSV Files", "*.csv")])
    pickle_file = open("abn.pkl", "wb")
    pickle.dump(file_path, pickle_file)
    pickle_file.close()
    abn.set(file_path)
    return file_path


def file_select():
    file_win = Tk()
    file_win.withdraw()
    file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    pickle_file = open("file.pkl", "wb")
    pickle.dump(file_path, pickle_file)
    pickle_file.close()
    file.set(file_path)
    return file_path


def dir_select():
    dir_win = Tk()
    dir_win.withdraw()
    dir_path = askdirectory()
    pickle_file = open("dir.pkl", "wb")
    pickle.dump(dir_path, pickle_file)
    pickle_file.close()
    directory.set(dir_path)
    return dir_path


def create_credit_note(df, dir_str):
    abn_const = 67004391422
    gst_const = 11

    # date ending
    date_ending = (df.iat[2, 17])
    date_ending = date_ending[12:]  # form of dd/mm/yyyy
    date_ending_str = (date_ending[:6] + date_ending[8:]).replace('/', "") # form of ddmmyy

    # last row (totals) retrieved as a series, not dataframe
    total_row = df.iloc[-1]

    total_invoice = total_row[17] # column R
    total_supp_inv = total_row[5] # column F
    amount = total_invoice - total_supp_inv

    total_gst = total_row[16] # column Q
    total_supp_inv_gst = total_row[6] # column G
    amount_gst = total_gst - total_supp_inv_gst

    html = f"""
        <h1>Metcash Food & Grocery</h1>
        <h1>ABN: {abn_const}</h1>
        <h1>Credit Note</h1>
        <p><b>Credit Note Number: </b>CT Disc {date_ending_str}</p>
        <p><b>Date: </b> {date_ending}</p>
        <br>
        <table style="border: 1px solid black;" border="0" align="center" width="85%">
        <thead>
            <tr style="outline: thin solid">
                <th height="5px" width="10%"><b>Code</b></th>
                <th height="5px" width="40%"><b>Description</b></th>
                <th height="5px" width="20%"><b>GST</b></th>
                <th height="5px" width="30%"><b>Amount</b></th>
            </tr>
        </thead>
        <tbody>
            <tr>
                    <td style="text-align:center;border-top: 1px solid black;">CTDISC</td>
                    <td style="text-align:center;border-top: 1px solid black;">Charge through discount</td>
                    <td style="text-align:center;border-top: 1px solid black">10</td>
                    <td style="text-align:center;border-top: 1px solid black">{"{:.2f}".format(gst_const*amount_gst)}</td>
                </tr>
                <tr>
                    <td style="text-align:center;border-bottom: 1px solid black;">CTDISC</td>
                    <td style="text-align:center;border-bottom: 1px solid black;">Charge through discount</td>
                    <td style="text-align:center;border-bottom: 1px solid black;">0</td>
                    <td style="text-align:center;border-bottom: 1px solid black;">{"{:.2f}".format(amount-(gst_const*amount_gst))}</td>
                </tr>
                <tr>
                    <td style="height:50px;"></td>
                    <td style="height:50px;"></td>
                    <td style="height:50px;"></td>
                    <td style="height:50px;"></td>
                </tr>
            <tr >
                    <td style="text-align:center;border-top: 1px solid black;"></td>
                    <td style="text-align:center;border-top: 1px solid black;"><b>Total GST</b></td>
                    <td style="text-align:center;border-top: 1px solid black;"></td>
                    <td style="text-align:center;border-top: 1px solid black;">{"{:.2f}".format(amount_gst)}</td>
                </tr>
                <tr>
                    <td style="text-align:center;"></td>
                    <td style="text-align:center;"><b>Total Credit Note</b></td>
                    <td style="text-align:center;"></td>
                    <td style="text-align:center;">{"{:.2f}".format(amount)}</td>
                </tr>

        </tbody>
        </table>"""
    # using pdfkit to write the html as a pdf, into the selected directory
    try:
        dir_path = Path(dir_str)
        pdfkit.from_string(html, f"{dir_path}/Credit Note.pdf", configuration=config)
    except IOError:
        print('continuing...')


def create_supplier_statements(df, dir_str):
    # Getting the list of ABNs
    try:
        pickle_file = open("abn.pkl", "rb")
        abn.set(pickle.load(pickle_file))
        pickle_file.close()
        abn_path = Path(abn.get())
    except (IOError, OSError, EOFError):
        raise MissingAbnFile

    try:
        csv_file = open(abn_path)
    except PermissionError:
        messagebox.showerror(title='', message='Please close the ABN CSV file on your desktop!')
    except FileNotFoundError:
        messagebox.showerror(title='', message='File not found, delete "abn.pkl" in src folder and reselect CSV')

    csvreader = csv.reader(csv_file)

    abn_header = []
    abn_header = next(csvreader)
    abn_rows = []
    # Checking through the rows
    for row in csvreader:
        abn_rows.append(row)
    csv_file.close()


    # date ending
    date_ending = (df.iat[2, 17])
    date_ending = date_ending[12:]  # form of dd/mm/yyyy

    # getting relevant subset of information needed
    subset_df = df.iloc[11:,[2,3,4,5,13]]
    subset_df = subset_df.dropna()
    # column0 = supp number
    # column1 = supp name
    # column2 = invoice num
    # column3 = supp inv. amount
    # column4 = invoice date

    # need mergesort for stability, do not change to faster alg. unless stable
    sorted_subset = subset_df.sort_values(by=subset_df.columns[0], kind='mergesort')
    sorted_subset.reset_index(drop=True)
    sorted_subset.dropna()

    # loop through from the beginning of the sorted dataset, making use of the stable sort.
    # Extract each supplier's relevant rows of information and send them to supp_pdf()
    # subset_df_sort above will put the NaN values at the bottom
    # once we reach NaN the algorithm will stop
    row_start = 0 # index, inclusive
    row_end = 0 # index, exclusive
    not_NaN = True # loop condition
    while not_NaN:
        # increment row_end until it doesn't match, or until row_end has reached the final index
        while sorted_subset.iat[row_end, 0] == sorted_subset.iat[row_start, 0]:
            row_end += 1
            if len(sorted_subset) == row_end:
                break

        supp_dataframe = sorted_subset.iloc[row_start:row_end,:]

        # call helper function here to write pdf
        supp_pdf(supp_dataframe, dir_str, date_ending, abn_rows)

        # update row_start and row_end
        row_start = row_end

        # check outer loop condition, will stop the loop once the last supplier is completed
        if len(sorted_subset) == row_end:
            not_NaN = False
        elif (sorted_subset.iat[row_start, 0]) == 'No':
            not_NaN = False

    messagebox.showinfo(title=None, message="Statements Finished")


def supp_pdf(supp_df, dir_str, date_ending_formatted, abn_rows):
    supplier_name = supp_df.iat[0, 1]
    supplier_num = supp_df.iat[0, 0]
    supplier_abn = 'N/A'

    # searching through abn data and matching supplier number to the ABN csv
    # if no match is found, supplier abn will be presented as 'n/a'
    for row in abn_rows:
        if row[0] == supplier_num:
            supplier_abn = row[2]
            break

    html = f"""
    <h1>{supplier_num} - {supplier_name}</h1>
    <h1>ABN: {supplier_abn}</h1>
    <h1>Statement</h1>
    <h2><b>Date: </b>{date_ending_formatted}</h2>
    <br>
    <table style="border: 1px solid black;" border="0" align="center" width="70%">
    <thead>
        <tr>
            <th style="border: 1px solid black;" width="20%"><b>Date</b></th>
            <th style="border: 1px solid black;" width="50%"><b>Invoice Number</b></th>
            <th style="border: 1px solid black;" width="30%"><b>Amount</b></th>
        </tr>
    </thead>
    <tbody>"""
    amount_sum = 0
    for i in range(len(supp_df)):
        # order: date, invoice num, amount
        amount_sum += supp_df.iat[i,3]
        html += f"""
            <tr>
                <td style="text-align:center;border: 1px solid black;">{str(supp_df.iat[i,4].strftime('%d/%m/%Y'))}</td>
                <td style="text-align:center;border: 1px solid black;">{str(supp_df.iat[i,2])}</td>
                <td style="text-align:center;border: 1px solid black;">{"{:.2f}".format(supp_df.iat[i,3])}</td>
            </tr>        
        """

    html += f"""
            <tr>
                <td style="height:30px;border: 1px solid black;"><b></b></td>
                <td style="height:30px;border: 1px solid black;"></td>
                <td style="height:30px;border: 1px solid black;"></td>
            </tr>
            <tr>
                <td style="text-align:center;border: 1px solid black;"><b>Total</b></td>
                <td style="border: 1px solid black;"></td>
                <td style="text-align:center;border: 1px solid black;">{"{:.2f}".format(amount_sum)}</td>
            </tr>
            
    </tbody>
    </table> """
    # using pdfkit to write the html as a pdf, into the selected directory
    try:
        dir_path = Path(dir_str)
        # completeName = os.path.join(dir_str, supplier_name.replace("/","") + ".pdf")
        # file1 = open(completeName, "w")
        # file1.close()
        pdfkit.from_string(html, f"{dir_path}/{supplier_name.replace('/','')}.pdf", configuration=config)
    except IOError as e:
        # print(supplier_name)
        # print(e)
        pass


def gui(frame, file, directory, abn):

    # !gui methods
    def clear_dir():

        confirm = askyesno(title='Clear Statements Directory?',
                           message='Are you sure you want to clear ALL files in selected statements directory?')
        if confirm:
            # if block will run if the user selects 'Yes'

            # remove GUI elements during operation / begin progress bar
            # progress_bar.start(10)
            create_statements_btn.grid_remove()
            clear_statements_btn.grid_remove()
            change_file_btn.grid_remove()
            change_dir_btn.grid_remove()

            # remove files
            try:
                for f in os.listdir(directory.get()):
                    os.remove(os.path.join(directory.get(), f))
            except PermissionError:
                messagebox.showerror(title="Sub-directories detected",message="Selected 'Statements Target Directory' cannot have sub-directories, please delete these sub-directories or choose new folder!")
            else:
                messagebox.showinfo(title="Completed Clear", message="Directory cleared successfully!")

            # put GUI elements back after completed
            # progress_bar.stop()
            create_statements_btn.grid()
            clear_statements_btn.grid()
            change_file_btn.grid()
            change_dir_btn.grid()

    def create_statements(file_str, dir_str):
        # remove GUI elements and display loading bar
        # progress_bar.start(10)
        create_statements_btn.grid_remove()
        clear_statements_btn.grid_remove()
        change_dir_btn.grid_remove()
        change_file_btn.grid_remove()

        # convert the tkinter variable into a valid Path
        file_str = Path(file_str)

        try:
            # create dataframe using pandas
            df = pd.read_excel(file_str)

            # create pdfs using the dataframe
            create_credit_note(df, dir_str)
            create_supplier_statements(df, dir_str)
        except (AttributeError, ValueError, KeyError, IndexError) as e:
            # print(e)
            messagebox.showerror(title="Invalid Excel File", message="Summary File is not formatted correctly, please choose valid Excel Sheet!")
        except PermissionError:
            messagebox.showerror(title="Excel file open", message="Please close the file you want to make summaries for!")
        except MissingAbnFile:
            messagebox.showerror(title="Missing ABN File",
                                 message="Please enter an ABN File!")

        # put GUI elements back
        # progress_bar.stop()
        create_statements_btn.grid()
        clear_statements_btn.grid()
        change_file_btn.grid()
        change_dir_btn.grid()

    def open_abn():
        # retrieving abn file
        try:
            pickle_file = open("abn.pkl", "rb")
            abn.set(pickle.load(pickle_file))
            pickle_file.close()
            abn_path = Path(abn.get())
            subprocess.call([path_to_notepad, abn_path])
        except (IOError, OSError, EOFError) as e:
            file_select_csv()
            messagebox.showinfo(title="", message="ABN File added successfully!")



    # !GUI configuration and layout
    Style().configure("TButton", padding=(0, 20, 0, 20), font='serif 10')

    frame.columnconfigure(0, pad=20, weight=1)
    frame.columnconfigure(1, pad=50, weight=1)
    frame.columnconfigure(2, pad=50, weight=1)
    frame.columnconfigure(3, pad=20, weight=1)
    frame.columnconfigure(4, pad=20, weight=1)

    frame.rowconfigure(0, pad=30, weight=1)
    frame.rowconfigure(1, pad=30, weight=1)
    frame.rowconfigure(2, pad=30, weight=1)
    frame.rowconfigure(3, pad=30, weight=1)
    frame.rowconfigure(4, pad=30, weight=1)

    label1 = Label(frame, text='Summary Filepath: ', anchor=W)
    label1.grid(row=0, column=0, sticky=W)

    file_path = Label(frame, textvariable=file, anchor=W)
    file_path.config(bg="white")
    file_path.grid(row=0, column=1, columnspan=3, sticky=W+E, ipady=5, ipadx=20)

    change_file_btn = Button(frame, text='Change File', command=file_select, width=20)
    change_file_btn.grid(row=0, column=4, sticky=W, padx=10)

    label2 = Label(frame, text='Statements Target Directory: ',anchor=W)
    label2.grid(row=1, column=0, sticky=W)

    dir_path = Label(frame, textvariable=directory, anchor=W)
    dir_path.config(bg='white')
    dir_path.grid(row=1, column=1, columnspan=3, sticky=W+E, ipady=5, ipadx=2)

    change_dir_btn = Button(frame, text='Change Directory', command=dir_select, width=20)
    change_dir_btn.grid(row=1, column=4, sticky=W, padx=10)

    clear_statements_btn = Button(frame, text='Clear Statements Folder', command=lambda: clear_dir(), width=20)
    clear_statements_btn.grid(row=2, column=4, sticky=W, padx=10)

    create_statements_btn = Button(frame, text='Create Statements', command=lambda: create_statements(file.get(), directory.get()), width=20)
    create_statements_btn.grid(row=3, column=4, sticky=W, padx=10)

    abn_btn = Button(frame, text='ABN File', command=lambda: open_abn(), width=20)
    abn_btn.grid(row=3, column=0, sticky=W, padx=10)

    # progress_bar = Progressbar(frame, length=100, mode='indeterminate')
    # progress_bar.grid(row=4, column=1, columnspan=3, sticky=W+E)

    frame.pack()


if __name__ == '__main__':
    main()
