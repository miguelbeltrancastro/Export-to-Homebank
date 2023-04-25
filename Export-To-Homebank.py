#Import libraries
import tkinter as tk
from tkinter import filedialog as fd
from datetime import datetime
import csv
import xlrd

#Export function
def export_file(selected_bank,filename):
    #Take the date to hava timestamp as part of the file name
    date = datetime.today().strftime('%Y%m%d')

    line_count = 0
    #For each bank the format of the input file is different
    match selected_bank:
        #Deutsche Postbank AG
        case "Postbank":
            with open(date + ' - Postbank_output.csv', mode='w', newline='') as output_file: #Open input file
                #Create CSV file
                output_writer = csv.writer(output_file, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                #Add header row to the export file
                output_writer.writerow(['date','paymode','info','payee','memo','amount','category','tags'])
                with open(filename) as csv_file:
                    csv_reader = csv.reader(csv_file, delimiter=';')
                    #Go through each row of the input file
                    for row in csv_reader:
                        if line_count >9:
                            #For each row with content, add a new row to the export CSV file ordering the value as appropriate
                            output_writer.writerow([f'{row[0]}', 6, '', '', f'{row[5]}'+' ('+f'{row[3]}'+')', f'{row[6]}'.replace('.','').replace(' €',''), '', ''])
                            line_count += 1
                        #Go throug ignoring header
                        else:
                            line_count += 1
                    #Pop-up message with information about the export done
                    tk.messagebox.showinfo(title="File processed", message='Fichero ' + filename + ' procesado como ' + selected_bank + ', '+ f'{line_count-9}' +' transacciones procesadas.')
       #Deutsche ING-DiBa AG
        case "ING DiBa":
            with open(date + ' - INGDiBa_output.csv', mode='w', newline='') as output_file:
                output_writer = csv.writer(output_file, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                output_writer.writerow(['date','paymode','info','payee','memo','amount','category','tags'])
                with open(filename) as csv_file:
                    csv_reader = csv.reader(csv_file, delimiter=';')
                    line_count = 0
                    for row in csv_reader:
                        if line_count >12:
                            output_writer.writerow([f'{row[0]}', 6, '', '', f'{row[2]}'+' ('+f'{row[5]}'+')', f'{row[8]}'.replace('.',''), '', ''])
                            line_count += 1
                        else:
                            line_count += 1
                    tk.messagebox.showinfo(title="File processed", message='Fichero ' + filename + ' procesado como ' + selected_bank + ', '+ f'{line_count-15}' +' transacciones procesadas.')
        #Spanish ING Direct
        case "ING Direct":
            with open(date + ' - INGDirect_output.csv', mode='w', newline='') as output_file:
                output_writer = csv.writer(output_file, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                output_writer.writerow(['date','paymode','info','payee','memo','amount','category','tags'])
                with open(filename) as excel_file:
                    book = xlrd.open_workbook(filename)
                    sheet = book.sheet_by_index(0)
                    line_count = 0
                    for rx in range(sheet.nrows):
                        if line_count >5:
                            output_writer.writerow([sheet.row(rx)[0].value, 6, '', '', sheet.row(rx)[3].value.replace(';','.'), sheet.row(rx)[6].value, '', ''])
                            line_count += 1
                        else:
                            line_count += 1
                    tk.messagebox.showinfo(title="File processed", message='Fichero ' + filename + ' procesado como ' + selected_bank + ', '+ f'{line_count-5}' +' transacciones procesadas.')
    window.quit()

#Creation of the main window
window = tk.Tk()
window.title("Export to Homebank")
window.geometry('640x480+50+50')

#Definition of the list of the banks for selection
banks = [
    "Postbank",
    "ING DiBa",
    "ING Direct"
]
#Creation of the bank drop down selector
selected_bank = tk.StringVar()
selected_bank.set("Postbank") 
bank = tk.OptionMenu( window , selected_bank , *banks )
bank.pack()

#Creation of the export button
export_button = tk.Button(window,text ="Export to Homebank",command=lambda: export_file(selected_bank.get(),filename))
export_button.pack()

#Creation of the exit button
exit_button = tk.Button(window,text ="Exit",command=lambda: window.quit())
exit_button.pack()

#Request to select the input file
filename = fd.askopenfilename(title="Select input file", filetypes =[('Excel Files', ('*.csv', '*.xlsx', '*.xls'))])

window.mainloop()