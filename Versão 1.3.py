# Created by: PyQt5 UI code generator 5.15.7
from PyQt5.QtGui import QMovie
import threading
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
import os
import sys
import time
import os
import pandas as pd
import openpyxl
import glob
import shutil
import numpy as np
import time
import seaborn as sns
import statistics
import matplotlib.pyplot as plt
from mpl_toolkits import mplot3d
from openpyxl.styles import Font, Color, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import NumberFormat
from mpl_toolkits.mplot3d import Axes3D

        
class Ui_Measurements(object):
    def setupUi(self, Measurements):
        Measurements.setObjectName("Measurements")
        Measurements.resize(400, 320)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("logo-orion-vertical-1.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Measurements.setWindowIcon(icon)
        Measurements.setStyleSheet("")
        self.centralwidget = QtWidgets.QWidget(Measurements)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setSpacing(0)
        self.gridLayout.setObjectName("gridLayout")
        self.widget = QtWidgets.QWidget(self.centralwidget)
        self.widget.setStyleSheet("background-color:rgb(0, 66, 107)")
        self.widget.setObjectName("widget")
        self.folder_select = QtWidgets.QPushButton(self.widget)
        self.folder_select.setGeometry(QtCore.QRect(25, 50, 350, 20))
        self.folder_select.setStyleSheet("background-color: rgb(0, 105, 165);\n"
"color: rgb(255,255, 255);\n"
"font: 75 10pt \"MS Shell Dlg 2\";")
        self.folder_select.setText("Select folder")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(175, 140, 100, 100))
        self.label.setMinimumSize(QtCore.QSize(100, 100))
        self.label.setMaximumSize(QtCore.QSize(100, 100))
        self.label.setObjectName("label")
        self.buttonBox = QtWidgets.QDialogButtonBox(self.widget)
        self.buttonBox.setGeometry(QtCore.QRect(120, 110, 160, 30))
        self.buttonBox.setStyleSheet("background-color: rgb(0, 105, 165);\n"
"color: rgb(255, 255, 255);\n"
"font: 75 10pt \"MS Shell Dlg 2\";")
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel|QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.lineEdit = QtWidgets.QLineEdit(self.widget)
        self.lineEdit.setGeometry(QtCore.QRect(100, 250, 200, 20))
        self.lineEdit.setStyleSheet("color: rgb(255, 255, 255);")
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit.setReadOnly(True)
        self.gridLayout.addWidget(self.widget, 0, 0, 1, 1)
        Measurements.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(Measurements)
        self.menubar.setObjectName("menubar")
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        Measurements.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(Measurements)
        self.statusbar.setObjectName("statusbar")
        Measurements.setStatusBar(self.statusbar)
        self.actionAbout_Minerva = QtWidgets.QAction(Measurements)
        self.actionAbout_Minerva.setObjectName("actionAbout_Minerva")
        self.actionAbout_Measurement = QtWidgets.QAction(Measurements)
        self.actionAbout_Measurement.setObjectName("actionAbout_Measurement")
        self.menuHelp.addAction(self.actionAbout_Measurement)
        self.menuHelp.addSeparator()
        self.menuHelp.addAction(self.actionAbout_Minerva)
        self.menubar.addAction(self.menuHelp.menuAction())
        self.actionAbout_Minerva.triggered.connect(self.open_about_minerva)
        self.actionAbout_Measurement.triggered.connect(self.open_about_measurement)

        self.retranslateUi(Measurements)
        QtCore.QMetaObject.connectSlotsByName(Measurements)
    
    def retranslateUi(self, Measurements):
        _translate = QtCore.QCoreApplication.translate
        Measurements.setWindowTitle(_translate("Measurements", "Measurements"))
        self.folder_select.setText(_translate("Measurements", "Select File Folder"))
        self.folder_select.clicked.connect(self.select_folder)
        self.lineEdit.setText(_translate("Measurements", "  Version 1.3 by Minerva Dev® & Orion"))
        self.menuHelp.setTitle(_translate("Measurements", "Help"))
        self.actionAbout_Minerva.setText(_translate("Measurements", "About Minerva"))
        self.actionAbout_Measurement.setText(_translate("Measurements", "About Measurement"))
        self.buttonBox.rejected.connect(self.close_program)
        self.buttonBox.accepted.connect(self.run_program)

    def open_about_minerva(self):
        about_Minerva = QtWidgets.QDialog()
        about_Minerva.setWindowTitle("About Minerva")
        about_Minerva.resize(280, 420)
        layout = QtWidgets.QVBoxLayout()
        label = QtWidgets.QLabel("Informações sobre Minerva")
        layout.addWidget(label)
        about_Minerva.setLayout(layout)
        about_Minerva.exec_()

    def open_about_measurement(self):
        about_measurement = QtWidgets.QDialog()
        about_measurement.setWindowTitle("About Measurement")
        about_measurement.resize(450, 420)
        layout = QtWidgets.QVBoxLayout()
        logo = QtWidgets.QLabel()
        logo.setPixmap(QtGui.QPixmap("logo-orion-vertical-1.png"))
        layout.addWidget(logo)
        separator = QtWidgets.QFrame()
        separator.setFrameShape(QtWidgets.QFrame.HLine)
        layout.addWidget(separator)
        label = QtWidgets.QLabel("Measurement is a program designed to Orion make it easy to take and store measurements."
                                 "With its user-friendly interface, you can perform accurate measurements with just a few clicks.\n\n"
                                 "Version: 1.3\n"
                                 "Author: Marcus Filgueiras and Gustavo Pessanha\n"
                                 "Copyright (c) 2023 Minerva Dev")
        label.setWordWrap(True)
        label.setFrameStyle(QtWidgets.QFrame.HLine | QtWidgets.QFrame.Sunken)
        layout.addWidget(label)
        about_measurement.setLayout(layout)
        about_measurement.exec_()

        
    def __init__(self):
        self.selected_folder = ""

    def select_folder(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        folder = QFileDialog.getExistingDirectory(Measurements, "Select folder", options=options)
        #Reset the text to "Select folder" if the user cancelled the file selection
        if folder == '':
            self.folder_select.setText("Select File Folder")
        else:
            self.folder_select.setText(folder)
            self.selected_folder = folder
    

    def run_program(self):
        if self.selected_folder == "":
            # Show an error message if no folder has been selected
            error_msg = "No folder selected. Please select a folder before running the program."
            msg = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning, "Error", error_msg)
            msg.setDetailedText("Maybe you didn't select the folder with the files and this error appeared on your screen, in this case it will not start the program, please go back and select the folder to run the program.")
            msg.exec_()
        else:
            # set qmovie as label
            self.movie = QMovie("Reload-1s-50px.gif")
            self.label.setMovie(self.movie)
            self.movie.start()
            # Run the program using the selected folder in a new thread
            program_instance = Program()
            program_instance = Program(selected_folder=self.selected_folder, movie=self.movie, label=self.label)
            t = threading.Thread(target=program_instance.programa)
            t.start()

    def close_program(self):
        #Close the main window and exit the application
        Measurements.close()
        sys.exit()

class Program:
    def __init__(self, selected_folder=None, movie=None, label=None):
        self.selected_folder = selected_folder
        self.count = 0
        self.movie = movie
        self.label = label

    def show_error_message(self, error_message):
        message = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Critical, "Error", error_message)
        message.exec_()


    def programa(self):
        if self.selected_folder:
            #Empty List
            filenames = []
            folder_names = []
            reference_filenames = []
            # Use glob to search for folders with names containing "pr5_files"
            folders = glob.glob(self.selected_folder +'/**/*set*', recursive=True)
            #only name file
            folders1 = [os.path.basename(folder) for folder in folders]
            
            #Extract the data from the pr5 files and transform it into txt with the data you need
            for folder in folders:
                folder_names.append(folder)
                for file in os.listdir(folder):
                    if file.endswith('.pr5'):
                        reference_filenames.append(file)
                        # Open PR5 file for reading
                        with open(os.path.join(folder, file), 'r') as f_in:
                          # Create a text file name based on the PR5 file name
                            txt_file = file[:-3] + 'txt'
                            filenames.append(txt_file)
                            # Open the text file for writing
                            with open(os.path.join(folder, txt_file), 'w') as f_out:
                              # Loop over the lines of the PR5 file
                              for line in f_in:
                                # Check if the current line starts with the string you want
                                if line.startswith('[Average Data]'):
                                 # Write all lines from here
                                  break
                              else:
                                # If we didn't find the string "[Average Data]", continue to the next iteration
                                continue
                              # Write all remaining lines to the text file
                              for line in f_in:
                                f_out.write(line.replace('.', ',')[:19]+ '\n')
                        
            ##########Pega todos os novos arquivos .txt e move ele para um pasta separada###########################
            # Loop over each folder containing the string "pr5_files"
            current_folder = os.path.basename(os.getcwd())
            for folder in folders:
                # Construa o caminho da pasta de destino
                dest_folder = os.path.join(folder, 'txt_files')
                # Crie a pasta de destino, se ela ainda não existir
                if not os.path.exists(dest_folder):
                    os.makedirs(dest_folder)
                # Loop over each file in the current folder
                for file in os.listdir(folder):
                    # Check if the file is a txt file
                    if file.endswith('.txt'):
                        # Build the source and destination paths for the file
                        src_path = os.path.join(folder, file)
                        dest_path = os.path.join(dest_folder, file)
                        # Move the file to the destination folder
                        shutil.move(src_path, dest_path)     
                    
            ####################### Create a new Excel workbook ########################################
            workbook = openpyxl.Workbook()
            arq  = os.path.basename(self.selected_folder)
            output_file = os.path.join(self.selected_folder, arq + ".xlsx")
            workbook.save(output_file)
            # Load the workbook
            workbook = openpyxl.load_workbook(output_file)
            # Get the first sheet
            sheet = workbook.worksheets[0]
            # Rename the sheet to the name of the first folder
            sheet.title = folders1[0]
            # Iterate over the remaining folders
            for folder in folders1[1:]:
                self.count += 1
                # Create a new sheet with the name of the current folder 
                if folder in workbook.sheetnames:
                    print(f'Sheet {folder} already exists')
                else:
                    # Create a new sheet
                    sheet = workbook.create_sheet(title=folder)      
            # Iterate over all sheets
            for worksheet in workbook.worksheets:
                # Set the width of the first eight columns to 130 px
                for i in 'ABCDEFGHIJK':
                    worksheet.column_dimensions[i].width = 18.59098065
                    # Get the cells in the first row
                    first_row_cells = worksheet[1]
            # Save the workbook
            workbook.save(output_file)
            ###################################################################
            # Set the path to the Excel file
            excel_file = output_file
            txt_folders = []
            for folder in folders:
                # Get the full path of the folder
                full_path = os.path.abspath(folder)
                # Get all the subfolders inside the current folder
                subfolders = os.listdir(full_path)
                # Iterate over the subfolders and check if their name ends with "txt_files"
                for subfolder in subfolders:
                    if subfolder.endswith("txt_files"):
                        # If it does, append the full path of the subfolder to txt_folders
                        txt_folders.append(os.path.join(full_path, subfolder))
                        # If it does, append the full path of the subfolder to txt_folders
                        full_path = os.path.join(full_path, subfolder)
                        if full_path not in txt_folders:
                            txt_folders.append(full_path)
            #################################################################################
                
            # Open the Excel workbook
            workbook = openpyxl.load_workbook(output_file)
            # Check if the subfolder ends with "txt_files"
            if subfolder.endswith("txt_files"):
            # If it does, append the full path of the subfolder to txt_folders
                txt_folders.append(os.path.join(full_path, subfolder))
                
            # Load the workbook
            # Set the path to the Excel file
            excel_file = output_file
            txt_folders = []
            for folder in folders:
                # Get the full path of the folder
                txt_folder = os.path.join(folder, 'txt_files')
                txt_folders.append(txt_folder)
            # Iterate over each folder containing txt files
            for txt_folder, sheet in zip(txt_folders, workbook.worksheets):
                # Iterate over each txt file in the folder
                for txt_file in os.listdir(txt_folder):
                    # Check if the file is a txt file
                    if txt_file.endswith('.txt'):
                        # Set the start_column variable to 1
                        # Build the full path of the txt file
                        txt_file_path = os.path.join(txt_folder, txt_file)
                        # Read the txt file into a Pandas DataFrame
                        df = pd.read_csv(txt_file_path, skiprows=6, sep='\s+', decimal=',', header=None)
                        # Crie um novo dataframe com os nomes das colunas
                        # Check if this is the first txt file in the folder
                        if txt_file == os.listdir(txt_folder)[0]:
                            # Select the first 2 columns of the DataFrame
                            df = df[df.columns[:2]]
                        else:
                            # Drop the first column of the DataFrame
                            df = df.drop(df.columns[0], axis=1) 
                        # Transpose the DataFrame
                        df = df.transpose()
                        # Convert the DataFrame to a list of rows
                        rows = dataframe_to_rows(df, index=False)
                        # Get the starting row and column for writing the data
                        # Set the column names
                        column_names = ['Deph (mm)', 'Orion Reference','Avg ID (mm)','Min ID (mm)', 'Max ID (mm)', 'OoR (mm)', 'Ova (%)', 'Min Radial (mm)','Pos of Min Radial (°)', 'Max Radial (mm)','Pos of Max Radial (°)']
                        # Iterate over the column names and set the values in the first row of the worksheet
                        for i, column_name in enumerate(column_names):   
                            sheet.cell(row=1, column=i+1).value = column_name
                        # Convert the DataFrame to a list of lists
                        rows = df.values.tolist()
                        # Iterate over the rows and append them to the worksheet
                        for row in rows:
                            sheet.append(row)
                            
            sheet = workbook.worksheets[0]
            cells = workbook.worksheets[0]
            
            for sheet in workbook.worksheets:
                #Move A2:WB para L2
                last_row_with_data = sheet.max_row
                sheet.move_range(f'A2:WB{last_row_with_data}', rows=-1, cols=11)
                #Formulas for excel
                for row in sheet.iter_rows(min_row=2):
                    #This code calculates the maximum and minimum sum of pairs of cells
                    max_sum = 0
                    min_sum = float('inf')
                    for i in range(12,312):
                    # selecionar a célula atual e a célula seguinte
                        cell1 = sheet.cell(row=row[0].row, column=i)
                        cell2 = sheet.cell(row=row[0].row, column=i+300)
                        # calcular a soma das células atuais
                        sum = cell1.value + cell2.value
                        # atualizar o máximo da soma se a soma atual for maior que o máximo atual
                        max_sum = max(max_sum, sum)
                        min_sum = min(min_sum, sum)
                        
                        
                    
                    #Set the formula
                    #Set the formula for the Avg ID value
                    row[2].value = f"=(sum(L{row[0].row}:WM{row[0].row})/300)"    
                    #Set the formula for the Min ID value
                    row[3].value = format(float(f"{min_sum}".replace(",", ".")), '.5f').replace(".", ",")
                    #Set the formula for the Max ID value
                    row[4].value = format(float(f"{max_sum}".replace(",", ".")), '.5f').replace(".", ",")
                    #Set the formula for the OoR value
                    row[5].value = f"=(E{row[0].row}-D{row[0].row})"
                    #Set the formula for the Ova value
                    row[6].value = f"=((F{row[0].row}/C{row[0].row})*100)"
                    # Set the formula for the maximum value
                    row[9].value = f"=MAX(L{row[0].row}:WM{row[0].row})"
                    # Set the formula for position degrees the maximum value
                    row[10].value=f"ÍNDICE($1:$1;CORRESP(J{row[0].row};L{row[0].row}:WM{row[0].row};0)+11)"
                    # Set the formula for the minimum value
                    row[7].value = f"=MIN(L{row[0].row}:WM{row[0].row})"
                    # Set the formula for position degrees the minimum value
                    row[8].value=f"ÍNDICE($1:$1;CORRESP(H{row[0].row};L{row[0].row}:WM{row[0].row};0)+11)"
            # initialize a counter variable to keep track of the current filename
            filename_counter = 0
            # iterate through the worksheets
            for sheet in workbook.worksheets:
                # get the number of rows in the sheet
                num_rows = sheet.max_row
                # iterate through the rows in the sheet
                for i in range(1,num_rows):
                    # only insert the filename if it exists in the list
                    if filename_counter < len(reference_filenames):
                        row = i + 1  # increment the row index by 1 to account for the header row
                        filename = reference_filenames[filename_counter]  # get the current filename
                        sheet.cell(row=row, column=2).value = filename  # set the value of the cell in the specified row and column
                        filename_counter += 1  # increment the counter
                    else:
                        # reset the counter to 0 if all of the filenames have been inserted
                        filename_counter = 0
            # Modify the font and fill properties of the cells in the first row of each worksheet in the workbook
            for sheet in workbook.worksheets:
                # Get the cells in the first row
                first_row_cells = sheet[1]
                # Iterate over the cells in the first row and set the font and fill
                for cell in first_row_cells:
                    cell.fill = openpyxl.styles.PatternFill(patternType='solid', fgColor='224C5A')
                    cell.font = Font(color='f08a04', bold=True)
                    # Set the alignment for all cells in the worksheet
                for row in sheet.rows:
                    for cell in row:
                        cell.alignment = Alignment(horizontal='center') 
            # Save the workbook
            workbook.save(output_file)
            # Load the workbook
            workbook = openpyxl.load_workbook(output_file)
            # Iterate over the sheet names
            for sheet_name in workbook.sheetnames:
                fig, axs = plt.subplots(1, 1, figsize=(10,10))
                sheet = workbook[sheet_name]
                last_row = sheet.max_row
                data = [row[12:] for row in sheet.iter_rows(min_row=2, max_row=last_row, values_only=True)]
                data = np.array(data)
                image = axs.imshow(data, cmap='viridis', aspect = 2)
                axs.set_title(sheet_name)
                cbar = plt.colorbar(image, ax=axs, shrink=0.2)
                figures_folder = os.path.join(self.selected_folder, "figures")
                if not os.path.exists(figures_folder):
                    os.makedirs(figures_folder)
                filename = os.path.join(figures_folder, f"{sheet_name}.png")
                if os.path.exists(filename):
                    os.remove(filename)
                plt.savefig(filename, bbox_inches='tight')
                       

        #Finish program        
        self.movie.stop()
        self.label.setGeometry(162, 140, 76, 10)
        self.label.setStyleSheet("color: white; font-size: 10pt; font-family: MS Shell Dlg 2; font-weight: bold")
        self.label.setText("COMPLETE")
                

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    Measurements = QtWidgets.QMainWindow()
    ui = Ui_Measurements()
    ui.setupUi(Measurements)
    Measurements.show()
    sys.exit(app.exec_())
