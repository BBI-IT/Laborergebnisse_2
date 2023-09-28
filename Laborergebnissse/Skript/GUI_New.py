# -*- coding: utf-8 -*-
"""
Created on Wed Jul  5 13:30:58 2023

@author: zeh
"""

import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QTextEdit
from PyQt5.QtGui import QPixmap
import subprocess

class MyWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Anlage 4")
        self.setGeometry(100, 100, 800, 600)

        self.label = QLabel("Moin, ich helfe dir beim Erstellen deiner Anlage 4!", self)
        self.button = QPushButton("Erstellen!", self)

        self.success_text = QTextEdit(self)
        self.success_text.setGeometry(50, 150, 700, 200)

        self.transfer_info_text = QTextEdit(self)
        self.transfer_info_text.setGeometry(50, 400, 700, 100)

        self.label.setGeometry(50, 50, 500, 30)
        self.button.setGeometry(50, 100, 100, 30)
        self.label.setStyleSheet("background-color: blue; color: white;")

        logo_label = QLabel(self)
        logo_label.setPixmap(QPixmap(r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\logo.png'))
        logo_label.setGeometry(self.width() - 0, 0, 600, 500)

        self.button.clicked.connect(self.on_button_click)

        self.show()

    def on_button_click(self):
        skript_pfad = [
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\1.Consolidate_Files.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\2.Formating_Cells.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\3.Transferred.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\4.EraseFreeRows.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\5.Anlagennummern.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\6.RenameBW.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\7.RenameBS.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\8.CopyUP.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\9.Attachementnumbers.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\10.SortinginGeneralNoInput.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\11.TransferbackNEW.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\12.Formating_final.py',
        ]

        success_message = ""
        transfer_info = ""

        for skript in skript_pfad:
            result = subprocess.run(['python', skript])
            if result.returncode == 0:
                success_message += f"Skript {skript} wurde erfolgreich ausgeführt.\n"
            else:
                success_message += f"Fehler beim Ausführen von Skript {skript}.\n"

        self.success_text.setPlainText(success_message)

        # Get information from TransferbackNEW.py
        try:
            with open(r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\TransferbackNEW_info.txt', 'r') as file:
                transfer_info = file.read()
        except FileNotFoundError:
            transfer_info = "TransferbackNEW_info.txt not found."

        self.transfer_info_text.setPlainText(transfer_info)

app = QApplication(sys.argv)
window = MyWindow()
sys.exit(app.exec_())
