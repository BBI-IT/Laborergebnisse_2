# -*- coding: utf-8 -*-
"""
Created on Mon Jul  3 12:09:04 2023

@author: zeh
"""

import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel
from PyQt5.QtGui import QPixmap
import subprocess

# Erstellen Sie eine Klasse für Ihr Hauptfenster
class MyWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # Erstellen Sie Widgets
        self.label = QLabel("Moin, ich helfe dir beim Erstellen deiner Anlage 4!", self)
        self.button = QPushButton("Erstellen!", self)
        
        # Konfigurieren Sie die Widget-Positionen und -Größen
        self.label.setGeometry(50, 50, 500, 30)
        self.button.setGeometry(50, 100, 100, 30)
        self.label.setStyleSheet("background-color: blue; color: white;")

        # Erstellen Sie ein QLabel-Widget für das Logo
        logo_label = QLabel(self)
        logo_label.setPixmap(QPixmap( r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\logo.png'))
        logo_label.setGeometry(self.width() - 0, 0, 600, 500)

        # Verknüpfen Sie ein Ereignis mit einer Methode
        self.button.clicked.connect(self.on_button_click)

        # Zeigen Sie das Hauptfenster an
        self.show()

    def on_button_click(self):
        # Definieren Sie den Pfad zu jedem Skript
        skript_pfad = [
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\1.Consolidate_Files.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\2.Formating_Cells.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\3.Transfered.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\4.EraseFreeRows.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\5.Anlagennummern.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\6.RenameBW.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\7.RenameBS.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\8.CopyUP.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\9.Attachementnumbers.py',
            r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\10.SortinginGeneralNoInput.py',
            #r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\11.TransferbackNEW.py',
            #r'C:\Users\zeh\Desktop\Laborergebnissse\Skript\12.Formating_final.py',
        ]

        # Führen Sie jedes Skript aus und überprüfen Sie den Rückgabewert
        for skript in skript_pfad:
            result = subprocess.run(['python', skript])
            if result.returncode == 0:
                self.label.setText(f"Skript {skript} wurde erfolgreich ausgeführt.")
            else:
                self.label.setText(f"Fehler beim Ausführen von Skript {skript}.")

# Erstellen Sie die Anwendung
app = QApplication(sys.argv)

# Erstellen Sie eine Instanz Ihres Hauptfensters
window = MyWindow()

# Starten Sie die Ereignisschleife der Anwendung
sys.exit(app.exec_())

