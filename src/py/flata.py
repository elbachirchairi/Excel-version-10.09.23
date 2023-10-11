from PyQt5.QtWidgets import *
from PyQt5 import uic
import pandas as pd
import openpyxl


class flata(QMainWindow):
    def __init__(self, parent):
        super(flata, self).__init__(parent)
        self.parent = parent
        uic.loadUi("src/ui/Flata.ui", self)

        