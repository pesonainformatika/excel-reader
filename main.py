import sys

import pandas as pd
from PyQt5.QtCore import QDir
from PyQt5.QtWidgets import QApplication, QWidget, QTableWidget, QTableWidgetItem, QPushButton, QVBoxLayout, QFileDialog


class MainWindow(QWidget):
    def __init__(self):
        super(MainWindow, self).__init__()

        # set window width
        self.window_width = 700
        self.window_height = 500

        # resize window
        self.resize(self.window_width, self.window_height)
        self.setWindowTitle("Excel Readers")

        # layout settings
        layout = QVBoxLayout()
        self.setLayout(layout)

        # Table Settings
        self.table = QTableWidget()
        layout.addWidget(self.table)

        # Button Config
        self.button = QPushButton("Load .xlsx or csv File")
        self.button.clicked.connect(self.load_excel_data)
        layout.addWidget(self.button)

    # create method to load excel file
    def load_excel_data(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.AnyFile)
        dialog.setFilter(QDir.Files)

        if dialog.exec_():
            filename = dialog.selectedFiles()
            if filename[0].endswith('.xlsx'):
                df = pd.read_excel(filename[0])
                if df.size == 0:
                    return
                else:
                    df.fillna('', inplace=True)
                    self.table.setRowCount(df.shape[0])
                    self.table.setColumnCount(df.shape[1])

                    # returns pandas array object
                    for row in df.iterrows():
                        values = row[1]
                        for col_index, value in enumerate(values):
                            if isinstance(value, (float, int)):
                                value = '{0:0,.0f}'.format(value)
                            tableitem = QTableWidgetItem(str(value))
                            self.table.setItem(row[0], col_index, tableitem)

                    self.table.setColumnWidth(2, 300)
            else:
                pass


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyleSheet('''
            QWidget {
                font-size: 17px;
            }
        ''')

    main_window = MainWindow()
    main_window.show()

    try:
        sys.exit(app.exec_())
    except SystemExit:
        print('Closing Window')
