import sys
import os
import threading

from PyQt5 import QtWidgets
from PyQt5 import QtCore
import xlsxwriter
import sqlite3


class Communicate(QtCore.QObject):

    """
    Communication between
    GUI and
    Converter component

    """

    done = QtCore.pyqtSignal()
    wait = QtCore.pyqtSignal()
    update_prog = QtCore.pyqtSignal(int)


class ConverterWindow(QtWidgets.QWidget):

    """

    Gui component

    """

    def __init__(self):
        super().__init__()

        self.file_path = None
        self.save_path = None
        self.conv = None

        self.communicate = Communicate()
        self.communicate.wait.connect(self.wait)
        self.communicate.done.connect(self.done)
        self.communicate.update_prog.connect(self.update_prog)

        self.prog = QtWidgets.QProgressBar(self)
        self.logs = QtWidgets.QTextEdit(self)
        self.logs.setReadOnly(True)

        self.layout = QtWidgets.QGridLayout(self)
        self.layout.addWidget(self.prog)
        self.layout.addWidget(self.logs)

        self.init()

    def init(self):

        self.set_file_path()
        self.set_save_path()
        self.init_UI()
        self.convert_data()

    def init_UI(self):

        self.prog.setValue(0)

        self.logs.append(
            "Selected file is '{file_path:}'"
            "\nExcel file is saved to '{save_path:}'".format(
                file_path=self.file_path,
                save_path=self.save_path
            )
        )

        self.setWindowTitle("SQLite to excel converter")
        self.show()

    def set_save_path(self):

        # get file extension
        check_extension = self.file_path.split(".")[-1]

        extension = \
            check_extension if isinstance(check_extension, list) else None

        if extension in ["db", "sqlite", "sqlite3", "db3"]:

            idx = len(extension)
            self.save_path = "{}xlsx".format(self.file_path[:-idx])

        else:

            self.save_path = "{}.xlsx".format(self.file_path)

    def set_file_path(self):

        directory = os.getenv("HOME")

        file_path = QtWidgets.QFileDialog.getOpenFileName(
            self,
            'Open file',
            directory,
            "SQLite files  (*.db *.sqlite *.sqlite3 *.db3)"
        )[0]

        if file_path:
            self.file_path = file_path

        else:
            self.show_error()

    def convert_data(self):

        self.conv = Converter(
            db_file=self.file_path,
            save_path=self.save_path,
            ui=self.communicate
        )

        self.conv.start()

    def done(self):

        self.prog.setValue(100)
        self.logs.append("Done!")

    def wait(self):

        self.logs.append("Saving excel file, please wait...")

    def update_prog(self, x):

        self.prog.setValue(x)

    def closeEvent(self, event):

        self.conv.stop()
        super().closeEvent(event)

    @staticmethod
    def show_error():

        msgbox = QtWidgets.QMessageBox()
        msgbox.setIcon(QtWidgets.QMessageBox.Critical)
        msgbox.setWindowTitle("Error")
        msgbox.setText("Selecting a file is required to proceed.")
        close = msgbox.addButton("Close", QtWidgets.QMessageBox.ActionRole)

        msgbox.exec_()

        if msgbox.clickedButton() == close:
            sys.exit()


class Converter(threading.Thread):

    """

    converts
    sqlite to xlsx

    """

    def __init__(self, db_file, save_path, ui):
        super().__init__()

        self.db_file = db_file
        self.save_path = save_path
        self.ui = ui
        self._stopped = False

    def run(self):

        """
        main method

        """

        # Create a workbook
        workbook = xlsxwriter.Workbook(self.save_path)

        # Some data we want to write to the worksheet.
        conn = self.create_connection(db_file=self.db_file)

        # get all table names
        cur = conn.cursor()
        data = cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table';"
        )
        names = [item for sublist in data for item in sublist]

        # remove unwanted tables
        cleaned_names = \
            [item for item in names if not item.startswith("sqlite")]

        for i, name in enumerate(cleaned_names):

            table = self.select_table(conn, name)

            workbook = self.write_table_to_workbook(
                workbook=workbook,
                table=table,
                worksheet_name=name
            )

            # update ui progress bar
            self.ui.update_prog.emit(i / len(cleaned_names) * 100)

            if self.stopped():
                sys.exit()

        # save and close
        self.ui.wait.emit()
        workbook.close()
        self.ui.done.emit()

    def stop(self):
        self._stopped = True

    def stopped(self):
        return self._stopped

    @staticmethod
    def create_connection(db_file):

        """
        create a database connection to the SQLite database
        specified by the db_file
        :param db_file: database file
        :return: Connection object or None
        """

        try:
            conn = sqlite3.connect(db_file)
            return conn
        except sqlite3.Error as e:
            print(e)

        return None

    @staticmethod
    def select_table(conn, table):

        """
        Query all rows in the table
        :param conn: the Connection object
        :param table: the content of the table
        :return: rows of the table
        """

        cur = conn.cursor()

        cur.execute("SELECT * FROM '{}'".format(table))

        rows = cur.fetchall()

        col_names = [description[0] for description in cur.description]

        rows.insert(0, col_names)

        return rows

    @staticmethod
    def write_table_to_workbook(workbook, table, worksheet_name):

        """
        Query all rows in the table
        :param conn: the Connection object
        :param table: the content of the table
        :param workbook: the workbook (excel file)
        :param worksheet_name: the worksheet name
        :return: rows of the table
        """

        worksheet = workbook.add_worksheet(worksheet_name)

        rows = table
        cols = range(len(table[0]))

        # Iterate over the data and write it out row by row.
        for row_number, row in enumerate(rows):
            for col_number in cols:
                worksheet.write(row_number, col_number, row[col_number])

        return workbook


if __name__ == '__main__':

    app = QtWidgets.QApplication(sys.argv)
    win = ConverterWindow()
    win.setGeometry(100, 200, 400, 300)
    sys.exit(app.exec_())
