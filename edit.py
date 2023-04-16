import os

from PyQt5.QtWidgets import QLineEdit, QTableWidget, QTableWidgetItem, QAbstractItemView, QMessageBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon

from funs import get_max_page


class MyTable(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.viewport().setAcceptDrops(True)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setDragDropMode(QAbstractItemView.InternalMove)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
            for url in event.mimeData().urls():
                file_path = str(url.toLocalFile())
                if file_path.endswith(('pdf', 'PDF', 'Pdf', 'PDf', 'pDf', 'pdF', 'pDF')):
                    pages = get_max_page(file_path)
                    row_position = self.rowCount()
                    self.insertRow(row_position)
                    icon = QIcon('icon/文件-pdf_file-pdf.ico')
                    self.setItem(row_position, 0, QTableWidgetItem(icon, os.path.basename(file_path)))
                    self.setItem(row_position, 1, QTableWidgetItem(f'1-{pages}'))
                    self.setItem(row_position, 2, QTableWidgetItem(os.path.dirname(file_path)))
                elif file_path.endswith(
                        ('jpg', 'JPG', 'JpG', 'Jpg', 'JPg', 'jpG', 'jPg', 'jPG', 'jpeg', 'JPEG', 'PNG', 'png',
                         'Png', 'Jpeg', 'jPeg', 'pNg', 'pNG', 'bmp', 'BMP', 'Bmp', 'BMp', 'bmP')):
                    row_position = self.rowCount()
                    self.insertRow(row_position)
                    icon = QIcon('icon/照片_pic.ico')
                    self.setItem(row_position, 0, QTableWidgetItem(icon, os.path.basename(file_path)))
                    self.setItem(row_position, 1, QTableWidgetItem('/'))
                    self.setItem(row_position, 2, QTableWidgetItem(os.path.dirname(file_path)))
                else:
                    QMessageBox.warning(self, '警告', '含不支持格式！')
        else:
            event.ignore()


class MyLineEdit(QLineEdit):
    def __init__(self, parent):
        super(MyLineEdit, self).__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):
        if e.mimeData().hasText():
            e.accept()
        else:
            e.ignore()

    def dropEvent(self, e):
        try:
            file_paths = e.mimeData().text()
            file_path = file_paths.split('\n')[0]
            path = file_path.replace('file:///', '', 1)
            self.setText(path)
        except Exception as f:
            print(str(f))
