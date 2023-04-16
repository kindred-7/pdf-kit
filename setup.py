import os.path
import sys

from PyQt5.QtWidgets import QApplication, QMainWindow, QHeaderView, QFileDialog
from PyQt5.QtCore import QRegExp
from PyQt5.QtGui import QRegExpValidator
from main import *
from funs import *


class MyWin(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.merge_path = ''

        self.my_regex = QRegExp("[0-9]+$")
        self.validator = QRegExpValidator()
        self.validator.setRegExp(self.my_regex)

        # 设置默认显示界面
        self.stackedWidget.setCurrentIndex(0)

        self.listWidget.clicked.connect(self.switch_fun)

        # 设置ListWidget样式
        self.listWidget.item(0).setSelected(True)
        self.listWidget.setSpacing(4)

        # 设备表格列宽
        header = self.tableWidget.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Stretch)

        # first stac-kwidget
        self.btn_odd.clicked.connect(lambda: self.get_file(self.sender()))
        self.btn_even.clicked.connect(lambda: self.get_file(self.sender()))
        self.btn_odd_merge.clicked.connect(self.run_odd_merge)

        # second stac-widget
        self.btn_reverse.clicked.connect(lambda: self.get_file(self.sender()))
        self.btn_reverse_1.clicked.connect(self.run_reverse)

        # extract-widget
        self.btn_extract.clicked.connect(lambda: self.get_file(self.sender()))
        self.btn_extract_1.clicked.connect(self.run_extract)
        self.extract_interval.setValidator(self.validator)
        self.extract_group.setValidator(self.validator)

        # rotate-widget
        self.btn_rotate.clicked.connect(self.get_rotate_file)
        self.btn_rotate_1.clicked.connect(self.run_rotate)
        self.radioButton_5.clicked.connect(self.get_pages)
        self.lineEdit_start.setValidator(self.validator)
        self.lineEdit_end.setValidator(self.validator)

        # merge_widget
        if self.tableWidget.rowCount() == 0:
            self.pushButton_10.setEnabled(False)
            self.pushButton_9.setEnabled(False)

        self.tableWidget.itemChanged.connect(self.update_button)

        self.btn_import.clicked.connect(self.open_files)
        self.tableWidget.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tableWidget.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tableWidget.horizontalHeader().setDefaultAlignment(Qt.AlignLeft)
        self.pushButton_9.clicked.connect(self.move_up)
        self.pushButton_10.clicked.connect(self.move_down)
        self.pushButton_11.clicked.connect(self.remove_item)
        # unlock
        self.btn_unlock.clicked.connect(self.unlock_files)

        self.btn_merge.clicked.connect(self.multi_merge)

        self.tableWidget.selectionModel().selectionChanged.connect(self.update_button)

        # deltet
        self.btn_dele.clicked.connect(lambda: self.get_file(self.sender()))
        self.btn_run_dele.clicked.connect(self.run_delete)
        self.delete_interval.setValidator(self.validator)
        self.delete_group.setValidator(self.validator)

        # replace
        self.btn_origin.clicked.connect(lambda: self.get_file(self.sender()))
        self.btn_place.clicked.connect(lambda: self.get_file(self.sender()))
        self.btn_run_replace.clicked.connect(self.run_replace)
        self.original_interval.setValidator(self.validator)
        self.original_group.setValidator(self.validator)
        self.place_interval.setValidator(self.validator)
        self.place_group.setValidator(self.validator)

        # insert
        self.btn_insert.clicked.connect(lambda: self.get_file(self.sender()))
        self.r_btn_page.clicked.connect(self.get_max)
        self.insert_pos.setValidator(self.validator)

        self.btn_insert_page.clicked.connect(lambda: self.get_file(self.sender()))

        self.btn_run_insert.clicked.connect(self.run_insert)

    def get_pages(self):
        if self.lineEdit_rotate.text():
            page = get_max_page(self.lineEdit_rotate.text())
            self.max_page.setText(str(page))

    def get_max(self):
        if self.lineEdit_insert.text():
            page = get_max_page(self.lineEdit_insert.text())
            self.insert_max_page.setText(str(page))

    def get_rotate_file(self):
        try:
            file, _ = QFileDialog.getOpenFileName(self, "选择", "../", "PDF文件(*.pdf)")
            if file:
                self.lineEdit_rotate.setText(file)
                max_page = get_max_page(self.lineEdit_rotate.text())
                self.max_page.setText(str(max_page))
        except (FileNotFoundError, ValueError) as e_:
            QMessageBox.warning(self, '警告', str(e_))

    def run_rotate(self):
        try:
            file = self.lineEdit_rotate.text()
            output_dir = make_dir(file)
            max_page = int(self.max_page.text())
            # 设定旋转角度
            if self.comboBox.currentIndex() == 0:
                angle = 90
            elif self.comboBox.currentIndex() == 1:
                angle = 270
            else:
                angle = 180
            # 设定奇偶页
            if self.comboBox_2.currentIndex() == 0:
                odd_mark = 0
            elif self.comboBox_2.currentIndex() == 1:
                odd_mark = 1
            else:
                odd_mark = 2

            if file:
                if self.radioButton_4.isChecked():
                    page_list = [n for n in range(0, max_page)]
                    rotate_page(file, page_list, odd_mark, angle, output_dir)
                if self.radioButton_5.isChecked():
                    start = int(self.lineEdit_start.text())
                    end = int(self.lineEdit_end.text())
                    if start <= end and 0 <= start <= max_page and 0 <= end <= max_page:
                        page_list = [n for n in range(start - 1, end)]
                        rotate_page(file, page_list, odd_mark, angle, output_dir)
                    elif start > end:
                        QMessageBox.warning(self, '警告', '起始页码不能大于截止页')
                    else:
                        QMessageBox.warning(self, '警告', '输入页码超范围')

                QMessageBox.information(self, '信息', '旋转完成！')
            else:
                QMessageBox.warning(self, '警告', '请输入完整路径')

        except Exception as _e:
            QMessageBox.warning(self, '警告', str(_e))

    def switch_fun(self, item):
        """
        点击列表切换功能窗口
        :param item: 指定列表项
        :return: none
        """
        r = item.row()
        self.stackedWidget.setCurrentIndex(r)

    def get_file(self, item):
        try:
            name = item.objectName().split('_')[1]
            file, _ = QFileDialog.getOpenFileName(self, "选择", "../", "PDF文件(*.pdf)")
            if file:
                eval(f"self.lineEdit_{name}.setText(eval(repr(file)))")
        except (FileNotFoundError, ValueError) as e_:
            QMessageBox.warning(self, '警告', str(e_))

    def open_files(self):
        self.merge_path, _ = QFileDialog.getOpenFileNames(self, '选择文件', '../',
                                                          "All Files(*);;PDF 文件(*.pdf);;PNG(*.png);;JPEG(*.jpeg *.jpg)"
                                                          ";;BMP(*.bmp)")
        if self.merge_path:
            try:
                icon = QIcon('icon/文件-pdf_file-pdf.ico')
                img_icon = QIcon('icon/照片_pic.ico')
                # files_name = os.listdir(self.merge_path)
                for file in self.merge_path:
                    row_count = self.tableWidget.rowCount()
                    self.tableWidget.insertRow(row_count)
                    if file.endswith(('pdf', 'PDF', 'Pdf', 'PDf', 'pDf', 'pdF', 'pDF')):
                        # file_path = file
                        pages = get_max_page(file)

                        file_item = QTableWidgetItem(icon, os.path.basename(file))
                        page_item = QTableWidgetItem(f'1-{pages}')
                        file_path_item = QTableWidgetItem(os.path.dirname(file))
                        self.tableWidget.setItem(row_count, 0, file_item)
                        self.tableWidget.setItem(row_count, 1, page_item)
                        self.tableWidget.setItem(row_count, 2, file_path_item)
                    elif file.endswith(
                            ('jpg', 'JPG', 'JpG', 'Jpg', 'JPg', 'jpG', 'jPg', 'jPG', 'jpeg', 'JPEG', 'PNG', 'png',
                             'Png', 'Jpeg', 'jPeg', 'pNg', 'pNG', 'bmp', 'BMP', 'Bmp', 'BMp', 'bmP')):

                        file_item = QTableWidgetItem(img_icon, os.path.basename(file))
                        file_path_item = QTableWidgetItem(os.path.dirname(file))
                        self.tableWidget.setItem(row_count, 0, file_item)
                        self.tableWidget.setItem(row_count, 1, QTableWidgetItem('/'))
                        self.tableWidget.setItem(row_count, 2, file_path_item)
                    else:
                        pass
            except Exception as ex:
                QMessageBox.warning(self, '警告', str(ex))

    def update_button(self):
        """向上/下按钮状态变更"""
        row = self.tableWidget.currentRow()
        rows = self.tableWidget.rowCount()
        if rows < 2:
            self.pushButton_10.setEnabled(False)
            self.pushButton_9.setEnabled(False)
        else:
            if row == 0:
                self.pushButton_10.setEnabled(True)
                self.pushButton_9.setEnabled(False)
            elif row == self.tableWidget.rowCount() - 1:
                self.pushButton_9.setEnabled(True)
                self.pushButton_10.setEnabled(False)
            elif row == -1:
                self.pushButton_9.setEnabled(False)
                self.pushButton_10.setEnabled(False)
            else:
                self.pushButton_9.setEnabled(True)
                self.pushButton_10.setEnabled(True)

    def move_up(self):
        """表格选中行向上移动"""
        row = self.tableWidget.currentRow()
        col = self.tableWidget.currentColumn()
        if row > 0:
            self.tableWidget.insertRow(row - 1)
            for i in range(self.tableWidget.columnCount()):
                self.tableWidget.setItem(row - 1, i, self.tableWidget.takeItem(row + 1, i))
            self.tableWidget.setCurrentCell(row - 1, col)
            self.tableWidget.removeRow(row + 1)

    def move_down(self):
        """表格选中行向下移动"""
        row = self.tableWidget.currentRow()
        col = self.tableWidget.currentColumn()
        if row < self.tableWidget.rowCount() - 1:
            self.tableWidget.insertRow(row + 2)
            for col in range(self.tableWidget.columnCount()):
                self.tableWidget.setItem(row + 2, col, self.tableWidget.takeItem(row, col))
            self.tableWidget.setCurrentCell(row + 2, col)
            self.tableWidget.removeRow(row)

    def remove_item(self):
        """移除表格行"""
        row = self.tableWidget.currentRow()
        if row != -1:
            self.tableWidget.removeRow(row)
        self.update_button()

    def multi_merge(self):
        try:
            files_list = []
            page_list = []
            rows = self.tableWidget.rowCount()
            for r in range(rows):
                file_path = os.path.join(self.tableWidget.item(r, 2).text(), self.tableWidget.item(r, 0).text())
                page_range = self.tableWidget.item(r, 1).text()
                files_list.append(file_path)
                page_list.append(page_range)

            out_path = make_dir(files_list[0])
            if files_list[0].endswith(('pdf', 'PDF', 'Pdf', 'PDf', 'pDf', 'pdF', 'pDF', 'PdF')):
                merge_pdfs(files_list, page_list, out_path)
            elif files_list[0].endswith(
                    ('jpg', 'JPG', 'JpG', 'Jpg', 'JPg', 'jpG', 'jPg', 'jPG', 'jpeg', 'JPEG', 'PNG', 'png',
                     'Png', 'Jpeg', 'jPeg', 'pNg', 'pNG', 'bmp', 'BMP', 'Bmp', 'BMp', 'bmP')):
                merge_imgs(files_list, out_path)
            else:
                QMessageBox.information(self, '信息', '不支持格式！')
            QMessageBox.information(self, '信息', '合并完成！')
        except Exception as exc:
            QMessageBox.warning(self, '警告', str(exc))

    def unlock_files(self):
        try:
            rows = self.tableWidget.rowCount()
            if rows:
                output_dir = self.tableWidget.item(0, 2).text()
                output_path = os.path.join(output_dir, 'unlock')
                if not os.path.exists(output_path):
                    os.mkdir(output_path)
                for r in range(rows):
                    file_path = os.path.join(self.tableWidget.item(r, 2).text(), self.tableWidget.item(r, 0).text())
                    if file_path.split('.')[1].lower() == 'pdf':
                        unlock_file(file_path, output_path)
                        QMessageBox.information(self, '信息', '解锁成功！')

                    else:
                        QMessageBox.warning(self, '信息', '不支持格式')

            else:
                QMessageBox.warning(self, '警告', '请导入PDF文件')

        except Exception as e:
            QMessageBox.warning(self, '警告', str(e))

    def run_odd_merge(self):
        """奇偶页合并"""
        try:
            odd_dir = self.lineEdit_odd.text()
            even_dir = self.lineEdit_even.text()
            output_dir = make_dir(odd_dir)

            if odd_dir and even_dir:
                odd_merge_even(odd_dir, output_dir, even_dir)
                QMessageBox.information(self, '信息', '合并完成！')
            else:
                QMessageBox.warning(self, '警告', '请输入完整路径')
        except Exception as a:
            QMessageBox.warning(self, '警告', str(a))

    def run_reverse(self):
        """页面倒序"""
        try:
            reverse_dir = self.lineEdit_reverse.text()
            output_dir = make_dir(reverse_dir)

            if reverse_dir:
                reverse_pdf(reverse_dir, output_dir)
                QMessageBox.information(self, '信息', '倒序完成！')
            else:
                QMessageBox.warning(self, '警告', '请输入完整路径')
        except Exception as a:
            QMessageBox.warning(self, '警告', str(a))

    def run_extract(self):
        """拆分/提取页面"""
        try:
            extract_dir = self.lineEdit_extract.text()
            output_dir = make_dir(extract_dir)

            if extract_dir:
                # 拆分
                if self.radioButton.isChecked():
                    num = int(self.spinBox.text())
                    split_pdf_by_page_number(extract_dir, output_dir, num)

                # 提取
                else:
                    page_range = self.lineEdit_5.text()
                    interval = int(self.extract_interval.text())
                    group = int(self.extract_group.text())
                    if page_range:
                        new_page = make_newpage(page_range, interval, group)
                        if self.radioButton_2.isChecked():
                            extract_pages(extract_dir, output_dir, new_page, 0)

                        if self.radioButton_3.isChecked():
                            extract_pages(extract_dir, output_dir, new_page, 1)
                    else:
                        QMessageBox.warning(self, '警告', '请输入完整路径')

                QMessageBox.information(self, '信息', '拆分/提取完成！')
            else:
                QMessageBox.warning(self, '警告', '请输入完整路径')
        except Exception as a:
            QMessageBox.warning(self, '警告', str(a))

    def run_delete(self):
        """执行删除页面"""
        try:
            file_dir = self.lineEdit_dele.text()
            page_range = self.lineEdit_2.text()
            max_page = get_max_page(file_dir)
            if file_dir and page_range:
                interval = int(self.delete_interval.text())
                group = int(self.delete_group.text())
                new_page = make_newpage(page_range, interval, group)
                page_list = make_page_list(new_page)

                if max(page_list) + 1 > max_page or min(page_list) < 0:
                    QMessageBox.warning(self, '警告', '页码超范围')

                else:
                    output_dir = make_dir(file_dir)
                    delete_pages(file_dir, page_list, output_dir)
                    QMessageBox.information(self, '信息', '删除成功！')

            else:
                QMessageBox.warning(self, '警告', '请输入完整信息')

        except Exception as e:
            QMessageBox.warning(self, '警告', str(e))

    def run_replace(self):
        """执行替换页"""
        try:
            origin_file = self.lineEdit_origin.text()
            origin_range = self.origin_range.text()
            place_file = self.lineEdit_place.text()
            place_range = self.place_range.text()

            if origin_file and origin_range and place_range and place_file:
                origin_max_page = get_max_page(origin_file)
                o_interval = int(self.original_interval.text())
                o_group = int(self.original_group.text())
                origin_new_page = make_newpage(origin_range, o_interval, o_group)
                origin_list = make_page_list(origin_new_page)

                place_max_page = get_max_page(place_file)
                p_interval = int(self.place_interval.text())
                p_group = int(self.place_group.text())
                place_new_page = make_newpage(place_range, p_interval, p_group)
                place_list = make_page_list(place_new_page)
                print(place_list)
                # 两文件页面范围数量需一致
                if len(origin_list) != len(place_list):
                    QMessageBox.warning(self, '警告', '页面数量不一致')
                else:
                    if max(origin_list) + 1 > origin_max_page or min(origin_list) < 0:
                        QMessageBox.warning(self, '警告', '页码超范围')
                    elif max(place_list) + 1 > place_max_page or min(place_list) < 0:
                        QMessageBox.warning(self, '警告', '页码超范围')
                    else:
                        out_dir = make_dir(origin_file)
                        replace_pages(origin_file, place_file, origin_list, place_list, out_dir)
                        QMessageBox.information(self, '信息', '替换页面完成！')

            else:
                QMessageBox.warning(self, '警告', '请输入完整信息')

        except Exception as e:
            QMessageBox.warning(self, '警告', str(e))

    def run_insert(self):
        try:
            original_file = self.lineEdit_insert.text()
            insert_file = self.lineEdit_insert_page.text()
            output_path = make_dir(original_file)

            origin_max_page = get_max_page(original_file)
            insert_max_page = get_max_page(insert_file)

            page_range = self.insert_page.text()
            interval = self.insert_interval.text()
            group = self.insert_group.text()
            pos = 0
            if original_file and insert_file and page_range and interval and group:
                insert_new_page = make_newpage(page_range, int(interval), int(group))
                page_list = make_page_list(insert_new_page)

                if max(page_list) + 1 > insert_max_page or min(page_list) < 0:
                    QMessageBox.warning(self, '警告', '页码超范围')
                else:
                    if self.r_btn_first.isChecked():
                        pos = 1
                    if self.r_btn_second.isChecked():
                        pos = origin_max_page
                    if self.r_btn_page.isChecked():
                        if self.insert_pos:
                            pos = int(self.insert_pos.text())
                            if pos > origin_max_page or pos < 1:
                                pos = -1
                        else:
                            QMessageBox.warning(self, '警告', "请输入完整信息")
                    if pos == -1:
                        QMessageBox.warning(self, '警告', '页码超范围')
                    else:
                        if self.insert_com.currentIndex() == 0:
                            pos = pos
                        elif self.insert_com.currentIndex() == 1:
                            pos = pos - 1
                        insert_page(original_file, insert_file, page_list, output_path, pos)
                        QMessageBox.information(self, '信息', '插入页面完成！')

            else:
                QMessageBox.warning(self, '警告', '请输入完整信息')

        except Exception as e:
            QMessageBox.warning(self, '警告', str(e))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = MyWin()
    win.show()
    sys.exit(app.exec_())
