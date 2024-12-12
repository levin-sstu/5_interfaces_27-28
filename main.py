import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QToolBar, QAction, QFileDialog,
    QInputDialog, QColorDialog
)
from PyQt5.QtGui import QFont, QImage
from PyQt5.QtCore import Qt
from PyQt5.QtPrintSupport import QPrinter, QPrintPreviewDialog
from odf.opendocument import OpenDocumentText
from odf.text import P
from odf.opendocument import load
import os

os.environ['QT_PLUGIN_PATH'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.venv', 'Lib', 'site-packages',
                                            'PyQt5', 'Qt5', 'plugins')

class TextEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Текстовый редактор")
        self.setGeometry(100, 100, 800, 600)

        # Основной текстовый редактор
        self.text_edit = QTextEdit(self)
        self.setCentralWidget(self.text_edit)

        # Панель инструментов
        toolbar = QToolBar("Инструменты")
        self.addToolBar(toolbar)

        # Действия для панели инструментов
        bold_action = QAction("Жирный", self)
        bold_action.triggered.connect(self.set_bold)
        toolbar.addAction(bold_action)

        italic_action = QAction("Курсив", self)
        italic_action.triggered.connect(self.set_italic)
        toolbar.addAction(italic_action)

        underline_action = QAction("Подчеркнутый", self)
        underline_action.triggered.connect(self.set_underline)
        toolbar.addAction(underline_action)

        color_action = QAction("Цвет текста", self)
        color_action.triggered.connect(self.set_text_color)
        toolbar.addAction(color_action)

        align_left_action = QAction("Выровнять влево", self)
        align_left_action.triggered.connect(self.align_left)
        toolbar.addAction(align_left_action)

        align_center_action = QAction("Выровнять по центру", self)
        align_center_action.triggered.connect(self.align_center)
        toolbar.addAction(align_center_action)

        align_right_action = QAction("Выровнять вправо", self)
        align_right_action.triggered.connect(self.align_right)
        toolbar.addAction(align_right_action)

        insert_image_action = QAction("Вставить изображение", self)
        insert_image_action.triggered.connect(self.insert_image)
        toolbar.addAction(insert_image_action)

        insert_table_action = QAction("Вставить таблицу", self)
        insert_table_action.triggered.connect(self.insert_table)
        toolbar.addAction(insert_table_action)

        save_action = QAction("Сохранить как ODF", self)
        save_action.triggered.connect(self.save_odf)
        toolbar.addAction(save_action)

        load_action = QAction("Загрузить из ODF", self)
        load_action.triggered.connect(self.load_odf)
        toolbar.addAction(load_action)

        preview_action = QAction("Предварительный просмотр PDF", self)
        preview_action.triggered.connect(self.preview_pdf)
        toolbar.addAction(preview_action)

        export_action = QAction("Экспорт в PDF", self)
        export_action.triggered.connect(self.export_pdf)
        toolbar.addAction(export_action)

    def set_bold(self):
        self.text_edit.setFontWeight(QFont.Bold if self.text_edit.fontWeight() != QFont.Bold else QFont.Normal)

    def set_italic(self):
        state = self.text_edit.fontItalic()
        self.text_edit.setFontItalic(not state)

    def set_underline(self):
        state = self.text_edit.fontUnderline()
        self.text_edit.setFontUnderline(not state)

    def set_text_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.text_edit.setTextColor(color)

    def align_left(self):
        self.text_edit.setAlignment(Qt.AlignLeft)

    def align_center(self):
        self.text_edit.setAlignment(Qt.AlignCenter)

    def align_right(self):
        self.text_edit.setAlignment(Qt.AlignRight)

    def insert_image(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Выберите изображение", "", "Images (*.png *.xpm *.jpg *.bmp)")
        if file_name:
            cursor = self.text_edit.textCursor()
            image = QImage(file_name)
            cursor.insertImage(image)

    def insert_table(self):
        rows, ok = QInputDialog.getInt(self, "Введите количество строк", "Строки:", 2, 1, 10)
        if ok:
            cols, ok = QInputDialog.getInt(self, "Введите количество столбцов", "Столбцы:", 2, 1, 10)
            if ok:
                cursor = self.text_edit.textCursor()
                table = cursor.insertTable(rows, cols)
                for row in range(rows):
                    for col in range(cols):
                        cell = table.cellAt(row, col)
                        cell_cursor = cell.firstCursorPosition()
                        cell_cursor.insertText(f"Строка {row + 1}, Столбец {col + 1}")

    def save_odf(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "Сохранить как ODF", "", "ODF files (*.odt)")
        if file_name:
            doc = OpenDocumentText()
            text = self.text_edit.toPlainText()
            p = P(text=text)
            doc.text.addElement(p)
            doc.save(file_name)

    def load_odf(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Загрузить из ODF", "", "ODF files (*.odt)")
        if file_name:
            doc = load(file_name)  # Загружаем документ из файла
            text = ""
            for element in doc.getElementsByType(P):
                text += element.firstChild.data + "\n"
            self.text_edit.setPlainText(text)

    def preview_pdf(self):
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        preview = QPrintPreviewDialog(printer, self)
        preview.paintRequested.connect(self.print_document)
        preview.exec_()

    def export_pdf(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "Экспорт в PDF", "", "PDF files (*.pdf)")
        if file_name:
            printer = QPrinter(QPrinter.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(file_name)
            self.print_document(printer)

    def print_document(self, printer):
        document = self.text_edit.document()
        document.print_(printer)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    editor = TextEditor()
    editor.show()
    sys.exit(app.exec_())