import sys
import os
import fitz  # PyMuPDF
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QPushButton, QVBoxLayout, QHBoxLayout, QFileDialog, QTextEdit, QLabel, QScrollArea, QSizePolicy, QFrame, QProgressDialog, QMessageBox
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import Qt

from PyPDF2 import PdfReader, PdfWriter

class PdfToolApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        self.setGeometry(100, 100, 1000, 600)
        self.setWindowTitle('PDF拆分合并工具 by WTQ')

        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()

        self.browse_button = QPushButton('打开PDF', self.central_widget)
        self.browse_button.clicked.connect(self.browse_pdf)

        self.split_button = QPushButton('自定义拆分PDF', self.central_widget)
        self.split_button.clicked.connect(self.split_pdf)

        self.merge_button = QPushButton('合并PDF', self.central_widget)
        self.merge_button.clicked.connect(self.merge_pdfs)

        self.quick_split_button = QPushButton('等分拆分', self.central_widget)
        self.quick_split_button.clicked.connect(self.quick_split_pdf)

        self.quick_split_layout = QHBoxLayout()
        self.quick_split_label = QLabel('指定每部分页数:', self.central_widget)
        self.quick_split_input = QTextEdit(self.central_widget)
        self.quick_split_layout.addWidget(self.quick_split_label)
        self.quick_split_layout.addWidget(self.quick_split_input)

        self.page_range_label = QLabel('每组包含的页数范围（例如：1-5）:', self.central_widget)
        self.page_range_input = QTextEdit(self.central_widget)

        self.pdf_info_label = QLabel(self.central_widget)
        self.pdf_info_label.setAlignment(Qt.AlignLeft)

        self.preview_label = QLabel('预览区域:', self.central_widget)
        self.preview_scroll = QScrollArea(self.central_widget)
        self.preview_frame = QFrame(self.central_widget)
        self.preview_layout = QVBoxLayout(self.preview_frame)
        self.preview_layout.setAlignment(Qt.AlignTop)
        self.preview_scroll.setWidgetResizable(True)
        self.preview_scroll.setWidget(self.preview_frame)

        self.zoom_in_button = QPushButton('放大', self.central_widget)
        self.zoom_in_button.clicked.connect(self.zoom_in)

        self.zoom_out_button = QPushButton('缩小', self.central_widget)
        self.zoom_out_button.clicked.connect(self.zoom_out)

        self.zoom_controls_layout = QHBoxLayout()
        self.zoom_controls_layout.addWidget(self.zoom_in_button)
        self.zoom_controls_layout.addWidget(self.zoom_out_button)

        self.layout.addWidget(self.browse_button)
        self.layout.addWidget(self.page_range_label)
        self.layout.addWidget(self.page_range_input)
        self.layout.addWidget(self.pdf_info_label)
        self.layout.addWidget(self.preview_label)
        self.layout.addWidget(self.preview_scroll)
        self.layout.addLayout(self.zoom_controls_layout)
        self.layout.addWidget(self.split_button)
        self.layout.addWidget(self.quick_split_button)
        self.layout.addLayout(self.quick_split_layout)
        self.layout.addWidget(self.merge_button)

        self.central_widget.setLayout(self.layout)

        self.max_preview_height = 300  #设定预览区初始值大小
        self.zoom_factor = 1.0

    def browse_pdf(self):
        file_dialog = QFileDialog.getOpenFileName(self, '选择PDF文件', '', 'PDF Files (*.pdf)')
        input_file = file_dialog[0]

        if input_file:
            self.pdf_path = input_file
            self.clear_preview()
            self.load_pdf_preview()
            self.display_pdf_info()

    def clear_preview(self):
        for i in reversed(range(self.preview_layout.count())):
            widget = self.preview_layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()

    def show_progress_dialog(self, text):
        progress_dialog = QProgressDialog(text, "取消", 0, 0, self)
        progress_dialog.setWindowTitle("加载中")
        progress_dialog.setWindowModality(Qt.WindowModal)
        progress_dialog.show()
        return progress_dialog

    def load_pdf_preview(self):
        pdf_document = fitz.open(self.pdf_path)

        max_preview_width = self.preview_scroll.viewport().width() - 20

        progress_dialog = self.show_progress_dialog("正在加载PDF页面...")
        progress_dialog.setMaximum(pdf_document.page_count)

        for page_num in range(pdf_document.page_count):
            if progress_dialog.wasCanceled():
                break

            page = pdf_document.load_page(page_num)
            page_width = page.rect.width
            scale_factor = max_preview_width / page_width
            scaled_height = int(self.max_preview_height * self.zoom_factor)
            pixmap = QPixmap.fromImage(self.get_page_image(page, scale_factor, scaled_height))
            label = QLabel()
            label.setPixmap(pixmap)
            label.setAlignment(Qt.AlignCenter)

            page_label = QLabel(f"第 {page_num + 1} 页")
            page_label.setAlignment(Qt.AlignCenter)

            layout = QVBoxLayout()
            layout.addWidget(label)
            layout.addWidget(page_label)

            container_widget = QWidget()
            container_widget.setLayout(layout)

            self.preview_layout.addWidget(container_widget)

            progress_dialog.setValue(page_num + 1)
            QApplication.processEvents()

        pdf_document.close()
        progress_dialog.close()

    def get_page_image(self, page, scale_factor=1.0, height=None):
        if height is None:
            height = self.max_preview_height * self.zoom_factor
        img = page.get_pixmap(matrix=fitz.Matrix(4 * scale_factor, 4 * scale_factor))
        q_img = QImage(img.samples, img.width, img.height, img.stride, QImage.Format_RGB888)
        return q_img.scaledToHeight(height, Qt.SmoothTransformation)

    def zoom_in(self):
        self.zoom_factor *= 1.2
        self.clear_preview()
        self.load_pdf_preview()

    def zoom_out(self):
        self.zoom_factor /= 1.2
        self.clear_preview()
        self.load_pdf_preview()

    def display_pdf_info(self):
        pdf_reader = PdfReader(self.pdf_path)
        total_pages = len(pdf_reader.pages)
        file_size = os.path.getsize(self.pdf_path) / 1024.0 / 1024.0

        pdf_info = f"页数: {total_pages}\n文件大小: {file_size:.2f} MB\n路径: {self.pdf_path}"
        self.pdf_info_label.setText(pdf_info)

    def check_duplicates(self, page_ranges, total_pages):
        used_pages = set()
        duplicates = []
        for start_page, end_page in page_ranges:
            if start_page <= 0 or end_page > total_pages or start_page > end_page:
                continue
            for page_num in range(start_page, end_page + 1):
                if page_num in used_pages:
                    duplicates.append(str(page_num))
                used_pages.add(page_num)
        return duplicates

    def check_missing(self, page_ranges, total_pages):
        all_pages = set(range(1, total_pages + 1))
        used_pages = set()
        for start_page, end_page in page_ranges:
            used_pages.update(range(start_page, end_page + 1))
        return [str(page) for page in all_pages - used_pages]

    def split_pdf(self):
        if not hasattr(self, 'pdf_path'):
            return

        page_range_text = self.page_range_input.toPlainText()
        page_ranges = [list(map(int, r.strip().split('-'))) for r in page_range_text.splitlines() if r.strip()]

        pdf_reader = PdfReader(self.pdf_path)
        total_pages = len(pdf_reader.pages)

        if not page_ranges:
            reply = QMessageBox.warning(self, "警告", "请输入拆分组合范围！", QMessageBox.Ok)
            return

        selected_pages = [i for i in range(total_pages)]

        duplicates = self.check_duplicates(page_ranges, total_pages)
        missing_pages = self.check_missing(page_ranges, total_pages)

        if duplicates or missing_pages:
            msg = ""
            if duplicates:
                msg += f"拆分组合范围有重复页面：{', '.join(duplicates)}\n"
            if missing_pages:
                msg += f"拆分组合范围有遗漏页面：{', '.join(missing_pages)}\n"
            msg += "是否继续拆分？"
            reply = QMessageBox.warning(self, "警告", msg, QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return

        for start_page, end_page in page_ranges:
            if start_page <= 0 or end_page > total_pages or start_page > end_page:
                reply = QMessageBox.warning(self, "警告", f"拆分组合范围错误：{start_page}-{end_page}", QMessageBox.Ok)
                return

            pdf_writer = PdfWriter()
            for page_num in selected_pages[start_page - 1:end_page]:
                pdf_writer.add_page(pdf_reader.pages[page_num])

            output_file = f"split_pages_{start_page}-{end_page}.pdf"
            with open(output_file, "wb") as output_pdf:
                pdf_writer.write(output_pdf)

        QMessageBox.information(self, "提示", "拆分完成！", QMessageBox.Ok)

    def quick_split_pdf(self):
        if not hasattr(self, 'pdf_path'):
            return

        page_count = len(PdfReader(self.pdf_path).pages)
        try:
            pages_per_part = int(self.quick_split_input.toPlainText())
            if pages_per_part <= 0:
                raise ValueError()
        except ValueError:
            reply = QMessageBox.warning(self, "警告", "请输入有效的每部分页数！", QMessageBox.Ok)
            return

        total_parts = (page_count + pages_per_part - 1) // pages_per_part

        if total_parts == 1:
            reply = QMessageBox.warning(self, "警告", "无需拆分，每部分页数过大！", QMessageBox.Ok)
            return

        start_pages = [(i * pages_per_part) + 1 for i in range(total_parts)]
        end_pages = [(i + 1) * pages_per_part for i in range(total_parts - 1)]
        end_pages.append(page_count)

        page_ranges = list(zip(start_pages, end_pages))

        self.page_range_input.setPlainText('\n'.join([f"{start}-{end}" for start, end in page_ranges]))
        self.split_pdf()

    def merge_pdfs(self):
        file_dialog = QFileDialog.getOpenFileNames(self, '选择要合并的PDF文件', '', 'PDF Files (*.pdf)')
        input_files = file_dialog[0]

        if input_files:
            pdf_writer = PdfWriter()

            for input_file in input_files:
                pdf_reader = PdfReader(input_file)
                for page in pdf_reader.pages:
                    pdf_writer.add_page(page)

            output_file_dialog = QFileDialog.getSaveFileName(self, '保存合并后的PDF文件', '', 'PDF Files (*.pdf)')
            output_file = output_file_dialog[0]

            if output_file:
                with open(output_file, 'wb') as output_pdf:
                    pdf_writer.write(output_pdf)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    pdf_tool = PdfToolApp()
    pdf_tool.show()
    sys.exit(app.exec_())
