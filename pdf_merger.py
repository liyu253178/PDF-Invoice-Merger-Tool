import sys
import logging
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QListWidget, QLabel, QFileDialog, QSpinBox,
                             QComboBox, QMessageBox, QScrollArea, QProgressBar)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QPainter
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape

import os
import tempfile
import fitz
import io
import traceback

# 配置日志记录
log_file = 'pdf_merger_error.log'
logging.basicConfig(
    filename=log_file,
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

class PDFMerger(QMainWindow):
    def __init__(self):
        super().__init__()
        self.files = []
        self.preview_label = None
        self.progress_bar = None
        self.initUI()

    def log_error(self, error_msg, exc_info=None):
        """记录错误信息到日志文件"""
        if exc_info:
            logging.error(f"{error_msg}\n{traceback.format_exc()}")
        else:
            logging.error(error_msg)

    def initUI(self):
        self.setWindowTitle('哲宇一号发票合并助手')
        self.setGeometry(100, 100, 1200, 800)

        # 创建主窗口部件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QHBoxLayout(central_widget)

        # 左侧布局（文件列表和按钮）
        left_layout = QVBoxLayout()
        
        # 文件列表
        self.file_list = QListWidget()
        left_layout.addWidget(QLabel('已选择的文件：'))
        left_layout.addWidget(self.file_list)

        # 按钮布局
        button_layout = QHBoxLayout()
        add_button = QPushButton('添加文件')
        remove_button = QPushButton('移除文件')
        remove_all_button = QPushButton('移除全部文件')
        button_layout.addWidget(add_button)
        button_layout.addWidget(remove_button)
        button_layout.addWidget(remove_all_button)
        left_layout.addLayout(button_layout)

        # 中间布局（设置选项）
        middle_layout = QVBoxLayout()
        
        # 页面方向设置
        middle_layout.addWidget(QLabel('页面方向：'))
        self.orientation = QComboBox()
        self.orientation.addItems(['纵向', '横向'])
        middle_layout.addWidget(self.orientation)

        # 每页排列设置
        middle_layout.addWidget(QLabel('每页文件数：'))
        layout_options = QHBoxLayout()
        self.rows = QSpinBox()
        self.cols = QSpinBox()
        self.rows.setMinimum(1)
        self.cols.setMinimum(1)
        self.rows.setValue(3)
        self.cols.setValue(2)
        layout_options.addWidget(QLabel('行数：'))
        layout_options.addWidget(self.rows)
        layout_options.addWidget(QLabel('列数：'))
        layout_options.addWidget(self.cols)
        middle_layout.addLayout(layout_options)

        # 合并按钮
        merge_button = QPushButton('合并文件')
        middle_layout.addWidget(merge_button)

        # 右侧布局（预览区域）
        right_layout = QVBoxLayout()
        right_layout.addWidget(QLabel('预览：'))
        
        # 创建预览区域
        preview_scroll = QScrollArea()

        # 添加进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setTextVisible(True)
        right_layout.addWidget(self.progress_bar)
        preview_scroll.setWidgetResizable(True)
        preview_container = QWidget()
        preview_layout = QVBoxLayout(preview_container)
        self.preview_label = QLabel()
        self.preview_label.setAlignment(Qt.AlignCenter)
        preview_layout.addWidget(self.preview_label)
        preview_scroll.setWidget(preview_container)
        right_layout.addWidget(preview_scroll)
        
        # 添加所有布局到主布局
        layout.addLayout(left_layout, 2)
        layout.addLayout(middle_layout, 1)
        layout.addLayout(right_layout, 3)

        # 连接信号和槽
        add_button.clicked.connect(self.add_files)
        remove_button.clicked.connect(self.remove_files)
        remove_all_button.clicked.connect(self.remove_all_files)
        merge_button.clicked.connect(self.merge_files)
        self.orientation.currentIndexChanged.connect(self.update_preview)
        self.rows.valueChanged.connect(self.update_preview)
        self.cols.valueChanged.connect(self.update_preview)

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择文件",
            "",
            "支持的文件 (*.pdf *.jpg *.jpeg *.png *.tif *.bmp)"
        )
        for file in files:
            if file not in self.files:
                self.files.append(file)
                self.file_list.addItem(os.path.basename(file))
        if files:
            self.update_preview()
        self.update_progress_bar()

    def remove_files(self):
        for item in self.file_list.selectedItems():
            idx = self.file_list.row(item)
            self.file_list.takeItem(idx)
            self.files.pop(idx)
        self.update_preview()
        self.update_progress_bar()

    def remove_all_files(self):
        self.files.clear()
        self.file_list.clear()
        self.update_preview()
        self.update_progress_bar()

    def update_progress_bar(self):
        total_files = len(self.files)
        if total_files == 0:
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("等待添加文件")
        else:
            self.progress_bar.setValue(100)
            self.progress_bar.setFormat(f"已选择 {total_files} 个文件")

    def convert_image_to_pdf(self, image_path):
        try:
            # 创建临时PDF文件
            temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            temp_pdf.close()
            
            # 打开并转换图片
            with Image.open(image_path) as img:
                # 转换为RGB模式
                if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                    background = Image.new('RGB', img.size, (255, 255, 255))
                    if img.mode == 'P':
                        img = img.convert('RGBA')
                    background.paste(img, mask=img.split()[3] if img.mode == 'RGBA' else None)
                    img = background
                elif img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # 保存为PDF
                img.save(temp_pdf.name, 'PDF', resolution=300.0)
            
            return temp_pdf.name
        except Exception as e:
            error_msg = f'图片转换失败（{os.path.basename(image_path)}）：{str(e)}'
            self.log_error(error_msg, exc_info=True)
            raise Exception(error_msg)

    def process_pdf_page(self, file_path):
        try:
            # 创建临时PDF文件
            temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            temp_pdf.close()
            
            # 读取原PDF文件
            reader = PdfReader(file_path)
            if len(reader.pages) > 0:
                # 创建新的PDF并添加第一页
                writer = PdfWriter()
                writer.add_page(reader.pages[0])
                with open(temp_pdf.name, 'wb') as f:
                    writer.write(f)
                return temp_pdf.name
            else:
                raise Exception('PDF文件为空')
        except Exception as e:
            error_msg = f'PDF处理失败（{os.path.basename(file_path)}）：{str(e)}'
            self.log_error(error_msg, exc_info=True)
            raise Exception(error_msg)

    def update_preview(self):
        if not self.files:
            self.preview_label.clear()
            return

        try:
            # 创建临时PDF用于预览
            temp_preview = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            temp_preview.close()
            temp_pdfs = []

            # 处理所有文件
            processed_files = []
            total_files = len(self.files)
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("正在生成预览...")
            QApplication.processEvents()

            for i, file in enumerate(self.files):
                try:
                    if file.lower().endswith('.pdf'):
                        temp_pdf = self.process_pdf_page(file)
                    else:
                        temp_pdf = self.convert_image_to_pdf(file)
                    temp_pdfs.append(temp_pdf)
                    processed_files.append(temp_pdf)
                    
                    # 更新进度条
                    progress_value = int(((i + 1) / total_files) * 50) # 前50%用于文件处理
                    self.progress_bar.setValue(progress_value)
                    self.progress_bar.setFormat(f'正在处理文件: %p% - {i + 1}/{total_files} 文件')
                    QApplication.processEvents()

                except Exception as e:
                    error_msg = f'处理文件时出错：{str(e)}'
                    self.log_error(error_msg)
                    QMessageBox.warning(self, '警告', error_msg)
                    continue

            if not processed_files:
                self.progress_bar.setValue(0)
                self.progress_bar.setFormat("预览生成失败")
                return

            # 创建预览PDF
            page_size = A4
            if self.orientation.currentText() == '横向':
                page_size = landscape(A4)

            c = canvas.Canvas(temp_preview.name, pagesize=page_size)
            page_width, page_height = page_size
            rows = self.rows.value()
            cols = self.cols.value()
            cell_width = page_width / cols
            cell_height = page_height / rows

            for i, file_path in enumerate(processed_files):
                if i >= rows * cols:
                    break
                
                # 更新进度条
                progress_value = int(50 + ((i + 1) / (rows * cols)) * 50) # 后50%用于渲染
                self.progress_bar.setValue(progress_value)
                self.progress_bar.setFormat(f'正在渲染预览: %p% - {i + 1}/{rows * cols} 页面')
                QApplication.processEvents()

                try:
                    row = i // cols
                    col = i % cols
                    x = col * cell_width
                    y = page_height - (row + 1) * cell_height

                    # 计算缩放和居中
                    doc = fitz.open(file_path)
                    if doc.page_count > 0:
                        page = doc[0]
                        pdf_width = page.rect.width
                        pdf_height = page.rect.height
                        
                        # 计算最佳缩放比例，保持纵横比
                        scale_x = (cell_width * 0.95) / pdf_width  # 留出5%边距
                        scale_y = (cell_height * 0.95) / pdf_height
                        scale = min(scale_x, scale_y)
                        
                        # 计算居中位置
                        scaled_width = pdf_width * scale
                        scaled_height = pdf_height * scale
                        centered_x = x + (cell_width - scaled_width) / 2
                        centered_y = y + (cell_height - scaled_height) / 2

                        # 将PDF页面转换为图像后绘制到画布上
                        # 提高DPI以改善图像质量
                        dpi = 1200  # 设置DPI为1200，提高图像质量
                        pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
                        img_data = pix.tobytes("png")
                        
                        # 创建临时图像文件
                        temp_img = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
                        temp_img.write(img_data)
                        temp_img.close()
                        
                        c.saveState()
                        c.translate(centered_x, centered_y)
                        c.scale(scale, scale)
                        c.drawImage(temp_img.name, 0, 0, width=pdf_width, height=pdf_height)
                        c.restoreState()
                        
                        # 删除临时图像文件
                        try:
                            os.unlink(temp_img.name)
                        except:
                            pass
                    doc.close()
                except Exception as e:
                    error_msg = f'预览文件时出错（{os.path.basename(file_path)}）：{str(e)}'
                    self.log_error(error_msg, exc_info=True)
                    QMessageBox.warning(self, '警告', error_msg)
                    continue

            c.save()

            # 使用PyMuPDF将PDF转换为图像进行预览
            try:
                doc = fitz.open(temp_preview.name)
                page = doc[0]
                pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))  # 缩放为50%以适应显示
                img_data = pix.tobytes("ppm")
                qimg = QPixmap()
                qimg.loadFromData(img_data)
                
                # 设置预览图像
                self.preview_label.setPixmap(qimg)
                doc.close()
            except Exception as e:
                error_msg = f'生成预览图像时出错：{str(e)}'
                self.log_error(error_msg, exc_info=True)
                QMessageBox.warning(self, '警告', error_msg)

        except Exception as e:
            error_msg = f'预览生成失败：{str(e)}'
            self.log_error(error_msg, exc_info=True)
            QMessageBox.warning(self, '警告', error_msg)
            self.preview_label.clear()
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("预览生成失败")
            QApplication.processEvents()

        finally:
            # 清理临时文件
            for temp_pdf in temp_pdfs:
                try:
                    os.unlink(temp_pdf)
                except:
                    pass
            try:
                os.unlink(temp_preview.name)
            except:
                pass
            self.progress_bar.setValue(100)
            self.progress_bar.setFormat("预览完成")
            QApplication.processEvents()

    def merge_files(self):
        if not self.files:
            QMessageBox.warning(self, '警告', '请先添加文件！')
            return

        output_file, _ = QFileDialog.getSaveFileName(
            self,
            "保存合并后的PDF",
            "",
            "PDF文件 (*.pdf)"
        )

        if not output_file:
            return

        try:
            # 创建临时PDF文件列表
            temp_pdfs = []
            processed_files = []

            # 处理所有输入文件
            total_files_to_process = len(self.files)
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("正在处理文件...")
            QApplication.processEvents()

            for i, file in enumerate(self.files):
                try:
                    if file.lower().endswith('.pdf'):
                        temp_pdf = self.process_pdf_page(file)
                    else:
                        temp_pdf = self.convert_image_to_pdf(file)
                    temp_pdfs.append(temp_pdf)
                    processed_files.append(temp_pdf)
                    
                    # 更新进度条
                    progress_value = int(((i + 1) / total_files_to_process) * 50) # 前50%用于文件处理
                    self.progress_bar.setValue(progress_value)
                    self.progress_bar.setFormat(f'正在处理文件: %p% - {i + 1}/{total_files_to_process} 文件')
                    QApplication.processEvents()

                except Exception as e:
                    error_msg = f'处理文件时出错：{str(e)}'
                    self.log_error(error_msg)
                    QMessageBox.warning(self, '警告', error_msg)
                    continue

            if not processed_files:
                QMessageBox.warning(self, '警告', '没有可处理的文件！')
                self.progress_bar.setValue(0)
                self.progress_bar.setFormat("合并失败")
                return

            # 设置页面大小和方向
            page_size = A4
            if self.orientation.currentText() == '横向':
                page_size = landscape(A4)

            # 创建输出PDF写入器
            output_writer = PdfWriter()
            
            # 计算每页的网格大小
            rows = self.rows.value()
            cols = self.cols.value()
            files_per_page = rows * cols
            
            # 按页处理文件
            for i in range(0, len(processed_files), files_per_page):
                page_files = processed_files[i:i + files_per_page]
                
                # 创建新的空白页面
                temp_page = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
                temp_page.close()
                c = canvas.Canvas(temp_page.name, pagesize=page_size)
                page_width, page_height = page_size
                
                # 计算每个文件的显示区域大小
                cell_width = page_width / cols
                cell_height = page_height / rows

                # 在页面上排列文件
                for j, pdf_file in enumerate(page_files):
                    try:
                        row = j // cols
                        col = j % cols
                        
                        # 计算当前单元格的位置
                        x = col * cell_width
                        y = page_height - (row + 1) * cell_height

                        # 使用PyMuPDF读取PDF
                        doc = fitz.open(pdf_file)
                        if doc.page_count > 0:
                            page = doc[0]
                            pdf_width = page.rect.width
                            pdf_height = page.rect.height
                            
                            # 计算最佳缩放比例，保持纵横比
                            scale_x = (cell_width * 0.95) / pdf_width  # 留出5%边距
                            scale_y = (cell_height * 0.95) / pdf_height
                            scale = min(scale_x, scale_y)
                            
                            # 计算居中位置
                            scaled_width = pdf_width * scale
                            scaled_height = pdf_height * scale
                            centered_x = x + (cell_width - scaled_width) / 2
                            centered_y = y + (cell_height - scaled_height) / 2

                            # 将PDF页面转换为图像后绘制到画布上
                            # 提高DPI以改善图像质量
                            dpi = 1200  # 设置DPI为1200，提高图像质量
                            pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
                            img_data = pix.tobytes("png")
                            
                            # 创建临时图像文件
                            temp_img = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
                            temp_img.write(img_data)
                            temp_img.close()

                        # 更新进度条
                        current_file_index = i + j
                        progress_value = int(50 + ((current_file_index + 1) / len(processed_files)) * 50) # 后50%用于合并
                        self.progress_bar.setValue(progress_value)
                        self.progress_bar.setFormat(f'正在合并: %p% - {current_file_index + 1}/{len(processed_files)} 文件')
                        QApplication.processEvents() # 允许UI更新

                            
                        if doc.page_count > 0:
                            c.saveState()
                            c.translate(centered_x, centered_y)
                            c.scale(scale, scale)
                            c.drawImage(temp_img.name, 0, 0, width=pdf_width, height=pdf_height)
                            c.restoreState()
                            
                            # 删除临时图像文件
                            try:
                                os.unlink(temp_img.name)
                            except:
                                pass
                        doc.close()
                    except Exception as e:
                        error_msg = f'处理文件时出错（{os.path.basename(pdf_file)}）：{str(e)}'
                        self.log_error(error_msg, exc_info=True)
                        QMessageBox.warning(self, '警告', error_msg)
                        continue

                # 保存当前页面
                c.save()
                
                # 将当前页面添加到输出PDF
                reader = PdfReader(temp_page.name)
                output_writer.add_page(reader.pages[0])
                temp_pdfs.append(temp_page.name)

            # 保存最终的PDF文件
            with open(output_file, 'wb') as output:
                output_writer.write(output)

            QMessageBox.information(self, '成功', '文件合并完成！')
            self.progress_bar.setValue(100)
            self.progress_bar.setFormat("合并完成")
            QApplication.processEvents() # 允许UI更新

        except Exception as e:
            error_msg = f'文件合并失败：{str(e)}'
            self.log_error(error_msg, exc_info=True)
            QMessageBox.critical(self, '错误', error_msg)
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("合并失败")
            QApplication.processEvents()

        finally:
            # 清理临时文件
            for temp_pdf in temp_pdfs:
                try:
                    os.unlink(temp_pdf)
                except:
                    pass

def main():
    app = QApplication(sys.argv)
    merger = PDFMerger()
    merger.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()