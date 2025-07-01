from PyQt5 import QtWidgets, uic
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
from PIL import Image, ImageEnhance, ImageQt
import fitz
import sys
import io

class PDFEditor(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi("editor.ui", self)

        self.pdf_doc = None
        self.original_image = None
        self.edited_image = None

        # Signals
        self.loadPdfBtn.clicked.connect(self.load_pdf)
        self.saveBtn.clicked.connect(self.save_image)
        self.pageSlider.valueChanged.connect(self.change_page_slider)
        self.pageBox.valueChanged.connect(self.change_page_box)
        self.brightnessSlider.valueChanged.connect(self.update_image)
        self.contrastSlider.valueChanged.connect(self.update_image)

    def load_pdf(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF Files (*.pdf)")
        if path:
            self.pdf_doc = fitz.open(path)
            total_pages = len(self.pdf_doc)
            self.pageSlider.setMaximum(total_pages - 1)
            self.pageBox.setMaximum(total_pages)
            self.pageSlider.setValue(0)
            self.pageBox.setValue(1)
            self.load_page(0)

    def load_page(self, index):
        if not self.pdf_doc:
            return
        page = self.pdf_doc.load_page(index)
        pix = page.get_pixmap(dpi=200)
        img = Image.open(io.BytesIO(pix.tobytes("ppm"))).convert("RGB")
        self.original_image = img
        self.update_image()

    def change_page_slider(self, value):
        self.pageBox.setValue(value + 1)
        self.load_page(value)

    def change_page_box(self, value):
        self.pageSlider.setValue(value - 1)
        self.load_page(value - 1)

    def update_image(self):
        if not self.original_image:
            return
        brightness = self.brightnessSlider.value() / 100
        contrast = self.contrastSlider.value() / 100
        img = ImageEnhance.Brightness(self.original_image).enhance(brightness)
        img = ImageEnhance.Contrast(img).enhance(contrast)
        self.edited_image = img
        qt_img = ImageQt.ImageQt(img)
        pixmap = QPixmap.fromImage(qt_img).scaled(800, 500, Qt.KeepAspectRatio)
        self.imageLabel.setPixmap(pixmap)

    def save_image(self):
        if self.edited_image:
            path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save Image", "", "PNG (*.png);;JPG (*.jpg)")
            if path:
                self.edited_image.save(path)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = PDFEditor()
    window.show()
    sys.exit(app.exec_())