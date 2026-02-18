from PyQt5.QtWidgets import QGraphicsDropShadowEffect
from PyQt5.QtGui import QColor

def apply_shadow(widget, blur=15, offset=(0, 2)):
    shadow = QGraphicsDropShadowEffect()
    shadow.setBlurRadius(blur)
    shadow.setXOffset(offset[0])
    shadow.setYOffset(offset[1])
    shadow.setColor(QColor(0, 0, 0, 40))
    widget.setGraphicsEffect(shadow)

APP_STYLE = """
    QMainWindow, QDialog {
        background-color: #fafafa;
    }
    
    #MainTitle {
        font-family: 'Malgun Gothic', 'Segoe UI';
        font-size: 28px;
        font-weight: 800;
        color: #1a237e;
        margin-top: 10px;
    }
    
    #SubTitle {
        font-family: 'Malgun Gothic';
        font-size: 15px;
        color: #7986cb;
        margin-bottom: 20px;
    }
    
    #DropArea {
        background-color: #ffffff;
        border: 2px dashed #90caf9;
        border-radius: 20px;
    }
    
    #DropArea:hover {
        background-color: #e3f2fd;
        border: 2px dashed #2196f3;
    }
    
    #DropArea QLabel {
        font-size: 18px;
        color: #64b5f6;
        font-weight: bold;
    }
    
    QPushButton {
        padding: 14px 28px;
        font-size: 15px;
        font-weight: bold;
        border-radius: 10px;
        background-color: #eeeeee;
        color: #424242;
        border: none;
    }
    
    QPushButton:hover {
        background-color: #e0e0e0;
    }
    
    QPushButton:pressed {
        background-color: #bdbdbd;
    }
    
    #PrimaryBtn {
        background-color: #283593;
        color: #ffffff;
    }
    
    #PrimaryBtn:hover {
        background-color: #1a237e;
    }
    
    #PrimaryBtn:disabled {
        background-color: #c5cae9;
        color: #ffffff;
    }
    
    #NoticeLabel {
        font-size: 12px;
        color: #e57373;
        line-height: 1.5;
        margin-top: 20px;
    }
    
    QProgressBar {
        border: none;
        border-radius: 5px;
        background-color: #e0e0e0;
        height: 8px;
        text-align: center;
    }
    
    QProgressBar::chunk {
        background-color: #3f51b5;
        border-radius: 5px;
    }
    
    QGroupBox {
        font-weight: bold;
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        margin-top: 1.5em;
        padding-top: 10px;
    }
    
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 3px 0 3px;
        color: #3f51b5;
    }
"""
