# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'tags_window_design.ui'
##
## Created by: Qt User Interface Compiler version 6.9.1
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QApplication, QFrame, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QPushButton, QSizePolicy,
    QWidget)

class Ui_TagsWindow(object):
    def setupUi(self, TagsWindow):
        if not TagsWindow.objectName():
            TagsWindow.setObjectName(u"TagsWindow")
        TagsWindow.resize(440, 320)
        sizePolicy = QSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(TagsWindow.sizePolicy().hasHeightForWidth())
        TagsWindow.setSizePolicy(sizePolicy)
        TagsWindow.setMinimumSize(QSize(440, 320))
        TagsWindow.setMaximumSize(QSize(440, 320))
        self.frame = QFrame(TagsWindow)
        self.frame.setObjectName(u"frame")
        self.frame.setGeometry(QRect(10, 0, 431, 311))
        self.frame.setFrameShape(QFrame.StyledPanel)
        self.frame.setFrameShadow(QFrame.Raised)
        self.tag_lineEdit = QLineEdit(self.frame)
        self.tag_lineEdit.setObjectName(u"tag_lineEdit")
        self.tag_lineEdit.setGeometry(QRect(0, 180, 121, 20))
        self.TagList = QListWidget(self.frame)
        self.TagList.setObjectName(u"TagList")
        self.TagList.setGeometry(QRect(0, 80, 191, 81))
        self.label = QLabel(self.frame)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(0, 10, 111, 16))
        self.add_tag = QPushButton(self.frame)
        self.add_tag.setObjectName(u"add_tag")
        self.add_tag.setGeometry(QRect(130, 180, 61, 21))
        self.delete_tag = QPushButton(self.frame)
        self.delete_tag.setObjectName(u"delete_tag")
        self.delete_tag.setGeometry(QRect(130, 161, 61, 20))
        self.label_2 = QLabel(self.frame)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setGeometry(QRect(0, 30, 351, 16))
        self.type_label = QLabel(self.frame)
        self.type_label.setObjectName(u"type_label")
        self.type_label.setGeometry(QRect(110, 10, 311, 16))
        font = QFont()
        font.setBold(True)
        font.setItalic(False)
        font.setUnderline(True)
        self.type_label.setFont(font)
        self.line = QFrame(self.frame)
        self.line.setObjectName(u"line")
        self.line.setGeometry(QRect(200, 60, 20, 171))
        self.line.setFrameShape(QFrame.Shape.VLine)
        self.line.setFrameShadow(QFrame.Shadow.Sunken)
        self.label_3 = QLabel(self.frame)
        self.label_3.setObjectName(u"label_3")
        self.label_3.setGeometry(QRect(0, 249, 191, 16))
        self.mask_lineEdit = QLineEdit(self.frame)
        self.mask_lineEdit.setObjectName(u"mask_lineEdit")
        self.mask_lineEdit.setGeometry(QRect(0, 269, 191, 21))
        self.save_mask = QPushButton(self.frame)
        self.save_mask.setObjectName(u"save_mask")
        self.save_mask.setGeometry(QRect(130, 290, 61, 20))
        self.label_5 = QLabel(self.frame)
        self.label_5.setObjectName(u"label_5")
        self.label_5.setGeometry(QRect(0, 60, 201, 16))
        self.TagList_2 = QListWidget(self.frame)
        self.TagList_2.setObjectName(u"TagList_2")
        self.TagList_2.setGeometry(QRect(230, 80, 191, 81))
        self.add_tag_2 = QPushButton(self.frame)
        self.add_tag_2.setObjectName(u"add_tag_2")
        self.add_tag_2.setGeometry(QRect(360, 180, 61, 21))
        self.tag_lineEdit_2 = QLineEdit(self.frame)
        self.tag_lineEdit_2.setObjectName(u"tag_lineEdit_2")
        self.tag_lineEdit_2.setGeometry(QRect(230, 180, 121, 20))
        self.delete_tag_2 = QPushButton(self.frame)
        self.delete_tag_2.setObjectName(u"delete_tag_2")
        self.delete_tag_2.setGeometry(QRect(360, 161, 61, 20))
        self.label_6 = QLabel(self.frame)
        self.label_6.setObjectName(u"label_6")
        self.label_6.setGeometry(QRect(230, 60, 201, 16))
        self.line_2 = QFrame(self.frame)
        self.line_2.setObjectName(u"line_2")
        self.line_2.setGeometry(QRect(0, 230, 431, 20))
        self.line_2.setFrameShape(QFrame.Shape.HLine)
        self.line_2.setFrameShadow(QFrame.Shadow.Sunken)

        self.retranslateUi(TagsWindow)

        QMetaObject.connectSlotsByName(TagsWindow)
    # setupUi

    def retranslateUi(self, TagsWindow):
        TagsWindow.setWindowTitle(QCoreApplication.translate("TagsWindow", u"\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0430 \u0442\u0435\u0433\u043e\u0432 \u0438 \u043c\u0430\u0441\u043a\u0438", None))
        self.label.setText(QCoreApplication.translate("TagsWindow", u"\u041f\u043e\u0438\u0441\u043a \u0444\u0430\u0439\u043b\u043e\u0432 \u0442\u0438\u043f\u0430 ", None))
        self.add_tag.setText(QCoreApplication.translate("TagsWindow", u"\u0434\u043e\u0431\u0430\u0432\u0438\u0442\u044c", None))
        self.delete_tag.setText(QCoreApplication.translate("TagsWindow", u"\u0443\u0434\u0430\u043b\u0438\u0442\u044c", None))
        self.label_2.setText(QCoreApplication.translate("TagsWindow", u"\u0432\u044b\u043f\u043e\u043b\u043d\u044f\u0435\u0442\u0441\u044f \u043f\u043e \u0441\u043b\u0435\u0434\u0443\u044e\u0449\u0438\u043c \u0442\u0435\u0433\u0430\u043c (\u0440\u0435\u0433\u0438\u0441\u0442\u0440\u043e\u043d\u0435\u0437\u0430\u0432\u0438\u0441\u0438\u043c\u043e):", None))
        self.type_label.setText(QCoreApplication.translate("TagsWindow", u"...", None))
        self.label_3.setText(QCoreApplication.translate("TagsWindow", u"\u041c\u0430\u0441\u043a\u0430 \u0434\u043b\u044f \u043d\u043e\u0432\u043e\u0433\u043e \u0438\u043c\u0435\u043d\u0438 :", None))
        self.save_mask.setText(QCoreApplication.translate("TagsWindow", u"\u0441\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c", None))
        self.label_5.setText(QCoreApplication.translate("TagsWindow", u"\u0432 \u0438\u043c\u0435\u043d\u0438 \u0444\u0430\u0439\u043b\u0430:", None))
        self.add_tag_2.setText(QCoreApplication.translate("TagsWindow", u"\u0434\u043e\u0431\u0430\u0432\u0438\u0442\u044c", None))
        self.delete_tag_2.setText(QCoreApplication.translate("TagsWindow", u"\u0443\u0434\u0430\u043b\u0438\u0442\u044c", None))
        self.label_6.setText(QCoreApplication.translate("TagsWindow", u"\u0432\u043d\u0443\u0442\u0440\u0438 \u0444\u0430\u0439\u043b\u0430 (Excel \u0438 PDF):", None))
    # retranslateUi

