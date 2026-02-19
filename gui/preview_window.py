import re

from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QListWidget, 
                             QListWidgetItem, QTextEdit, QLabel, QPushButton, 
                             QSplitter, QFrame, QMessageBox)
from PyQt5.QtCore import Qt

from core.models import ExamDocument


class PreviewWindow(QDialog):
    def __init__(self, exam_document: ExamDocument, parent=None):
        super().__init__(parent)
        self.exam_document = exam_document
        self.current_index = -1

        self.setWindowTitle("파싱 결과 미리보기 및 보정")
        self.resize(1000, 700)
        self.initUI()
        self.bindEvents()
        self.loadQuestions()

    def initUI(self):
        main_layout = QVBoxLayout(self)
        
        # 안내 문구
        header_label = QLabel("왼쪽 목록에서 문제를 선택하여 내용을 확인하고 수정할 수 있습니다.")
        header_label.setStyleSheet("font-weight: bold; color: #333; margin-bottom: 10px;")
        main_layout.addWidget(header_label)

        # 스플리터 (좌: 목록, 우: 상세편집)
        splitter = QSplitter(Qt.Horizontal)
        
        # 좌측 리스트
        self.list_widget = QListWidget()
        self.list_widget.setObjectName("QuestionList")
        splitter.addWidget(self.list_widget)
        
        # 우측 상세 영역
        detail_container = QFrame()
        detail_layout = QVBoxLayout(detail_container)
        
        detail_layout.addWidget(QLabel("문제 및 지문 영역:"))
        self.question_edit = QTextEdit()
        detail_layout.addWidget(self.question_edit)
        
        detail_layout.addWidget(QLabel("정답:"))
        self.answer_edit = QTextEdit()
        self.answer_edit.setMaximumHeight(50)
        detail_layout.addWidget(self.answer_edit)
        
        detail_layout.addWidget(QLabel("해설 영역:"))
        self.explanation_edit = QTextEdit()
        detail_layout.addWidget(self.explanation_edit)
        
        splitter.addWidget(detail_container)
        splitter.setStretchFactor(1, 1) # 우측 영역이 더 넓게
        
        main_layout.addWidget(splitter)
        
        # 하단 제어 버튼
        btn_layout = QHBoxLayout()
        self.save_btn = QPushButton("변경사항 저장")
        self.save_btn.setProperty("class", "success")
        
        self.convert_btn = QPushButton("선택한 설정으로 변환 시작")
        self.convert_btn.setProperty("class", "primary")
        
        self.close_btn = QPushButton("닫기")
        self.close_btn.setProperty("class", "danger")
        
        btn_layout.addStretch()
        btn_layout.addWidget(self.save_btn)
        btn_layout.addWidget(self.convert_btn)
        btn_layout.addWidget(self.close_btn)
        
        main_layout.addLayout(btn_layout)

    def bindEvents(self):
        self.list_widget.currentRowChanged.connect(self.onQuestionSelected)
        self.save_btn.clicked.connect(self.saveCurrentQuestion)
        self.convert_btn.clicked.connect(self.convertAndClose)
        self.close_btn.clicked.connect(self.reject)

    def loadQuestions(self):
        self.list_widget.clear()
        for question in self.exam_document.questions:
            status = "✅"
            if self.exam_document.file_type == "TYPE_A" and not question.answer:
                status = "❌"
            elif self.exam_document.file_type == "TYPE_A" and not question.explanation:
                status = "⚠️"

            preview_line = question.question_text.splitlines()[0] if question.question_text else "(본문 없음)"
            item_text = f"{status} {question.number:02d}. {preview_line}"
            self.list_widget.addItem(QListWidgetItem(item_text))

        if self.exam_document.questions:
            self.list_widget.setCurrentRow(0)

    def onQuestionSelected(self, index: int):
        self.current_index = index
        if index < 0 or index >= len(self.exam_document.questions):
            self.question_edit.clear()
            self.answer_edit.clear()
            self.explanation_edit.clear()
            return

        question = self.exam_document.questions[index]
        body_parts = [question.question_text]
        if question.sub_items:
            body_parts.append("\n".join(question.sub_items))
        if question.choices:
            body_parts.append("\n".join(question.choices))

        self.question_edit.setPlainText("\n".join(part for part in body_parts if part))
        self.answer_edit.setPlainText(question.answer_line or question.answer or "")
        self.explanation_edit.setPlainText(question.explanation or "")

    def saveCurrentQuestion(self, show_message: bool = True) -> bool:
        index = self.current_index
        if index < 0 or index >= len(self.exam_document.questions):
            return False

        question = self.exam_document.questions[index]
        edited_lines = [
            line.rstrip()
            for line in self.question_edit.toPlainText().splitlines()
            if line.strip()
        ]

        choices: list[str] = []
        sub_items: list[str] = []
        body: list[str] = []
        for line in edited_lines:
            stripped = line.strip()
            if stripped[:1] in {"①", "②", "③", "④", "⑤"}:
                choices.append(stripped)
            elif stripped[:1] in {"㉠", "㉡", "㉢", "㉣", "㉤", "㉥"}:
                sub_items.append(stripped)
            else:
                body.append(stripped)

        question_text = "\n".join(body).strip()
        if not question_text:
            QMessageBox.warning(
                self,
                "저장 실패",
                "문제 본문이 비어 있습니다.\n문제 제목/질문 문장을 입력한 뒤 저장해 주세요.",
            )
            self.question_edit.setFocus()
            return False

        question.question_text = question_text
        question.choices = choices
        question.sub_items = sub_items
        answer_text = self.answer_edit.toPlainText().strip()
        answer_symbol = None
        circle_match = re.search(r"[①②③④⑤]", answer_text)
        if circle_match:
            answer_symbol = circle_match.group(0)
        else:
            digit_match = re.search(r"(?<!\d)([1-5])(?!\d)", answer_text)
            if digit_match:
                answer_symbol = {"1": "①", "2": "②", "3": "③", "4": "④", "5": "⑤"}[digit_match.group(1)]
        question.answer = answer_symbol or (answer_text or None)
        if answer_text:
            question.answer_line = answer_text if "정답" in answer_text else f"정답 {answer_text}"
        else:
            question.answer_line = None
        question.explanation = self.explanation_edit.toPlainText().strip() or None

        self.loadQuestions()
        self.list_widget.setCurrentRow(index)
        if show_message:
            QMessageBox.information(self, "저장 완료", "현재 문항 변경사항을 저장했습니다.")
        return True

    def convertAndClose(self):
        if self.saveCurrentQuestion(show_message=False):
            self.accept()
