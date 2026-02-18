import json

from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QLineEdit, QPushButton, QFormLayout, QGroupBox,
                             QComboBox, QDoubleSpinBox, QSpinBox, QTabWidget,
                             QWidget, QFileDialog, QMessageBox, QListWidget,
                             QCheckBox,
                             QListWidgetItem)

from core.config_manager import ConfigManager


class SettingsWindow(QDialog):
    def __init__(self, config_manager: ConfigManager, parent=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.setWindowTitle("설정")
        self.resize(620, 760)
        self.initUI()
        self.loadConfig()

    def initUI(self):
        layout = QVBoxLayout(self)

        tabs = QTabWidget()

        # ── 일반 설정 탭 ──
        general_tab = QWidget()
        gen_layout = QVBoxLayout(general_tab)

        # 출력 폴더
        path_group = QGroupBox("출력 폴더 설정")
        path_layout = QHBoxLayout()
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("C:/모의고사_출력/")
        self.path_btn = QPushButton("변경")
        path_layout.addWidget(self.path_edit)
        path_layout.addWidget(self.path_btn)
        path_group.setLayout(path_layout)
        gen_layout.addWidget(path_group)

        # 폰트 및 서식
        font_group = QGroupBox("폰트 및 서식")
        font_form = QFormLayout()

        self.q_font_combo = QComboBox()
        self.q_font_combo.setEditable(True)
        self.q_font_combo.addItems(["중고딕", "HY중고딕", "맑은 고딕", "돋움"])
        font_form.addRow("문제 폰트:", self.q_font_combo)

        self.p_font_combo = QComboBox()
        self.p_font_combo.setEditable(True)
        self.p_font_combo.addItems(["휴먼명조", "HY신명조", "바탕", "함초롬바탕"])
        font_form.addRow("지문 폰트:", self.p_font_combo)

        self.size_spin = QDoubleSpinBox()
        self.size_spin.setRange(6.0, 20.0)
        self.size_spin.setSingleStep(0.5)
        self.size_spin.setValue(9.5)
        self.size_spin.setSuffix(" pt")
        font_form.addRow("글자 크기:", self.size_spin)

        self.char_width_spin = QSpinBox()
        self.char_width_spin.setRange(50, 200)
        self.char_width_spin.setValue(95)
        self.char_width_spin.setSuffix(" %")
        font_form.addRow("장평:", self.char_width_spin)

        self.char_spacing_spin = QSpinBox()
        self.char_spacing_spin.setRange(-50, 50)
        self.char_spacing_spin.setValue(-5)
        self.char_spacing_spin.setSuffix(" %")
        font_form.addRow("자간:", self.char_spacing_spin)

        font_group.setLayout(font_form)
        gen_layout.addWidget(font_group)

        # 문단 설정
        para_group = QGroupBox("문단 설정")
        para_form = QFormLayout()

        self.line_spacing_spin = QSpinBox()
        self.line_spacing_spin.setRange(100, 300)
        self.line_spacing_spin.setValue(140)
        self.line_spacing_spin.setSuffix(" %")
        para_form.addRow("줄 간격:", self.line_spacing_spin)

        self.indent_spin = QDoubleSpinBox()
        self.indent_spin.setRange(0.0, 50.0)
        self.indent_spin.setSingleStep(0.1)
        self.indent_spin.setValue(13.8)
        self.indent_spin.setSuffix(" pt")
        para_form.addRow("내어쓰기:", self.indent_spin)

        para_group.setLayout(para_form)
        gen_layout.addWidget(para_group)

        style_group = QGroupBox("HWP 스타일 연동")
        style_form = QFormLayout()

        self.style_enabled_check = QCheckBox("스타일 적용 사용")
        style_form.addRow(self.style_enabled_check)

        template_row = QHBoxLayout()
        self.style_template_edit = QLineEdit()
        self.style_template_edit.setPlaceholderText("C:/path/to/template.hwp")
        self.style_template_btn = QPushButton("찾기")
        template_row.addWidget(self.style_template_edit)
        template_row.addWidget(self.style_template_btn)
        style_form.addRow("템플릿 HWP:", template_row)

        module_row = QHBoxLayout()
        self.module_dll_edit = QLineEdit()
        self.module_dll_edit.setPlaceholderText("C:/path/to/FilePathCheckerModule*.dll")
        self.module_dll_btn = QPushButton("찾기")
        module_row.addWidget(self.module_dll_edit)
        module_row.addWidget(self.module_dll_btn)
        style_form.addRow("보안 모듈 DLL:", module_row)

        self.question_style_edit = QLineEdit()
        self.question_style_edit.setPlaceholderText("Normal")
        style_form.addRow("문제 스타일:", self.question_style_edit)

        self.passage_style_edit = QLineEdit()
        self.passage_style_edit.setPlaceholderText("Body")
        style_form.addRow("지문 스타일:", self.passage_style_edit)

        self.choice_style_edit = QLineEdit()
        self.choice_style_edit.setPlaceholderText("Body")
        style_form.addRow("선지 스타일:", self.choice_style_edit)

        self.sub_items_style_edit = QLineEdit()
        self.sub_items_style_edit.setPlaceholderText("Body")
        style_form.addRow("소문항 스타일:", self.sub_items_style_edit)

        self.explanation_style_edit = QLineEdit()
        self.explanation_style_edit.setPlaceholderText("Body")
        style_form.addRow("해설 스타일:", self.explanation_style_edit)

        style_notice = QLabel(
            "※ 템플릿 문서 안에 동일한 스타일명이 있어야 적용됩니다.\n"
            "※ HwpAutomation Modules 보안 모듈이 미등록이면 템플릿 열기가 차단됩니다."
        )
        style_notice.setWordWrap(True)
        style_form.addRow(style_notice)

        style_group.setLayout(style_form)
        gen_layout.addWidget(style_group)

        gen_layout.addStretch()
        tabs.addTab(general_tab, "일반")

        # ── 파싱 패턴 탭 ──
        advanced_tab = QWidget()
        adv_layout = QVBoxLayout(advanced_tab)

        adv_layout.addWidget(QLabel("문제 시작 패턴:"))
        self.q_pattern_list = QListWidget()
        adv_layout.addWidget(self.q_pattern_list)

        adv_layout.addWidget(QLabel("정답 감지 패턴:"))
        self.a_pattern_list = QListWidget()
        adv_layout.addWidget(self.a_pattern_list)

        adv_layout.addWidget(QLabel("해설 감지 패턴:"))
        self.e_pattern_list = QListWidget()
        adv_layout.addWidget(self.e_pattern_list)

        adv_layout.addWidget(QLabel("부정문 키워드:"))
        self.neg_list = QListWidget()
        adv_layout.addWidget(self.neg_list)

        adv_layout.addWidget(
            QLabel("※ 패턴 수정은 config/default_config.json 파일을 직접 편집하세요.")
        )

        tabs.addTab(advanced_tab, "고급 (패턴)")

        layout.addWidget(tabs)

        # 하단 버튼
        btn_layout = QHBoxLayout()
        self.export_btn = QPushButton("설정 내보내기")
        self.import_btn = QPushButton("설정 가져오기")
        self.save_btn = QPushButton("저장")
        self.save_btn.setStyleSheet("background-color: #1a237e; color: white;")
        self.cancel_btn = QPushButton("취소")
        btn_layout.addWidget(self.export_btn)
        btn_layout.addWidget(self.import_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.save_btn)
        btn_layout.addWidget(self.cancel_btn)
        layout.addLayout(btn_layout)

        self.path_btn.clicked.connect(self.pickOutputDirectory)
        self.style_template_btn.clicked.connect(self.pickStyleTemplateFile)
        self.module_dll_btn.clicked.connect(self.pickModuleDllFile)
        self.style_enabled_check.toggled.connect(self.setStyleControlsEnabled)
        self.export_btn.clicked.connect(self.exportConfig)
        self.import_btn.clicked.connect(self.importConfig)
        self.save_btn.clicked.connect(self.saveConfig)
        self.cancel_btn.clicked.connect(self.reject)
        self.setStyleControlsEnabled(self.style_enabled_check.isChecked())

    def loadConfig(self):
        config = self.config_manager.all()

        # 경로
        output_path = config.get("paths", {}).get("output_directory", "")
        self.path_edit.setText(output_path)

        # 폰트
        fmt = config.get("format", {})
        q_font = fmt.get("question_font", "중고딕")
        p_font = fmt.get("passage_font", "휴먼명조")
        font_size = float(fmt.get("font_size", 9.5))
        char_width = int(fmt.get("char_width", 95))
        char_spacing = int(fmt.get("char_spacing", -5))

        if self.q_font_combo.findText(q_font) == -1:
            self.q_font_combo.addItem(q_font)
        if self.p_font_combo.findText(p_font) == -1:
            self.p_font_combo.addItem(p_font)

        self.q_font_combo.setCurrentText(q_font)
        self.p_font_combo.setCurrentText(p_font)
        self.size_spin.setValue(font_size)
        self.char_width_spin.setValue(char_width)
        self.char_spacing_spin.setValue(char_spacing)

        # 문단
        para = config.get("paragraph", {})
        self.line_spacing_spin.setValue(int(para.get("line_spacing", 140)))
        self.indent_spin.setValue(float(para.get("indent_value", 13.8)))

        style = config.get("style", {})
        self.style_enabled_check.setChecked(bool(style.get("enabled", True)))
        template_path = str(style.get("template_path", "")).strip() or str(style.get("style_map_source", "")).strip()
        self.style_template_edit.setText(template_path)
        self.module_dll_edit.setText(str(style.get("module_dll_path", "")).strip())
        self.question_style_edit.setText(str(style.get("question_style", "Normal")))
        self.passage_style_edit.setText(str(style.get("passage_style", "Body")))
        self.choice_style_edit.setText(str(style.get("choice_style", "Body")))
        self.sub_items_style_edit.setText(str(style.get("sub_items_style", "Body")))
        self.explanation_style_edit.setText(str(style.get("explanation_style", "Body")))
        self.setStyleControlsEnabled(self.style_enabled_check.isChecked())

        # 패턴 (읽기 전용 표시)
        parsing = config.get("parsing", {})
        self.q_pattern_list.clear()
        for p in parsing.get("question_patterns", []):
            self.q_pattern_list.addItem(QListWidgetItem(p))
        self.a_pattern_list.clear()
        for p in parsing.get("answer_patterns", []):
            self.a_pattern_list.addItem(QListWidgetItem(p))
        self.e_pattern_list.clear()
        for p in parsing.get("explanation_patterns", []):
            self.e_pattern_list.addItem(QListWidgetItem(p))
        self.neg_list.clear()
        for kw in config.get("negative_keywords", []):
            self.neg_list.addItem(QListWidgetItem(kw))

    def pickOutputDirectory(self):
        selected = QFileDialog.getExistingDirectory(
            self,
            "출력 폴더 선택",
            self.path_edit.text() or "",
        )
        if selected:
            self.path_edit.setText(selected)

    def pickStyleTemplateFile(self):
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "스타일 템플릿 HWP 선택",
            self.style_template_edit.text() or "",
            "HWP Files (*.hwp);;All Files (*)",
        )
        if selected:
            self.style_template_edit.setText(selected)

    def pickModuleDllFile(self):
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "보안 모듈 DLL 선택",
            self.module_dll_edit.text() or "",
            "DLL Files (*.dll);;All Files (*)",
        )
        if selected:
            self.module_dll_edit.setText(selected)

    def setStyleControlsEnabled(self, enabled: bool):
        self.style_template_edit.setEnabled(enabled)
        self.style_template_btn.setEnabled(enabled)
        self.module_dll_edit.setEnabled(enabled)
        self.module_dll_btn.setEnabled(enabled)
        self.question_style_edit.setEnabled(enabled)
        self.passage_style_edit.setEnabled(enabled)
        self.choice_style_edit.setEnabled(enabled)
        self.sub_items_style_edit.setEnabled(enabled)
        self.explanation_style_edit.setEnabled(enabled)

    def exportConfig(self):
        target_path, _ = QFileDialog.getSaveFileName(
            self,
            "설정 내보내기",
            "hwp_exam_editor_settings.json",
            "JSON Files (*.json);;All Files (*)",
        )
        if not target_path:
            return

        try:
            config = self.config_manager.all()
            with open(target_path, "w", encoding="utf-8") as file:
                json.dump(config, file, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "내보내기 완료", f"설정을 저장했습니다.\n{target_path}")
        except Exception as exc:
            QMessageBox.warning(self, "내보내기 실패", f"설정 파일 저장에 실패했습니다.\n{exc}")

    def importConfig(self):
        source_path, _ = QFileDialog.getOpenFileName(
            self,
            "설정 가져오기",
            "",
            "JSON Files (*.json);;All Files (*)",
        )
        if not source_path:
            return

        try:
            with open(source_path, "r", encoding="utf-8") as file:
                payload = json.load(file)
            if not isinstance(payload, dict):
                raise ValueError("루트 JSON 객체는 딕셔너리여야 합니다.")

            self.config_manager.update(payload)
            self.loadConfig()
            QMessageBox.information(self, "가져오기 완료", "설정을 불러와 화면에 반영했습니다.")
        except Exception as exc:
            QMessageBox.warning(self, "가져오기 실패", f"설정 파일을 읽지 못했습니다.\n{exc}")

    def saveConfig(self):
        template_path = self.style_template_edit.text().strip()
        if self.style_enabled_check.isChecked() and not template_path:
            QMessageBox.warning(
                self,
                "스타일 연동 경고",
                "스타일 적용이 켜져 있지만 템플릿 경로가 비어 있습니다.\n"
                "저장 후에도 스타일 적용이 제한될 수 있습니다.",
            )

        partial = {
            "paths": {
                "output_directory": self.path_edit.text().strip(),
            },
            "format": {
                "question_font": self.q_font_combo.currentText(),
                "passage_font": self.p_font_combo.currentText(),
                "font_size": self.size_spin.value(),
                "char_width": self.char_width_spin.value(),
                "char_spacing": self.char_spacing_spin.value(),
            },
            "paragraph": {
                "line_spacing": self.line_spacing_spin.value(),
                "indent_value": self.indent_spin.value(),
            },
            "style": {
                "enabled": self.style_enabled_check.isChecked(),
                "template_path": template_path,
                "style_map_source": template_path,
                "module_dll_path": self.module_dll_edit.text().strip(),
                "question_style": self.question_style_edit.text().strip() or "Normal",
                "passage_style": self.passage_style_edit.text().strip() or "Body",
                "choice_style": self.choice_style_edit.text().strip() or "Body",
                "sub_items_style": self.sub_items_style_edit.text().strip() or "Body",
                "explanation_style": self.explanation_style_edit.text().strip() or "Body",
            },
        }
        self.config_manager.update(partial)
        QMessageBox.information(self, "저장 완료", "설정을 저장했습니다.")
        self.accept()
