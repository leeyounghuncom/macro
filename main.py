import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QToolBar, QTableWidget,
    QMessageBox, QListWidget, QVBoxLayout, QWidget, QDialog, QComboBox, QPushButton, QLabel, QTableWidgetItem,
    QLineEdit, QFileDialog
)
from PyQt6.QtGui import QAction, QDesktopServices, QIcon
from PyQt6.QtCore import QUrl, QSize
from PyQt6.QtCore import QRegularExpression
from PyQt6.QtGui import QRegularExpressionValidator


class MacroRecorderApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # 메인 윈도우 설정
        self.setWindowTitle("이영훈닷컴 매크로")  
        self.setGeometry(100, 100, 800, 600)  # 창의 위치(x=100, y=100)와 크기(너비=800, 높이=600) 설정

        # UI 구성요소 초기화
        self.create_sidebar()  # 사이드바 생성
        self.create_menubar()  # 메뉴바 생성

        # 툴바 설정
        self.toolbar = QToolBar("메인 툴바")
        self.addToolBar(self.toolbar)

        # 툴바 액션 생성 및 추가
        self.create_actions()
        self.toolbar.addAction(self.keybaord_action)  # 키보드 액션 추가
        self.toolbar.addAction(self.delay_action)  # 딜레이 액션 추가

        # 중앙 테이블 위젯 설정
        self.table = QTableWidget(self)
        self.table.setColumnCount(3)  # 3개의 열 생성
        self.table.setHorizontalHeaderLabels(["명령", "", ""])  # 열 헤더 설정
        self.setCentralWidget(self.table)

    def create_sidebar(self):
        # 사이드바 위젯 및 레이아웃 설정
        sidebar_widget = QWidget(self)
        sidebar_layout = QVBoxLayout(sidebar_widget)

        # 리스트 위젯 생성 및 설정
        sidebar = QListWidget()
        sidebar.setIconSize(QSize(24, 24))  # 아이콘 크기 설정
        sidebar.setSpacing(10)  # 항목 간 간격 설정
        sidebar_layout.addWidget(sidebar)

        # 사이드바를 메인 윈도우의 왼쪽에 배치
        self.setMenuWidget(sidebar_widget)

    def create_menubar(self):
        # 메뉴바 생성
        menubar = self.menuBar()

        # 파일 메뉴 생성 및 액션 추가
        file_menu = menubar.addMenu("파일")
        
        # 새로 만들기 액션
        new_action = QAction("새로 만들기", self)
        new_action.triggered.connect(self.new_file)
        new_action.setIcon(QIcon("images/new.png"))
        file_menu.addAction(new_action)

        # 파일 열기 액션
        open_action = QAction("열기", self)
        open_action.triggered.connect(self.open_file)
        open_action.setIcon(QIcon("images/open.png"))
        file_menu.addAction(open_action)

        # 저장 액션
        save_action = QAction("저장", self)
        save_action.triggered.connect(self.save_file)
        save_action.setIcon(QIcon("images/save.png"))
        file_menu.addAction(save_action)

        file_menu.addSeparator()  # 구분선 추가

        # 종료 액션
        exit_action = QAction("종료", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # 도움말 메뉴 생성
        help_menu = menubar.addMenu("도움말")
        
        # 정보 액션
        info_action = QAction("정보", self)
        info_action.triggered.connect(self.show_info)
        info_action.setIcon(QIcon("images/info.png"))
        help_menu.addAction(info_action)

        # 문서 액션
        doc_action = QAction("문서", self)
        doc_action.triggered.connect(self.show_documentation)
        doc_action.setIcon(QIcon("images/doc.png"))
        help_menu.addAction(doc_action)

        help_menu.addSeparator()

        # 소개 액션
        about_action = QAction("소개", self)
        about_action.triggered.connect(self.show_about_dialog)
        about_action.setIcon(QIcon("images/about.png"))
        help_menu.addAction(about_action)

    def create_actions(self):
        # 키보드 액션 생성 및 설정
        self.keybaord_action = QAction("키보드", self)
        self.keybaord_action.triggered.connect(self.keybaord)
        self.keybaord_action.setShortcut("F5")  # 단축키 설정

        # 딜레이 액션 생성 및 설정
        self.delay_action = QAction("딜레이", self)
        self.delay_action.triggered.connect(self.delay)
        self.delay_action.setShortcut("F6")  # 단축키 설정

    # 메뉴 액션 메서드들
    def new_file(self):
        # 테이블 초기화 (모든 행 제거)
        self.table.setRowCount(0)
        QMessageBox.information(self, "새로 만들기", "새 매크로 파일이 생성되었습니다.")

    def open_file(self):
        # 파일 열기 다이얼로그 표시
        file_path, _ = QFileDialog.getOpenFileName(self, "파일 열기", "", "Macro Files (*.mcr);;All Files (*)")
        if file_path:
            self.load_from_file(file_path)

    def load_from_file(self, file_path):
        try:
            with open(file_path, 'r') as file:
                self.table.setRowCount(0)  # 기존 테이블 데이터 초기화
                for line in file:
                    parts = line.strip().split(',')  # 파일의 각 줄을 쉼표로 분리
                    row_position = self.table.rowCount()  # 현재 테이블 행 수 가져오기
                    self.table.insertRow(row_position)  # 새로운 행 삽입
                    for col, part in enumerate(parts):
                        self.table.setItem(row_position, col, QTableWidgetItem(part))  # 셀에 데이터 삽입
            QMessageBox.information(self, "열기 완료", f"파일이 {file_path}에서 성공적으로 로드되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "열기 실패", f"파일 열기 중 오류가 발생했습니다: {e}")

    def save_file(self):
        # 파일 저장 다이얼로그 표시
        options = QFileDialog.Option.DontUseNativeDialog
        file_path, _ = QFileDialog.getSaveFileName(self, "파일 저장", "", "Macro Files (*.mcr);;All Files (*)",
                                                   options=options)
        if file_path:
            if not file_path.endswith(".mcr"):
                file_path += ".mcr"  # 확장자 자동 추가
            self.save_to_file(file_path)

    def save_to_file(self, file_path):
        try:
            with open(file_path, 'w') as file:
                for row in range(self.table.rowCount()):
                    command = self._safe_item_text(row, 0)
                    param1 = self._safe_item_text(row, 1)
                    param2 = self._safe_item_text(row, 2)
                    file.write(f"{command},{param1},{param2}\n")  # 행 데이터 파일에 저장
            QMessageBox.information(self, "저장 완료", f"파일이 {file_path}에 저장되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "저장 실패", f"파일 저장 중 오류가 발생했습니다: {e}")

    def _safe_item_text(self, row, col):
        item = self.table.item(row, col)
        return item.text() if item else ""  # 셀 데이터 안전하게 반환

    def show_info(self):
        # 특정 URL 열기 (정보 제공)
        QDesktopServices.openUrl(QUrl("https://leeyounghun.com"))

    def show_documentation(self):
        # 문서 URL 열기
        QDesktopServices.openUrl(QUrl("https://github.com/leeyounghuncom/macro"))

    def keybaord(self):
        try:
            # 키보드 명령 설정 대화상자 실행
            dialog = KeyboardCommandDialog(self)  # 대화상자 생성
            if dialog.exec() == QDialog.DialogCode.Accepted:  # 대화상자 확인 버튼 클릭 시
                event_type, key, keycode = dialog.get_command()  # 명령과 키 값 가져오기
                if event_type and key:
                    row_position = self.table.rowCount()  # 현재 테이블 행 수 가져오기
                    self.table.insertRow(row_position)  # 새로운 행 삽입
                    self.table.setItem(row_position, 0, QTableWidgetItem("Keyboard"))  # 명령 설정
                    self.table.setItem(row_position, 1, QTableWidgetItem(event_type))  # 이벤트 유형 설정

                    if key == "KeyCode" and keycode:  # 처리할 keycode가 있으면 추가
                        self.table.setItem(row_position, 2, QTableWidgetItem(keycode))  # 3열에 숫자 추가
                    else:
                        self.table.setItem(row_position, 2, QTableWidgetItem(key))  # 3열에 키 추가
                else:
                    print("Invalid command input")
        except Exception as e:
            print(f"Error in keyboard action: {e}")

    def delay(self):
        # 딜레이 명령 설정 대화상자 실행
        dialog = DelayCommandDialog(self)  # 대화상자 생성
        dialog.setParent(None)  # 메모리 누수 방지
        if dialog.exec() == QDialog.DialogCode.Accepted:  # 대화상자 확인 버튼 클릭 시
            delay_value = dialog.get_delay()  # 딜레이 값 가져오기
            row_position = self.table.rowCount()  # 현재 테이블 행 수 가져오기
            self.table.insertRow(row_position)  # 새로운 행 삽입
            self.table.setItem(row_position, 0, QTableWidgetItem("Delay"))  # 명령 설정
            self.table.setItem(row_position, 1, QTableWidgetItem(delay_value))  # 딜레이 값 설정

    def show_about_dialog(self):
        # 소개 대화 상자 생성 및 표시
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("이영훈닷컴 매크로 소개")  # 대화 상자 제목
        msg_box.setText(
            "이영훈닷컴 매크로\n"
            "Version 0.0.0\n"
            "Copyright © 2025 이영훈닷컴 소프트웨어\n"
            "Registered to: 이영훈닷컴\n"
        )
        msg_box.setIcon(QMessageBox.Icon.Information)  # 정보 아이콘 설정
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)  # 확인 버튼 설정
        msg_box.exec()  # 대화 상자 실행


class KeyboardCommandDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("키보드 명령 설정")
        self.setGeometry(200, 200, 300, 200)

        layout = QVBoxLayout(self)

        self.event_type_label = QLabel("이벤트 선택:")
        self.event_type_combo = QComboBox(self)
        self.event_type_combo.addItems(["KeyUp", "KeyDown", "SystemKeyUp", "SystemKeyDown"])

        self.key_label = QLabel("키 선택:")
        self.key_combo = QComboBox(self)
        self.key_combo.addItems([  # 키 목록 추가
            "KeyCode", "Back", "Tab", "Clear", "Return", "ShiftLeft", "ControlLeft", "ShiftRight", "ControlRight",
            "AltLeft",
            "AltRight", "Space", "Left", "Up", "Right", "Down", "Print", "Insert", "Delete", "End", "Home", "PageUP",
            "PageDown",
            "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u",
            "v", "w", "x", "y", "z",
            "LWindows", "RWindows",
            "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12"])

        self.keycode_input_label = QLabel("KeyCode 값:")
        self.keycode_input = QLineEdit(self)
        self.keycode_input.setEnabled(False)

        self.ok_button = QPushButton("OK", self)
        self.ok_button.clicked.connect(self.accept)

        layout.addWidget(self.event_type_label)
        layout.addWidget(self.event_type_combo)
        layout.addWidget(self.key_label)
        layout.addWidget(self.key_combo)
        layout.addWidget(self.keycode_input_label)
        layout.addWidget(self.keycode_input)
        layout.addWidget(self.ok_button)
        # 기본적으로 "a" 키를 선택
        self.key_combo.setCurrentText("a")
        self.key_combo.currentTextChanged.connect(self.on_key_combo_change)
        self.keycode_input.textChanged.connect(self.validate_keycode)  # Connect textChanged to validation
        self.commands_list = []  # List to store the commands

    def on_key_combo_change(self, text):
        # 키 조합이 변경되었을 때 KeyCode 입력 필드 활성화 여부 설정
        if text == "KeyCode":
            self.keycode_input.setEnabled(True)
        else:
            self.keycode_input.setEnabled(False)

    def validate_keycode(self):
        keycode_value = self.keycode_input.text()

        if self.key_combo.currentText() == "KeyCode" and not keycode_value.isdigit():
            QMessageBox.critical(self, "입력 오류", "숫자만 가능합니다.")
            self.keycode_input.clear()

    def get_command(self):
        event_type = self.event_type_combo.currentText()  # 이벤트 유형
        key = self.key_combo.currentText()  # 선택된 키
        keycode = self.keycode_input.text()  # KeyCode 값
        return event_type, key, keycode

class DelayCommandDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("딜레이 설정")  # 대화상자 제목 설정
        self.setGeometry(200, 200, 300, 150)  # 대화상자 크기 설정

        # 레이아웃 설정
        layout = QVBoxLayout(self)

        # 밀리초 입력 필드
        self.milliseconds_label = QLabel("초:")
        self.milliseconds_input = QLineEdit(self)  # 밀리초 입력 필드
        self.milliseconds_input.setValidator(QRegularExpressionValidator(QRegularExpression(r'^\d+$')))

        # 확인 버튼
        self.ok_button = QPushButton("OK", self)
        self.ok_button.clicked.connect(self.accept)


        # 레이아웃에 위젯 추가
        layout.addWidget(self.milliseconds_label)
        layout.addWidget(self.milliseconds_input)
        layout.addWidget(self.ok_button)

    def get_delay(self):
        # 입력된 딜레이 값 반환
        return self.milliseconds_input.text()

if __name__ == "__main__":
    # 애플리케이션 실행
    app = QApplication(sys.argv)
    window = MacroRecorderApp()
    window.show()
    sys.exit(app.exec())