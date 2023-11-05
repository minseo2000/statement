import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout,
    QWidget, QTextEdit, QLineEdit, QLabel, QHBoxLayout
)


class ServerWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # 윈도우 설정
        self.setWindowTitle('서버 설정')
        self.setGeometry(100, 100, 400, 300)

        # 레이아웃과 위젯
        layout = QVBoxLayout()

        # 호스트 IP 입력 필드
        host_layout = QHBoxLayout()
        self.host_label = QLabel('호스트 IP:')
        self.host_input = QLineEdit('0.0.0.0')  # 기본값으로 localhost 설정
        host_layout.addWidget(self.host_label)
        host_layout.addWidget(self.host_input)
        layout.addLayout(host_layout)

        # 포트 번호 입력 필드
        port_layout = QHBoxLayout()
        self.port_label = QLabel('포트 번호:')
        self.port_input = QLineEdit('50000')  # 기본 포트 번호 설정
        port_layout.addWidget(self.port_label)
        port_layout.addWidget(self.port_input)
        layout.addLayout(port_layout)

        # 서버 상태를 보여주는 텍스트 에디트 위젯
        self.server_status = QTextEdit()
        self.server_status.setReadOnly(True)
        layout.addWidget(self.server_status)

        # 서버 시작 버튼
        self.btn_start = QPushButton('서버 시작')
        self.btn_start.clicked.connect(self.start_server)
        layout.addWidget(self.btn_start)

        # 서버 중지 버튼
        self.btn_stop = QPushButton('서버 중지')
        self.btn_stop.clicked.connect(self.stop_server)
        layout.addWidget(self.btn_stop)

        # 중앙 위젯 설정
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def start_server(self):
        # 사용자가 입력한 호스트 IP와 포트 번호를 가져옵니다.
        host = self.host_input.text()
        port = self.port_input.text()
        self.server_status.append(f'서버가 {host}:{port} 에서 시작되었습니다.')
        # 여기에 서버 시작 로직을 구현합니다.

    def stop_server(self):
        # 서버 중지 로직을 구현합니다.
        self.server_status.append('서버가 중지되었습니다.')
        # 여기에 서버 중지 로직을 구현합니다.


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = ServerWindow()
    win.show()
    sys.exit(app.exec_())
