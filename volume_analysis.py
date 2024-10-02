import os
import subprocess
import re
import csv
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QLineEdit, QListWidget, QTableWidget, QTableWidgetItem, \
    QVBoxLayout, QHBoxLayout, QProgressDialog
from PyQt5.QtCore import QCoreApplication, Qt


class VolumeAnalysisApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        # UI 구성 요소
        self.setWindowTitle('MP3 & MP4 Volume Analysis')
        self.setGeometry(300, 300, 1000, 600)  # 전체 창 크기 조정

        # Layouts
        main_layout = QVBoxLayout(self)
        top_layout = QHBoxLayout()
        bottom_layout = QVBoxLayout()
        button_layout = QHBoxLayout()  # 버튼 레이아웃

        # 폴더 선택 라벨
        self.label = QtWidgets.QLabel('리소스 파일이 있는 폴더를 선택하세요:', self)
        top_layout.addWidget(self.label)

        # 폴더 경로 표시 텍스트 필드
        self.folder_path_display = QLineEdit(self)
        self.folder_path_display.setReadOnly(True)
        top_layout.addWidget(self.folder_path_display)

        # 폴더 선택 버튼
        self.browse_button = QtWidgets.QPushButton('폴더 선택', self)
        self.browse_button.clicked.connect(self.select_folder)
        top_layout.addWidget(self.browse_button)

        # 분석 버튼
        self.analyze_button = QtWidgets.QPushButton('분석', self)
        self.analyze_button.clicked.connect(self.analyze_folder)
        top_layout.addWidget(self.analyze_button)

        main_layout.addLayout(top_layout)

        # 파일 리스트 표시 위젯
        self.file_list = QListWidget(self)

        # 분석 결과를 표시할 테이블 위젯
        self.table_widget = QTableWidget(self)
        self.table_widget.setColumnCount(3)
        self.table_widget.setHorizontalHeaderLabels(['Filename', 'Mean Volume (dB)', 'Comment'])
        self.table_widget.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)

        # Set column widths
        self.table_widget.setColumnWidth(0, 300)  # Filename column width
        self.table_widget.setColumnWidth(1, 150)  # Mean Volume column width
        self.table_widget.setColumnWidth(2, 350)  # Comment column width

        # 하단 레이아웃에 파일 리스트와 테이블 추가
        bottom_layout.addWidget(self.file_list)
        bottom_layout.addWidget(self.table_widget)

        main_layout.addLayout(bottom_layout)

        # 버튼 레이아웃에 내보내기 및 종료 버튼 추가
        self.export_button = QtWidgets.QPushButton('내보내기', self)
        self.export_button.clicked.connect(self.export_to_csv)
        button_layout.addWidget(self.export_button)

        self.exit_button = QtWidgets.QPushButton('종료', self)
        self.exit_button.clicked.connect(self.close)
        button_layout.addWidget(self.exit_button)

        main_layout.addLayout(button_layout)

        self.setLayout(main_layout)

        self.selected_folder = None

    def get_volume_from_ffmpeg(self, file_path):
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            ffmpeg_path = os.path.join(current_dir, 'ffmpeg', 'ffmpeg.exe')
            result = subprocess.run(
                [
                    ffmpeg_path,
                    '-i', file_path,
                    '-af', 'volumedetect',
                    '-f', 'null',
                    'NUL' if os.name == 'nt' else '/dev/null'
                ],
                stderr=subprocess.PIPE,
                text=False
            )
            output = result.stderr.decode('utf-8')

            mean_volume_match = re.search(r'mean_volume: ([\d.-]+) dB', output)
            if mean_volume_match:
                return float(mean_volume_match.group(1))
            else:
                return None
        except Exception as e:
            QMessageBox.critical(f"Error processing file {file_path}: {e}")
            return None

    def process_folder(self, folder_path, progress_dialog):
        self.table_widget.setRowCount(0)  # Clear existing rows
        files = os.listdir(folder_path)
        total_files = len(files)

        # Initialize a dictionary to hold results and errors
        results = {}
        errors = {}

        allowed_extensions = ('.mp3', '.mp4')  # 허용된 파일 확장자 목록

        for index, filename in enumerate(files):
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path) and file_path.lower().endswith(allowed_extensions):
                try:
                    mean_volume = self.get_volume_from_ffmpeg(file_path)
                    if mean_volume is not None:
                        # 음량 범위 검토
                        if -19.00 <= mean_volume <= -15.00:
                            comment = ''
                        else:
                            comment = '가이드 범위(-15.00 dB ~ -19.00 dB)에 맞지 않습니다.'
                        results[filename] = (f"{mean_volume:.2f} dB", comment)
                    else:
                        errors[filename] = '에러가 발생했습니다.'
                except Exception as e:
                    errors[filename] = f"{str(e)} 에러가 발생했습니다."

            # 프로그레스 다이얼로그 업데이트
            progress_dialog.setValue(int((index + 1) / total_files * 100))
            QCoreApplication.processEvents()  # UI 강제 갱신

            # 사용자가 취소 버튼을 누른 경우 작업을 중단
            if progress_dialog.wasCanceled():
                break

        # 결과와 에러를 원래 파일 순서대로 테이블에 추가
        combined_results = []
        for filename in files:
            if filename in results:
                volume, comment = results[filename]
                combined_results.append((filename, volume, comment))
            elif filename in errors:
                combined_results.append((filename, '-', errors[filename]))

        if combined_results:
            self.table_widget.setRowCount(len(combined_results))
            for row, (filename, mean_volume, comment) in enumerate(combined_results):
                self.table_widget.setItem(row, 0, QTableWidgetItem(filename))
                self.table_widget.setItem(row, 1, QTableWidgetItem(mean_volume))
                self.table_widget.setItem(row, 2, QTableWidgetItem(comment))
        else:
            self.table_widget.setRowCount(1)
            self.table_widget.setItem(0, 0, QTableWidgetItem("No audio files found or unable to calculate volume."))
            self.table_widget.setItem(0, 1, QTableWidgetItem(''))
            self.table_widget.setItem(0, 2, QTableWidgetItem(''))

    def select_folder(self):
        dialog = QFileDialog(self)
        dialog.setFileMode(QFileDialog.Directory)  # Allow both files and folders to be shown
        dialog.setOption(QFileDialog.ShowDirsOnly, True)  # Show only directories

        if dialog.exec_() == QFileDialog.Accepted:
            self.selected_folder = dialog.selectedFiles()[0]
            self.folder_path_display.setText(self.selected_folder)  # 폴더 경로 표시
            self.load_files_in_folder(self.selected_folder)

    def load_files_in_folder(self, folder_path):
        self.file_list.clear()  # 기존 파일 목록을 초기화
        self.table_widget.hide()  # 분석 테이블 미노출
        allowed_extensions = {'.mp3', '.mp4'}  # 허용된 파일 확장자 목록
        files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))
                 and os.path.splitext(f)[1].lower() in allowed_extensions]  # mp3, mp4 파일만 포함

        # 파일 목록을 번호와 함께 표시
        for index, filename in enumerate(files, start=1):
            item_text = f"{index}. {filename}"  # 번호와 파일 이름을 포함한 텍스트
            self.file_list.addItem(item_text)

    def analyze_folder(self):
        if self.selected_folder:
            # 로딩 화면 (QProgressDialog) 생성
            progress_dialog = QProgressDialog("평균 음량을 분석 중입니다...", "취소", 0, 100, self)
            progress_dialog.setWindowTitle("분석 중")
            progress_dialog.setWindowModality(QtCore.Qt.WindowModal)

            # 창 크기 조정
            progress_dialog.resize(400, 100)  # 가로 400, 세로 100으로 크기 설정

            # 물음표 버튼 제거
            progress_dialog.setWindowFlags(progress_dialog.windowFlags() & ~Qt.WindowContextHelpButtonHint)

            progress_dialog.show()

            # 분석 프로세스 실행
            self.process_folder(self.selected_folder, progress_dialog)

            # 분석 완료 후 프로그레스 바를 100%로 설정
            progress_dialog.setValue(100)

            # 분석 테이블 노출
            self.table_widget.show()
        else:
            QMessageBox.warning(self, "경고", "분석할 폴더를 먼저 선택하세요.")

    def export_to_csv(self):
        if self.table_widget.rowCount() == 0:
            QMessageBox.warning(self, "경고", "테이블에 데이터가 없습니다.")
            return

        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "엑셀로 내보내기", "", "CSV Files (*.csv)", options=options)

        if file_name:
            try:
                with open(file_name, mode='w', newline='', encoding='utf-8-sig') as file:
                    writer = csv.writer(file)
                    writer.writerow(['Filename', 'Mean Volume (dB)', 'Comment'])  # 헤더 작성
                    for row in range(self.table_widget.rowCount()):
                        filename = self.table_widget.item(row, 0).text()
                        mean_volume = self.table_widget.item(row, 1).text()
                        comment = self.table_widget.item(row, 2).text()
                        writer.writerow([filename, mean_volume, comment])
                QMessageBox.information(self, "완료", "CSV 파일이 성공적으로 저장되었습니다.")
            except Exception as e:
                QMessageBox.critical(self, "오류", f"파일 저장 중 오류가 발생했습니다: {e}")


def main():
    import sys
    app = QtWidgets.QApplication(sys.argv)
    volume_analysis_app = VolumeAnalysisApp()
    volume_analysis_app.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
