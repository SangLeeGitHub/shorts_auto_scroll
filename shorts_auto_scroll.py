import os
import math
import re
import sys
import time
import pyautogui
from PyQt6.QtNetwork import QNetworkCookie
from PyQt6.QtWebEngineCore import QWebEngineProfile
from cryptography.fernet import Fernet
import subprocess

if os.name == 'nt':
    import win32con
    import win32api
    import win32gui
else:
    import Quartz
    from AppKit import NSRunningApplication, NSApplicationActivateIgnoringOtherApps

from PyQt6.QtCore import QUrl, QTimer, QEvent, Qt, QByteArray, QDateTime
from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QMessageBox, QMenu, QMenuBar, QLabel
from PyQt6.QtWebEngineWidgets import QWebEngineView
from googleapiclient.discovery import build


def is_app_in_accessibility(program_name):
    script = f'''
        tell application "System Events"
            set isAppInAccessibility to false
            repeat with theProcess in (every application process whose name is not "Shorts Auto Scroll")
                try
                    if app name of theProcess is not equal to "" then
                        set isAppInAccessibility to true
                        exit repeat
                    end if
                end try
            end repeat
        end tell
        return isAppInAccessibility
    '''

    result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
    if result.returncode == 0:
        return result.stdout.strip() == 'true'
    else:
        # AppleScript 실행 중 오류 발생
        print("AppleScript 실행 중 오류 발생:", result.stderr)
        return False


def iso8601_duration_to_seconds(duration):
    # ISO 8601 지속 시간을 파싱하기 위한 정규 표현식
    pattern = r'PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?'
    hours, minutes, seconds = re.match(pattern, duration).groups()

    # None 값은 0으로 변환
    hours = int(hours) if hours else 0
    minutes = int(minutes) if minutes else 0
    seconds = int(seconds) if seconds else 0

    # 시간, 분, 초를 모두 초로 변환
    return hours * 3600 + minutes * 60 + seconds


def decrypt_message(encrypted_message, key):
    f = Fernet(key)
    decrypted_message = f.decrypt(encrypted_message.encode()).decode('utf-8')
    return decrypted_message


def get_api_key_from_file(filename="api.key"):
    api_key_path = ''
    try:
        # For PyInstaller
        if getattr(sys, 'frozen', False):
            bundle_dir = os.path.dirname(sys.executable)
            resource_path = os.path.join(os.path.dirname(bundle_dir), 'Resources')
        else:
            resource_path = os.path.dirname(os.path.abspath(__file__))

        api_key_path = os.path.join(resource_path, filename)

        decryption_key = b"YOUR_GENERATED_KEY_HERE"

        with open(api_key_path, 'r') as file:
            encrypted_api_key = file.read().strip()
            os.system(f'logger "API key loaded "')
            return decrypt_message(encrypted_api_key, decryption_key)
    except FileNotFoundError:
        os.system(f'logger "Error: {api_key_path} file not found!"')
        sys.exit(1)


def get_video_duration(api_key, video_id):
    youtube = build("youtube", "v3", developerKey=api_key)

    request = youtube.videos().list(
        part="contentDetails",
        id=video_id
    )
    response = request.execute()

    if "items" in response and len(response["items"]) > 0:
        content_details = response["items"][0]["contentDetails"]
        duration_in_iso8601 = content_details["duration"]
        return iso8601_duration_to_seconds(duration_in_iso8601)
    else:
        os.system(f'logger "Error fetching video details."')
        return None


def activate_app():
    current_app = NSRunningApplication.currentApplication()
    current_app.activateWithOptions_(NSApplicationActivateIgnoringOtherApps)


if os.name == 'nt':
    def send_key_to_window(window_title, key_code):
        hwnd = win32gui.FindWindow(None, window_title)
        if hwnd:
            win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, key_code, 0)
            win32api.PostMessage(hwnd, win32con.WM_KEYUP, key_code, 0)
        else:
            os.system(f'logger "{window_title} not found!"')
else:
    def send_key_to_app(title_to_find, key_code):
        event = Quartz.CGEventCreateKeyboardEvent(None, key_code, True)
        Quartz.CGEventPost(Quartz.kCGHIDEventTap, event)
        event = Quartz.CGEventCreateKeyboardEvent(None, key_code, False)
        Quartz.CGEventPost(Quartz.kCGHIDEventTap, event)


MAX_FILE_SIZE = 300 * 1024  # 300k


def maintain_file_size():
    if os.path.getsize(cookies_file_path) > MAX_FILE_SIZE:
        with open(cookies_file_path, 'r') as file:
            lines = file.readlines()

        while sum(len(line) for line in lines) > MAX_FILE_SIZE and lines:
            lines.pop(0)

        with open(cookies_file_path, 'w') as file:
            file.writelines(lines)


def save_cookie(cookie):
    domain = cookie.domain()
    if ".google.com" in domain:
        cookie_str = cookie.toRawForm().data().decode()

        try:
            with open(cookies_file_path, "r") as file:
                existing_cookies = file.readlines()
        except FileNotFoundError:
            existing_cookies = []

        for existing_cookie in existing_cookies:
            if cookie_str.strip() == existing_cookie.strip():
                return

        with open(cookies_file_path, "a") as file:
            file.write(cookie_str + "\n")

        maintain_file_size()


def load_cookies():
    try:
        valid_cookies = []
        with open(cookies_file_path, "r") as file:
            for line in file:
                parts = line.strip().split(";")
                cookie = QNetworkCookie()

                if "=" not in parts[0]:
                    continue  # skip this iteration

                # name and value
                name, value = parts[0].split("=", 1)
                cookie.setName(name.strip().encode())
                cookie.setValue(value.strip().encode())

                is_expired = False

                # other attributes
                for part in parts[1:]:
                    if "=" not in part:
                        continue  # skip this part

                    key, value = part.split("=", 1)
                    if "domain" in key:
                        cookie.setDomain(value.strip())
                    elif "path" in key:
                        cookie.setPath(value.strip())
                    elif "expires" in key:
                        expiration_date = QDateTime.fromString(value.strip(), "ddd, dd MMM yyyy HH:mm:ss GMT")
                        if expiration_date.isValid():
                            cookie.setExpirationDate(expiration_date)
                            if expiration_date < QDateTime.currentDateTime():
                                is_expired = True
                                break

                if not is_expired:
                    valid_cookies.append(line.strip())
                    cookie_store.setCookie(cookie)

        # 만료된 쿠키 제거 후 유효한 쿠키만 파일에 다시 작성
        with open(cookies_file_path, "w") as file:
            file.write("\n".join(valid_cookies))

    except Exception as e:
        if isinstance(e, FileNotFoundError):
            os.system("Warning: cookies.txt file not found. Continuing without loading cookies.")
        else:
            os.system(f'logger "Another error occurred: {e}"')


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.api_key = get_api_key_from_file()
        self.video_id = None
        self.original_geometry = None
        self.init_ui()
        self.init_menu()
        self.google_login()
        self.title = "Shorts Auto Scroll"
        self.setWindowTitle(self.title)
        self.remaining_time = None
        self.remaining_timer = QTimer(self)
        self.remaining_timer.timeout.connect(self.update_remaining_time)
        self.is_key_from_function = False
        self.web_view.installEventFilter(self)
        # self.setWindowOpacity(.2)

    def init_ui(self):
        self.layout = QVBoxLayout()

        self.web_view = QWebEngineView()
        self.layout.addWidget(self.web_view)

        self.container = QWidget()
        self.container.setLayout(self.layout)
        self.setCentralWidget(self.container)
        self.video_id_label = QLabel('Video ID: None')
        self.length_label = QLabel('Remaining: Unknown')
        self.statusBar().addWidget(self.video_id_label)  # Left-aligned by default
        self.statusBar().addPermanentWidget(self.length_label)  # Right-aligned
        self.resize(360, 685)

    def init_menu(self):
        # Create a menu bar
        menu_bar = QMenuBar(self)

        # Create an "About" menu item
        menu = QMenu("Menu", self)
        about_action = menu.addAction("About")
        about_action.triggered.connect(self.show_about_popup)

        # Create an "Exit" menu item
        exit_action = menu.addAction("Exit")
        exit_action.triggered.connect(self.close_application)

        # Add the "About" menu to the menu bar
        menu_bar.addMenu(menu)
        self.setMenuBar(menu_bar)

    def show_about_popup(self):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("About")
        msg_box.setText('Shorts Auto Scroll 0.7<br><a href="https://nore.co/app/sas">by Slime</a><br>')
        msg_box.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextBrowserInteraction)  # Enable interaction with the link
        msg_box.setIcon(QMessageBox.Icon.Information)
        msg_box.exec()

    def close_application(self):
        self.close()

    def eventFilter(self, source, event):
        if source == self.web_view and event.type() == QEvent.Type.ShortcutOverride:
            if event.key() == Qt.Key.Key_Down:
                # 만약 이 키 이벤트가 함수에 의한 것이 아니라면 start_loop()를 호출
                if not self.is_key_from_function:
                    QTimer.singleShot(1000, self.start_loop)
                    return True  # 이벤트 처리 완료. 추가로 아래 키 입력을 허용하지 않음.
                # 만약 이 키 이벤트가 함수에 의한 것이면, 플래그를 초기화하고 이벤트를 통과시킴.
                self.is_key_from_function = False
        return super(MainWindow, self).eventFilter(source, event)

    def changeEvent(self, event):
        if event.type() == QEvent.Type.WindowStateChange:
            if self.windowState() & Qt.WindowState.WindowMinimized:
                # 창이 다른 창들 뒤에 머물도록 설정
                self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnBottomHint)
                # 창 상태를 다시 정상 상태로 설정하여 최소화 효과를 제거
                self.setWindowState(Qt.WindowState.WindowNoState)
                self.show()

            else:
                # WindowStaysOnBottomHint 속성 해제
                self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowStaysOnBottomHint)
                self.show()

    def google_login(self):
        self.web_view.setUrl(QUrl("https://accounts.google.com"))
        self.web_view.loadFinished.connect(self.on_load_finished)
        self.setCentralWidget(self.web_view)

    def on_load_finished(self, ok):
        current_url = self.web_view.url().toString()

        if "youtube.com/shorts" in current_url:
            QTimer.singleShot(500, self.send_tab_and_enter)
        elif "accounts.google.com" not in current_url:
            self.web_view.setUrl(QUrl("https://youtube.com/shorts"))

    def send_tab_and_enter(self):
        QTimer.singleShot(500, self.press_tab_then_enter)

    def press_tab_then_enter(self):
        pyautogui.press('tab')
        pyautogui.press('enter')
        pyautogui.press('esc')
        self.press_down()

    def start_loop(self):
        js_code = "window.location.href"
        self.web_view.page().runJavaScript(js_code, self.on_url_retrieved)

    def handle_video(self):
        if self.video_id:
            start_time = time.time()

            if not self.api_key:
                os.system(f'logger "API key not available yet."')
                return

            duration_in_seconds = get_video_duration(self.api_key, self.video_id)
            execution_time = time.time() - start_time

            if duration_in_seconds is None:
                duration_in_seconds = 15  # 에러 발생 시 기본 값

            self.video_id_label.setText(f'Video ID: {self.video_id}')

            # remaining_time 초기화 및 statusBar 업데이트
            self.remaining_time = duration_in_seconds - 2
            self.length_label.setText(f'Remaining: {self.remaining_time} sec.')

            duration_in_seconds -= execution_time  # 실행 시간을 빼 줌
            duration_in_seconds = math.ceil(duration_in_seconds) + 1

            # 1초마다 remaining_time 업데이트
            self.remaining_timer.start(1000)

            if duration_in_seconds:
                QTimer.singleShot((duration_in_seconds - 2) * 1000, self.press_down)

    def update_remaining_time(self):
        if self.remaining_time is not None and self.remaining_time > 0:
            self.remaining_time -= 1
            self.length_label.setText(f'Remaining: {self.remaining_time} sec.')
            if self.remaining_time == 0:
                self.remaining_timer.stop()

    def press_down(self):
        self.is_key_from_function = True  # 플래그 설정
        if os.name == 'nt':
            send_key_to_window(self.title, win32con.VK_DOWN)
        else:
            activate_app()
            send_key_to_app(self.title, 125)

        QTimer.singleShot(2000, self.start_loop)  # 다시 루프 시작

    def on_url_retrieved(self, url):
        if 'shorts/' not in url:
            QTimer.singleShot(500, self.start_loop)
            return

        match = re.search(r'shorts/([\w\-]+)', url)
        if match:
            self.video_id = match.group(1)
            self.handle_video()  # URL이 검색되면 비디오를 처리
        else:
            self.video_id_label.setText('Video ID: None')
            self.length_label.setText('Length: Unknown')


if os.name != 'nt':
    program_name = "/Applications/Shorts Auto Scroll.app"
    is_app_in_accessibility(program_name)

app = QApplication(sys.argv)

os.system('logger "Getting started with the Shorts Auto Scroll."')

support_path = os.path.expanduser('~/Library/Application Support/ShortsAutoScroll')
if not os.path.exists(support_path):
    os.makedirs(support_path)
cookies_file_path = os.path.join(support_path, 'cookies.txt')

try:
    profile = QWebEngineProfile.defaultProfile()
    cookie_store = profile.cookieStore()
    cookie_store.cookieAdded.connect(save_cookie)

    load_cookies()

    window = MainWindow()
    window.show()

    sys.exit(app.exec())

except Exception as e:
    error_dialog = QMessageBox.critical(None, "Error", str(e))
