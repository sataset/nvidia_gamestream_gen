import os
import sys
import re
import PyQt5.QtWidgets as qtw
from PyQt5.QtCore import Qt as qtc

# get lnk path in Windows
import win32com.client
import urllib.parse
from shutil import copy

# get requests
from bs4 import BeautifulSoup
import urllib3
import json

# download
import urllib3.request


class MainWindow(qtw.QMainWindow):
    def __init__(self, app):
        super().__init__()
        # ref[1]
        self.path_shield_lib = os.getenv('USERPROFILE') \
            + '\\AppData\\Local\\NVIDIA Corporation\\Shield Apps'
        self.path_shield_thumbs = self.path_shield_lib + '\\StreamingAssets'

        self.setAcceptDrops(True)
        
        screen = app.primaryScreen()
        screen_w, screen_h = screen.size().width(), screen.size().height()
        window_w, window_h = 400, 400
        window_posx, window_posy = (screen_w / 2 - window_w / 2,
                                    screen_h / 2 - window_h / 2)
        self.setGeometry(window_posx, window_posy, window_w, window_h)

        self.setWindowTitle('Nvidia Gamestream Library Creation Tool')
        label = qtw.QLabel("Drag'n'Drop game shortcuts here")
        label.setAlignment(qtc.AlignCenter)
        self.setCentralWidget(label)
    
    def dragEnterEvent(self, e):
        if e.mimeData().hasFormat('text/uri-list'):
            self.lnks_comp = []
            self.lnks_incomp = []
            for lnk in e.mimeData().urls():
                lnk_path = lnk.toString().split('///')
                if lnk_path[0] == 'file:':
                    # if it is a link to file => translate url codes to unicode
                    lnk_path = urllib.parse.unquote(lnk_path[1])
                else:
                    continue
                # check if it is a link to the 3rd party launcher
                if lnk_path.split('.')[-1] == 'url':
                    name = re.search("/([^\\/]+).url", lnk_path).group(1)
                    self.lnks_incomp.append((name, lnk_path))
                    continue
                # check if it is not an executable
                check_shell = win32com.client.Dispatch('WScript.Shell')
                check_shortcut = check_shell.CreateShortCut(lnk_path)
                if check_shortcut.Targetpath[-4:] != '.exe':
                    continue

                name = re.search("/([^\\/]+).lnk", lnk_path).group(1)
                self.lnks_comp.append((name, lnk_path))
            e.accept()
        else:
            e.ignore()
    
    def dropEvent(self, e):
        self.show_lists()
        e.accept()
    
    def show_lists(self):
        self.list_comp = qtw.QListWidget()
        self.list_incomp = qtw.QListWidget()
        self.button_generate = qtw.QPushButton('Generate')

        self.list_comp.addItems([i[0] for i in self.lnks_comp])
        self.list_incomp.addItems([i[0] for i in self.lnks_incomp])
        self.button_generate.clicked.connect(self.generatePress)
        
        self.setCentralWidget(qtw.QWidget(self))
        self.layout = qtw.QVBoxLayout()  # ref[6]
        self.centralWidget().setLayout(self.layout)

        self.layout.addWidget(qtw.QLabel("Compatible games"))
        self.layout.addWidget(self.list_comp)
        self.layout.addWidget(qtw.QLabel("Incompatible games (Steam, GOG, etc. links)"))
        self.layout.addWidget(self.list_incomp)
        self.layout.addWidget(self.button_generate)
    
    def generatePress(self):
        steamgriddb_api_key = 'effdb93fd27500ab1e9e6e1128e2b12f'
        # http = urllib.connection_from_url('https://www.steamgriddb.com/')
        http = urllib3.PoolManager()
        headers = {'Authorization': f'Bearer {steamgriddb_api_key}'}

        def get_json_data(connection, url, headers=None):
            response = connection.request('GET', url, headers=headers)
            soup = BeautifulSoup(response.data, features='html.parser')
            return json.loads(soup.text)['data']

        for name, lnk in self.lnks_comp:
            query = '+'.join(name.split())

            # fetch a game_id from steamgriddb by name
            url = f'https://www.steamgriddb.com/api/v2/search/autocomplete/{query}'
            json_data = get_json_data(http, url, headers)
            game_id = json_data[0]['id']

            # fetch a game cover
            url = f'https://www.steamgriddb.com/api/v2/grids/game/{game_id}'
            json_data = get_json_data(http, url, headers)

            # download thumbnail (ref[11])
            if not json_data:
                print(f'{name} game is not found in SteamGridDB.')
                continue
            thumb = http.request('GET', json_data[0]['thumb'])
            path_thumb = self.path_shield_thumbs + f'\\{name}'
            try:
                os.mkdir(path_thumb)
            except FileExistsError:
                pass
            with open(path_thumb + '\\box-art.png', 'wb') as f:
                f.write(thumb.data)

            copy(lnk, self.path_shield_lib)

        self.button_generate.setText('Success!')


if __name__ == '__main__':
    app = qtw.QApplication(sys.argv)
    window = MainWindow(app)
    window.show()
    app.exec_()
