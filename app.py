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
    def __init__(self, app, args):
        super().__init__()
        # ref[1]
        self.path_shield_lib = os.getenv('USERPROFILE') \
            + '\\AppData\\Local\\NVIDIA Corporation\\Shield Apps'
        self.path_shield_thumbs = self.path_shield_lib + '\\StreamingAssets'
        self.setAcceptDrops(True)
        
        screen = app.primaryScreen()
        screen_w, screen_h = screen.size().width(), screen.size().height()
        window_w, window_h = 400, 400
        
        if len(args) > 1:
            window_posx, window_posy = int(args[1]), int(args[2])
        else:
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

                # if it is [not a link to file] OR [not a link], then skip
                if lnk_path[0] != 'file:' \
                   or lnk_path[1].split('.')[-1] not in ['url', 'lnk']:
                    continue
                
                # translate url codes to unicode
                lnk_path = urllib.parse.unquote(lnk_path[1])
                
                # if [it is a link to the 3rd party launcher]
                # add to `lnks_incomp` list and skip
                if lnk_path.split('.')[-1] == 'url':
                    name = re.search("/([^\\/]+).url", lnk_path).group(1)
                    self.lnks_incomp.append((name, lnk_path))
                    continue
                
                # checking where link points
                check_shell = win32com.client.Dispatch('WScript.Shell')
                check_shortcut = check_shell.CreateShortCut(lnk_path)
                
                # check if link source [exists] OR [is not an executable]
                if not os.path.isfile(check_shortcut.Targetpath) \
                   or check_shortcut.Targetpath.split('.')[-1] != 'exe':
                    continue

                name = re.search("/([^\\/]+).lnk", lnk_path).group(1)
                self.lnks_comp.append((name, lnk_path))
            e.accept()
        else:
            e.ignore()
    
    def dropEvent(self, e):
        # e.setDropAction(qtc.MoveAction)
        self.show_lists()
        e.accept()
    
    def show_lists(self):
        self.list_comp = qtw.QListWidget()
        self.list_incomp = qtw.QListWidget()
        self.button_generate = qtw.QPushButton('Generate')

        self.list_comp.addItems(i[0] for i in self.lnks_comp)
        self.list_incomp.addItems(i[0] for i in self.lnks_incomp)
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
        for directory in [self.path_shield_lib, self.path_shield_thumbs]:
            if not os.path.exists(directory):
                os.makedirs(directory)

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

            # creating directory for thumbnails (box-art)
            path_thumb = f'{self.path_shield_thumbs}\\{name}'
            if not os.path.exists(path_thumb):
                os.mkdir(path_thumb)
            
            # download thumbnail (ref[11])
            if not json_data:
                print(f'{name} game is not found in SteamGridDB.')
                print('Adding default thumbnails.')
                copy('.\\assets\\box-art.png', path_thumb)
            else:
                thumb = http.request('GET', json_data[0]['thumb'])
                with open(path_thumb + '\\box-art.png', 'wb') as f:
                    f.write(thumb.data)
            
            if not os.path.isfile(f"{self.path_shield_lib}\\{lnk.split('/')[-1]}"):
                copy(lnk, self.path_shield_lib)
        self.button_generate.setText('Success!')


if __name__ == '__main__':
    app = qtw.QApplication(sys.argv)
    window = MainWindow(app, args=sys.argv)
    window.show()
    app.exec_()
