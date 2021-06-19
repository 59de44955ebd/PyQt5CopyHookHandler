import os
import sys
import shutil

from ctypes import Structure, POINTER, windll, c_void_p, byref, cast
from ctypes.wintypes import DWORD, LPVOID, MSG, LPCSTR
from win32con import WM_COPYDATA, SW_RESTORE

from comtypes import CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER, IUnknown, GUID
from comtypes.hresult import S_OK

from PyQt5.QtCore import Qt, QSize, QMimeData, QUrl
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem,
        QMessageBox, QAbstractItemView, QFileIconProvider)

APP_NAME = 'Drag-to-Explorer Demo'
APP_VERSION = 1
APP_DIR = os.path.dirname(os.path.realpath(__file__))

TYPE_FILE = 0
TYPE_FOLDER = 1

CLSID_ShellExtension = '{8CAB387B-2F65-45B3-A903-ADC360479B7C}'


class COPYDATASTRUCT(Structure):
    _fields_ = (
        ('dwData', POINTER(DWORD)),
        ('cbData', DWORD),
        ('lpData', LPVOID)
    )


class MyTreeWidget(QTreeWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self._drop_dir = os.path.join(os.environ['TMP'], '__qt5drop__')
        if not os.path.isdir(self._drop_dir):
            os.mkdir(self._drop_dir)

    def closeEvent(self, e):
        if os.path.isdir(self._drop_dir):
            os.rmdir(self._drop_dir)

    def mimeData(self, tree_items):
        mime_data = QMimeData()
        mime_data.setUrls([QUrl.fromLocalFile(self._drop_dir)])
        return mime_data

    def mimeTypes(self):
        return ['text/uri-list']

    def startDrag(self, dropActions):
        ''' force copy mode '''
        super().startDrag(Qt.CopyAction)

    def copyToExplorer(self, target_dir):
        for tree_item in self.selectedItems():
            src = tree_item.data(0, Qt.UserRole)
            if tree_item.type() == TYPE_FOLDER:
                src_basename = os.path.basename(src)
                dest = os.path.join(target_dir, src_basename)
                if os.path.isdir(dest):
                    if QMessageBox.question(
                            self,
                            'Delete existing directory?',
                            'Directory "{}" already exists. Delete it?'.format(src_basename)
                            ) != QMessageBox.Yes:
                        continue
                    shutil.rmtree(dest)
                shutil.copytree(src, dest)
            else:
                src_basename = os.path.basename(src)
                if os.path.isfile(os.path.join(target_dir, src_basename)):
                    if QMessageBox.question(
                            self,
                            'Overwrite?',
                            'File "{}" already exists. Overwrite it?'.format(src_basename)
                            ) != QMessageBox.Yes:
                        continue
                shutil.copy(src, target_dir)


class Main(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)

        self._tree_widget = MyTreeWidget(self)
        self.setCentralWidget(self._tree_widget)

        self._tree_widget.setHeaderHidden(True)
        self._tree_widget.setIconSize(QSize(40, 30))
        self._tree_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self._tree_widget.setDragDropMode(QAbstractItemView.DragOnly)

        image_dir = os.path.join(APP_DIR, r'resources\images')
        icon_folder = QFileIconProvider().icon(QFileIconProvider.Folder)

        folder_item = QTreeWidgetItem(['images'], TYPE_FOLDER)
        folder_item.setIcon(0, icon_folder)
        folder_item.setData(0, Qt.UserRole, image_dir)
        self._tree_widget.addTopLevelItem(folder_item)

        for f in os.listdir(image_dir):
            file_item = QTreeWidgetItem([f], TYPE_FILE)
            file_item.setIcon(0, QIcon(os.path.join(image_dir, f)))
            file_item.setData(0, Qt.UserRole, os.path.join(image_dir, f))
            folder_item.addChild(file_item)
            folder_item.setExpanded(True)

        self.resize(240, 520)
        self.show()

    def nativeEvent (self, event_type, message):
        if event_type == 'windows_generic_MSG':
            msg = MSG.from_address(message.__int__())
            if msg.message == WM_COPYDATA:
                ds = cast(msg.lParam, POINTER(COPYDATASTRUCT))
                data = cast(ds.contents.lpData, LPCSTR)
                target_dir = os.path.dirname(data.value.decode())
                self._tree_widget.copyToExplorer(target_dir)
        return False, 0


if __name__ == '__main__':
    app = QApplication(sys.argv)

    # check if shell extension is registered
    shell_ext = c_void_p(None)
    hr = windll.ole32.CoCreateInstance(
            byref(GUID(CLSID_ShellExtension)),
            0,
            CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
            byref(IUnknown._iid_),
            byref(shell_ext))
    if hr != S_OK:
        # not registered, register it now?
        if QMessageBox.question(
                None,
                'Shell extension not registered',
                "Shell extension 'PyQt5CopyHookHandler.dll' is not registered in the system. Register it now?"
                ) != QMessageBox.Yes:
            sys.exit(0)
        ret = windll.shell32.ShellExecuteW(
                None,
                'runas',
                'regsvr32.exe',
                os.path.join(APP_DIR, r'resources\shellext\PyQt5CopyHookHandler.dll'),
                None,
                1)
        # restart explorer
        if QMessageBox.question(
                None,
                'Restart Explorer?',
                'Before the shell extension can be used, Explorer must be restarted. Restart it now?'
                ) != QMessageBox.Yes:
            sys.exit(0)
        windll.shell32.ShellExecuteW(
                None,
                'open',
                'cmd.exe',
                '/c taskkill /IM explorer.exe /F && start explorer.exe',
                None,
                0)

    main = Main()
    sys.exit(app.exec_())
