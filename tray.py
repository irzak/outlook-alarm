import sys

from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QAction, QApplication, QColorDialog, QMenu, QSystemTrayIcon

app = QApplication(sys.argv)
app.setQuitOnLastWindowClosed(False)

# Create the icon
icon = QIcon("bell.png")

clipboard = QApplication.clipboard()
dialog = QColorDialog()


def copy_color_hex():
    if dialog.exec_():
        color = dialog.currentColor()
        clipboard.setText(color.name())


def copy_color_rgb():
    if dialog.exec_():
        color = dialog.currentColor()
        clipboard.setText(
            "rgb(%d, %d, %d)" % (color.red(), color.green(), color.blue())
        )


def copy_color_hsv():
    if dialog.exec_():
        color = dialog.currentColor()
        clipboard.setText(
            "hsv(%d, %d, %d)" % (color.hue(), color.saturation(), color.value())
        )


# Create the tray
tray = QSystemTrayIcon()
tray.setIcon(icon)
tray.setVisible(True)

# Create the menu
menu = QMenu()
action1 = QAction("Hex")
action1.triggered.connect(copy_color_hex)

action2 = QAction("RGB")
action2.triggered.connect(copy_color_rgb)

action3 = QAction("HSV")
action3.triggered.connect(copy_color_hsv)

quit = QAction("Quit")
quit.triggered.connect(app.quit)

menu.addAction(action1)
menu.addAction(action2)
menu.addAction(action3)
menu.addAction(quit)

# Add the menu to the tray
tray.setContextMenu(menu)

app.exec_()
