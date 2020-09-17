import logging
import sys
import threading
import time
import winsound
from datetime import datetime

import win32com.client
from apscheduler.schedulers.qt import QtScheduler
from decouple import config
from PyQt5.QtCore import QUrl
from PyQt5.QtGui import QIcon
from PyQt5.QtMultimedia import QMediaContent, QMediaPlayer
from PyQt5.QtWidgets import QAction, QApplication, QMenu, QSystemTrayIcon

GROUP_FOLDER = config("GROUP_FOLDER")
TARGET_FOLDER = config("TARGET_FOLDER")
AUDIO = config("AUDIO")
INTERVAL = 5

mediaPlayer = QMediaPlayer(None, flags=QMediaPlayer.VideoSurface)
mediaPlayer.setMedia(QMediaContent(QUrl.fromLocalFile(AUDIO)))

logging.basicConfig(
    filename='running.log', 
    format="%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s"
)
logger = logging.getLogger(__name__)


def checkUnreadEmail(folder, target, player, logger):
    logger.warning("Start checking email")

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        japanFolder = outlook.folders(folder)
        japanInbox = japanFolder.folders(target)
        mails = japanInbox.Items

        unreadEmails = mails.Restrict("[UnRead] = true")

        if unreadEmails.Count > 0:
            if player.state() != QMediaPlayer.PlayingState:
                player.play()
            else:
                pass

    except Exception as ex:
        logger.critical(str(ex))
        player.play()


scheduler = QtScheduler()
scheduler.add_job(
    checkUnreadEmail,
    "interval",
    seconds=INTERVAL,
    args=[GROUP_FOLDER, TARGET_FOLDER, mediaPlayer, logger],
)
scheduler.start(paused=True)


def runClicked(scheduler, runBtn, stopBtn, stopSoundBtn):
    runBtn.setEnabled(False)
    stopBtn.setEnabled(True)
    stopSoundBtn.setEnabled(True)
    scheduler.resume()


def stopClicked(scheduler, runBtn, stopBtn, stopSoundBtn):
    runBtn.setEnabled(True)
    stopBtn.setEnabled(False)
    stopSoundBtn.setEnabled(False)
    scheduler.pause()


def stopSoundClicked(player, runBtn, stopBtn, stopSoundBtn):
    player.stop()


app = QApplication(sys.argv)
app.setQuitOnLastWindowClosed(False)

# Create the icon
icon = QIcon("bell.png")

clipboard = QApplication.clipboard()

# Create the tray
tray = QSystemTrayIcon()
tray.setIcon(icon)
tray.setVisible(True)

# Create the menu
menu = QMenu()

# Create actions.
# User will click this button when interaction with program
run = QAction("Run...")
stop = QAction("Stop Checking Email")
stopSound = QAction("Stop Sound")
exitProgram = QAction("Exit Program")

stop.setEnabled(False)
stopSound.setEnabled(False)

run.triggered.connect(lambda x: runClicked(scheduler, run, stop, stopSound))
stop.triggered.connect(lambda x: stopClicked(scheduler, run, stop, stopSound))
stopSound.triggered.connect(lambda x: stopSoundClicked(mediaPlayer, run, stop, stopSound))
exitProgram.triggered.connect(app.quit)

# Add actions to menu
menu.addAction(run)
menu.addAction(stop)
menu.addAction(stopSound)
menu.addAction(exitProgram)

# Add the menu to the tray
tray.setContextMenu(menu)

app.exec_()
