
import sys, os
from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMenu
from PyQt5.QtGui import QIcon
from pynput import keyboard
from threading import Thread
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from pygame import mixer
from time import sleep
from tendo import singleton
from operator import truediv

# Files
soundSwitchOff = "assets/switchOff.mp3"
soundSwitchOn = "assets/switchOn.mp3"
micOnIcon = "assets/iconMicOn.png"
micOffIcon = "assets/iconMicOff.png"
micOnIcon2 = "assets/iconMicOn2.png"
micOffIcon2 = "assets/iconMicOff2.png"
exitIcon = "assets/iconExit.png"

# ---------- Misc --------

def soundPlayer(file):
    global currentSoundThread,mixer,soundVolume
    mixer.music.load(file)
    mixer.music.set_volume(soundVolume)
    mixer.music.play()

def playSound(file):
    global currentSoundThread
    if mixer.music:
        mixer.music.stop()
    currentSoundThread = Thread(target=soundPlayer, args=[file])
    currentSoundThread.start()

# ---------- Mic Functions --------

def getDeviceVolume():
    import pythoncom
    pythoncom.CoInitialize()
    devices = AudioUtilities.GetMicrophone()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    return cast(interface, POINTER(IAudioEndpointVolume))

def setMicMuted(status):
    getDeviceVolume().SetMute(status, None)

def isMuted():
    return getDeviceVolume().GetMute()

def setMutedIcon(muted):
    global trayIcon
    if muted:
        trayIcon.setIcon(QIcon(micOffIcon))
        togleMicAction.setIcon(QIcon(micOnIcon2))
    else: 
        trayIcon.setIcon(QIcon(micOnIcon))
        togleMicAction.setIcon(QIcon(micOffIcon2))

def muteMicrophone():
    print("Turning microphone On...")
    setMicMuted(True)
    setMutedIcon(True)
    playSound(soundSwitchOff)

def unmuteMicrophone():
    print("Turning microphone OFF...")
    setMicMuted(False)
    setMutedIcon(False)
    playSound(soundSwitchOn)

def toggleMicrophone():
    if isMuted(): unmuteMicrophone()
    else: muteMicrophone()

# ---------- Key Listener --------

def execute(keyFunc):
    action = keyFunkOptions.get(keyFunc)
    if(action): keyFunkOptions[keyFunc]()

def keyName(key):
    import pynputDictionary
    nKey = ""
    if isinstance(key, keyboard.Key):
        nKey = "Key."+key.name
    elif isinstance(key, keyboard.KeyCode):
        try:
            nKey = "Key."+key.char
        except:  
            nKey = key
    if nKey == "":
        print("Error trying to get the type of the key... " + type(key))
        nKey = "<?>"
    return nKey

def onPress(key):
    global currentKeys
    key = keyName(key)
    if key == "" or key == "<?>": return
    currentKeys.add(key)
    for keyFunc in combinations:
        if set(combinations[keyFunc]) == currentKeys:
            try:
                execute(keyFunc)
            except:
               print("error trying to add the key '{0}'".format(key))

def onRelease(key):
    global currentKeys
    key = keyName(key)
    if key not in currentKeys: return
    try:
        currentKeys.remove(key)
    except:
        print("error trying to remove the key '{0}'".format(key))

def keyListenerThreadFunction():
    global keyboardListener
    import pythoncom
    pythoncom.CoInitialize()
    keyboardListener = keyboard.Listener(on_press=onPress, on_release=onRelease)  
    keyboardListener.start()
    
def initHotKeys():
    global combinations, currentKeys
    combinations = {}
    currentKeys = set()
    import pynputDictionary
    listFromFile = {
        "muteKeys": "Key.ctrl_l Key.f1",
        "toggleMicKeys": "Key.ctrl_l Key.f2",
        "unmuteKeys": "Key.ctrl_l Key.f3",
    }
    
    for keyFunc in listFromFile:
        stringKeys = listFromFile[keyFunc].split()
        keysCombination = []
        for k in stringKeys:
            if len(k) == 1:
                try:
                    key = keyboard.KeyCode(char=k)
                except:
                    print("Error converting '", k , " in to a key")   
            else:
                try:
                    key = pynputDictionary.keyDictionary.get(k)
                except:
                    print("Error converting '", k , " in to a key")   
            if key:
                keysCombination.append(k)
        combinations[keyFunc] = keysCombination

def initKeyListener():
    global currentKeys, keyFunkOptions
    currentKeys = set()
    keyFunkOptions = {
        "muteKeys" : muteMicrophone,
        "toggleMicKeys" : toggleMicrophone,
        "unmuteKeys" : unmuteMicrophone
    }
    initHotKeys()
    startListener()

def startListener():
    keyboardListener = keyboard.Listener(on_press=onPress, on_release=onRelease)  
    keyboardListener.start()

# ---------- Tray Icon --------

def quitApp():
    import pythoncom
    pythoncom.CoInitialize()
    print("Exiting...")
    trayIcon.hide()
    os._exit(0)

def initTrayApp():
    global app, trayIcon, togleMicAction, unmuteAction, muteAction, exitAction
    app = QApplication(sys.argv)
    trayIcon = QSystemTrayIcon(QIcon(micOnIcon), parent=app)
    trayIcon.setToolTip("MicToggler")
    trayIcon.show()

    menu = QMenu()

    togleMicAction = menu.addAction("TogleMic")
    togleMicAction.triggered.connect(toggleMicrophone)
    togleMicAction.setIcon(QIcon(micOnIcon2))

    unmuteAction = menu.addAction("Unmute")
    unmuteAction.triggered.connect(unmuteMicrophone)
    unmuteAction.setIcon(QIcon(micOnIcon2))

    muteAction = menu.addAction("Mute")
    muteAction.triggered.connect(muteMicrophone)
    muteAction.setIcon(QIcon(micOffIcon2))

    exitAction = menu.addAction("Exit")
    exitAction.triggered.connect(quitApp)
    exitAction.setIcon(QIcon(exitIcon))

    trayIcon.setContextMenu(menu)

def initialIcon():
    sleep(0.1)
    global trayIcon
    if isMuted(): trayIcon.setIcon(QIcon(micOffIcon))
    else: trayIcon.setIcon(QIcon(micOnIcon))

# ---------- Manteniment --------

def singleInstance():
    global me
    try:
        me = singleton.SingleInstance()
    except:
        import pyautogui
        from sys import exit
        msg = "MicToggle is already started"
        print(msg)
        pyautogui.alert(msg)
        os._exit(0)
        
def trdChecker():
    global keyboardListener
    sleep(0.1)
    initialIcon()

# ---------- Start --------
def startConfigs():
    global soundVolume
    soundVolume = 0.35

if __name__ == "__main__":
    global app, trdThreadChecker
    print("Starting MicToggler...")
    singleInstance()
    startConfigs()
    initTrayApp()
    initKeyListener()
    mixer.init()
    trdThreadChecker = Thread(target=trdChecker)
    trdThreadChecker.start()
    exit(app.exec())
