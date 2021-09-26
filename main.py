import cv2
import time
import numpy as np
import math
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from HandTrackingModule import HandDetector

##############
# Video Size
wCam, hCam = 800, 600
##############

cap = cv2.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)
pTime = 0

detector = HandDetector(detection_con=0.7)

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
# volume.GetMute()
# volume.GetMasterVolumeLevel()
volRange = volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]
# volume.SetMasterVolumeLevel(-20.0, None)


while True:
    success, img = cap.read()

    img = detector.find_hands(img)
    lmList = detector.find_position(img, draw=False)
    if len(lmList) != 0:
        # print(lmList[4], lmList[8])
        x1, y1 = lmList[4][1], lmList[4][2]
        x2, y2 = lmList[8][1], lmList[8][2]
        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
        # print(x1, y1, x2, y2)

        cv2.circle(img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)
        cv2.circle(img, (x2, y2), 15, (255, 0, 255), cv2.FILLED)
        cv2.line(img, (x1, y1), (x2, y2), (255, 0, 255), 3)
        cv2.circle(img, (cx, cy), 15, (255, 0, 255), cv2.FILLED)

        length = math.hypot(x2 - x1, y2 - y1)
        # print(length)
        # Hand Range is 20 to 220
        # Volume range is -65 - 0
        if length < 50:
            cv2.circle(img, (cx, cy), 15, (0, 255, 0), cv2.FILLED)
        vol = np.interp(length, [20, 220], [minVol, maxVol])
        # print(vol)
        volume.SetMasterVolumeLevel(vol, None)

        volBar = np.interp(length, [20, 220], [300, 150])
        volPercent = np.interp(length, [20, 220], [0, 100])
        cv2.rectangle(img, (50, 150), (75, 300), (0, 255, 0), 2)
        cv2.rectangle(img, (50, int(volBar)), (75, 300), (0, 255, 0), cv2.FILLED)
        cv2.putText(img, f"{int(volPercent)} %", (85, 320), cv2.FONT_HERSHEY_PLAIN, 2, (255, 0, 255), 2)

    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime

    cv2.putText(img, f"{int(fps)}", (20, 40), cv2.FONT_HERSHEY_PLAIN, 2, (255, 0, 255), 2)
    cv2.imshow("Img", img)

    if cv2.waitKey(1) & 0xFF == ord('q'):
        break
