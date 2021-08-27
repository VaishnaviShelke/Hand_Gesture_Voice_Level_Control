import math
import cv2
import numpy as np
import HandTracking as htm
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume





devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
# volume.GetMute()
# volume.GetMasterVolumeLevel()
volRange = volume.GetVolumeRange()
# volume.SetMasterVolumeLevel(0, None)
minVol=volRange[0]
maxVol=volRange[1]
vol=0
volBar=400
volPer=0

wCam,hCam=640,480
detector =htm.handDetector(detectionCon=0.7)


cap=cv2.VideoCapture(0)
cap.set(3,wCam)
cap.set(4,hCam)

while True:
    sucess , img =cap.read()
    img=detector.findHands(img)
    lmList = detector.findPosition(img ,draw=False)
    if len(lmList)!=0:
        # print(lmList[4],lmList[8])
        x1,y1 = lmList[4][1],lmList[4][2]
        x2, y2 = lmList[8][1], lmList[8][2]
        cx,cy = (x1+x2)//2 , (y1+y2)//2

        cv2.circle(img, (x1,y1),10,(255,0,255),cv2.FILLED)
        cv2.circle(img, (x2, y2), 10, (255, 0, 255), cv2.FILLED)
        cv2.circle(img, (cx, cy),10, (255, 0, 255), cv2.FILLED)
        cv2.line(img, (x1,y1) , (x2,y2),(255,0,255),3)

        length = math.hypot(x2-x1,y2-y1)
        # print(length)
        # Hand Range is in 50 - 200
        # Volume Range is in 65 - 0
        vol=np.interp(length,[50,200],[minVol,maxVol])
        volBar=np.interp(length,[50,200],[400,150])
        volPer = np.interp(length, [50, 200], [0, 100])

        # print(int (length),vol)
        volume.SetMasterVolumeLevel(vol, None)


        if length<50:
            cv2.circle(img, (cx, cy), 10, (0, 255, 0), cv2.FILLED)
        if length > 200:
            cv2.circle(img, (x1, y1), 10, (0, 255, 0), cv2.FILLED)
            cv2.circle(img, (x2, y2), 10, (0, 255, 0), cv2.FILLED)
        cv2.putText(img ,f'{int (volPer)}%',(00,450),cv2.FONT_HERSHEY_COMPLEX,1,(255,0,0),3)
        cv2.rectangle(img,(10,150),(40,400),(255,0,0),3)
        cv2.rectangle(img, (10, int(volBar)), (40, 400), (255,0,0), cv2.FILLED)
    cv2.imshow("Img",img)
    k = cv2.waitKey(1)
    # print(k)
    if k == 13:  #
        cv2.destroyAllWindows()
        break
