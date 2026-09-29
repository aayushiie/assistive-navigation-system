import numpy as np
import cv2 as cv
import time
import threading
import queue
import csv
import torch
import pandas as pd
import pythoncom
import win32com.client

# load MiDaS 
midas = torch.hub.load("intel-isl/MiDaS", "MiDaS_small")
midas.eval()

transforms = torch.hub.load("intel-isl/MiDaS", "transforms")
transform = transforms.small_transform

# TTS 
tts_queue = queue.Queue(maxsize=1)

def tts_worker():
    pythoncom.CoInitialize()
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Volume = 100

    while True:
        text, rate = tts_queue.get()
        if text is None:
            break
        speaker.Rate = rate
        speaker.Speak(text, 0)
        tts_queue.task_done()

    pythoncom.CoUninitialize()

threading.Thread(target=tts_worker, daemon=True).start()

cooldown = {}

def speak(text, cls, rate=0):
    now = time.time()
    if now - cooldown.get(cls, 0) < 3:
        return
    cooldown[cls] = now

    try:
        tts_queue.put_nowait((text, rate))
    except:
        pass

# priority
PRIORITY = {"person":1.0,"car":1.0,"bus":1.0,"truck":1.0,"dog":0.6,"chair":0.3}

def get_priority(cls):
    return PRIORITY.get(cls,0.4)

# distance
REAL_HEIGHTS = {"person":1.7,"car":1.5,"bus":3.0,"truck":2.5,"dog":0.5}

def estimate_distance(cls, h):
    F = 615
    H = REAL_HEIGHTS.get(cls,1.0)
    return (F*H)/(h+1e-3)

# motion
prev_pos = {}

def compute_motion(cls, center):
    if cls in prev_pos:
        m = np.linalg.norm(np.array(center)-np.array(prev_pos[cls]))
    else:
        m = 0
    prev_pos[cls]=center
    return m

# danger 
def danger(dist, motion, cls):
    return 0.5*(1/(dist+1e-3)) + 0.3*motion + 0.2*get_priority(cls)

# position 
def get_pos(x,w):
    if x<w*0.33: return "left"
    elif x<w*0.66: return "center"
    return "right"

# CSV 
csv_file=open("results.csv","w",newline="")
writer=csv.writer(csv_file)
writer.writerow(["class","mono","depth_raw","depth_cal","gt"])

# calibration
depth_samples=[]
real_samples=[]
k=1.0

# load model
with open("coco.names") as f:
    classes=[c.strip() for c in f]

net=cv.dnn.readNetFromTensorflow(
    "frozen_inference_graph.pb",
    "ssd_mobilenet_v3_large_coco_2020_01_14.pbtxt"
)

cap=cv.VideoCapture(0)

print("Keys: s=save | c=calibrate | q=quit")

# main loop
while True:
    ret,frame=cap.read()
    if not ret: break

    h,w=frame.shape[:2]

    # MiDaS 
    img=cv.cvtColor(frame,cv.COLOR_BGR2RGB)
    inp=transform(img)

    with torch.no_grad():
        pred=midas(inp)
        pred=torch.nn.functional.interpolate(
            pred.unsqueeze(1),
            size=frame.shape[:2],
            mode="bicubic",
            align_corners=False
        ).squeeze()

    depth_map=pred.cpu().numpy()

    # detection 
    blob=cv.dnn.blobFromImage(cv.resize(frame,(300,300)),0.007843,(300,300),127.5)
    net.setInput(blob)
    det=net.forward()

    objs=[]

    for i in range(det.shape[2]):
        conf=det[0,0,i,2]
        if conf>0.5:
            cid=int(det[0,0,i,1])-1
            if cid<0 or cid>=len(classes): continue

            cls=classes[cid]

            x1=int(det[0,0,i,3]*w)
            y1=int(det[0,0,i,4]*h)
            x2=int(det[0,0,i,5]*w)
            y2=int(det[0,0,i,6]*h)

            bh=y2-y1
            center=((x1+x2)//2,(y1+y2)//2)

            # safe indexing
            cx = min(max(center[0],0), w-1)
            cy = min(max(center[1],0), h-1)

            # distances
            mono=estimate_distance(cls,bh)
            depth=depth_map[cy, cx]
            depth_cal=k * depth   

            # motion + danger
            m=compute_motion(cls,center)
            score=danger(mono,m,cls)
            pos=get_pos(center[0],w)

            objs.append((score,cls,mono,depth,depth_cal,pos))

            cv.rectangle(frame,(x1,y1),(x2,y2),(0,255,0),2)
            cv.putText(frame,f"{cls} {mono:.1f}m",(x1,y1-10),
                       cv.FONT_HERSHEY_SIMPLEX,0.5,(0,255,0),2)

    if objs:
        objs.sort(reverse=True)
        score,cls,mono,depth,depth_cal,pos=objs[0]

        # TTS
        if score>1.5:
            speak(f"Warning {cls} very close on your {pos}",cls,3)
        elif score>0.8:
            speak(f"{cls} approaching on your {pos}",cls,1)
        else:
            speak(f"{cls} on your {pos}",cls,-1)

    cv.imshow("Assistive System",frame)

    key=cv.waitKey(1)&0xFF

    # save 
    if key==ord('s') and objs:
        gt=float(input("Enter distance (meters): "))
        writer.writerow([cls,mono,depth,depth_cal,gt])
        print("Saved")

    # calibrate
    if key==ord('c') and objs:
        real=float(input("Enter distance (meters): "))
        depth_samples.append(depth)
        real_samples.append(real)

        if len(depth_samples)>=5:
            k=np.mean(np.array(real_samples)/np.array(depth_samples))
            print("Updated scale k =",k)

    if key==ord('q'):
        break

# clean csv
csv_file.close()
cap.release()
cv.destroyAllWindows()

# metrics calculation
data=pd.read_csv("results.csv")

mae_mono=np.mean(abs(data["mono"]-data["gt"]))
rmse_mono=np.sqrt(np.mean((data["mono"]-data["gt"])**2))

mae_depth=np.mean(abs(data["depth_cal"]-data["gt"]))
rmse_depth=np.sqrt(np.mean((data["depth_cal"]-data["gt"])**2))

print("\n=== RESULTS ===")
print("Monocular MAE:",mae_mono)
print("Monocular RMSE:",rmse_mono)
print("Calibrated Depth MAE:",mae_depth)
print("Calibrated Depth RMSE:",rmse_depth)
