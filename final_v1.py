import numpy as np
import cv2 as cv
import time
import threading
import queue
import pythoncom
import win32com.client

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

tts_thread = threading.Thread(target=tts_worker, daemon=True)
tts_thread.start()

object_cooldown = {}

def speak(text, cls, rate=-2):
    now = time.time()
    if now - object_cooldown.get(cls, 0) < 3.0:
        return
    object_cooldown[cls] = now
    try:
        tts_queue.put_nowait((text, rate))
    except queue.Full:
        pass

PRIORITY = {
    "person": 1.0, "bicycle": 0.9, "car": 1.0, "bus": 1.0, "truck": 1.0,
    "dog": 0.6, "cat": 0.5,
    "chair": 0.3, "couch": 0.3,
    "bottle": 0.2
}

def get_priority(cls):
    return PRIORITY.get(cls, 0.4)

REAL_HEIGHTS = {
    "person": 1.7, "car": 1.5, "truck": 2.5, "bus": 3.0,
    "bicycle": 1.2, "chair": 1.0, "dog": 0.5
}

def estimate_distance(cls, box_height):
    FOCAL_LENGTH = 615
    real_height = REAL_HEIGHTS.get(cls, 1.0)
    if box_height == 0:
        return 999
    return (FOCAL_LENGTH * real_height) / box_height

previous_positions = {}

def compute_motion(cls, center):
    if cls in previous_positions:
        prev = previous_positions[cls]
        motion = np.linalg.norm(np.array(center) - np.array(prev))
    else:
        motion = 0

    previous_positions[cls] = center
    return motion

def compute_danger(distance, motion, cls):
    priority = get_priority(cls)

    distance_score = 1 / (distance + 1e-3)

    w1, w2, w3 = 0.5, 0.3, 0.2

    score = w1 * distance_score + w2 * motion + w3 * priority
    return score

def get_position(x_center, frame_width):
    if x_center < frame_width * 0.33:
        return "left"
    elif x_center < frame_width * 0.66:
        return "center"
    else:
        return "right"

with open("coco.names", "r") as f:
    classes = [line.strip() for line in f.readlines()]

net = cv.dnn.readNetFromTensorflow(
    "frozen_inference_graph.pb",
    "ssd_mobilenet_v3_large_coco_2020_01_14.pbtxt"
)

cap = cv.VideoCapture(0)

while True:
    ret, frame = cap.read()
    if not ret:
        break

    height, width = frame.shape[:2]

    blob = cv.dnn.blobFromImage(
        cv.resize(frame, (300, 300)),
        0.007843, (300, 300), 127.5
    )

    net.setInput(blob)
    detections = net.forward()

    objects = []

    for i in range(detections.shape[2]):
        conf = detections[0, 0, i, 2]
        if conf > 0.5:

            class_id = int(detections[0, 0, i, 1]) - 1
            if class_id < 0 or class_id >= len(classes):
                continue

            cls = classes[class_id]

            x1 = int(detections[0, 0, i, 3] * width)
            y1 = int(detections[0, 0, i, 4] * height)
            x2 = int(detections[0, 0, i, 5] * width)
            y2 = int(detections[0, 0, i, 6] * height)

            box_height = y2 - y1
            center = ((x1 + x2)//2, (y1 + y2)//2)

            distance = estimate_distance(cls, box_height)
            motion = compute_motion(cls, center)
            score = compute_danger(distance, motion, cls)

            position = get_position(center[0], width)

            objects.append({
                "class": cls,
                "distance": distance,
                "score": score,
                "position": position
            })

            cv.rectangle(frame, (x1, y1), (x2, y2), (0,255,0), 2)
            cv.putText(frame,
                f"{cls} {distance:.1f}m S:{score:.2f}",
                (x1, y1-10),
                cv.FONT_HERSHEY_SIMPLEX, 0.5, (0,255,0), 2
            )

    if objects:
        objects = sorted(objects, key=lambda x: x["score"], reverse=True)
        top = objects[0]

        cls = top["class"]
        dist = top["distance"]
        pos = top["position"]
        score = top["score"]

        if score > 1.5:
            msg = f"Warning! {cls} very close on your {pos}"
            speak(msg, cls, rate=3)
        elif score > 0.8:
            msg = f"{cls} approaching on your {pos}"
            speak(msg, cls, rate=1)
        else:
            msg = f"{cls} on your {pos}, {dist:.1f} meters"
            speak(msg, cls, rate=-1)

    cv.imshow("Assistive Navigation", frame)

    if cv.waitKey(1) & 0xFF == ord('q'):
        break

tts_queue.put((None, None))
tts_thread.join()
cap.release()
cv.destroyAllWindows()