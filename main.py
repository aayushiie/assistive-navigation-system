import cv2 as cv
import tensorflow as tf
from tensorflow import keras
from keras.applications import MobileNet
from keras.applications.mobilenet import preprocess_input, decode_predictions
import numpy as np
import win32com.client

# load model
model = MobileNet(
    weights='imagenet',
    include_top=True,
    classes=1000,
    classifier_activation='softmax'
)

# text-to-speech 
speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Rate = 0
speaker.Volume = 100
last_spoken = "" # to avoid reinitialisation

# preprocessing and predict
def process_frame(frame, model):
    frame = cv.resize(frame, (224,224))
    frame = cv.cvtColor(frame, cv.COLOR_BGR2RGB)
    frame = np.expand_dims(frame, axis=0)
    frame = preprocess_input(frame)

    predictions = model.predict(frame, verbose=0)
    decoded = decode_predictions(predictions, top=3)[0]

    return decoded

# webcam
vid = cv.VideoCapture(0)

if not vid.isOpened():
    print("camera not found")
    exit()

frame_count = 0

while True:
    ret, frame = vid.read()

    if not ret or frame is None:
        break

    frame_count+=1

    predictions = process_frame(frame, model)

    print(f"\nFrame {frame_count}")

    # draw predictions
    y_offset=30
    for imagenet_id, label, score in predictions:
        if score>=0.3: #30% confidence threshold
            print(f" {label}: {score:.2%}")
            text = f"{label}: {score:.2%}"
            cv.putText(
                frame, 
                text,
                (10, y_offset),
                cv.FONT_HERSHEY_SIMPLEX,
                0.8,
                (0, 0, 255),
                2)
            y_offset+=30

            #t2s
            # speaker.speak(label)
            if label != last_spoken:
                speaker.Speak(label, 1)
                last_spoken = label
            
    cv.imshow('Video', frame)

    if cv.waitKey(1) & 0xff==ord('q'):
        break

vid.release()
cv.destroyAllWindows()