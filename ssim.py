import cv2 as cv
import tensorflow as tf
from tensorflow import keras
from keras.applications import MobileNet
from keras.applications.mobilenet import preprocess_input, decode_predictions
import numpy as np
import win32com.client
from skimage.metrics import structural_similarity as ssim

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

# preprocessing and predict
def process_frame(frame, model):
    img = cv.resize(frame, (224,224))
    img = cv.cvtColor(img, cv.COLOR_BGR2RGB)
    img = np.expand_dims(img, axis=0)
    img = preprocess_input(img)

    predictions = model.predict(img, verbose=0)
    decoded = decode_predictions(predictions, top=3)[0]

    return decoded

# webcam
vid = cv.VideoCapture(0)

if not vid.isOpened():
    print("camera not found")
    exit()

prev_gray = None
frame_count = 0

while True:
    ret, frame = vid.read()

    if not ret or frame is None:
        break

    frame_count += 1

    # convert to gray for SSIM
    gray = cv.cvtColor(frame, cv.COLOR_BGR2GRAY)

    run_prediction = True

    if prev_gray is not None:
        score, _ = ssim(prev_gray, gray, full=True)

        # if frames almost identical skip prediction
        if score > 0.95:
            run_prediction = False

    prev_gray = gray

    if run_prediction:
        predictions = process_frame(frame, model)

        print(f"\nFrame {frame_count}")

        y_offset = 30
        for imagenet_id, label, score in predictions:
            if score >= 0.3:

                print(f"{label}: {score:.2%}")
                text = f"{label}: {score:.2%}"

                cv.putText(
                    frame,
                    text,
                    (10, y_offset),
                    cv.FONT_HERSHEY_SIMPLEX,
                    0.8,
                    (0,0,255),
                    2
                )

                y_offset += 30

                # speak only when prediction changes
                speaker.Speak(label)

    cv.imshow("Video", frame)

    if cv.waitKey(1) & 0xFF == ord('q'):
        break

vid.release()
cv.destroyAllWindows()