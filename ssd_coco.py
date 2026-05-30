import numpy as np
import cv2 as cv
import win32com.client


speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Rate = 0
speaker.Volume = 100
last_spoken = ""

# minimum confidence level for detection
min_confidence = 0.5

with open("coco.names", "r") as f:
    classes = [line.strip() for line in f.readlines()]

# bounding box colors
np.random.seed(543210)
colors = np.random.uniform(0, 255, size=(len(classes), 3))

# load network
net = cv.dnn.readNetFromTensorflow("frozen_inference_graph.pb", "ssd_mobilenet_v3_large_coco_2020_01_14.pbtxt")

# webcam
cap = cv.VideoCapture(0)

while True:
    ret, image = cap.read()
    if not ret:
        break

    height, width = image.shape[:2]

    # create blob (binary large object)
    blob = cv.dnn.blobFromImage(
        cv.resize(image, (300, 300)),
        0.007843,
        (300, 300),
        127.5
    )
    # new_pixel = (pixel - 127.5) * 0.007843  centering data around 0

    # feed blob to network
    net.setInput(blob)
    detected_objects = net.forward()

    for i in range(detected_objects.shape[2]):
        confidence = detected_objects[0, 0, i, 2]  # [image_id, class_id, confidence, x1, y1, x2, y2]

        if confidence > min_confidence:
            class_index = int(detected_objects[0, 0, i, 1]) - 1 

            upper_left_x = int(detected_objects[0, 0, i, 3] * width)
            upper_left_y = int(detected_objects[0, 0, i, 4] * height)
            lower_right_x = int(detected_objects[0, 0, i, 5] * width)
            lower_right_y = int(detected_objects[0, 0, i, 6] * height)

            # convert to percentage
            prediction_text = classes[class_index]

            cv.rectangle(
                image,
                (upper_left_x, upper_left_y),
                (lower_right_x, lower_right_y),
                colors[class_index],
                2
            )

            cv.putText(
                image,
                f"{prediction_text}: {confidence*100:.1f}%",
                (upper_left_x, upper_left_y - 10 if upper_left_y > 20 else upper_left_y + 20),
                cv.FONT_HERSHEY_SIMPLEX,
                0.6,
                colors[class_index],
                2
            )

            #t2s
            # speaker.speak(prediction_text)
            if prediction_text != last_spoken:
                speaker.Speak(prediction_text, 1)
                last_spoken = prediction_text

    cv.imshow("Detected Objects", image)

    if cv.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv.destroyAllWindows()