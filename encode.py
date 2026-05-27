
import cv2
import face_recognition
import os
import pickle

IMAGES_FOLDER = "images"   # folder name yahi rakho
ENCODE_FILE = "EncodeFile.p"

images = []
classNames = []

if not os.path.exists(IMAGES_FOLDER):
    raise FileNotFoundError(f"{IMAGES_FOLDER} folder nahi mila")

myList = os.listdir(IMAGES_FOLDER)
print("Images:", myList)

for cl in myList:
    img_path = os.path.join(IMAGES_FOLDER, cl)
    curImg = cv2.imread(img_path)

    if curImg is None:
        print(f"Skipping invalid image: {cl}")
        continue

    images.append(curImg)
    classNames.append(os.path.splitext(cl)[0])

print("Names:", classNames)

def findEncodings(images):
    encodeList = []
    validNames = []

    for img, name in zip(images, classNames):
        img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        encodes = face_recognition.face_encodings(img)

        if len(encodes) > 0:
            encodeList.append(encodes[0])
            validNames.append(name)
        else:
            print(f"No face found in: {name}")

    return encodeList, validNames

print("Encoding Started...")
encodeListKnown, validClassNames = findEncodings(images)

with open(ENCODE_FILE, 'wb') as file:
    pickle.dump([encodeListKnown, validClassNames], file)

print("Encoding Complete ✅")
