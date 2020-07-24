# Importing all the required libraries
import cv2
import threading
import random, time
import win32com.client as wincl
import curselist

# Using Haar cascade classifiers for object detection
face_cascade = cv2.CascadeClassifier(
    cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
)
smile_cascade = cv2.CascadeClassifier(
    "C:/Users/tyagi/AppData/Local/Programs/Python/Python37-32/Lib/site-packages/cv2/data/haarcascade_smile.xml"
)


# Creating an VideoCapture object used for image acquisition 
cap = cv2.VideoCapture(0)
# Setting frames per second
cap.set(cv2.CAP_PROP_FPS, 60)
# initializing text to speech
speak = wincl.Dispatch("SAPI.SpVoice")



class insult_smile:
    
    show_Boxes = True
    insult_count=0

    #Function to fetch random insults
    def curse(self):
        speak.Speak(random.choice(curselist.insults))


    #Function to detect face and smile
    def insult(self):

        while True:  
            ret, img = cap.read() # Capturing the video frame by frame (ret-> returns true/false if image is available & img-> image array vector)
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            faces = face_cascade.detectMultiScale(gray, scaleFactor=1.3, minNeighbors=5)

            for (x, y, w, h) in faces:
                if self.show_Boxes:
                    cv2.rectangle(img, (x, y), (x + w, y + h), (255, 0, 0), 2) #(image;startPoint as top,left;endPoint as bottom,right;colour as BGR;thickness)

                # Getting the region of intrests
                roi_gray = gray[y : y + h, x : x + h]
                roi_color = img[y : y + h, x : x + h]

                smiles = smile_cascade.detectMultiScale(
                    roi_gray, scaleFactor=1.8, minNeighbors=20
                ) # Increased minNeighbors to get more accuracy

                for (sx, sy, sw, sh) in smiles:
                    cv2.rectangle(roi_color, (sx, sy), (sx + sw, sy + sh), (0, 255, 255), 2)
                    if threading.active_count() < 5 :
                        #start insulting when smile detected
                        self.insult_count+=1
                        print("Here Comes insult",self.insult_count)
                        c = threading.Thread(target=self.curse)
                        c.start()
                

            # Display the resulting frame 
            cv2.imshow("img", img)
        
            # If escape key is pressed break from the loop
            k = cv2.waitKey(30) & 0xFF
            if k == 27:
                break

    
if __name__=="__main__":
    obj=insult_smile()
    
    # Enable threading
    t = threading.Thread(target=obj.insult)
    t.start()
    t.join()

    # Releasing the cap object 
    cap.release()
    # Destroy all the windows
    cv2.destroyAllWindows()
