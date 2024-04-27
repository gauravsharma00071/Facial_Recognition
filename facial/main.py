from tkinter import *
from tkinter import ttk, messagebox
from PIL import Image,ImageTk
from tkinter.simpledialog import askstring
from tkinter.messagebox import showinfo
from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import streamlit as st
import pandas as pd
import time
from datetime import datetime
#from test import FaceRecognition

from win32com.client import Dispatch
#from add_faces import addfaces




class Face_Recognition_System:
    def __init__(self, root):
        self.root=root
        self.root.geometry("1250x670+0+0")
        self.root.title("Face Recoginition System")

        headLine = Label(root, text="Silence is Dangerous",
                         font=("times new roman", 30, "bold"), bg="white", fg="blue")
        headLine.place(x=0, y=5, width=1250, height=45)



#background image
        img3 = Image.open(r"images/bg.jpg")
        img3 = img3.resize((1250, 870), Image.MESH)
        self.photoimg3 = ImageTk.PhotoImage(img3)
        bg_img = Label(self.root, image=self.photoimg3)
        bg_img.place(x=0, y=130, width=1250, height=870)

        title_lbl=Label(bg_img,text="FACE RECOGNITION ATTENDANCE SYSTEM SOFTWARE",font=("times new roman",30,"bold"),bg="white",fg="red")
        title_lbl.place(x=0,y=0,width=1250,height=45)

#adding Face Details on bg image
        img4 = Image.open(r"images/add.png")
        img4 = img4.resize((200, 200), Image.MESH)
        self.photoimg4=ImageTk.PhotoImage(img4)

        b1=Button(bg_img,image=self.photoimg4,cursor="hand2",command=self.adding_FaceData)
        b1.place(x=150,y=100,width=200,height=200)

        b1_1 = Button(bg_img, text="Add Face Details",command=self.adding_FaceData, cursor="hand2",font=("times new roman",14,"bold"),bg="white",fg="blue")
        b1_1.place(x=150, y=300, width=200, height=40)



#adding Face Detector Button on bg image
        img5 = Image.open(r"images/faceDetect.png")
        img5 = img5.resize((200, 200), Image.MESH)
        self.photoimg5=ImageTk.PhotoImage(img5)

        b2=Button(bg_img,image=self.photoimg5,command=self.testing_face, cursor="hand2")
        b2.place(x=370,y=100,width=200,height=200)

        b2_1 = Button(bg_img, text="Detect Face",command=self.testing_face, cursor="hand2",font=("times new roman",14,"bold"),bg="white",fg="blue")
        b2_1.place(x=370, y=300, width=200, height=40)

# adding Attendance Button on bg image
        img6 = Image.open(r"images/attendance.png")
        img6 = img6.resize((200, 200), Image.MESH)
        self.photoimg6 = ImageTk.PhotoImage(img6)

        b3 = Button(bg_img, image=self.photoimg6, cursor="hand2",command=self.attendanceCheck)
        b3.place(x=590, y=100, width=200, height=200)

        b3_1 = Button(bg_img, text="Attendance", cursor="hand2",command=self.attendanceCheck, font=("times new roman", 14, "bold"), bg="white",
                      fg="blue")
        b3_1.place(x=590, y=300, width=200, height=40)

# adding Help Desk Button on bg image
        img7 = Image.open(r"images/exit.png")
        img7 = img7.resize((200, 200), Image.MESH)
        self.photoimg7 = ImageTk.PhotoImage(img7)

        b4 = Button(bg_img, image=self.photoimg7, cursor="hand2",command=root.destroy)
        b4.place(x=810, y=100, width=200, height=200)

        b4_1 = Button(bg_img, text="Exit", cursor="hand2",command=root.destroy, font=("times new roman", 14, "bold"), bg="white",
                      fg="blue")
        b4_1.place(x=810, y=300, width=200, height=40)

        name_lbl = Label(bg_img, text="Gaurav Sharma Software Developer",
                          font=("times new roman", 30, "bold"), bg="white", fg="red")
        name_lbl.place(x=0, y=460, width=1250, height=45)


#=====================functions button===============

    def adding_FaceData(self):
        self.app = Toplevel(self.root)
        #self.app = Train(self.new__window)
        video = cv2.VideoCapture(0)
        facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

        faces_data = []

        i = 0
        name = askstring('Name', 'What is your name?')
        showinfo('Hello!', 'Hi, {}'.format(name))

        while True:
            ret, frame = video.read()
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = facedetect.detectMultiScale(gray, 1.3, 5)
            for (x, y, w, h) in faces:
                crop_img = frame[y:y + h, x:x + w, :]
                resized_img = cv2.resize(crop_img, (50, 50))
                if len(faces_data) <= 100 and i % 10 == 0:
                    faces_data.append(resized_img)
                i = i + 1
                cv2.putText(frame, str(len(faces_data)), (50, 50), cv2.FONT_HERSHEY_COMPLEX, 1, (50, 50, 255), 1)
                cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 1)
            cv2.imshow("Frame", frame)
            k = cv2.waitKey(1)
            if k == ord('q') or len(faces_data) == 100:
                break
        video.release()
        cv2.destroyAllWindows()
        messagebox.showinfo("Successfully","Your face data has been successfully detected")
        #print("")

        faces_data = np.asarray(faces_data)
        faces_data = faces_data.reshape(100, -1)

        if 'names.pkl' not in os.listdir('data/'):
            names = [name] * 100
            with open('data/names.pkl', 'wb') as f:
                pickle.dump(names, f)
        else:
            with open('data/names.pkl', 'rb') as f:
                names = pickle.load(f)
            names = names + [name] * 100
            with open('data/names.pkl', 'wb') as f:
                pickle.dump(names, f)

        if 'faces_data.pkl' not in os.listdir('data/'):
            with open('data/faces_data.pkl', 'wb') as f:
                pickle.dump(faces_data, f)
        else:
            with open('data/faces_data.pkl', 'rb') as f:
                faces = pickle.load(f)
            faces = np.append(faces, faces_data, axis=0)
            with open('data/faces_data.pkl', 'wb') as f:
                pickle.dump(faces, f)
    def testing_face(self):
        self.new__window = Toplevel(self.root)

        def speak(str1):
            speak = Dispatch(("SAPI.SpVoice"))
            speak.Speak(str1)

        video = cv2.VideoCapture(0)
        facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

        with open('data/names.pkl', 'rb') as w:
            LABELS = pickle.load(w)
        with open('data/faces_data.pkl', 'rb') as f:
            FACES = pickle.load(f)

        knn = KNeighborsClassifier(n_neighbors=5)
        knn.fit(FACES, LABELS)

        imgBackground = cv2.imread("background.png")

        COL_NAMES = ['NAME', 'TIME']

        exist = False  # Move the exist variable declaration outside the loop
        while True:
            ret, frame = video.read()
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = facedetect.detectMultiScale(gray, 1.3, 5)
            for (x, y, w, h) in faces:
                crop_img = frame[y:y + h, x:x + w, :]
                resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
                output = knn.predict(resized_img)
                ts = time.time()
                date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
                timestamp = datetime.fromtimestamp(ts).strftime("%H:%M-%S")
                exist = os.path.isfile("Attendance/Attendance_" + date + ".csv")
                cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 1)
                cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 2)
                cv2.rectangle(frame, (x, y - 40), (x + w, y), (50, 50, 255), -1)
                cv2.putText(frame, str(output[0]), (x, y - 15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
                cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 1)
                attendance = [str(output[0]), str(timestamp)]
            imgBackground[162:162 + 480, 55:55 + 640] = frame
            cv2.imshow("Frame", imgBackground)
            k = cv2.waitKey(1)
            if k == ord('o'):
                speak("Attendance Taken..")
                time.sleep(5)
                if exist:
                    with open("Attendance/Attendance_" + date + ".csv", "+a") as csvfile:
                        writer = csv.writer(csvfile)
                        writer.writerow(attendance)
                else:
                    with open("Attendance/Attendance_" + date + ".csv", "+a") as csvfile:
                        writer = csv.writer(csvfile)
                        writer.writerow(COL_NAMES)
                        writer.writerow(attendance)
            if k == ord('q'):
                break
        video.release()
        cv2.destroyAllWindows()
        #self.app = FaceRecognition(self.new__window)

    def attendanceCheck(self):
        messagebox.showinfo("Attendance","You are not authorized to check attendance. Check Your Attendance in Your Database and Excel File.")
        '''
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M-%S")

        from streamlit_autorefresh import st_autorefresh

        count = st_autorefresh(interval=2000, limit=100, key="fizzbuzzcounter")

        if count == 0:
            st.write("Count is zero")
        elif count % 3 == 0 and count % 5 == 0:
            st.write("FizzBuzz")
        elif count % 3 == 0:
            st.write("Fizz")
        elif count % 5 == 0:
            st.write("Buzz")
        else:
            st.write(f"Count: {count}")

        df = pd.read_csv("Attendance/Attendance_" + date + ".csv")

        st.dataframe(df.style.highlight_max(axis=0))
        '''



if __name__ == "__main__":
    root=Tk()
    obj=Face_Recognition_System(root)
    root.mainloop()
