#!/usr/bin/env python
# coding: utf-8

# In[1]:


from keras.preprocessing.image import img_to_array
from keras.models import load_model
import numpy as np
import argparse
import pickle
import cv2
import imutils
import os
import xlwt
import xlrd
import xlutils
from xlwt import Formula
from xlutils.copy import copy
import schedule 
import time 
detector = cv2.CascadeClassifier('haarcascade_frontalface_default.xml') #model detect face
model = load_model('3_model_data.h5') # model face recognition


#name_class = ["an","canh","chi","chung","cong","duy","hiep","hieu","huy","lan","nhung","phien","phuong","tam","thu"]
name_class=["delanie","hang","huisman","hung","linh","nam","phuong","qui","thu","tram"]
#name_class=["Chau","duy","lan","nam","nhung","phuong","qui","suong","thuy","tram"]
#def work_am():
p=1
camera = cv2.VideoCapture(0)
path="the_nhan_vien"

while True:
    rb=xlrd.open_workbook('bang1.xlsx')
    wb=copy(rb)
    w_sheet=wb.get_sheet(0)
    sheet = rb.sheet_by_index(0)
    (grabbed, frame) = camera.read() # doc frame video


    frame = imutils.resize(frame, width=500) #hinh thi kich thuoc webcam
    frame_clone = frame.copy() #copy ra 1 ban de xu ly sao nay

    # detect face
    faces = detector.detectMultiScale(frame, scaleFactor=1.1, minNeighbors=5, minSize=(30, 30),
                                      flags=cv2.CASCADE_SCALE_IMAGE)

    # Loop over the face bounding boxes
    for (fX, fY, fW, fH) in faces:

        roi = frame[fY:fY + fH, fX:fX + fW] # cat khuon mat
        roi = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY) # chuyen anh thanh anh xam
        roi = cv2.resize(roi, (200, 200)) #resize khich thuoc
        roi = roi.astype("float") / 255.0 # chuan hoa
        roi = img_to_array(roi) # chuyen thanh mang
        roi = np.expand_dims(roi, axis=0) # them 1 chieu vi dau vao CNN la (None,hight,with,dept)


        predictions = model.predict(roi) # du doan
        i = predictions.argmax(axis=1)[0] #lay chi so cao nhat
        accur = "{:.2f}%".format(predictions[0][i]*100) #tinh do chinh xac
        print("do chinh xac: " + accur)    
        result = predictions[0][i]
        if result > 0.8: # thay doi do chinh xac o day neu predict lon hon 80% thi hien ten nguoc lai hien unknown
            text = "{}".format(name_class[i]) #lay ten class
    # Show our detected faces along with smiling/not smiling labels
            img1 = cv2.imread(os.path.join(path,"{}".format(str(i) + ".png")), cv2.IMREAD_COLOR)
            winname="anh_nhan_vien"
            cv2.namedWindow(winname)        # Create a named window
            cv2.moveWindow(winname, 1000,70)  # Move it to (40,30)
            cv2.imshow(winname,img1)
        else:
            text = "Unknown"
            i=10
            img1 = cv2.imread(os.path.join(path,"{}".format(str(i) + ".png")), cv2.IMREAD_COLOR)
            winname="anh_nhan_vien"
            cv2.namedWindow(winname)        # Create a named window
            cv2.moveWindow(winname, 1000,70)  # Move it to (40,30)
            cv2.imshow(winname,img1)
        #hien thi ten class
        cv2.putText(frame_clone, text, (fX, fY - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.45, (0, 0, 255), 2)
        # dong khung mat
        cv2.rectangle(frame_clone, (fX, fY), (fX + fW, fY + fH), (0, 0, 255), 2)
        if cv2.waitKey(1) & 0xFF == ord("a"):
            if i==0:
                a=1+sheet.cell_value(5, 3)
                w_sheet.write(5,3,a)
                w_sheet.write(5,5,Formula('IF(D6>=E6,E6,D6)'))
                w_sheet.write(5,7,Formula('266000*F6+G6'))

            elif i==1:
                a=1+sheet.cell_value(6, 3)
                w_sheet.write(6,3,a)
                w_sheet.write(6,5,Formula('IF(D7>=E7,E7,D7)'))
                w_sheet.write(6,7,Formula('266000*F7+G7'))
                
                
            elif i==2:
                a=1+sheet.cell_value(7, 3)
                w_sheet.write(7,3,a)
                w_sheet.write(7,5,Formula('IF(D8>=E8,E8,D8)'))
                w_sheet.write(7,7,Formula('266000*F8+G8'))
                
            elif i==3:
                a=1+sheet.cell_value(8, 3)
                w_sheet.write(8,3,a)
                w_sheet.write(8,5,Formula('IF(D9>=E9,E9,D9)'))
                w_sheet.write(8,7,Formula('266000*F9+G9')) 
                            
                
            elif i==4:
                a=1+sheet.cell_value(9, 3)
                w_sheet.write(9,3,a)
                w_sheet.write(9,5,Formula('IF(D10>=E10,E10,D10)'))
                w_sheet.write(9,7,Formula('266000*F10+G10'))
                
            elif i==5:
                a=1+sheet.cell_value(10, 3)
                w_sheet.write(10,3,a)
                w_sheet.write(10,5,Formula('IF(D11>=E11,E11,D11)'))
                w_sheet.write(10,7,Formula('266000*F11+G11'))
                
            elif i==6:
                a=1+sheet.cell_value(11, 3)
                w_sheet.write(11,3,a)
                w_sheet.write(11,5,Formula('IF(D12>=E12,E12,D12)'))
                w_sheet.write(11,7,Formula('266000*F12+G12'))
                
            elif i==7:
                a=1+sheet.cell_value(12, 3)
                w_sheet.write(12,3,a)
                w_sheet.write(12,5,Formula('IF(D13>=E13,E13,D13)'))
                w_sheet.write(12,7,Formula('266000*F13+G13'))
                
            elif i==8:
                a=1+sheet.cell_value(13, 3)
                w_sheet.write(13,3,a)
                w_sheet.write(13,5,Formula('IF(D14>=E14,E14,D14)'))
                w_sheet.write(13,7,Formula('266000*F14+G14'))
                
            elif i==9:
                a=1+sheet.cell_value(14, 3)
                w_sheet.write(14,3,a)
                w_sheet.write(14,5,Formula('IF(D15>=E15,E15,D15)'))
                w_sheet.write(14,7,Formula('266000*F15+G15'))
            wb.save('bang1.xlsx')
                             

    cv2.imshow("Face", frame_clone)

    # If the 'q' key is pressed, stop the loop
    if cv2.waitKey(1) & 0xFF == ord("q"):
        #if p==2:
        break

# Cleanup the camera and close any open windows
camera.release()
cv2.destroyAllWindows()


# In[ ]:




