import cv2

class SimpleProcess:
    def __init__(self,widht,height): # chieu rong chieu cao anh
        self.width = widht
        self.height = height

    def process(self,image):

        return cv2.resize(image,(self.width,self.height))