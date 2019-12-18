import cv2
import numpy as np
import os

class dataload:
    def __init__(self,process=None):
        self.process = process
        if self.process is None:
            self.process=[]

    def load(self,ImagePaths):
        data= []
        labels= []
        print("load image......")
        for ImagePath in ImagePaths:

            image = cv2.imread(ImagePath,cv2.IMREAD_GRAYSCALE)
            label = ImagePath.split(os.path.sep)[-2]
            if self.process is not None:
                for p in self.process:
                    image = p.process(image)

            data.append(image)
            labels.append(label)

        return (np.array(data),np.array(labels)) # tuple anh va nhan tuong ung
