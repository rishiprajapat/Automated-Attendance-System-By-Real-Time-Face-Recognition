import cv2 
  
# import Numpy 
import numpy as np 
  
# read a image using imread 
img = cv2.imread('abc.jpg') 
img = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY) 
  
# creating a Histograms Equalization 
# of a image using cv2.equalizeHist() 
equ = cv2.equalizeHist(img) 
  
# stacking images side-by-side 
# res = np.hstack((img, equ)) 
#final = cv2.medianBlur(equ, 3)
# show image input vs output 
cv2.imwrite('image1.jpg',equ) 

  
cv2.waitKey(0) 
cv2.destroyAllWindows() 