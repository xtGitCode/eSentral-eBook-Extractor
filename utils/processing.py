# https://github.com/yardstick17/image_text_reader

import cv2
import numpy as np
from PIL import Image

def preprocessing(img):
  img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
  #resize
  resized = cv2.resize(img, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)
  # blur
  blur = cv2.GaussianBlur(resized, (0,0), sigmaX=33, sigmaY=33)
  # divide
  divide = cv2.divide(resized, blur, scale=255) 
  #sharpen
  kernel = np.array([[0, -1, 0], [-1,5,-1],[0,-1,0]])
  sharpened_image = cv2.filter2D(divide, -1, kernel)
  #remove noise
  filtered = cv2.adaptiveThreshold(sharpened_image, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 41,3)
  kernel = np.ones((1, 1), np.uint8)
  opening = cv2.morphologyEx(filtered, cv2.MORPH_OPEN, kernel)
  closing = cv2.morphologyEx(opening, cv2.MORPH_CLOSE, kernel)
  or_image = cv2.bitwise_or(sharpened_image, closing)
  return or_image

def splitTwo(img):
  img = cv2.imread(img)
  h,w,c = img.shape
  w1 = w / -2
  w1= int(w1)
  crop_img1 = img[60:-20, 48:w1] 
  crop_img2 = img[60:-20, w1:-40]
  return crop_img1, crop_img2