import cv2
import numpy as np
import pytesseract
import os

# C:\Program Files\Tesseract-OCR

per = 25
pixelThreshold = 50
roi = [[(314,637), (1065,697), 'text', 'Name'],
       [(154,844), (508,901), 'text', 'ID'],
       [(585,845), (1062,905), 'text', 'Mobile'],
       [(154,1047), (742,1105), 'text', 'Branch'],
       [(841,1049), (1030,1107), 'text', 'Admission Year'],
       [(400,1185), (1082,1243), 'text', 'Email'],
       [(400,1297), (1081,1355), 'text', 'Address'],
       [(400,1390), (1081,1448), 'text', 'Pincode'],
       [(189,1460), (232,1502), 'box', 'Declaration']]

with open('DataOutput.csv', 'a+') as f:
    for data in roi:
        f.write((str(data[3]) + ','))
    f.write('\n')

pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

imgQ = cv2.imread('Query0.jpg')
h,w,c = imgQ.shape
imgQ = cv2.resize(imgQ,(w,h))

orb = cv2.ORB_create(1200)
kp1, des1 = orb.detectAndCompute(imgQ,None)
# imgkp1 = cv2.drawKeypoints(imgQ,kp1,None)

path = "UseForms"
myPickList = os.listdir(path)
print(myPickList)


for j,y in enumerate(myPickList):
    img = cv2.imread(path + "/" + y)
    # cv2.imshow(y, img)
    kp2, des2 = orb.detectAndCompute(img, None)
    bf = cv2.BFMatcher(cv2.NORM_HAMMING)
    matches = bf.match(des2, des1)
    matches=tuple(sorted(matches,key = lambda x:x.distance))
    good = matches[:int(len(matches)*(per/100))]
    imgMatch = cv2.drawMatches(img,kp2,imgQ,kp1, good[:10],None,flags=2)
    # cv2.imshow(y, imgMatch)

    srcPoints = np.float32([kp2[m.queryIdx].pt for m in good]).reshape(-1,1,2)
    dstPoints = np.float32([kp1[m.trainIdx].pt for m in good]).reshape(-1,1,2)

    M,_ = cv2.findHomography(srcPoints, dstPoints, cv2.RANSAC, 5.0)
    imgScan = cv2.warpPerspective(img, M,(w,h))
    # cv2.imshow(y, imgScan)

    imgShow = imgScan.copy()
    imgMask = np.zeros_like(imgShow)

    myData = []
    print(f'############## Extraction Data from Form {j+1} ############')

    for x,r in enumerate(roi):

        cv2.rectangle(imgMask, (r[0][0],r[0][1]),(r[1][0],r[1][1]),(0,255,0),cv2.FILLED)
        imgShow = cv2.addWeighted(imgShow,0.99,imgMask,0.1,0)

        imgCrop = imgScan[r[0][1]:r[1][1], r[0][0]:r[1][0]]
        # cv2.imshow(str(x),imgCrop)
        # print(pytesseract.image_to_string(imgCrop))

        if r[2] == 'text':
            print('{} :{}'.format(r[3], pytesseract.image_to_string(imgCrop)))
            str1 = pytesseract.image_to_string(imgCrop)
            str1=str1.replace(",","")
            myData.append(str1.rstrip())

        if r[2] == 'box':
            imgGray = cv2.cvtColor(imgCrop,cv2.COLOR_BGR2GRAY)
            imgThresh = cv2.threshold(imgGray,170,255,cv2.THRESH_BINARY_INV)[1]
            totalPixels = cv2.countNonZero(imgThresh)
            # print(totalPixels)
            if totalPixels>pixelThreshold: totalPixels = 1;
            else: totalPixels = 0
            print(f'{r[3]} : {totalPixels}')
            myData.append(totalPixels)
        cv2.putText(imgShow,str(myData[x])+" ",(r[0][0],r[0][1]),cv2.FONT_HERSHEY_PLAIN,2,(0,0,255),2)
    with open('DataOutput.csv','a+') as f:
        for data in myData:
            f.write((str(data)+','))
        f.write('\n')

    imgShow = cv2.resize(imgShow,(w//3,h//3))
    print(myData)
    cv2.imshow(y+"2",imgShow)
# cv2.imshow("Key points for query",imgkp1)
# cv2.imshow("Output",imgQ)
cv2.waitKey(0)









