import tkinter as tk
from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
from PIL import Image, ImageTk
import numpy as np
import cv2

import pytesseract
from PIL import ImageFont, ImageDraw 


class CornerBox():
    def __init__(self, pos, canvas, c_height, c_width):
        if pos == 'NW':
            self.x, self.y = 5, 5
        elif pos == 'NE':
            self.x, self.y = c_width-5, 5
        elif pos == 'SE':
            self.x, self.y = c_width-5, c_height-5
        elif pos == 'SW':
            self.x, self.y = 5, c_height-5

        self.c_height, self.c_width = c_height, c_width
        self.canvas = canvas

        self.cb_id = canvas.create_rectangle(
            self.x-5, self.y-5, self.x+5, self.y+5, fill="blue")

        self.reset_x, self.reset_y = self.x, self.y

        # canvas.tag_bind(self.cb_id, "<Button-1>", self.grab)
        # canvas.tag_bind(self.cb_id, "<B1-Motion>", self.drag)

    def grab(self, event):
        widget = event.widget
        self.x = widget.canvasx(event.x)
        self.y = widget.canvasy(event.y)
        self.widget = self.cb_id

    def drag(self, event):
        widget = event.widget
        xc = widget.canvasx(event.x)
        yc = widget.canvasy(event.y)
        if xc < 5 or yc < 5 or xc > self.c_width-5 or yc > self.c_height-5:
            return

        self.canvas.move(self.widget, xc-self.x, yc-self.y)
        self.x, self.y = xc, yc

        x1, y1, x2, y2 = self.canvas.coords(self.cb_id)
        self.x = (x1+x2)/2
        self.y = (y1+y2)/2

    @property
    def coords(self):
        return (self.x, self.y)

    @property
    def id(self):
        return self.cb_id

    def reset(self):
        xc, yc = self.reset_x, self.reset_y
        self.canvas.move(self.cb_id, xc-self.x, yc-self.y)
        self.x, self.y = xc, yc

        x1, y1, x2, y2 = self.canvas.coords(self.cb_id)
        self.x = (x1+x2)/2
        self.y = (y1+y2)/2

    @coords.setter
    def coords(self, xc, yc):
        if xc < 5 or yc < 5 or xc > self.c_width-5 or yc > self.c_height-5:
            return
        self.canvas.move(self.cb_id, xc-self.x, yc-self.y)
        self.x, self.y = xc, yc

        x1, y1, x2, y2 = self.canvas.coords(self.cb_id)
        self.x = (x1+x2)/2
        self.y = (y1+y2)/2


class CropApp:
    def __init__(self, master, img):
        self.master = master

        self.screen_width = root.winfo_screenwidth()
        self.screen_height = root.winfo_screenheight()

        self.img_file_name = img
        self.im = Image.open(img)
        self.img_width, self.img_height = self.im.size

        f1 = 0
        f2 = 0
        self.scale_factor = 1

        if self.img_width > self.screen_width:
            f1 = self.img_width // self.screen_width

        if self.img_height > self.screen_height:
            f2 = self.img_height // self.screen_height

        if f1 != 0 or f2 != 0:
            self.scale_factor = max(f1, f2)
            self.scale_factor += 1
            self.im = self.im.resize(
                (self.img_width//self.scale_factor, self.img_height//self.scale_factor), Image.ANTIALIAS)
            self.img_width, self.img_height = self.im.size

        self.c_width = self.img_width + 10
        self.c_height = self.img_height + 10

        self.canvas = tk.Canvas(
            self.master, width=self.c_width, height=self.c_height)
        self.canvas.pack()

        self.img_canvas = ImageTk.PhotoImage(self.im)
        self.img_canvas_img_id = self.canvas.create_image(
            5, 5, image=self.img_canvas, anchor=tk.NW)

        self.NW = CornerBox('NW', self.canvas, self.c_height, self.c_width)
        self.NE = CornerBox('NE', self.canvas, self.c_height, self.c_width)
        self.SE = CornerBox('SE', self.canvas, self.c_height, self.c_width)
        self.SW = CornerBox('SW', self.canvas, self.c_height, self.c_width)

        self.but_frame = tk.Frame(self.master)
        self.but_frame.pack()

        # self.coord_butt = tk.Button(
        #     self.but_frame, text="Show Cordinates", command=self.printBoxDetails)
        # self.coord_butt.pack(side=tk.LEFT)

        self.reset_butt = tk.Button(
            self.but_frame, text="Reset", command=self.restCorners)
        self.reset_butt.pack(side=tk.LEFT)

        self.crop_butt = tk.Button(
            self.but_frame, text="Crop", command=self.cropImage)
        self.crop_butt.pack(side=tk.LEFT)

        self.save = tk.Button(
            self.but_frame, text="Save", command=self.saveImage)
        self.save.pack(side=tk.LEFT)

        self.box_id = None
        self.p1_id = None
        self.p2_id = None
        self.p3_id = None
        self.p4_id = None

        self.drawBox()

        for i in [self.NW, self.NE, self.SE, self.SW]:
            self.canvas.tag_bind(i.cb_id, "<Button-1>", i.grab)
            self.canvas.tag_bind(i.cb_id, "<B1-Motion>", i.drag)

    def printBoxDetails(self):
        print(self.NW.coords, self.NE.coords, self.SE.coords, self.SW.coords)

    def restCorners(self):
        self.NW.reset()
        self.NE.reset()
        self.SE.reset()
        self.SW.reset()

    def cropImage(self):
        A = self.NW.coords
        B = self.NE.coords
        C = self.SE.coords
        D = self.SW.coords

        A = (A[0]*self.scale_factor, A[1]*self.scale_factor)
        B = (B[0]*self.scale_factor, B[1]*self.scale_factor)
        C = (C[0]*self.scale_factor, C[1]*self.scale_factor)
        D = (D[0]*self.scale_factor, D[1]*self.scale_factor)

        pts1 = np.float32(
            [list(A),
             list(B),
             list(D),
             list(C)]
        )

        widthA = np.sqrt(((D[0] - C[0])
                          ** 2) + ((D[1] - C[1]) ** 2))
        widthB = np.sqrt(((A[0] - B[0])
                          ** 2) + ((A[1] - B[1]) ** 2))
        maxWidth = max(int(widthA), int(widthB))

        heightA = np.sqrt(((A[0] - D[0]) ** 2) + (
            (A[1] - D[1]) ** 2))
        heightB = np.sqrt(((B[0] - C[0]) ** 2) + (
            (B[1] - C[1]) ** 2))
        maxHeight = max(int(heightA), int(heightB))

        pts2 = np.float32(
            [[0, 0],
             [maxWidth, 0],
             [0, maxHeight],
             [maxWidth, maxHeight]]
        )

        img = cv2.imread(self.img_file_name)
        matrix = cv2.getPerspectiveTransform(pts1, pts2)
        self.res = cv2.warpPerspective(img, matrix, (maxWidth, maxHeight))
        resized_image = cv2.resize(
            self.res, (0, 0), fx=1/self.scale_factor, fy=1/self.scale_factor)
        cv2.imshow("Result", resized_image)


    def saveImage(self):
        cv2.imwrite('result.jpg', self.res)

        cv2.destroyAllWindows()

        # Images
        uncropped = filename
        cropped = 'result.jpg'

        # Group Images
        img1 = cv2.imread(uncropped)
        img2 = cv2.imread(cropped)
        h1, w1 = img1.shape[:2]
        h2, w2 = img2.shape[:2]
        img3 = np.zeros((h1+h2, w1,3), dtype=np.uint8)
        img3[:,:] = (255,255,255)
        img3[:h1,:w1,:3]=img1
        img3[h1:h1+h2,:w2,:3]= img2

        # Img 4
        img4 = np.zeros((h1+h2, w1 + 800,3), dtype=np.uint8)
        img4[:,:] = (255,255,255)
        img4[:h1,:w1,:3]=img1
        img4[h1:h1+h2,w1-w2:w1,:3]= img2


        # Show
        ##cv.imshow("Cropped/Uncropped", img3)


        # Pytesseract
        pytesseract.pytesseract.tesseract_cmd = r'C:\Users\efu\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
        uncropped_ocr = pytesseract.image_to_string(Image.open(uncropped), config='--psm 1')# --oem 4')        
        cropped_ocr = pytesseract.image_to_string(Image.open(cropped), config='--psm 1')# --oem 4')        

        # OCR
        print('\n\n################################################################')
        print('#  Reading text from Uncropped Image:')
        print('################################################################')
        print('-----------------------------------------------------------------')
        print(uncropped_ocr)
        print('-----------------------------------------------------------------')


        print('\n\n################################################################')
        print('#  Reading text from Cropped Image:')
        print('################################################################')
        print('-----------------------------------------------------------------')
        print(cropped_ocr)
        print('-----------------------------------------------------------------')

        # PIL
        pil_img = Image.fromarray(img4)
        font = ImageFont.truetype("arial.ttf", 24)
        spacing = 5
        draw = ImageDraw.Draw(pil_img)


        # Write OCR
        draw.text((w1+20, int(h1/2)-50), uncropped_ocr, fill='black', font=font, spacing=spacing, align='left')
        draw.text((w1+20, int(h2+(h1/2))+50), cropped_ocr, fill='black', font=font, spacing=spacing, align='left')

        np_img = np.array(pil_img)
        cv2.imshow('Cropped/Uncropped', np_img)

        

    def drawBox(self, event=None):
        if self.box_id != None:
            self.canvas.delete(self.box_id)
        if self.p1_id != None:
            self.canvas.delete(self.p1_id)
        if self.p2_id != None:
            self.canvas.delete(self.p2_id)
        if self.p3_id != None:
            self.canvas.delete(self.p3_id)
        if self.p4_id != None:
            self.canvas.delete(self.p4_id)

        self.box_id = self.canvas.create_line(
            *self.NW.coords,
            *self.NE.coords,
            *self.SE.coords,
            *self.SW.coords,
            *self.NW.coords,
            fill="blue", width=2)

        self.p1_id = self.canvas.create_polygon(
            0, 0, *self.NW.coords, *self.SW.coords, 0, self.c_height, fill="blue", stipple="gray25")
        self.p2_id = self.canvas.create_polygon(
            0, 0, *self.NW.coords, *self.NE.coords, self.c_width, 0, fill="blue", stipple="gray25")
        self.p3_id = self.canvas.create_polygon(
            self.c_width, 0, *self.NE.coords, *self.SE.coords, self.c_width, self.c_height, fill="blue", stipple="gray25")
        self.p4_id = self.canvas.create_polygon(
            self.c_width, self.c_height, *self.SE.coords, *self.SW.coords, 0, self.c_height, fill="blue", stipple="gray25")

        self.canvas.tag_raise(self.NW.cb_id, 'all')
        self.canvas.tag_raise(self.NE.cb_id, 'all')
        self.canvas.tag_raise(self.SE.cb_id, 'all')
        self.canvas.tag_raise(self.SW.cb_id, 'all')

        self.canvas.after(1, self.drawBox)


if __name__ == "__main__":
    root = tk.Tk()
    filename = askopenfilename()
    img_file_name = filename
    
    App = CropApp(root, img_file_name)

    root.mainloop()
