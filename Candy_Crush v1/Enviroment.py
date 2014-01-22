import os, sys
import Image, ImageGrab, ImageOps
import time, random
from random import randrange
import win32api
import win32com.client,win32con
from numpy import *
import time
from PIL import Image
import operator
from Board import Board
from win32api import GetSystemMetrics

class ExternalEnviroment(object):

    def __init__(self):
        self.baseUrl = 'http://media.pog.com/system/contents/54288/original/candy-crush.swf?1361832848'
        self.shell =  win32com.client.Dispatch("WScript.Shell")
        self.screen = Screen()
        self.mouse = Mouse()
        #print('Hello world from External Controller')
   
    def getBaseUrl(self):
        return self.baseUrl

    def setBaseUrl(self,url):
        self.baseUrl = url

    def setShell(self,sh):
        self.shell = sh

    def getSheel(self):
        return self.shell

    def runApp(self,app):
        try:
            self.getSheel().Run(app)
          
            win32api.Sleep(1000)
            self.getSheel().AppActivate(app)
            win32api.Sleep(1000)
            for c in self.getBaseUrl():
                self.getSheel().SendKeys(c)
                win32api.Sleep(10)
            self.getSheel().SendKeys("{ENTER}")
            win32api.Sleep(100)
            self.getSheel().SendKeys("{F11}")
            win32api.Sleep(100)

        except Exception as e:
            raise Exception('Error on reunning browser.',e)
    
    def delay(self,seconds):
        for i in range(0,seconds/1000):
            time.sleep(1)

    def getScreen(self):
        return self.screen

    def setScreen(self,sc):
        self.screen = sc

    def getMousen(self):
        return self.mouse

    def setMouse(self,ms):
        self.mouse = ms

class Screen:
   
    def __init__(self):
        self.Top = 12
        self.Buttom = 582
        self.Right = 742
        self.Left = 108
   
    def getArea(self):
        return Screen()

    def captureScreen(self):
        return ImageGrab.grab()

    def captureArea(self,area = None):
        if area is None:
            area = self
        box = (area.Left,area.Top,area.Right,area.Buttom)
        return ImageGrab.grab(box)
   
    def getTop(self):
        return self.Top

    def getButtom(self):
        return self.Buttom

    def getRight(self):
        return self.Right

    def getLeft(self):
        return self.Left

    def getParts(self,input,height,width,i,k):
      
      try:
        board = Board()
        im = Image.open(input)
        imgwidth, imgheight = im.size
        x =0
        y = 0
        temp1 =0

        yAxis =0
        for i in range(0,imgheight,height):
            y =0
            temp2 =0     
            xAxis =0
            yAxis =(x+1)*width
            for j in range(0,imgwidth,width):
                box = (j+temp2, i-temp1, j+width, (i+abs(temp1))+height)
                a = im.crop(box)
                aw,ah = a.size
                pixel = a.getpixel((aw/2,ah/2))
                #if pixel == (0,0,0):
                    #print('empty image')
                if pixel != (0,0,0):
                    xAxis =(y+1)*height
                    name = os.getcwd()+'\\parts\\'+str(x)+str(y)+'.png'
                    a.save(name)
                    board.getItems()[x][y].setImName(name)
                    board.getItems()[x][y].setX(self.getLeft()+xAxis)
                    board.getItems()[x][y].setY(self.getTop()+yAxis)
                    board.getItems()[x][y].setWidth(width)
                    board.getItems()[x][y].setHeigh(height)
                    y +=1
                    k +=1
                    temp2+=1
                
            temp1 -=3
            x +=1
      except Exception as e:
            raise Exception('Error in partionaing image.',e)
      return board
        
    def comapre(self,im1 , im2):
       try:
           h1 = im1.histogram()
           h2 = im2.histogram()

           rms = math.sqrt(reduce(operator.add,
              map(lambda a,b: (a-b)**2, h1, h2))/len(h1))
           return rms
       except Exception as e:
            raise Exception('Error compring to images.',e)
       
    def comparePix(self,im1,im2):
        try:
            p1 = im1.getpixel((22,23))  
            p2 = im2.getpixel((22,23))  
        
            t1 = math.sqrt(reduce(operator.add,
                  map(lambda a,b: (a-b)**2, p1, p2))/len(p1))

            p3 = im1.getpixel((49,49)) 
            p4 = im2.getpixel((49,49))  
        
            t2 = math.sqrt(reduce(operator.add,
                  map(lambda a,b: (a-b)**2, p3, p4))/len(p3))

            if p1 == p2 and p3 == p4:
                return 1
            else:
                return 0
        except Exception as e:
               raise Exception('Error compring to pixsels.',e)   
          
    def compareSe(self,im1,im2):

        try:
            aw,ah = im1.size
            l1 = im1.crop(((aw/2)-4,(ah/2)-4,((aw/2)+4),(ah/2)+4))
            aw,ah = im2.size
            l2 = im2.crop(((aw/2)-4,(ah/2)-4,((aw/2)+4),(ah/2)+4))
            h1= l1.histogram()
            h2= l2.histogram()
        
            rms = math.sqrt(reduce(operator.add,
                  map(lambda a,b: (a-b)**2, h1, h2))/len(h1))

            return rms
        except Exception as e:
            raise Exception('Error compring to images.',e)

class Mouse(object):

    def __init__(self, *args, **kwargs):
        return super(Mouse, self).__init__(*args, **kwargs)

    def clickLeft(self,x,y):
        self.move(x,y)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
        time.sleep(0.5)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)

    def clickRight(self,x,y):
        self.move(x,y)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,x,y,0,0)
        time.sleep(0.5)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,x,y,0,0)

    def move(self,x,y):
        win32api.SetCursorPos((x,y))
    
    def hide(self):
        self.move(GetSystemMetrics(0),GetSystemMetrics(1))