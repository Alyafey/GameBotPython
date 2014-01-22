import os, sys
import Image, ImageGrab, ImageOps
import time, random
from random import randrange
import win32api
import win32com.client,win32con
from numpy import *
from PIL import Image
import operator
from win32api import GetSystemMetrics
from Board  import Board,Square
from PIL import ImageChops
import win32gui
from DesktopWindow import DesktopWindow

class CandyCrush(object):     

    def __init__(self):
        self.Url ='http://media.pog.com/system/contents/54288/original/candy-crush.swf?1361832848'
        self.App = 'chrome.exe'
        self.StartX = 107
        self.StartY = 17
        self.isExit=False
        self.Window = DesktopWindow()

    def compare(self,im1 , im2):
       try:
           x =im1-im2
           x = math.pow(x,2)
           x = x/100
           return x == 0
       except Exception as e:
            raise Exception('Error compring to images.',e)

    def getPixel(self,x,y):
         return self.Window.get_pixel_color(x,y)
   
    def getBoard(self,im):
      
      try:
        board = Board()
        y = self.StartY
        for i in range(0,9):
            x = self.StartX
            for j in range(0,9):
                xx = x+35
                yy = y+40
                c=0
                try:
                    pix = im[xx,yy]
                    c += pix[0]
                    c += pix[1]
                    c += pix[2]
                except Exception as e:
                    c=0
                board.getItems()[i][j].setX(x)
                board.getItems()[i][j].setY(y)
                board.getItems()[i][j].setColor(c)
                x += 71
            y+=63
      except Exception as e:
            raise Exception('Error in partionaing image.',e)
      return board
                
    def startApp(self):
        shell =  win32com.client.Dispatch("WScript.Shell")
        shell.Run(self.App)
        win32api.Sleep(1000)
        shell.AppActivate(self.App)
        win32api.Sleep(1000)

        shell.SendKeys(self.Url)
        shell.SendKeys("{ENTER}")
        win32api.Sleep(100)
        shell.SendKeys("{F11}")
        time.sleep(26)

    def grabImage(self):
        return ImageGrab.grab()

    def move(self,x,y):
        win32api.SetCursorPos((x,y))
    
    def hideMouse(self):
        self.move(GetSystemMetrics(0),GetSystemMetrics(1))

    def clickLeft(self,x,y):
        self.move(x,y)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
        time.sleep(0.1)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)

    def findMoves(self,board):
        r1 = 0
        r2 = 0
        for i in reversed(range(r1,9)):
            for j in  range(r1,9):
                    #if(j<8):
                    #  c1 = self.compare(board.getItems()[i][j].getColor(),200)
                    #  return ((board.getItems()[i][j].getX(),board.getItems()[i][j].getY()),
                    #                (board.getItems()[i][j+1].getX(),board.getItems()[i][j+1].getY()))
                #start condition for rows
                    # first condistions
                    if(i <=7 and j <=6):

                        c1 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i][j+1].getColor())
                        c2 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i+1][j+2].getColor())
                        if( c1   and c2 ):
                           return ((board.getItems()[i+1][j+2].getX(),board.getItems()[i+1][j+2].getY()),
                                   (board.getItems()[i][j+2].getX(),board.getItems()[i][j+2].getY()))

                        c1 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i+1][j+1].getColor())
                        c2 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i][j+2].getColor())
                        if( c1   and c2 ):
                           return ((board.getItems()[i+1][j+1].getX(),board.getItems()[i+1][j+1].getY()),   
                                   (board.getItems()[i][j+1].getX(),board.getItems()[i][j+1].getY()))

                        c1 = self.compare(board.getItems()[i+1][j].getColor(),board.getItems()[i][j+1].getColor())
                        c2 = self.compare(board.getItems()[i+1][j].getColor(),board.getItems()[i][j+2].getColor())
                        if( c1   and c2 ):
                           return ((board.getItems()[i+1][j].getX(),board.getItems()[i+1][j].getY()),
                                   (board.getItems()[i][j].getX(),board.getItems()[i][j].getY()))

                     #second conditions
                    if(i >0 and j <=6):
                        c1 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i][j+1].getColor())
                        c2 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i-1][j+2].getColor())
                        if( c1   and c2 ):
                           return ((board.getItems()[i-1][j+2].getX(),board.getItems()[i-1][j+2].getY()),
                                   (board.getItems()[i][j+2].getX(),board.getItems()[i][j+2].getY()))

                        c1 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i-1][j+1].getColor())
                        c2 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i][j+2].getColor())
                        if( c1   and c2 ):
                           return ((board.getItems()[i-1][j+1].getX(),board.getItems()[i-1][j+1].getY()),
                                   (board.getItems()[i][j+1].getX(),board.getItems()[i][j+1].getY()))

                        c1 = self.compare(board.getItems()[i-1][j].getColor(),board.getItems()[i][j+1].getColor())
                        c2 = self.compare(board.getItems()[i-1][j].getColor(),board.getItems()[i][j+2].getColor())
                        if( c1   and c2 ):
                           return ((board.getItems()[i-1][j].getX(),board.getItems()[i-1][j].getY()),
                                   (board.getItems()[i][j].getX(),board.getItems()[i][j].getY()))

                     #third conditions
                    if(j <=5):
                        c1 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i][j+1].getColor())
                        c2 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i][j+3].getColor())
                        if( c1   and c2 ):
                           return ((board.getItems()[i][j+3].getX(),board.getItems()[i][j+3].getY()),
                                   (board.getItems()[i][j+2].getX(),board.getItems()[i][j+2].getY()))

                        c1 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i][j+2].getColor())
                        c2 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i][j+3].getColor())
                        if( c1   and c2 ):
                           return ((board.getItems()[i][j].getX(),board.getItems()[i][j].getY()),
                                   (board.getItems()[i][j+1].getX(),board.getItems()[i][j+1].getY()))

                    #end conditions for rows
                    #start conditons for coulmns

                    if(i<=6 and j<8):
                        c1 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i+1][j].getColor())
                        c2 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i+2][j+1].getColor())
                        if( c1   and c2 ):
                           return ((board.getItems()[i+2][j+1].getX(),board.getItems()[i+2][j+1].getY()),
                                   (board.getItems()[i+2][j].getX(),board.getItems()[i+2][j].getY()))

                        c1 = self.compare(board.getItems()[i][j+1].getColor(),board.getItems()[i+1][j+1].getColor())
                        c2 = self.compare(board.getItems()[i][j+1].getColor(),board.getItems()[i+2][j].getColor())
                        if( c1   and c2 ):
                           return ((board.getItems()[i+2][j].getX(),board.getItems()[i+2][j].getY()),
                                   (board.getItems()[i+2][j+1].getX(),board.getItems()[i+2][j+1].getY()))

                        c1 = self.compare(board.getItems()[i][j+1].getColor(),board.getItems()[i+1][j].getColor())
                        c2 = self.compare(board.getItems()[i][j+1].getColor(),board.getItems()[i+2][j+1].getColor())
                        if( c1   and c2 ):
                           return ((board.getItems()[i+1][j].getX(),board.getItems()[i+1][j].getY()),
                                   (board.getItems()[i+1][j+1].getX(),board.getItems()[i+1][j+1].getY()))
                       
                        c1 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i+1][j+1].getColor())
                        c2 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i+2][j].getColor())
                        if( c1   and c2 ):
                           return ((board.getItems()[i+1][j+1].getX(),board.getItems()[i+1][j+1].getY()),
                                   (board.getItems()[i+1][j].getX(),board.getItems()[i+1][j].getY()))  

                        c1 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i+1][j+1].getColor())
                        c2 = self.compare(board.getItems()[i][j].getColor(),board.getItems()[i+2][j+1].getColor())
                        if( c1   and c2 ):
                           return ((board.getItems()[i][j].getX(),board.getItems()[i][j].getY()),
                                   (board.getItems()[i][j+1].getX(),board.getItems()[i][j+1].getY()))  

                        c1 = self.compare(board.getItems()[i][j+1].getColor(),board.getItems()[i+1][j].getColor())
                        c2 = self.compare(board.getItems()[i][j+1].getColor(),board.getItems()[i+2][j].getColor())
                        if( c1   and c2 ):
                           return ((board.getItems()[i][j+1].getX(),board.getItems()[i][j+1].getY()),
                                   (board.getItems()[i][j].getX(),board.getItems()[i][j].getY())) 

                     #end conditons for coulmns

    def isLoaded(self,b):
        c = self.comapre(b.getItems()[7][3].getImg(),b.getItems()[7][4].getImg())
        c2 = self.comapre(b.getItems()[7][3].getImg(),b.getItems()[7][5].getImg())
        if( c and c2 ):
           print ('Still loading....')
           return False
        else:
           print ('Game Loaded....')
           return True

    def go(self):
        self.hideMouse()
        self.startApp()
        print('Starting program!......')
        while True:
            if(self.isExit == True):
                break
            im =self.grabImage().load()
            #time.clock()
            b = self.getBoard(im)
            #print(time.clock())
            moves = self.findMoves(b)
            if(moves != None):
                self.clickLeft(moves[0][0]+30,moves[0][1]+30)
                time.sleep(.15)
                self.clickLeft(moves[1][0]+30,moves[1][1]+30)