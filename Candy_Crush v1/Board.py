class Board(object):

    def __init__(self):
        self.items = [[Square() for x in xrange(9)] for x in xrange(9)]

    def getItems(self):
        return self.items

    def setItems(self,coll):
        self.items = coll
        

class Square(object):

    def __init__(self):
        self.X = 0
        self.Y = 0
        self.width = 0
        self.heigh = 0
        self.color=0
        self.img =[]
   
    def getX(self):
        return self.X

    def getY(slef):
        return slef.Y

    def getName(self):
        return self.imName
     
    def setX(self,xNew):
        self.X = xNew

    def setY(self,yNew):
        self.Y = yNew

    def setImName(self,im):
        self.imName = im

    def getWidth(self):
        return self.width

    def getHeigh(self):
        return self.heigh


    def setWidth(self,w):
        self.width = w

    def setHeigh(self,h):
        self.heigh = h

    def setImg(self,im):
        self.img = im

    def getImg(self):
        return self.img

    def setColor(self,c):
        self.color = c

    def getColor(self):
        return self.color