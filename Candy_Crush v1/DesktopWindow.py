import gc
import os
import time
import win32api, win32con, win32gui
import sys
import win32com.client

class DesktopWindow(object):
    def __init__(self, *args, **kwargs):
        self.window_id = win32gui.GetDesktopWindow()
        self.window_dc = win32gui.GetWindowDC(self.window_id)
        pass
    def get_pixel_color(self, i_x, i_y):
        long_colour = win32gui.GetPixel(self.window_dc, i_x, i_y)
        i_colour = int(long_colour)
        return (i_colour & 0xff, (i_colour >> 8) & 0xff,
                (i_colour >> 16) & 0xff)