from ctypes import *
import win32clipboard as clip

import win32con
import os
import subprocess

from io import BytesIO
from PIL import ImageGrab

from pynput import keyboard,mouse

import datetime
import tkinter as tk
from threading import Thread



INSTRUCTION_STRING = 'waiting for key press. num0 for help'


class SnippingWindow():

    def __init__(self,drag_enabled):
        self.drag_enabled = drag_enabled 
        self.clicks=0
        self.stop_window=False

        self.x = 0
        self.y = 0
        self.start_x = None
        self.start_y = None
        self.rect = None


        #set main window
        self.root = tk.Tk()
        self.root.attributes('-alpha', 0.3)
        # self.root.geometry('1920x1080')
        self.root.attributes("-fullscreen", True) 
        self.root.attributes("-topmost", True)  #bring window to the front
        self.root.overrideredirect(1)
        # self.root.update()
        # self.root.attributes("-topmost", False)
        
        #set drawing canvas
        self.canvas = tk.Canvas(self.root,width=1920,height=1080,cursor="tcross")
        self.canvas['background']='black'
        self.canvas.pack()  
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_move_press)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)  

       
    
    def on_button_press(self, mouse):
        #save mouse start position
        self.start_x = mouse.x
        self.start_y = mouse.y
        self.rect = self.canvas.create_rectangle(self.x, self.y, 1, 1, fill="white")
       
        
        if not self.drag_enabled:   #just moving the mouse draws a rectangle, if in clicking mode
            self.canvas.bind("<Motion>",self.on_move_press)


    def on_move_press(self, mouse):
        #current mouse position
        current_x, current_y = (mouse.x, mouse.y)

        #expand rectangle as you drag the mouse
        self.canvas.coords(self.rect, self.start_x, self.start_y, current_x, current_y)



    def on_button_release(self, mouse): #release ends the drawing
        self.clicks += 1
        if self.drag_enabled or self.clicks==2:
            self.stop_window = True


       

class SnippingThingy():

    def __init__(self):
        self.click_count=0
        self.save_enabled = False
        self.drag_enabled = True    #default mode is drag like win snipping tool
                                    #False = click to select corners

        #default screenshot bbox = 1920x1080
        (self.x1,self.y1) = 0,0   
        (self.x2,self.y2) = 1920,1080

        print(INSTRUCTION_STRING)


    def open_window_mask(self): 
        window = SnippingWindow(drag_enabled=self.drag_enabled)
        while True:
            if window.stop_window:
                window.root.destroy()
                break
            window.root.update()

        

    #callback for mouse listener
    def on_click_click(self,x,y,button,pressed): 
        #click in 2 corners
    
        if pressed and self.click_count==0: #first time callback gets called
            print('corner 1 clicked at {0},{1}'.format(x,y))
            self.x1,self.y1 = x,y
            self.click_count+=1

        elif pressed and self.click_count==1: #second click
            print('corner 2 clicked at {0},{1}'.format(x,y))
            self.x2,self.y2=x,y
            self.click_count+=1

        elif self.click_count==2: #after second click stop listening, continue with program
            return False
        
    #callback for mouse listener
    def on_click_drag(self,x,y,button,pressed):
        #drag screenshot area like snipping tool

        if button.name == 'left' and pressed:
            print('corner 1 clicked at {0},{1}'.format(x,y))
            self.x1,self.y1 = x,y
        elif button.name == 'left' and not pressed:
            print('corner 2 clicked at {0},{1}'.format(x,y))
            self.x2,self.y2 = x,y
            return False
        

    def handle_selection(self):
        #selection must be from top left to bottom right
        current_x1,current_y1 = self.x1,self.y1
        current_x2,current_y2 = self.x2,self.y2

        #swap coordinates to make selection work with ImageGrab.grab()
        if self.x1>self.x2 and self.y1>self.y2: #selection from bottom right to top left
            self.x1,self.x2 = current_x2,current_x1
            self.y1,self.y2 = current_y2,current_y1

        elif self.x1>self.x2 and self.y1<self.y2: #selection from top right > bottom left
            self.x1,self.y1 = current_x2,current_y1
            self.x2,self.y2 = current_x1,current_y2

        elif self.x2>self.x1 and self.y2<self.y1: #selection from lower left to top right
            self.x1,self.y1 = current_x1,current_y2
            self.x2,self.y2 = current_x2,current_y1
        
        elif self.x1==self.x2 or self.y1==self.y2: #0 pixel wide/high selection
            print('0 pixel wide/high selection, using full size instead')
            self.x1,self.y1 = 0,0
            self.x2,self.y2 = 1920,1080



    def setup(self): #define corners of screenshot bbox by clicking upper and lower corner
       
        #open a transparent window on top of everything
        #to enable selection without interfering the currently active program
        t = Thread(target=self.open_window_mask, )
        t.start()

        print('drag corners') if self.drag_enabled else print('click corners')
    
        self.click_count=0
        #select area by clicking 2 corners or dragging like snipping tool
        with mouse.Listener(on_click=self.on_click_drag if self.drag_enabled else self.on_click_click) as listener: 
            listener.join() #wait for mouse click
            

        #selection might be bad
        self.handle_selection()
        
        t.join() #wait for the transparent window to close
        if self.drag_enabled:
            self.take_screenshot() #screenshot right after dragging the area

        
        
        

    def save_screenshot(self,image):
        try:
            now=datetime.datetime.now()
            (d,m,y)=now.day, now.month, now.year
            (hr,mi,s)=now.hour, now.minute, now.second
            screen_name = "C:/Users/Robse/Documents/Screeen/screenshot{0}-{1}-{2}-{3}-{4}-{5}.png".format(y,m,d,hr,mi,s)
            image.save(screen_name)
            print('saved as ' + screen_name)
                    
        except OSError:
            print('couldnt save correctly')



    def take_screenshot(self):
            try:
                #take screenshot and add it to the clipboard, save to file if save_enabled == true
                image = ImageGrab.grab(bbox=(self.x1,self.y1,self.x2,self.y2))
                # image.resize((896,640))
                output = BytesIO()
                image.convert('RGB').save(output, 'BMP')
                data = output.getvalue()[14:]
                output.close()

                clip.OpenClipboard()
                clip.EmptyClipboard()
                clip.SetClipboardData(win32con.CF_DIB, data)
                clip.CloseClipboard()

                print('screenshot taken')

                if self.save_enabled:
                    self.save_screenshot(image)
                    

            except OSError:
                print('couldnt copy to clipboard')
            
            except SystemError:
                print('something went wrong when taking a picture')


    def set_full_area(self):
        (self.x1,self.y1) = 0,0
        (self.x2,self.y2) = 1920,1080
        print('Area set to full size')


    def open_folder(self):
        path="C:\\Users\\Robse\\Documents\\Screeen\\"
        subprocess.Popen(f'explorer {os.path.realpath(path)}')
        print('opened folder')


    def print_help(self):
        mode = 'drag' if self.drag_enabled else 'click'
        print(
            f'save: {self.save_enabled} \n' \
            f'mode: {mode} \n' \
            f'selection: {self.x1,self.y1}   {self.x2,self.y2}\n'\
            'num0 - help \n' \
            'num1 - take screenshot \n' \
            'num2 - toggle selection mode: click/drag \n' \
            'num3 - toggle saving \n' \
            'num5 - open screenshot folder \n'\
            'num7 - select corners \n' \
            'num8 - set full area \n' \
            'num9 - close')


    def on_press(self,key):  
        if hasattr(key, 'vk'):

            #num0
            if key.vk==96: 
                self.print_help()
                
            #num1
            elif key.vk==97: 
                self.take_screenshot()
                
            #num2
            elif key.vk==98: 
                self.drag_enabled = not self.drag_enabled
                if self.drag_enabled: print('drag mode')
                else: print('click mode')
                
            #num3
            elif key.vk == 99: 
                self.save_enabled=not self.save_enabled
                print('toggled saving')
                
            #num7
            elif key.vk==103:
                self.setup()
                
            #num5
            elif key.vk==101: 
                self.open_folder()
                
            # num8
            elif key.vk == 104: 
                self.set_full_area()  
                
            #num9
            elif key.vk==105: 
                return False #stops the listener

            else: pass

            print(INSTRUCTION_STRING)
            

    def run(self):
        #run until num9 gets pressed
        with keyboard.Listener(on_press = self.on_press) as listener:
            listener.join()



def main():
    snips = SnippingThingy()
    snips.run()
    

if __name__=='__main__':
    main()
