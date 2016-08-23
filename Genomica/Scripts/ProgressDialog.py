# -*- coding: cp1252 -*-

import win32gui
import win32api
import win32con
import struct
from Diccionario import *
from MyExceptions import *
import logging



class ProgressDialog:

    counter = 0

    def __init__(self,thread):

        logging.debug("Estoy en ProgressDialog.__init__")

        win32gui.InitCommonControls()
        self.hinst = win32gui.dllhandle
        self.dic = Diccionario()
        self.thread = thread
        self.progress = 0

        self._DoCreate()


    def _RegisterWndClass(self):

        className = "GenomicaWinClass"

        if not ProgressDialog.counter:
        
            message_map = {}
            wc = win32gui.WNDCLASS()
            wc.SetDialogProc()
            wc.hInstance = self.hinst
            wc.lpszClassName = className
            wc.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW
            wc.hCursor = win32gui.LoadCursor( 0, win32con.IDC_ARROW )
            wc.hbrBackground = win32con.COLOR_WINDOW + 1
            wc.lpfnWndProc = message_map
            wc.cbWndExtra = win32con.DLGWINDOWEXTRA + struct.calcsize("Pi")
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            classAtom = win32gui.RegisterClass(wc)
            ProgressDialog.counter = 1
            self.classAtom = classAtom
            
        return className


    def _GetDialogTemplate(self, dlgClassName):

        style = win32con.WS_POPUP | win32con.WS_VISIBLE | win32con.WS_CAPTION | win32con.DS_SETFONT 
        cs = win32con.WS_CHILD | win32con.WS_VISIBLE
        title = self.dic.voc["exportingDlg.title"]
        size_x = 300
        size_y = 60
        dlg = [ [title, (0, 0, size_x, size_y), style, None, (8, "MS Sans Serif"), None, dlgClassName], ]
        dlg.append([130, self.dic.voc["exporting"], -1, (5, 10, size_x-5, 9), cs | win32con.SS_CENTER])
        self.buttonid = 1026
        s = win32con.BS_PUSHBUTTON | cs | win32con.WS_TABSTOP | win32con.SS_CENTER
        dlg.append([128, self.dic.voc["btn.cancel"], self.buttonid, (size_x/3, 35, size_x/3, 14), s])

        return dlg


    def _DoCreate(self):
        self.message_map = {
            win32con.WM_INITDIALOG: self.OnInitDialog,
            win32con.WM_COMMAND: self.OnCommand,
            win32con.WM_DESTROY: self.OnDestroy,
        }
        dlgClassName = self._RegisterWndClass()
        self.template = self._GetDialogTemplate(dlgClassName)


    def OnInitDialog(self, hwnd, msg, wparam, lparam):

        self.hwnd = hwnd
        desktop = win32gui.GetDesktopWindow()
        l,t,r,b = win32gui.GetWindowRect(self.hwnd)
        dt_l, dt_t, dt_r, dt_b = win32gui.GetWindowRect(desktop)
        centre_x, centre_y = win32gui.ClientToScreen( desktop, ( (dt_r-dt_l)/2, (dt_b-dt_t)/2) )
        win32gui.MoveWindow(hwnd, centre_x-(r/2), centre_y-(b/2), r-l, b-t, 0)


    def OnCommand(self, hwnd, msg, wparam, lparam):
        
        id = win32api.LOWORD(wparam)
        if id == self.buttonid:
            win32gui.PostMessage(hwnd,win32con.WM_DESTROY,0,0)
            #self.OnDestroy()


    def OnDestroy(self,hwnd, msg, wparam, lparam):
        self.thread.kill()
        win32gui.DestroyWindow(hwnd)
        win32gui.PostQuitMessage(0)


    def Show(self):
        logging.debug("Estoy en ProgressDialog.Show")
        win32gui.CreateDialogIndirect(self.hinst, self.template, 0, self.message_map)
        self.thread.start()
        win32gui.PumpMessages()


    def notifyProgress(self,progress):

        if progress == 1:

            logging.info("He terminado así que mando cerrar la ventana")
            win32gui.PostMessage(self.hwnd,win32con.WM_DESTROY,0,0)

        elif not (progress >= 0 and progress < 1):

            logging.info(str(type(progress)) + "Sucedió un error en el proceso de exportación")
            win32gui.PostMessage(self.hwnd,win32con.WM_DESTROY,0,0)


    def Unregister(self):

        ProgressDialog.counter = 0
        #win32gui.UnregisterClass("GenomicaWinClass",None)

        try:
            win32gui.UnregisterClass(self.classAtom,0)

        except:
            print "dio error"





    
if __name__=='__main__':

    # Esto no funciona porque hace falta pasarle la tarea:
    w = ProgressWindow()
    w.CreateWindow()
    win32gui.PumpMessages()
