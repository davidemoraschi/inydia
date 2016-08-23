# -*- coding: cp1252 -*-

import win32gui
import win32api
import win32con
import struct
from Diccionario import *
import logging


class FinishWindow:

    counter = 0

    def __init__(self):

        win32gui.InitCommonControls()
        self.hinst = win32gui.dllhandle
        self.dic = Diccionario()
        self.outputcode = 0


    def _RegisterWndClass(self):

        className = "GenomicaFinishWinClass"

        if not FinishWindow.counter:
        
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
            FinishWindow.counter = 1
            self.classAtom = classAtom
            
        return className


    def _GetDialogTemplate(self, dlgClassName):

        style = win32con.WS_POPUP | win32con.WS_VISIBLE | win32con.WS_CAPTION | win32con.DS_SETFONT 
        cs = win32con.WS_CHILD | win32con.WS_VISIBLE
        title = self.dic.voc["exportingDlg.title"]
        size_x = 300
        size_y = 60
        dlg = [ [title, (0, 0, size_x, size_y), style, None, (8, "MS Sans Serif"), None, dlgClassName], ]

        try:
            self.outputcode = float(self.outputcode)

            if self.outputcode == 1:

                message = self.dic.voc["export.done"]

            else:

                message = self.dic.voc["export.aborted"]
                
        except:
            
            message = self.dic.voc[self.outputcode.getMyErrorMessage()]


        dlg.append([130, message, -1, (0, 10, size_x, 9), cs | win32con.SS_CENTER])
        self.buttonid = 1026
        s = win32con.BS_PUSHBUTTON | cs | win32con.WS_TABSTOP | win32con.SS_CENTER
        dlg.append([128, self.dic.voc["btn.accept"], self.buttonid, (size_x/3, 35, size_x/3, 14), s])

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
            win32gui.DestroyWindow(hwnd)
            self.OnDestroy(hwnd, msg, wparam, lparam)


    def OnDestroy(self, hwnd, msg, wparam, lparam): 
        win32gui.PostQuitMessage(0)


    def Show(self,outputcode):
        self.outputcode = outputcode
        self._DoCreate()
        logging.debug("FinishWindow.Show")
        win32gui.CreateDialogIndirect(self.hinst, self.template, 0, self.message_map)
        win32gui.PumpMessages()


    def Unregister(self):

        FinishWindow.counter = 0
        #win32gui.UnregisterClass("GenomicaFinishWinClass",None)

        try:
            win32gui.UnregisterClass(self.classAtom,0)

        except:
            print "dio error2"

        
    




    
if __name__=='__main__':

    w = FinishWindow()
    w.Show(1)
