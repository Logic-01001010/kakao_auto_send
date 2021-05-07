#-*- coding: utf-8 -*-

import time, win32con, win32api, win32gui, ctypes
import json



delay = 0.3

PBYTE256 = ctypes.c_ubyte * 256
_user32 = ctypes.WinDLL("user32")
GetKeyboardState = _user32.GetKeyboardState
SetKeyboardState = _user32.SetKeyboardState
PostMessage = win32api.PostMessage
SendMessage = win32gui.SendMessage
FindWindow = win32gui.FindWindow
IsWindow = win32gui.IsWindow
GetCurrentThreadId = win32api.GetCurrentThreadId
GetWindowThreadProcessId = _user32.GetWindowThreadProcessId
AttachThreadInput = _user32.AttachThreadInput

MapVirtualKeyA = _user32.MapVirtualKeyA
MapVirtualKeyW = _user32.MapVirtualKeyW

MakeLong = win32api.MAKELONG
w = win32con


def enter(handle):
	time.sleep(0.01)
	win32api.PostMessage(handle, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
	time.sleep(0.01)
	win32api.PostMessage(handle, win32con.WM_KEYUP, win32con.VK_RETURN, 0)




def PostKeyEx( hwnd, key, shift, specialkey):
    if IsWindow(hwnd):
        
        ThreadId = GetWindowThreadProcessId(hwnd, None)
        
        lparam = MakeLong(0, MapVirtualKeyA(key, 0))
        msg_down=w.WM_KEYDOWN
        msg_up=w.WM_KEYUP
        
        if specialkey:
            lparam = lparam | 0x1000000
            
        if len(shift) > 0: #Если есть модификаторы - используем PostMessage и AttachThreadInput
            pKeyBuffers = PBYTE256()
            pKeyBuffers_old = PBYTE256()
            
            SendMessage(hwnd, w.WM_ACTIVATE, w.WA_ACTIVE, 0)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, True)
            GetKeyboardState( ctypes.byref(pKeyBuffers_old ))
            
            for modkey in shift:
                if modkey == w.VK_MENU:
                    lparam = lparam | 0x20000000
                    msg_down=w.WM_SYSKEYDOWN
                    msg_up=w.WM_SYSKEYUP
                pKeyBuffers[modkey] |= 128
    
            SetKeyboardState( ctypes.byref(pKeyBuffers) )
            time.sleep(0.01)
            PostMessage( hwnd, msg_down, key, lparam)
            time.sleep(0.01)
            PostMessage( hwnd, msg_up, key, lparam | 0xC0000000)
            time.sleep(0.01)
            SetKeyboardState( ctypes.byref(pKeyBuffers_old) )
            time.sleep(0.01)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, False)
            
        else: #Если нету модификаторов - используем SendMessage
            SendMessage( hwnd, msg_down, key, lparam)
            SendMessage( hwnd, msg_up, key, lparam | 0xC0000000)




def login(password):

	time.sleep(delay)

	handle = win32gui.FindWindow(None, "카카오톡")

	EXhandle = win32gui.FindWindowEx(handle, 0, 'EVA_ChildWindow_Dblclk', None)

	EXhandle = win32gui.FindWindowEx(EXhandle, 0, 'Edit', None)

	req = win32api.SendMessage(EXhandle, 0xC, 0, password)
	
	enter(EXhandle)

	if req:
		print("Tryed Login.")

def OpenChat(room):

	time.sleep(delay)


	handle = win32gui.FindWindow(None, "카카오톡")
	EXhandle = win32gui.FindWindowEx(handle, 0, 'EVA_ChildWindow', None)


	first = win32gui.FindWindowEx(EXhandle, 0, 'EVA_Window', None)
	EXhandle = win32gui.FindWindowEx(EXhandle, first, 'EVA_Window', None)


	EXhandle = win32gui.FindWindowEx(EXhandle, 0, 'Edit', None)
	
	req = win32api.SendMessage(EXhandle, win32con.WM_SETTEXT, 0, room)

	print("Searched "+room+".")

	time.sleep(delay)

	enter(EXhandle)

	print("Open "+room+".")


def send(room, message):

	time.sleep(delay)


	handle = win32gui.FindWindow(None, room)
	EXhandle = win32gui.FindWindowEx(handle, 0, "RICHEDIT50W", None)

	win32api.SendMessage(EXhandle, 0xc, 0, message)

	enter(EXhandle)

	print("Sended message => \""+ message +"\" At "+room+".")


def condition(date):

	now = time.localtime()

	if ((date['year'] == str(now.tm_year)) or (date['year'] == "*")) and ((date['month'] == str(now.tm_mon)) or (date['month'] == "*")) and ((date['day'] == str(now.tm_mday)) or (date['day'] == "*")) and ((date['hour'] == str(now.tm_hour)) or (date['hour'] == "*")) and ((date['min'] == str(now.tm_min)) or (date['min'] == "*")):
		
		return True

	else:

		return False






if __name__ == '__main__':


	while True:

		


		f = open('format.json', 'r', encoding='UTF8')
		json_data = json.load(f)


		password = json_data['password']
		room = json_data['room']
		date = json_data['date']
		message = json_data['message']


		if condition(date):

			now = time.localtime()
			print("%04d/%02d/%02d %02d:%02d:%02d" % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec))

			login(password)

			OpenChat(room)

			send(room, message)

			time.sleep(60)
