#include "windows.bi"
#include "win\shellapi.bi"


private Sub PutCode(byval ccode As long)
    Dim         As long ishift
    Dim         As long ialt
    Dim         As long ictrl
    Dim         As long icod
    Dim         As long kcod
    Dim         As long lcod
    Dim         As long hcod
    icod = VkKeyScan(ccode)
    If icod = - 1 Then
        Keybd_event(0 , ccode , KEYEVENTF_UNICODE , 0)
        sleep 10
        Keybd_event(0 , ccode , KEYEVENTF_UNICODE Or KEYEVENTF_KEYUP , 0)
        Exit Sub
    End If
    lcod = LoByte(icod)
    hcod = HiByte(icod)
    kcod = MapVirtualKey(lcod , 0)
    If hcod = 7 Then
        ishift = 1 : ialt = 1 : ictrl = 1
    ElseIf hcod = 6 Then
        ishift = 0 : ialt = 1 : ictrl = 1
    ElseIf hcod = 5 Then
        ishift = 1 : ialt = 1 : ictrl = 0
    ElseIf hcod = 4 Then
        ishift = 0 : ialt = 1 : ictrl = 0
    ElseIf hcod = 3 Then
        ishift = 1 : ialt = 0 : ictrl = 1
    ElseIf hcod = 2 Then
        ishift = 0 : ialt = 0 : ictrl = 1
    ElseIf hcod = 1 Then
        ishift = 1 : ialt = 0 : ictrl = 0
    ElseIf hcod = 0 Then
        ishift = 0 : ialt = 0 : ictrl = 0
    End If
    If ialt > 0 Then Keybd_event(VK_MENU , 0 , 0 , 0)
    If ictrl > 0 Then Keybd_event(VK_CONTROL , 0 , 0 , 0)
    If ishift > 0 Then Keybd_event(VK_SHIFT , 0 , 0 , 0)

    Keybd_event(lcod , kcod , 0 , 0)
    sleep 10
    Keybd_event(lcod , kcod , KEYEVENTF_KEYUP , 0)

    If ishift > 0 Then Keybd_event(VK_SHIFT , 0 , KEYEVENTF_KEYUP , 0)
    If ictrl > 0 Then Keybd_event(VK_CONTROL , 0 , KEYEVENTF_KEYUP , 0)
    If ialt > 0 Then Keybd_event(VK_MENU , 0 , KEYEVENTF_KEYUP , 0)
End Sub



private Sub SendKeys(ByVal sText As String)
    Dim VK                As WCHAR
    Dim As long VK_BAK = &H8
    Dim As long VK_WIN = &H5B
    Dim sChar             As String
    Dim i                 As long
    Dim         As Short mlret

    mlret = LoByte(GetKeyState(VK_CAPITAL))
    If mlret = 1 Then
        Keybd_event(VK_CAPITAL , 0 , 0 , 0)
        Keybd_event(VK_CAPITAL , 0 , KEYEVENTF_KEYUP , 0)
    End If
    For i = 1 To Len(sText)
        sChar = Mid(sText , i , 1)
        If sChar = "{" Then
            If UCase(Mid(sText , i + 1 , 6)) = "ENTER}" Then
                VK = VK_RETURN
                i = i + 6
             /' ElseIf UCase(Mid(sText , i + 1 , 10)) = "BACKSPACE}" Then
                VK = VK_BAK
                i = i + 10
            ElseIf UCase(Mid(sText , i + 1 , 3)) = "BS}" Then
                VK = VK_BAK
                i = i + 3
            ElseIf UCase(Mid(sText , i + 1 , 5)) = "BKSP}" Then
                VK = VK_BAK
                i = i + 5
            ElseIf UCase(Mid(sText , i + 1 , 6)) = "BREAK}" Then
                VK = VK_PAUSE
                i = i + 6
            ElseIf UCase(Mid(sText , i + 1 , 9)) = "CAPSLOCK}" Then
                VK = VK_CAPITAL
                i = i + 9
            ElseIf UCase(Mid(sText , i + 1 , 7)) = "DELETE}" Then
                VK = VK_DELETE
                i = i + 7
            ElseIf UCase(Mid(sText , i + 1 , 4)) = "DEL}" Then
                VK = VK_DELETE
                i = i + 4
            ElseIf UCase(Mid(sText , i + 1 , 5)) = "DOWN}" Then
                VK = VK_DOWN
                i = i + 5
            ElseIf UCase(Mid(sText , i + 1 , 3)) = "UP}" Then
                VK = VK_UP
                i = i + 3
            ElseIf UCase(Mid(sText , i + 1 , 5)) = "LEFT}" Then
                VK = VK_LEFT
                i = i + 5
            ElseIf UCase(Mid(sText , i + 1 , 6)) = "RIGHT}" Then
                VK = VK_RIGHT
                i = i + 6
            ElseIf UCase(Mid(sText , i + 1 , 4)) = "END}" Then
                VK = VK_END
                i = i + 4
            ElseIf UCase(Mid(sText , i + 1 , 5)) = "HOME}" Then
                VK = VK_HOME
                i = i + 5
            ElseIf UCase(Mid(sText , i + 1 , 4)) = "ESC}" Then
                VK = VK_ESCAPE
                i = i + 4
            ElseIf UCase(Mid(sText , i + 1 , 5)) = "HELP}" Then
                VK = VK_HELP
                i = i + 5
            ElseIf UCase(Mid(sText , i + 1 , 7)) = "INSERT}" Then
                VK = VK_INSERT
                i = i + 7
            ElseIf UCase(Mid(sText , i + 1 , 4)) = "INS}" Then
                VK = VK_INSERT
                i = i + 4
            ElseIf UCase(Mid(sText , i + 1 , 8)) = "NUMLOCK}" Then
                VK = VK_NUMLOCK
                i = i + 8
            ElseIf UCase(Mid(sText , i + 1 , 5)) = "PGUP}" Then
                VK = VK_PRIOR
                i = i + 5
            ElseIf UCase(Mid(sText , i + 1 , 5)) = "PGDN}" Then
                VK = VK_NEXT
                i = i + 5
            ElseIf UCase(Mid(sText , i + 1 , 11)) = "SCROLLLOCK}" Then
                VK = VK_SCROLL
                i = i + 11
            ElseIf UCase(Mid(sText , i + 1 , 4)) = "TAB}" Then
                VK = VK_TAB
                i = i + 4
            ElseIf UCase(Mid(sText , i + 1 , 3)) = "F1}" Then
                VK = VK_F1
                i = i + 3
            ElseIf UCase(Mid(sText , i + 1 , 3)) = "F2}" Then
                VK = VK_F2
                i = i + 3
            ElseIf UCase(Mid(sText , i + 1 , 3)) = "F3}" Then
                VK = VK_F3
                i = i + 3
            ElseIf UCase(Mid(sText , i + 1 , 3)) = "F4}" Then
                VK = VK_F4
                i = i + 3
            ElseIf UCase(Mid(sText , i + 1 , 3)) = "F5}" Then
                VK = VK_F5
                i = i + 3
            ElseIf UCase(Mid(sText , i + 1 , 3)) = "F6}" Then
                VK = VK_F6
                i = i + 3
            ElseIf UCase(Mid(sText , i + 1 , 3)) = "F7}" Then
                VK = VK_F7
                i = i + 3
            ElseIf UCase(Mid(sText , i + 1 , 3)) = "F8}" Then
                VK = VK_F8
                i = i + 3
            ElseIf UCase(Mid(sText , i + 1 , 3)) = "F9}" Then
                VK = VK_F9
                i = i + 3
            ElseIf UCase(Mid(sText , i + 1 , 4)) = "F10}" Then
                VK = VK_F10
                i = i + 4
            ElseIf UCase(Mid(sText , i + 1 , 4)) = "F11}" Then
                VK = VK_F11
                i = i + 4
            ElseIf UCase(Mid(sText , i + 1 , 4)) = "F12}" Then
                VK = VK_F12
                i = i + 4
            ElseIf UCase(Mid(sText , i + 1 , 4)) = "F13}" Then
                VK = VK_F13
                i = i + 4
            ElseIf UCase(Mid(sText , i + 1 , 4)) = "F14}" Then
                VK = VK_F14
                i = i + 4
            ElseIf UCase(Mid(sText , i + 1 , 4)) = "F15}" Then
                VK = VK_F15
                i = i + 4
            ElseIf UCase(Mid(sText , i + 1 , 4)) = "F16}" Then
                VK = VK_F16
                i = i + 4
                ' NEU! Windows-Taste
            ElseIf UCase(Mid(sText , i + 1 , 4)) = "WIN}" Then
                VK = VK_WIN
                i = i + 4
                ' NEU! Kontextmenü
            ElseIf UCase(Mid(sText , i + 1 , 5)) = "APPS}" Then
                VK = VK_APPS
                i = i + 5
                ' NEU! PrintScreen-Taste (DRUCK)
            ElseIf UCase(Mid(sText , i + 1 , 6)) = "PRTSC}" Then
                VK = VK_SNAPSHOT
                i = i + 6
            ElseIf UCase(Mid(sText , i + 1 , 6)) = "ALT_D}" Then
                Keybd_event VK_MENU , 0 , 0 , 0
                i = i + 6
                Continue For
            ElseIf UCase(Mid(sText , i + 1 , 6)) = "ALT_U}" Then
                Keybd_event VK_MENU , 0 , KEYEVENTF_KEYUP , 0
                i = i + 6
                Continue For
            ElseIf UCase(Mid(sText , i + 1 , 7)) = "CTRL_D}" Then
                Keybd_event VK_CONTROL , 0 , 0 , 0
                i = i + 7
                Continue For
            ElseIf UCase(Mid(sText , i + 1 , 7)) = "CTRL_U}" Then
                Keybd_event VK_CONTROL , 0 , KEYEVENTF_KEYUP , 0
                i = i + 7
                Continue For
            ElseIf UCase(Mid(sText , i + 1 , 8)) = "SHIFT_D}" Then
                Keybd_event VK_SHIFT , 0 , 0 , 0
                i = i + 8
                Continue For
            ElseIf UCase(Mid(sText , i + 1 , 8)) = "SHIFT_U}" Then
                Keybd_event VK_SHIFT , 0 , KEYEVENTF_KEYUP , 0
                i = i + 8
                Continue For
            ElseIf UCase(Mid(sText , i + 1 , 2)) = "}}" Then
                VK = - 1
                PutCode(Asc( "}"))
                i = i + 2
                Continue For
            ElseIf UCase(Mid(sText , i + 1 , 2)) = "{}" Then
                VK = - 1
                PutCode(Asc( "{"))
                i = i + 2
                Continue For
            ElseIf UCase(Mid(sText , i + 1 , 2)) = "+}" Then
                VK = - 1
                PutCode(Asc( "+"))
                i = i + 2
                Continue For
            ElseIf UCase(Mid(sText , i + 1 , 2)) = "%}" Then
                VK = - 1
                PutCode(Asc( "%"))
                i = i + 2
                Continue For
            ElseIf UCase(Mid(sText , i + 1 , 2)) = "^}" Then
                VK = - 1
                PutCode(Asc( "^"))
                PutCode(Asc( " "))
                i = i + 2
                Continue For
            ElseIf UCase(Mid(sText , i + 1 , 2)) = "~}" Then
                VK = - 1
                PutCode(Asc( "~"))
                PutCode(Asc( " "))
                i = i + 2
                Continue For '/
            End If
         /' ElseIf sChar = "+" Then                  ' SHIFT
            Keybd_event VK_SHIFT , 0 , 0 , 0
            Sleep 10
            Keybd_event VK_SHIFT , 0 , KEYEVENTF_KEYUP , 0
            Continue For
        ElseIf sChar = "%" Then                  ' ALT
            Keybd_event VK_MENU , 0 , 0 , 0
            Sleep 10
            Keybd_event VK_MENU , 0 , KEYEVENTF_KEYUP , 0
            Continue For
        ElseIf sChar = "^" Then                  ' CONTROL
            Keybd_event VK_CONTROL , 0 , 0 , 0
            Sleep 10
            Keybd_event VK_CONTROL , 0 , KEYEVENTF_KEYUP , 0
            Continue For
        ElseIf sChar = "~" Then                  ' ENTER
            Keybd_event VK_RETURN , 0 , 0 , 0
            Sleep 10
            Keybd_event VK_RETURN , 0 , KEYEVENTF_KEYUP , 0
            Continue For
        ElseIf sChar = "¨" Then
            VK = - 1
            PutCode(Asc( "¨"))
            PutCode(Asc( " "))
            Continue For
        ElseIf sChar = "`" Then
            VK = - 1
            PutCode(Asc( "`"))
            PutCode(Asc( " "))
            Continue For '/
        Else
            VK = - 1
            PutCode(Asc(sChar))
        End If
        If vk <> - 1 Then
            Keybd_event VK , 0 , 0 , 0
            'Sleep 10
            Keybd_event VK , 0 , KEYEVENTF_KEYUP , 0
            vk = - 1
        End If
    Next i
    'Sleep 10
    If mlret = 1 Then
        Keybd_event(VK_CAPITAL , 0 , 0 , 0)
        'Sleep 10
        Keybd_event(VK_CAPITAL , 0 , KEYEVENTF_KEYUP , 0)
    End If
End Sub

private function do_graph(byref exefile as string, byref graphfile as string, byref winname as string) as hwnd
	sleep 100
	shellexecute(0 , 0 , """" & exefile & """" , "" , "" , SW_HIDE)
	sleep 100
	dim as hwnd hw1 = findwindow( "cgraph" , "cgraph")
	if hw1 THEN
		print hw1
		postmessage(hw1 , WM_COMMAND , 15552 , 0)
		sleep 50
		SendKeys( graphfile & "{ENTER}")
		setwindowtext(hw1 , winname)
		SetActiveWindow(hw1)
	END IF
	return hw1
end function



dim as hwnd hw1 = do_graph(exepath & $"\graphit.exe", exepath & $"\explicitx4.txt", "My_Graph1")
dim as hwnd hw2 = do_graph(exepath & $"\graphit.exe", exepath & $"\implicitx3.txt", "My_Graph2")

dim buffer As String
dim f As long
buffer = "ImplicitX" & chr(13,10)
buffer &= "Title ""Testing New Curve""" & chr(13,10)
buffer &= "HorzAxisLabel ""The Horz Axis Label""" & chr(13,10)
buffer &= "SetLabels ""Curve1""" & chr(13,10)& chr(13,10)

dim y as double
dim i as long
For i  = 0 To 1000

    y = (Sin(6.28*i/200)*Exp(-i/200))*100
    '? i , y
	buffer &= trim(str(y)) & chr(13,10)
Next i


dim as hwnd hw3

f = FreeFile
Open  exepath & $"\file.txt" For Binary As #f
If Err = 0 Then
	Put #f, , buffer
	Close
	hw3 = do_graph(exepath & $"\graphit.exe", exepath & $"\file.txt", "My_New_Graph")
	sleep 1000
	kill exepath & $"\file.txt"
else
	Print "Error opening the file"
end if



print "waiting 20s or press any key to continue"
sleep 20000
hw1 = findwindow( "cgraph" , "My_Graph1")
print hw1

if hw1 THEN sendmessage(hw1 , WM_CLOSE , 0 , 0)

hw2 = findwindow( "cgraph" , "My_Graph2")
print hw2

if hw2 THEN sendmessage(hw2 , WM_CLOSE , 0 , 0)

hw3 = findwindow( "cgraph" , "My_New_Graph")
print hw3
if hw3 THEN sendmessage(hw3 , WM_CLOSE , 0 , 0)
print "Press any key or close console window to finish"
sleep
