Attribute VB_Name = "modMain" 
' ************************************************************************************************************
' ************************************************************************************************************
' **LL************************************bb***********************AA************RRRRRRRRRRRR*****TTTTTTTTTT**
' **LL************************************bb**********************AAAA***********RR*********RR********TT******
' **LL************************************bb*********************AA**AA**********RR**********RR*******TT******
' **LL****************eeeeeeeee***********bb********************AA****AA*********RR*********RR********TT******
' **LL***************ee*******ee**********bb*******************AA******AA********RRRRRRRRRRRR*********TT******
' **LL**************ee*********ee*********bbbbbbbbbbb*********AAAAAAAAAAAA*******RR******RR***********TT******
' **LL**************eeeeeeeeeeeee*********bb********bb*******AA**********AA******RR*******RR**********TT******
' **LL**************ee********************bb*********bb*****AA************AA*****RR********RR*********TT******
' **LL**************ee*********ee***..****bb********bb*****AA**************AA****RR*********RR********TT******
' **LLLLLLLLLLLLL****eeeeeeeeeee****..****bbbbbbbbbbb*****AA****************AA***RR**********RR*******TT******
' ************************************************************************************************************
' ************************************************************************************************************
'
' ************************************************************************************************************
' * Programmer Name  : Arturo Najera Diocton, Junior
' * Web Site         : http://www.facebook.com/ZhadukurFalcon
' * ICQ UNI/Nick     :
' * E-Mail           :
' * Creation Date    : February 29, 2000
' * Time             : 1056 hrs
' * Module Filename  : modMain.bas
' * Module Name      : modMain
' * Module Version   : 2.6.6 - July 31, 2011 14:07
' ************************************************************************************************************
' *
' * Comments         : I use this on the majority of my VB Projects / Applications
' *                  : This is a work in progress. Not in anyway finished
' *
' ************************************************************************************************************
' *
' * Modified             : June 04, 2000     -   2.0.0
' * Modified             : July 09, 2000     -   2.5.0
' * Sub Added - Graffiti : November 21, 2000 -   2.5.5
' * Function added Rubiks Cube Scrambler : May 16, 2006 - 2.6.0
' * Function added Display Time using API timGetTime computation : May 17, 2006 - 2.6.1
' * Constant Pi  - July 31, 2011 1407  -  2.6.5
' * Constant Phi - July 31, 2011 1407  -  2.6.5
' * Function added RadToDeg - July 31, 2011 1425  -  2.6.6
' * Function added DegToRad - July 31, 2011 1425  -  2.6.6
' *
' ************************************************************************************************************
' *
' * Subs and Functions Included:
' *
' * 01 Dice
' * 02 GradientPlus
' * 03 Graffiti
' * 04 MidScreen
' * 05 QTRScreenLL
' * 06 QTRScreenLR
' * 07 QTRScreenUL
' * 08 QTRScreenUR
' * 09 RePosCtrl
' * 10 RePosCtrlXY
' * 11 ShowFrm
' * 12 GradientPlus
' * 13 SpinGradient
' * 14 txtSelect
' * 15 UnloadFormRoutine
' * 16 RadToDeg
' * 17 DegToRad
' *
' ************************************************************************************************************
' *
' * FUTURE Updates
' *
' * Comprehensive File Creation Utility - Binary / Random Access
' * File Maintenance Routines - Binary / Random Access
' *
' * Comprehensive Database Creation Utility - ADO
' * Database Maintenance Routines - ADO
' *
' ************************************************************************************************************
' *
' * Module Notes:
' *
' * RePosCtrl & RePosCtrlXY - Needs finer refinement in regards to form resizing. - May 07, 2001
' *
' ************************************************************************************************************
Option Explicit

' ---------------------------------------------------------------------------
' Define the Time constants in milliseconds
' ---------------------------------------------------------------------------
Public Const constYear As Long = 31536000000 ' 31,536,000,000 milliseconds = 1 year        x 100  = 1 century
Public Const constDay  As Long = 86400000    '     86,400,000 milliseconds = 1 day         x 365  = 1 year
Public Const constHr   As Long = 3600000     '      3,600,000 milliseconds = 1 hour        x 24   = 1 day
Public Const constMin  As Long = 60000       '         60,000 milliseconds = 1 minute      x 60   = 1 hour
Public Const constSec  As Long = 1000        '          1,000 milliseconds = 1 second      x 60   = 1 minute
Public Const constTSec As Long = 1           '              1 milliseconds = 1 millisecond x 1000 = 1 second

Public Const constPI   As Single = 3.14159265358979     ' 4 x ArcTan(1)
'3.1415926535897932384626433832795028841971693993751058209749445923078164062862089986280348253421170679

Public Const constPhi  As Single = 1.61803398874989     ' (1 + Sqrt(5)) / 2
'1.6180339887498948482045868343656381177203091798057628621354486227052604628189024497072072041893911374

Public Const intObjSpc = 30     ' Used in conjunction with Functions RePosCtrl and RePosCtrlXY
Public intX, intY, intLeft, intTop, intWidth, intHeight As Integer
Public intLeftOld, intTopOld, intWidthOld, intHeightOld As Integer

'*************************************************************************************************************
'* Loop thru the forms collection and unload those forms from memory
'*************************************************************************************************************
'
Public Sub UnloadFormRoutine()
   Dim frm As Form
   For Each frm In Forms
      Unload frm          ' deactivate the form
      Set frm = Nothing   ' remove the object from memory
   Next
End Sub

'*************************************************************************************************************
'* Closes all form and shows a form stated in the parameter
'*************************************************************************************************************
'
Public Sub ShowFrm(frm As Form)
   UnloadFormRoutine
   Load frm
   frm.Show
End Sub

'*************************************************************************************************************
'* Converts Radians to Degrees
'*************************************************************************************************************
'
Function RadToDeg(sngRadians As Single)
   RadToDeg = sngRadians * (180 / constPI)
End Function

'*************************************************************************************************************
'* Converts Dedgrees to Radians
'*************************************************************************************************************
'
Function DegToRad(sngDegrees As Single)
   DegToRad = sngRadians * (constPI / 180)
End Function

'*************************************************************************************************************
'* Selects all text in the Textbox
'*************************************************************************************************************
'
Function txtSelect(txtbx As TextBox)
   txtbx.SelStart = 0
   txtbx.SelLength = Len(txtbx.Text)
End Function

'*************************************************************************************************************
'* Centers the Form on the Screen
'*************************************************************************************************************
'
Function MidScreen(frm As Form)
   frm.Left = (Screen.Width - frm.Width) / 2
   frm.Top = (Screen.Height - frm.Height) / 2
End Function

'*************************************************************************************************************
'* Center the Form on the Upper-Left Quadrant of the Screen
'*************************************************************************************************************
'
Function QTRScreenUL(frm As Form)
   frm.Left = (Screen.Width - frm.Width) / 4
   frm.Top = (Screen.Height - frm.Height) / 4
End Function

'*************************************************************************************************************
'* Center the Form on the Upper-Right Quadrant of the Screen
'*************************************************************************************************************
'
Function QTRScreenUR(frm As Form)
   frm.Left = (((Screen.Width - frm.Width) / 4) * 3)
   frm.Top = (Screen.Height - frm.Height) / 4
End Function

'*************************************************************************************************************
'* Center the Form on the Lower-Left Quadrant of the Screen
'*************************************************************************************************************
'
Function QTRScreenLL(frm As Form)
   frm.Left = ((Screen.Width - frm.Width) / 4)
   frm.Top = ((Screen.Height - frm.Height) / 4) * 3
End Function

'*************************************************************************************************************
'* Center the Form on the Lower-Right Quadrant of the Screen
'*************************************************************************************************************
'
Function QTRScreenLR(frm As Form)
   frm.Left = ((Screen.Width - frm.Width) / 4) * 3
   frm.Top = ((Screen.Height - frm.Height) / 4) * 3
End Function

'*************************************************************************************************************
'* Reposition a Control on a form
'*************************************************************************************************************
'
Function RePosCtrl(frm As Form, Ctrl As Control)
   Ctrl.Left = frm.ScaleWidth - Ctrl.Width
   Ctrl.Top = frm.ScaleHeight - Ctrl.Height
End Function

'*************************************************************************************************************
'* Reposition a Control on a form as to it's X & Y
'*************************************************************************************************************
'
Function RePosCtrlXY(frm As Form, Ctrl As Control, X As Integer, Y As Integer)
   Ctrl.Left = frm.ScaleWidth - (Ctrl.Width + X)
   Ctrl.Top = frm.ScaleHeight - (Ctrl.Height + Y)
End Function

'*************************************************************************************************************
'* Vari-Sided / Vari-Number Die Roll Simulator
'*************************************************************************************************************
'
Function Dice(Die, Sides As Integer)
   Dim X, DieRoll As Integer
   Do While DieRoll > (Die * Sides) Or DieRoll < Die
      DieRoll = 0
      For X = 1 To Die
         DieRoll = DieRoll + (Rnd * Sides) + 1
      Next
   Loop
   Dice = DieRoll
End Function

'*************************************************************************************************************
'* Add a colored gradient to a form
'*************************************************************************************************************
'
Function GradientPlus(frm As Form, StartRed As Integer, StartGreen As Integer, StartBlue As Integer, EndRed As Integer, EndGreen As Integer, EndBlue As Integer)
   On Error Resume Next
   Dim RedChange, GreenChange, BlueChange, X
   frm.DrawStyle = 6 ' Inside Solid
   frm.ScaleMode = 3 ' Pixels
   frm.DrawMode = 13 ' Copy Pen
   frm.DrawWidth = 2
   frm.ScaleHeight = 256
   For X = 0 To 255 'Start Loop
      frm.Line (0, X)-(Screen.Width, X - 1), RGB(StartRed + RedChange, StartGreen + GreenChange, StartBlue + BlueChange), B  'Draws Line With correct color
      RedChange = RedChange + (EndRed - StartRed) / 255 '
      GreenChange = GreenChange + (EndGreen - StartGreen) / 255 ' Sets Next Loops Color
      BlueChange = BlueChange + (EndBlue - StartBlue) / 255 '
   Next X
End Function

'*************************************************************************************************************
'* Add a spin colored gradient to a form
'*************************************************************************************************************
'
Function SpinGradient(frm As Form, rs%, gs%, bs%, re%, ge%, be%, smooth As Boolean)
   Dim ri, gi, bi, rc, gc, bc, X
   If frm.WindowState = vbMinimized Then Exit Function
   frm.BackColor = RGB(rs, gs, bs)
   If smooth = True Then
      frm.DrawStyle = 6
   Else
      frm.DrawStyle = 0
   End If
   If frm.ScaleWidth <> 255 Then frm.ScaleWidth = 255
   If frm.ScaleHeight <> 255 Then frm.ScaleHeight = 255
   frm.DrawWidth = 5
   frm.Refresh
   ri = (rs - re) / 255 / 2
   gi = (gs - ge) / 255 / 2
   bi = (bs - be) / 255 / 2
   rc = rs: bc = bs: gc = gs
   For X = 0 To 255
      frm.Line (X, 0)-(255 - X, 255), RGB(rc, gc, bc)
      rc = rc - ri
      gc = gc - gi
      bc = bc - bi
   Next X
   For X = 0 To 255
      frm.Line (255, X)-(0, 255 - X), RGB(rc, gc, bc)
      rc = rc - ri
      gc = gc - gi
      bc = bc - bi
   Next X
End Function

'*************************************************************************************************************
'* Add a graffiti like lines on the background of the form
'*************************************************************************************************************
'
Sub Graffiti(frm As Form)
   Dim colo, X1, X1, X2, Y1, Y2 As Integer
   If frm.WindowState = vbMinimized Then Exit Sub
   frm.BackColor = QBColor(7)
   frm.ScaleHeight = 75
   frm.ScaleWidth = 75
   For X = 0 To 200
      DoEvents
      X1 = Int(Rnd * 76)
      X2 = Int(Rnd * 76)
      Y1 = Int(Rnd * 76)
      Y2 = Int(Rnd * 76)
      colo = Int(Rnd * 15)
      frm.Line (X1, Y1)-(X2, Y2), QBColor(colo)
      colo = Int(Rnd * 15)
      frm.Line (X1, Y2)-(X2, Y1), QBColor(colo)
      colo = Int(Rnd * 15)
      frm.Line (X2, Y1)-(X1, Y2), QBColor(colo)
      colo = Int(Rnd * 15)
      frm.Line (Y1, Y2)-(X1, X2), QBColor(colo)
   Next X
End Sub

'*************************************************************************************************************
'* Centers the Caption on the Titlebar
'*************************************************************************************************************
'
Public Sub CenterC(frm As Form)
   Dim SpcF As Integer 'How many spaces can fit
   Dim clen As Integer 'caption length
   Dim oldc As String 'oldcaption
   Dim i As Integer 'not important
   'remove any spaces at the ends of the caption
   'very easy if you read it carefully
   oldc = frm.Caption
   Do While Left(oldc, 1) = Space(1)
      DoEvents
      oldc = Right(oldc, Len(oldc) - 1)
   Loop
   Do While Right(oldc, 1) = Space(1)
      DoEvents
      oldc = Left(oldc, Len(oldc) - 1)
   Loop
   '____________________________
   clen = Len(oldc)
   If InStr(oldc, "!") <> 0 Then
      If InStr(oldc, " ") <> 0 Then
         clen = clen * 1.5
      Else
         clen = clen * 1.4
      End If
   Else
      If InStr(oldc, " ") <> 0 Then
         clen = clen * 1.4
      Else
         clen = clen * 1.3
      End If
   End If
   '____________________________
   'see how many characters can fit
   'how many space can fit it the caption
   SpcF = frm.Width / 61.2244
   SpcF = SpcF - clen
   'How many spaces can fit-How much space the
   'caption takes up
   'Now the tricky part
   If SpcF > 1 Then
      DoEvents 'speed up the program
      frm.Caption = Space(Int(SpcF / 2)) + oldc
   Else
      'if the form is too small for spaces
      frm.Caption = oldc 
   End If
End Sub
'end
