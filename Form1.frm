VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Window Handler"
   ClientHeight    =   5805
   ClientLeft      =   1080
   ClientTop       =   1230
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5805
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   1080
      Top             =   4800
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Net Send Log"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Const BM_CLICK = &HF5
Const WM_GETTEXTLENGTH = &HE
'#############################################
'#Change MyName to the name of your computer.#
'#############################################
Const MyName = "MyComputerName"
Private Const WM_GETTEXT = &HD
Dim Messages As New Collection
Dim VarNum As Integer
Dim message(5000) As String
Dim msg$, from$, user$
Dim Tmessage(5000) As String
Dim Fmessage(5000) As String

Private Sub Form_Load()
VarNum = 0
AlwaysOnTop Form1, True
Me.Left = Screen.Width - Me.Width
Me.Top = 500
List1.Height = Me.Height - 500
End Sub

Private Sub List1_DblClick()
Dim varnum3 As String
varnum3 = Replace(Right(List1.text, 4), " ", "0")
MsgBox "Message from " & Fmessage(varnum3) & " to " & MyName & " on " & Date & " " & Tmessage(varnum3) & vbNewLine & vbNewLine & message(varnum3), vbOKOnly, "Messenger Service"
End Sub

Private Sub Timer1_Timer()
Dim lWin As Long
Dim Control As Long
Dim text As String
Dim textlen As Long
lWin = FindWindow(vbNullString, "Messenger Service ")
If lWin <> 0 Then
  Control = FindWindowEx(lWin, 0, "Static", vbNullString)
      textlen = SendMessage(Control, WM_GETTEXTLENGTH, vbNull, vbNull)
      textlen = textlen + 1
  text = Space$(textlen)
  Call SendMessage(Control&, WM_GETTEXT, textlen, ByVal text)
  
  Control = FindWindowEx(lWin, 0, "Button", vbNullString)
  SendMessage Control, BM_CLICK, 0, 0
  VarNum = VarNum + 1
  
  Tmessage(VarNum) = Time
Dim sVar1 As String
Dim sVar2 As String
sVar1 = text
sVar1 = Right(sVar1, Len(sVar1) - 13)
sVar1 = Replace(sVar1, Date, "")
    Dim counter%, t%

        For t% = 1 To Len(sVar1)
    If Mid$(sVar1, t%, 3) = " to" Then

    sVar2 = Left(sVar1, Int(counter))
    sVar1 = Right(sVar1, Len(sVar1) - Len(sVar2) - 4)
    GoTo endloop3
     Else
            counter% = counter% + 1
        End If
    Next t%
endloop3:
sVar1 = Right(sVar1, Len(sVar1) - Len(MyName))

counter = 0
t = 0
        For t% = 1 To Len(sVar1)
    If Mid$(sVar1, t%, 1) = "M" Then
    sVar1 = Right(sVar1, Len(sVar1) - t - 4)
    GoTo endloop

    
     Else
            counter% = counter% + 1
        End If
    Next t%
endloop:

Fmessage(VarNum) = sVar2
message(VarNum) = sVar1
      Dim varnum2 As String
    varnum2 = Right("000" & Str(VarNum), 4)
    'Cheap, but effective.
    List1.AddItem (sVar2 & ": " & Time & "                                     " & varnum2)

  
End If
End Sub
