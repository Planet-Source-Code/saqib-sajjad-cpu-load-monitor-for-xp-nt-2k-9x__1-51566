VERSION 5.00
Begin VB.Form pic 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "CPU Load monitor by saqibsajjad@yahoo.com"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4560
      Top             =   840
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1290
   End
   Begin VB.Menu mnu 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu interval 
         Caption         =   "Interval"
      End
      Begin VB.Menu lg 
         Caption         =   "Line Graph"
      End
      Begin VB.Menu bg 
         Caption         =   "Bar Graph"
      End
      Begin VB.Menu ot 
         Caption         =   "On top"
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "pic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Saqib Sajjad
'saqibsajjad@yahoo.com
'www.craftspakistan.com
'
'
'


Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'The PdhVbOpenQuery function creates and initializes a unique query structure that is used to manage the collection of performance data
Private Declare Function PdhVbOpenQuery Lib "PDH.DLL" (ByRef QueryHandle As Long) As Long
Private Declare Function PdhVbAddCounter Lib "PDH.DLL" (ByVal QueryHandle As Long, ByVal CounterPath As String, ByRef CounterHandle As Long) As Long
Private Declare Function PdhCollectQueryData Lib "PDH.DLL" (ByVal QueryHandle As Long) As Long
Private Declare Function PdhVbGetDoubleCounterValue Lib "PDH.DLL" (ByVal CounterHandle As Long, ByRef CounterStatus As Long) As Double
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const lType = 4
Private Const lSize = 4
Private Const HKEY_DYN_DATA As Long = &H80000006

Dim sk As Long
Dim HQ As Long ' handle to Query
Dim counter As Long ' hand
Dim once As Boolean

Dim px As Integer
Dim py As Integer
Dim nx As Integer
Dim ny As Integer
Dim graph As Boolean ' true means graph otherwise line
Dim a As Integer
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const OOMPS = SWP_NOSIZE Or SWP_NOMOVE


Private Const OFFSET = 500


Private Function IsOsWinXP() As Boolean

    Dim vi As OSVERSIONINFO
    vi.dwOSVersionInfoSize = Len(vi)
    Call GetVersionEx(vi)
    IsOsWinXP = (vi.dwPlatformId = VER_PLATFORM_WIN32_NT)
    
End Function



Private Sub Form_Load()
once = True
pic.ForeColor = vbGreen

px = 0
py = pic.Height
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then PopupMenu mnu

End Sub

Private Sub Form_Resize()
If pic.Height <= label1.Height + 500 Then pic.Height = label1.Height + 500
pic.Cls
px = 0
py = pic.Height
nx = 0
ny = pic.Height
a = 0
End Sub



Private Sub Timer1_Timer()
Dim lData As Long
Dim hKey As Long
Dim r As Long


If once = True Then 'init function is called once to initilize settings
    init
    once = False
End If

If IsOsWinXP Then


        Call PdhCollectQueryData(HQ)
         r = CLng(PdhVbGetDoubleCounterValue(counter, lData))
    label1.Caption = r & "%"
Else

        Call RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartStat", hKey)
        Call RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)
        Call RegCloseKey(hKey)
        Call RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", sk)
     
        Call RegQueryValueEx(sk, "KERNEL\CPUUsage", 0, lType, lData, lSize)
        r = CLng(lData)
End If


a = a + 35      'increment on x-axis

If r >= 70 Then ' if load is >70 then use red color
                    pic.ForeColor = vbRed
                    label1.ForeColor = vbRed
                Else
                    ' if load is less than 70 then use green color
                    pic.ForeColor = vbGreen
                    label1.ForeColor = vbGreen
            End If



If graph = False Then
            
   If r <> 0 Then  'if load is <>0 then draw a line otherwise draw a point only
        pic.Line (a, pic.Height - OFFSET)-(a, pic.Height - OFFSET - (r / 100) * pic.Height)
   Else
        pic.PSet (a, pic.Height - OFFSET)
    End If
Else

nx = a
ny = pic.Height - (r / 100) * pic.Height  '(r / 100) * pic.Height =%age w.r.t picture box height

    pic.Line (px, py - OFFSET)-(nx, ny - OFFSET)
px = nx
py = ny

End If


        If a > pic.Width Then
                a = 0
                pic.Cls
                nx = 0
                ny = 0
                px = 0
                py = 0
        End If

End Sub


Private Sub init()
Dim lData As Long
Dim hKey As Long
Dim r As Long
    If IsOsWinXP Then
        Call PdhVbOpenQuery(HQ)
        Call PdhVbAddCounter(HQ, "\Processor(0)\% Processor Time", counter)
        
        Call PdhCollectQueryData(HQ)
        Call PdhVbGetDoubleCounterValue(counter, lData)
    End If
End Sub

Private Sub interval_Click()
Dim ival As String

ival = InputBox("Enter Interval in milli second : (default is 200) ")

If ival <> "" And IsNumeric(ival) Then Timer1.Interval = ival
End Sub

Private Sub lg_click()
graph = True
 a = 0
                pic.Cls
                nx = 0
                ny = 0
                px = 0
                py = pic.Height
End Sub

Private Sub bg_click()
                graph = False
                a = 0
                pic.Cls
                nx = 0
                ny = 0
                px = 0
                py = pic.Height
End Sub

Private Sub ot_click()

If ot.Checked = False Then
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OOMPS)
    ot.Checked = True
Else
    Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, OOMPS)
    ot.Checked = False
End If
End Sub

Private Sub about_click()
   MsgBox "Coded by Saqib Sajjad " & vbCrLf & "Comments to saqibsajjad@yahoo.com " & vbCrLf & "My website www.craftspakistan.com", vbInformation, "CPU LOAD"
   ShellExecute pic.hwnd, vbNullString, "www.craftspakistan.com", vbNullString, "C:\", SW_SHOWNORMAL
End Sub
Private Sub exit_click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
   MsgBox "Coded by Saqib Sajjad " & vbCrLf & "Comments to saqibsajjad@yahoo.com " & vbCrLf & "My website www.craftspakistan.com", vbInformation, "CPU LOAD"
ShellExecute pic.hwnd, vbNullString, "www.craftspakistan.com", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

