VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2520
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4275
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPopUp.frx":0000
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   285
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1440
      Left            =   2745
      Picture         =   "frmPopUp.frx":240042
      ScaleHeight     =   1380
      ScaleWidth      =   1380
      TabIndex        =   2
      Top             =   750
      Width           =   1440
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   465
      Left            =   1485
      TabIndex        =   1
      Top             =   1965
      Width           =   1065
   End
   Begin VB.Timer tmrMenuPopup 
      Interval        =   10
      Left            =   330
      Top             =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by: Anuj sharma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   375
      TabIndex        =   0
      Top             =   225
      Width           =   3285
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------
'       Pop-up window like Nortan Anti-virus.
'------------------------------------------------------------------------
'       Developed by :  Anuj sharma
'       E-mail :        anujsharrma@yahoo.com
'------------------------------------------------------------------------



Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Const CONVERT_TO_TWIP_HEIGHT = 14.5
Const CONVERT_TO_TWIP_WIDTH = 14.7

Dim m_iVisible              As Integer
Dim m_lfrmTop               As Long
Dim m_lfrmLeft              As Long
Dim m_lTrayWindowHeight     As Long
Dim m_lWindowDesktopBottom  As Long

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo 0
Dim iCOunt          As Integer
Dim sClass          As String
Dim hwnd            As Long
Dim lpRect          As RECT
Dim lpDeskTopRect   As RECT

    sClass = "shell_TrayWnd"
    hwnd = FindWindow(sClass, "")
    GetClientRect hwnd, lpRect
    hwnd = GetDesktopWindow()
    GetClientRect hwnd, lpDeskTopRect
    m_iVisible = 1
    m_lTrayWindowHeight = CONVERT_TO_TWIP_HEIGHT * lpRect.Bottom
    m_lWindowDesktopBottom = CONVERT_TO_TWIP_HEIGHT * lpDeskTopRect.Bottom
    Me.Top = (m_lWindowDesktopBottom - m_lTrayWindowHeight)
    Me.Left = CONVERT_TO_TWIP_WIDTH * (lpDeskTopRect.Right - Me.ScaleWidth)
    m_lfrmTop = m_lWindowDesktopBottom - ((Me.Height * m_iVisible) + 100)
    m_lfrmLeft = Me.Left
    Me.Visible = True
   'g_FormAlertUnloaded = False
    tmrMenuPopup.Enabled = True
Exit Sub
Errhandler:
    MsgBox ("Form_Load in frmMenuPopUp")
End Sub

Private Sub tmrMenuPopup_Timer()
On Error GoTo 0
Static lCount As Long
    Me.Visible = True
    If Me.Top <= (m_lfrmTop) Then
        tmrMenuPopup.Enabled = False
        Exit Sub
    End If
    Me.Top = Me.Top - 20
Exit Sub
Errhandler:
    MsgBox ("tmrMenuPopup_Timer in frmMenuPopUp")
End Sub

