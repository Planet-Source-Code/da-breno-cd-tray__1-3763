VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   600
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Main"
      Visible         =   0   'False
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Force declare of all variables in this project
Option Explicit

'Window messages that identify mouse action
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209


' declare a new instance of the clsCDTray class
Public MyCdTray As New clsCDTray

'declare a new instance of CSysTrayIcon
Public MySysTray As New CSystrayIcon



Private Sub Form_Load()
        
    ' This funtion block initializes cd to ready and open to commands
    MyCdTray.InitCD
      
    ' set the Status flag to False to indcate the CD Tray is closed
    MyCdTray.CdTrayOpen = False
    
    ' change tooltip message to closed
    MySysTray.PopUpMessage = "CD Tray (Closed)"
    
    ' add the icon to the System Tray
    MySysTray.Initialize hWnd, Picture1.Picture, MySysTray.PopUpMessage
    MySysTray.ShowIcon
    
    
   
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'When the callback message of CSystrayIcon is WM_MOUSEMOVE,
  'the X of Form_MouseMove is used to see what happen to the
  'icon in the systray.
  Dim msgCallBackMessage As Long
  
  'To be able to compare the callback value to the window message,
  'we must divide X by Screen.TwipsPerPixelX. That represent the
  'horizontal number of twips in the screen. (1 pixel ~= 15 twips)
  msgCallBackMessage = X / Screen.TwipsPerPixelX
   
  Select Case msgCallBackMessage
    Case WM_MOUSEMOVE
      MySysTray.TipText = MySysTray.PopUpMessage
      
    Case WM_LBUTTONDOWN
       ' very simple - check status of tray and perform appropriate action
    
    If MyCdTray.CdTrayOpen = False Then    ' Tray is closed
        
        MyCdTray.OpenCDTray             ' open/eject  tray
        
        ' set the Status flag to True to indcate the CD Tray is open
        MyCdTray.CdTrayOpen = True
     
        ' change tooltip message to open
        MySysTray.PopUpMessage = "CD Tray (Open)"
        
    Else
        MyCdTray.CloseCDTray            ' close em up
        
        ' set the Status flag to False to indcate the CD Tray is closed
        MyCdTray.CdTrayOpen = False
    
        ' change tooltip message to closed
        MySysTray.PopUpMessage = "CD Tray (Closed)"
        
    End If
     
     Case WM_RBUTTONDOWN
    
        Form1.PopupMenu mnuMain
    
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
MySysTray.HideIcon
End Sub

Private Sub mnuExit_Click()
End
End Sub
