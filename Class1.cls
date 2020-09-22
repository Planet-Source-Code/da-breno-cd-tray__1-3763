VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public CdTrayOpen As Boolean


Function OpenCDTray()
   mciSendString "set cd door open", 0, 0, 0
End Function



Function CloseCDTray()
    mciSendString "set cd door closed", 0, 0, 0
End Function



Function UnloadAll()
    mciSendString "close all", 0, 0, 0
End Function

Function SetCDReady()
    mciSendString "open cdaudio alias cd wait shareable", 0, 0, 0
End Function
Function SetFormat()
    mciSendString "set cd time format tmsf wait", 0, 0, 0
End Function

Function InitCD()
    CloseCDTray
    SetCDReady
    SetFormat
End Function

