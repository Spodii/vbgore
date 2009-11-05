Attribute VB_Name = "FilePaths"
Option Explicit

Public DataPath As String
Public Data2Path As String
Public GrhPath As String
Public GrhMapPath As String
Public MapPath As String
Public MapEXPath As String
Public MusicPath As String
Public ServerDataPath As String
Public SfxPath As String
Public MessagePath As String
Public LogPath As String
Public ServerTempPath As String
Public SignsPath As String

Public Sub InitFilePaths()
'*****************************************************************
'Set the common file paths
'More info: http://www.vbgore.com/CommonCode.FilePaths.InitFilePaths
'*****************************************************************

    DataPath = App.Path & "\Data\"
    Data2Path = App.Path & "\Data2\"
    GrhPath = App.Path & "\Grh\"
    GrhMapPath = App.Path & "\GrhMapEditor\"
    MapPath = App.Path & "\Maps\"
    MapEXPath = App.Path & "\MapsEX\"
    MusicPath = App.Path & "\Music\"
    ServerDataPath = App.Path & "\ServerData\"
    SfxPath = App.Path & "\Sfx\"
    MessagePath = DataPath & "Messages\"
    SignsPath = DataPath & "Signs\"
    LogPath = App.Path & "\Logs\"
    ServerTempPath = ServerDataPath & "_temp\"

End Sub
