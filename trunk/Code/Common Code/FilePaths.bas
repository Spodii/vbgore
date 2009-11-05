Attribute VB_Name = "FilePaths"
Option Explicit

Public CharPath As String
Public DataPath As String
Public Data2Path As String
Public GrhPath As String
Public GrhMapPath As String
Public MailPath As String
Public MapPath As String
Public MapEXPath As String
Public MusicPath As String
Public NPCsPath As String
Public OBJsPath As String
Public QuestsPath As String
Public ServerDataPath As String
Public SfxPath As String

Public Sub InitFilePaths()
'***************************************
'Set the file paths
'***************************************

    CharPath = App.Path & "\Charfile\"
    DataPath = App.Path & "\Data\"
    Data2Path = App.Path & "\Data2\"
    GrhPath = App.Path & "\Grh\"
    GrhMapPath = App.Path & "\GrhMapEditor\"
    MailPath = App.Path & "\Mail\"
    MapPath = App.Path & "\Maps\"
    MapEXPath = App.Path & "\MapsEX\"
    MusicPath = App.Path & "\Music\"
    NPCsPath = App.Path & "\NPCs\"
    OBJsPath = App.Path & "\OBJs\"
    QuestsPath = App.Path & "\Quests\"
    ServerDataPath = App.Path & "\ServerData\"
    SfxPath = App.Path & "\Sfx\"

End Sub
