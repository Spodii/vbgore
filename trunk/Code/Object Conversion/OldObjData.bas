Attribute VB_Name = "OldObjData"
'*******************************************************************************
'Place in this module the old variables (so that way it can load correctly)
'*******************************************************************************
Option Explicit

Public Const NumStats As Byte = 31
Public Const MAX_INVENTORY_OBJS = 9999  'Maximum number of objects per slot (same obj)
Public Const MAX_INVENTORY_SLOTS = 49   'Maximum number of slots
Public Type ObjData
    Name As String              'Name
    ObjType As Byte             'Type (armor, weapon, item, etc)
    GrhIndex As Integer         'Graphic index
    SpriteBody As Integer       'Index of the body sprite to change to
    SpriteWeapon As Integer     'Index of the weapon sprite to change to
    SpriteHair As Integer       'Index of the hair sprite to change to
    SpriteHead As Integer       'Index of the head sprite to change to
    SpriteHelm As Integer       'Index of the helmet sprite to change to
    WeaponType As Byte          'What type of weapon, if it is a weapon
    Price As Long               'Price of the object
    RepHP As Long               'How much HP to replenish
    RepMP As Long               'How much MP to replenish
    RepSP As Long               'How much SP to replenish
    RepHPP As Integer           'Percentage of HP to replenish
    RepMPP As Integer           'Percentage of MP to replenish
    RepSPP As Integer           'Percentage of SP to replenish
    AddStat(1 To NumStats) As Long  'How much to add to the stat by the SID
End Type
Public ObjData() As ObjData
Public Type Obj 'Holds info about a object
    ObjIndex As Integer     'Index of the object
    Amount As Integer       'Amount of the object
End Type
