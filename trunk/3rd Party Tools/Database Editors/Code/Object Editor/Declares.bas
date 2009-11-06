Attribute VB_Name = "Declares"
Option Explicit
Public Const ServerId As Integer = 0

Public Type udtObjData
    Name As String                  'Name
    ObjType As Byte                 'Type (armor, weapon, item, etc)
    ClassReq As Integer             'Class requirement
    GrhIndex As Long                'Graphic index
    SpriteBody As Integer           'Index of the body sprite to change to
    SpriteWeapon As Integer         'Index of the weapon sprite to change to
    SpriteHair As Integer           'Index of the hair sprite to change to
    SpriteHead As Integer           'Index of the head sprite to change to
    SpriteWings As Integer          'Index of the wings sprite to change to
    WeaponType As Byte              'What type of weapon, if it is a weapon
    WeaponRange As Byte             'Range of the weapon (only applicable if a ranged WeaponType)
    UseGrh As Long                  'Grh of the object when used (projectile for ranged, slash for melee, effects for use-once)
    UseSfx As Byte                  'Sound effect played when the object is used (based on the .wav's number)
    ProjectileRotateSpeed As Byte   'How fast the projectile rotates (if at all)
    Value As Long                   'Value of the object
    RepHP As Long                   'How much HP to replenish
    RepMP As Long                   'How much MP to replenish
    RepSP As Long                   'How much SP to replenish
    RepHPP As Integer               'Percentage of HP to replenish
    RepMPP As Integer               'Percentage of MP to replenish
    RepSPP As Integer               'Percentage of SP to replenish
    ReqStr As Long                  'Required strength to use the item
    ReqAgi As Long                  'Required agility
    ReqMag As Long                  'Required magic
    ReqLvl As Long                  'Required level
    Stacking As Integer             'How much the item can be stacked up (-1 for no limit, 0 for
    AddStat(FirstModStat To NumStats) As Long   'How much to add to the stat by the SID
    Pointer As Integer
End Type

Public Objnumber As Integer


