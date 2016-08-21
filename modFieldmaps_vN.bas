Attribute VB_Name = "modFieldmaps_vN"
Option Base 0
Option Explicit

'***I_THROUGH_N*** (search on this comment for all the vN changes)
'set this to true if this version works with v1.11h to v1.11n (set to true for n, false for o)
Public Const WorksWithN = True
'Public Const WorksWithN = False

Const GetNextExDataBufSize = 2000
Public GetNextExDataBuf As GetNextExDatabufType
Public Type GetNextExDatabufType
    buf(1 To GetNextExDataBufSize) As Byte
End Type

'====> RACES
Public Racerec As RaceRecType
Public Racedatabuf As RaceDatabufType
Public RacePosBlock As RacePosBlockType
Public RaceKeyBuffer As String * 255
Public Type RaceRecType
    Number As Integer       '####    -   Number as Integer        2
    Name As String * 29     'NNNN    -   Name as String * 29     29
    nothing1 As Byte        '  00    -   Nothing as Byte          1
    MinInt As Integer       '-INT    -   Min Int as Integer       2
    MinWil As Integer       '-Wil    -   Min Wil as Integer       2
    MinStr As Integer       '-Str    -   Min Str as Integer       2
    MinHea As Integer       '-Hea    -   Min Hea as Integer       2
    MinAgl As Integer       '-Agl    -   Min Agl as Integer       2
    MinChm As Integer       '-Chm    -   Min Chm as Integer       2
    HPBonus As Integer      'HP00    -   HP Bonus as Integer      2
    nothing2 As Long        '0000    -   Nothing as Long          4
    AbilityA(9) As Integer  'AAA1    -   AbilityA1 as Integer         20
    CP As Integer           'CP00    -   Starting CP as Integer       2
    AbilityB(9) As Integer  'BBB1    -   AbilityB1 as Integer         20
    nothing3 As Long        '0020    -   Nothing as Long          4
    nothing4 As Integer     '0000    -   Nothing as Integer       2
    ExpChart As Integer     'EXPE    -   Exp Chart as Integer         2
    nothing5 As Integer     '0000    -   Nothing as Integer       2
    MaxInt As Integer       '+Int    -   Max Int as Integer       2
    MaxWil As Integer       '+Wil    -   Max Wil as Integer       2
    MaxStr As Integer       '+Str    -   Max Str as Integer       2
    MaxHea As Integer       '+Hea    -   Max Hea as Integer       2
    MaxAgl As Integer       '+Agl    -   Max Agl as Integer       2
    MaxChm As Integer       '+Chm    -   Max Chm as Integer       2
    Nothing6 As Long        '0000    -   Nothing as Long          4
    nothing7 As Long        '0000    -   Nothing as Long          4
    nothing8 As Long        '0000    -   Nothing as Long          4
End Type
Const RaceDataBufSize = 126
Public RaceFldMap(0 To 44) As FieldMap
Public Type RaceDatabufType
    buf(1 To RaceDataBufSize) As Byte
End Type
Public Type RacePosBlockType
    buf(1 To 128) As Byte
End Type

'====> CLASSES
Public Classrec As ClassRecType
Public Classdatabuf As ClassDatabufType
Public ClassPosBlock As ClassPosBlockType
Public ClassKeyNum As Integer
Public ClassKeyBuffer As String * 255
Public Type ClassRecType
    Number     As Integer
    Name       As String * 29
    AfterName  As Byte
    MinHp      As Integer
    MaxHP      As Integer
    Exp        As Integer
    nothing1   As Integer
    nothing2   As Integer
    nothing3   As Integer
    AbilityA(9)   As Integer
    MagicType  As Integer
    MagicLvL   As Integer
    Weapon     As Integer
    Armour     As Integer
    Combat     As Integer
    AbilityB(9)   As Integer
    nothing4   As Integer
    nothing5   As Integer
    Nothing6   As Integer
    TitleText  As Long
End Type
Const ClassDataBufSize = 156
Public ClassFldMap(0 To 37) As FieldMap
Public Type ClassDatabufType
    buf(1 To ClassDataBufSize) As Byte
End Type
Public Type ClassPosBlockType
    buf(1 To 128) As Byte
End Type


'====> SPELLS
Public Spellrec As SpellRecType
Public Spelldatabuf As SpellDatabufType
Public SpellPosBlock As SpellPosBlockType
Public SpellKeyBuffer As String * 255
Public Type SpellRecType
    Number As Integer
    Name As String * 29
    AfterName  As Byte
    DescA As String * 50
    AfterDescA As Byte
    DescB As String * 50
    AfterDescB As Byte
    N01 As Integer
    CastMsgA As Long
    N02(10) As Integer
    LevelCap As Byte
    N03 As Byte
    MsgStyle As Byte
    N04(2) As Byte
    AbilityB(9) As Integer
    Energy As Integer
    Level As Integer
    Min As Integer
    Max As Integer
    SpellType As Integer
    TypeOfResists As Integer
    Difficulty As Integer
    UNDEFINED01 As Integer
    Target As Integer
    duration As Integer
    TypeOfAttack As Integer
    UNDEFINED02 As Integer
    ResistAbility As Integer
    MageryA As Integer
    AbilityA(9) As Integer
    CastMsgB As Long
    'N05 As Integer
    Mana As Integer
    MaxIncrease As Byte
    LVLSMaxIncr As Byte
    MageryB As Integer
    MinIncrease As Byte 'u3
    LVLSMinIncr As Byte 'u4
    DurIncrease As Byte 'u5
    LVLSDurIncr As Byte 'u6
    ShortName As String * 5
    AfterShortName As Byte
    N06 As Long
End Type
Const SpellDataBufSize = 260
Public SpellFldMap(0 To 74) As FieldMap
Public Type SpellDatabufType
    buf(1 To SpellDataBufSize) As Byte
End Type
Public Type SpellPosBlockType
    buf(1 To 128) As Byte
End Type


'====> MONSTERS
Public Monsterrec As MonsterRecType
Public Monsterdatabuf As MonsterDatabufType
Public MonsterPosBlock As MonsterPosBlockType
Public MonsterKeyBuffer As String * 255
Public Type MonsterRecType
    Number          As Long      '4
    EmptySpace      As String * 50  '54
    Name            As String * 29  '83
    nothing1        As Byte
    Group           As Integer
    nothingXX1      As Integer
    ExpMulti        As Long
    'nothingXX2      As Integer
    Index           As Integer
    nothingXX3      As Integer
    Something2      As Long
    WeaponNumber    As Long         '105
    DR              As Integer
    AC              As Integer
    Something3      As Integer
    Follow          As Integer
    MR              As Integer
    BSDefence       As Integer      '116
    Experience      As Long
    'nothingXX4      As Integer
    Hitpoints       As Integer
    Energy          As Integer
    HPRegen         As Integer      '126
    AbilityA(9)     As Integer      '146
    AbilityB(9)     As Integer      '166
    GameLimit       As Integer
    Active          As Integer
    Type            As Integer
    nothing2        As Byte         '173
    Undead          As Byte
    Alignment       As Integer
    nothing3        As Integer
    RegenTime       As Integer
    DateKilled      As Integer
    TimeKilled      As Integer
    'Nothing6        As Integer      '186
    MoveMsg         As Long
    DeathMsg        As Long         '194
    ItemNumber(9)   As Long         '234
    ItemUses(9)     As Integer      '254
    ItemDropPer(9)  As Byte         '264
    nothing9        As Integer
    Runic           As Long
    Platinum        As Long
    Gold            As Long
    Silver          As Long
    Copper          As Long
    GreetTxt        As Long         '286
    CharmLvL        As Integer
    Nothing16       As Integer      '290
    DescTxt         As Long         '294
    AttackType(4)   As Byte         '299
    Nothing22       As Byte         '300
    AttackAccuSpell(4) As Integer   '310
    AttackPer(4)    As Byte         '315
    Nothing17           As Byte     '316
    AttackMinHCastPer(4) As Integer
    AttackMaxHCastLvl(4) As Integer '326
    Nothing18           As Integer  '328
    AttackHitMsg(4)       As Long
    AttackDodgeMsg(4)     As Long
    AttackMissMsg(4)      As Long   '388
    AttackEnergy(4)       As Integer    '398
    Nothing19           As Integer  '400
    TalkTxt             As Long     '404
    CharmRes            As Integer
    Nothing21           As Integer  '408
    AttackHitSpell(4)     As Integer    '418
    DeathSpellNumber    As Integer
    Nothing23           As Integer
    Nothing24           As Integer
    Nothing25           As Integer
    Nothing26           As Integer
    Nothing27           As Integer
    Nothing28           As Integer
    Nothing29           As Integer
    CreateSpellNumber   As Integer  '436
    SpellNumber(4)        As Integer    '446
    SpellCastPer(4)       As Byte
    SpellCastLvl(4)       As Byte   '456
    DescLine1       As String * 70
    nothing10       As Byte
    DescLine2       As String * 70
    Nothing11       As Byte
    DescLine3       As String * 70
    Nothing12       As Byte
    DescLine4       As String * 70
    Nothing13       As Byte
    Gender          As Byte
    Nothing14       As Byte
    Nothing15       As Integer  'section total: buf 300, 300+454=744 .. fields 184
          
End Type
Const MonsterDataBufSize = 756
Public MonsterFldMap(0 To 184) As FieldMap
Public Type MonsterDatabufType
    buf(1 To MonsterDataBufSize) As Byte
End Type
Public Type MonsterPosBlockType
    buf(1 To 128) As Byte
End Type

'====> ITEMS
Public Itemrec As ItemRecType
Public Itemdatabuf As ItemDatabufType
Public ItemPosBlock As ItemPosBlockType
Public ItemKeyBuffer As String * 255
Public Type ItemRecType
    Number          As Long                     '4
    unknown1        As Integer                  '6
    GameLimit       As Integer                  '8
    unknown2        As Integer                  '10
    unknown3        As Integer                  '12
    unknown4        As Integer                  '14
    unknown5        As Integer                  '16
    EmptySpace1     As String * 156             '172
    EmptySpace2     As Byte                     '173
    Name            As String * 29              '202
    nothing1        As Byte                     '203
    Desc1           As String * 60              '263 .... 9x61 = 549 ... 203+549=752
    nothing2        As Byte                     '264
    Desc2           As String * 60              '324
    nothing3        As Byte                     '325
    Desc3           As String * 60              '385
    nothing4        As Byte                     '386
    Desc4           As String * 60              '446
    nothing5        As Byte                     '447
    Desc5           As String * 60              '507
    Nothing6        As Byte                     '508
    Desc6           As String * 60              '568
    nothing7        As Byte                     '569
    Desc7           As String * 60              '629  'these last three lines aren't actually descriptions
    nothing8        As Byte                     '630
    Desc8           As String * 60              '690
    nothing9        As Byte                     '691
    Desc9           As String * 60              '751
    nothing10       As Byte                     '752
    unknown6        As Integer                  '754
    Weight          As Integer  'en             '756
    Type            As Integer  'item type      '758
    AbilityA(19)    As Integer                  '798 +40
    Uses            As Integer  'uses           '800
    unknown7        As Integer  'unknown        '802
    Cost            As Integer  'cost           '804
    Class(9)        As Integer                  '824 +20
    unknown8        As Integer                  '826
    unknown9        As Integer                  '828
    unknown10       As Integer                  '830
    Minhit          As Integer                  '832
    Maxhit          As Integer                  '834
    AC              As Integer                  '836
    Race(9)         As Long                     '876 +40
    Negate(19)      As Integer                  '916 +40
    Weapon          As Integer                  '918
    Armour          As Integer                  '920
    WornOn          As Integer                  '922
    Accuracy        As Integer                  '924
    DR              As Integer                  '926
    Gettable        As Byte
    unknown12       As Byte                     '928
    ReqStr          As Integer                  '930
    unknown13a(6)   As Integer
    OpenRunic       As Long
    OpenPlatinum    As Long
    OpenGold        As Long
    OpenSilver      As Long
    OpenCopper      As Long
    unknown13b(12)   As Integer                  '980 +60  '29 -- changes the numbers here 6-19-03
    Speed           As Integer                  '982
    unknown14       As Integer                  '984
    AbilityB(19)    As Integer                  '1024 +40
    unknown15       As Integer                  '1026
    HitMsg          As Long
    MissMsg         As Long
    ReadTB          As Long
    DistructMsg     As Long                     '1042
    unknown16(5)    As Integer                  '1054 +12
    NotDroppable    As Byte
    CostType        As Byte
    RetainAfterUses As Byte
    Robable       As Byte
    DestroyOnDeath  As Byte
    unknown19       As Byte                     '1060
    'unknown20(1)   As Byte
End Type
Const ItemDataBufSize = 1072 '1948 '1072
Public ItemFldMap(0 To 174) As FieldMap '188
Public Type ItemDatabufType
    buf(1 To ItemDataBufSize) As Byte
End Type
Public Type ItemPosBlockType
    buf(1 To 128) As Byte
End Type

'====> SHOPS
Public Shoprec As ShopRecType
Public Shopdatabuf As ShopDatabufType
Public ShopPosBlock As ShopPosBlockType
Public ShopKeyBuffer As String * 255
Public Type ShopRecType
    Number              As Long                 '4              '1
    Name                As String * 39          '43
    ShopAfterName           As Integer          '45
    ShopDescriptionA        As String * 52      '97
    ShopNothing1            As Byte             '98             '5
    ShopDescriptionB        As String * 52      '150
    ShopNothing2            As Byte             '151
    ShopDescriptionC        As String * 52      '203
    ShopNothing3            As Byte             '204
    ShopType                As Integer          '206            '10
    ShopMinLvL              As Integer          '208
    ShopMaxLvl              As Integer          '210
    ShopMarkUp              As Integer          '212
    ShopNothing4            As Integer          '214
    ShopClassLimit          As Byte          '215
    ShopNothingAA           As Byte             '216            '16
    ShopItemNumber(19)      As Long             '20*4=80, 296   '36
    ShopMax(19)             As Integer          '20*2=40, 336   '56
    ShopNow(19)             As Integer          '40, 376        '76
    ShopRgnTime(19)         As Integer          '40, 416        '96
    ShopRgnNumber(19)       As Integer          '40, 456        '116
    ShopRgnPercentage(19)   As Byte             '20, 476        '136
    
End Type
Const ShopDataBufSize = 504
Public ShopFldMap(0 To 135) As FieldMap
Public Type ShopDatabufType
    buf(1 To ShopDataBufSize) As Byte
End Type
Public Type ShopPosBlockType
    buf(1 To 128) As Byte
End Type


'====> ROOMS
Public Roomrec As RoomRecType
Public Roomdatabuf As RoomDatabufType
Public RoomPosBlock As RoomPosBlockType
Public RoomKeyBuffer As String * 255
Public Type RoomRecType
    MapNumber           As Long                 '4
    RoomNumber          As Long                 '8
    EmptySpace          As String * 253         '261
    Name                As String * 53          '314
    Desc(6)             As String * 71          '811
    AnsiMap             As String * 13          '824
    RoomExit(9)         As Long                 '864
    RoomType(9)         As Integer              '884
    Para1(9)            As Long                 '924
    Para2(9)            As Integer              '944
    Para3(9)            As Long                 '984
    Para4(9)            As Long                 '1024
    'CurrentRoomMon(14)  As Long              '1054             '***I_THROUGH_N*** (comment for n)
    CurrentRoomMon(14)  As Integer              '1054          '***I_THROUGH_N*** (UNcomment for n)
    Type                As Integer              '1056
    'NewSpot             As Integer                             '***I_THROUGH_N*** (comment for n)
    ShopNum             As Long              '1058
    nothing1(14)        As Integer                 '1090
    MinIndex            As Integer              '1092
    MaxIndex            As Integer              '1094
    ByNumber            As Long              '1100
    dontknow            As Integer
    Light               As Integer              '1102
    GangHouseNumber     As Integer              '1104
    RoomItems(16)       As Long
    RoomItemUses(16)    As Integer              ' -9
    nothing4            As Integer
    InvisItems(14)      As Long
    InvisItemUses(14)   As Integer              ' -8
    nothing5            As Integer              ' -8
    Runic               As Long
    Platinum            As Long
    Gold                As Long
    Silver              As Long
    Copper              As Long
    InvisRunic          As Long
    InvisPlatinum       As Long
    InvisGold           As Long
    InvisSilver         As Long
    InvisCopper         As Long
    'nothing7(4)         As Long             ' -5
    MaxRegen            As Long
    MonsterType         As Integer
    unknown69           As Integer
    Attributes          As Long
    nothing9            As Long
    DeathRoom           As Long
    ExitRoom            As Long
    RoomItemQty(16)     As Integer
    InvisItemQty(14)    As Integer
    CmdText             As Long
    nothing10           As Long             ' -1
    Delay               As Integer
    MaxArea             As Integer
    Nothing11           As Long
    ControlRoom         As Long
    PermNPC             As Long
    PlacedItems(9)      As Long
    Nothing12(1)        As Long             ' -3
    Something1          As Long
    Spell               As Long
    unknown70           As Integer
    NumMons             As Byte
    unknown71           As Byte
End Type

'Const RoomDataBufSize = 1544 '***I_THROUGH_N*** (comment for n)
Const RoomDataBufSize = 1512 '***I_THROUGH_N*** (UNcomment for n)

'Public RoomFldMap(0 To 249) As FieldMap     '***I_THROUGH_N*** (comment for n)
Public RoomFldMap(0 To 248) As FieldMap    '***I_THROUGH_N*** (UNcomment for n)
Public Type RoomDatabufType
    buf(1 To RoomDataBufSize) As Byte
End Type
Public Type RoomPosBlockType
    buf(1 To 128) As Byte
End Type

Public RoomKeyStruct As RoomKeyStructType
Public Type RoomKeyStructType
    MapNum As Long
    RoomNum As Long
End Type

'====> MESSAGES
Public Messagerec As MessageRecType
Public Messagedatabuf As MessageDatabufType
Public MessagePosBlock As MessagePosBlockType
Public MessageKeyBuffer As String * 255
Public Type MessageRecType
    Number              As Long
    MessageLine1        As String * 74
    MessageNothing02    As Integer
    MessageNothing03    As Integer
    MessageNothing04    As Integer
    MessageNothing05    As Integer
    MessageLine2        As String * 74
    MessageNothing06    As Integer
    MessageNothing07    As Integer
    MessageNothing08    As Integer
    MessageNothing09    As Integer
    MessageLine3        As String * 74
End Type
Const MessageDataBufSize = 256
Public MessageFldMap(0 To 11) As FieldMap
Public Type MessageDatabufType
    buf(1 To MessageDataBufSize) As Byte
End Type
Public Type MessagePosBlockType
    buf(1 To 128) As Byte
End Type

'====> TEXTBLOCKS
Public Const TextblockFixedSize = 24
Public PreviewLoaded As Boolean
Public TextblockRec As TextblockRecType
Public TextblockDataBuf As TextblockDataBufType
Public TextblockPosBlock As TextblockPosBlockType
Public TextblockKeyBuffer As String * 255
Public Type TextblockRecType
    PartNum As Integer
    LeadIn(1 To 14) As Byte
    Number As Long
    LinkTo As Long
    Data As String * 2000
End Type
Public Const TextblockMaxBufSize = 2024
Public TextblockFldMap(0 To 17) As FieldMap
Public Type TextblockDataBufType
    buf(1 To TextblockMaxBufSize) As Byte
End Type
Public Type TextblockPosBlockType
    buf(1 To 128) As Byte
End Type

Public TextblockKey As TextblockKeyType
Public Type TextblockKeyType
    PartNum As Integer
    LeadIn(1 To 14) As Byte
    Number As Long
End Type

'====> USERS
Public Userrec As UserRecType
Public Userdatabuf As UserDatabufType
Public UserPosBlock As UserPosBlockType
Public UserKeyBuffer As String * 255
Public Type UserRecType
    BBSName As String * 30              '1
    FirstName As String * 10            '2
    AfterFirstName As Byte              '3
    LastName As String * 18             '4
    AfterLastName As Byte               '5
    NotExperience As Long               '6
    SpellCasted(9) As Integer           '16
    SpellValue(9) As Integer            '26
    SpellRoundsLeft(9) As Integer       '36
    Title As String * 20                '37
    Race As Integer                     '38
    Class As Integer                    '39
    Level As Integer                    '40
    Stat(11) As Integer                 '52
    MaxHP As Integer                    '53
    CurrentHP As Integer                '54
    MaxENC As Integer                   '55
    CurrentENC As Integer               '56
    Energy(2) As Integer                '59
    unknown1(1) As Integer  'unknown -- seems like it's 125 for everyone    '61
    MagicRes As Integer                 '62
    MagicRes2 As Integer                '63
    MapNumber As Long                   '64
    RoomNum As Long                     '65
    nothing2 As Integer                 '66
    unknown2(1) As Integer  'unknown    '68
    nothing3 As Integer                 '69
    unknown3(1) As Byte     'unknown    '71
    nothing4 As Integer                 '72
    Item(99) As Long                    '172
    ItemUses(99) As Integer             '272
    nothing5 As Long        'always 32(hex) '273
    Key(49) As Long
    KeyUses(49) As Integer
    unknown4(3) As Long     'unknown
    BillionsOfExperience As Long
    MillionsOfExperience As Long
    Nothing6 As Integer     'always 64(hex)
    Spell(99) As Integer
    EvilPoints As Integer
    nothing7(2) As Long
    LastMap(19) As Long
    LastRoom(19) As Long
    nothing8 As Integer
    BroadcastChan As Integer
    unknown5 As Long        'seems to always be 513
    Perception As Integer
    Stealth As Integer
    MartialArts As Integer
    Thievery As Integer
    MaxMana As Integer
    CurrentMana As Integer
    SpellCasting As Integer
    Traps As Integer
    unknown6 As Integer      'some ppl this value is the same as traps, some it's close to it
    Picklocks As Integer
    Tracking As Integer
    nothing9 As Integer
    Runic As Long
    Platinum As Long
    Gold As Long
    Silver As Long
    Copper As Long
    WeaponHand As Long
    nothing10 As Long
    WornItem(19) As Long
    unknown7(19) As Integer 'unknown
    unknown8 As Integer
    LivesRemaining As Integer
    unknown9(15) As Integer  'unknown
    GangName As String * 19
    AfterGangName As Byte
    unknown11(5) As Byte    'unknown
    CPRemaining As Integer
    SuicidePassword As String * 8
    
    unknown12a(7) As Integer
    bEDITED As Byte
    unknown12c As Byte
    unknown12d(29) As Integer        'unknown
    
    Ability(29) As Integer
    AbilityModifier(29) As Integer
    unknown13a As Integer
    unknown13b As Integer
    unknown13c As Integer
    unknown13d As Integer
    unknown13e As Integer
    unknown13f As Integer
    unknown13g As Integer
    CharLife As Long
    unknown13(8) As Integer 'unknowns
    Bitmask1 As Byte
    Bitmask2 As Byte
    TestFlag1 As Byte
    TestFlag2 As Byte
    'TestFlag3 As Integer
    unknown14 As Integer
    unknown15(3) As Long
End Type
Const UserDataBufSize = 2028
Public UserFldMap(0 To 739) As FieldMap
Public Type UserDatabufType
    buf(1 To UserDataBufSize) As Byte
End Type
Public Type UserPosBlockType
    buf(1 To 128) As Byte
End Type

'====> BANKS
Public Bankrec As BankrecType
Public BankDatabuf As BankDatabufType
Public BankPosBlock As BankPosBlockType
Public BankKeyBuffer As String * 255
Public Type BankrecType
    BBSName As String * 30
    nothing1 As Integer
    ShopNumber As Long
    Cash As Long
End Type
Const BankDatabufSize = 76
Public BankFldMap(0 To 3) As FieldMap
Public Type BankDatabufType
    buf(1 To BankDatabufSize) As Byte
End Type
Public Type BankPosBlockType
    buf(1 To 128) As Byte
End Type

Public BankKey As BankKeyType
Public BankKeyDataBuf As BankKeyDataBufType
Public Type BankKeyType
    BBSName As String * 30
    nothing1 As Byte
    nothing2 As Byte
    ShopNumber As Long
End Type
Const BankKeyDataBufSize = 36
Public BankKeyFldMap(0 To 3) As FieldMap
Public Type BankKeyDataBufType
    buf(1 To BankKeyDataBufSize) As Byte
End Type

'====> ACTIONS
Public Actionrec As ActionRecType
Public ActionDatabuf As ActionDatabufType
Public ActionPosBlock As ActionPosBlockType
Public ActionKeyBuffer As String * 255
Public Type ActionRecType
    Name As String * 29
    AfterName As Byte
    SingleToUser As String * 74
    nothing2(2) As Integer
    SingleToRoom As String * 74
    nothing3(2) As Integer
    UserToUser As String * 74
    nothing4(2) As Integer
    UserToOtherUser As String * 74
    nothing5(2) As Integer
    UserToRoom As String * 74
    Nothing6(2) As Integer
    MonsterToUser As String * 74
    nothing7(2) As Integer
    MonsterToRoom As String * 74
    nothing8(2) As Integer
    InventoryToUser As String * 74
    nothing9(2) As Integer
    InventoryToRoom As String * 74
    nothing10(2) As Integer
    FloorItemToUser As String * 74
    Nothing11(2) As Integer
    FloorItemToRoom As String * 74
    Offset As Integer
End Type
Const ActionDataBufSize = 1010
Public ActionFldMap(0 To 43) As FieldMap
Public Type ActionDatabufType
    buf(1 To ActionDataBufSize) As Byte
End Type
Public Type ActionPosBlockType
    buf(1 To 128) As Byte
End Type



'====> GANGS
Public Gangrec As GangrecType
Public GangDatabuf As GangDatabufType
Public GangPosBlock As GangPosBlockType
Public GangKeyBuffer As String * 255
Public Type GangrecType
    KeyName As String * 20
    DisplayName As String * 20
    Exp As Long
    Leader As String * 30
    DateCreated As Integer
    unknown1 As Integer
    Members As Integer
    unknown2 As Long
    unknown3 As Long
    RollOver As Long
    RollTimes As Long
    unknown5 As String * 160
End Type
    
'    nothing1 As Integer
'    ShopNumber As Long
'    Cash As Long

Const GangDatabufSize = 256
Public GangFldMap(0 To 11) As FieldMap
Public Type GangDatabufType
    buf(1 To GangDatabufSize) As Byte
End Type
Public Type GangPosBlockType
    buf(1 To 128) As Byte
End Type


'====> UPDATEFILE
Public UpdateFileSpec As UpdateFileSpecType
Public Type UpdateFileSpecType
    RecordLength As Integer     'file specs
    PageSize As Integer
    IndexCount As Byte
    FileVersion As Byte
    Reserved(3) As Byte
    FileFlags As Integer
    DuplicatePointCount As Byte
    NotUsed As Byte
    Allocation As Integer
    KeyPosition As Integer      'key specs
    KeyLength As Integer
    KeyFlags As Integer
    Reserved2(3) As Byte
    ExtDataType As Byte
    NullValue As Byte
    NotUsed2(1) As Byte
    ManKeyNumber As Byte
    ACSNumber As Byte
End Type

Public Const UpdateDataBufSize = 2024
Public UpdateKeyBuffer As String * 255
Public Updatebuf As UpdateBufType
Public UpdatePosBlock As UpdatePosBlockType
Public Updaterec As UpdateRecType

Public Type UpdateBufType
    Data(1 To UpdateDataBufSize) As Byte
End Type

Public Type UpdateRecType
    recnumber As Long
    filenum As Long
    Data(1 To UpdateDataBufSize) As Byte
End Type

Public Type UpdatePosBlockType
    buf(1 To 128) As Byte
End Type


'====> STAT
Public DBStat As DBStatType
Public Type DBStatType
    RecLen As Integer
    PageSize As Integer
    nIndexes As Integer
    nRecords As Long
    FileFlags As Integer
    ReservedWord As Integer
    UnusedPages As Integer
    KeyPosition As Integer
    KeyLength As Integer
    KeyFlags As Integer
    nUniqueKeys As Long
    ExtendedDataType As Byte '27
    TheRest As String * 1893
End Type
Public Const DBStatBufSize = 1920
Public DBStatFldMap(0 To 12) As FieldMap
Public Type DBStatDatabufType
    buf(1 To DBStatBufSize) As Byte
End Type
Public DBStatDatabuf As DBStatDatabufType


Sub IntFieldMaps()
    AddRaceFieldMap RaceFldMap, 0
    AddClassFieldMap ClassFldMap, 0
    AddSpellFieldMap SpellFldMap, 0
    AddMonsterFieldMap MonsterFldMap, 0
    AddItemFieldMap ItemFldMap, 0
    AddShopFieldMap ShopFldMap, 0
    AddRoomFieldMap RoomFldMap, 0
    AddMessageFieldMap MessageFldMap, 0
    AddTextblockFieldMap TextblockFldMap, 0
    AddUserFieldMap UserFldMap, 0
    AddActionFieldMap ActionFldMap, 0
    AddBankFieldMap BankFldMap, 0
    AddBankKeyFieldMap BankKeyFldMap, 0
    AddGangFieldMap GangFldMap, 0
    AddDBStatFieldMap DBStatFldMap, 0
End Sub

Sub AddRoomFilterFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    AddField Map, ctr, FLD_INTEGER, 4 'map
    AddField Map, ctr, FLD_INTEGER, 4 'room
    AddField Map, ctr, FLD_INTEGER, 4 'value
End Sub
Sub AddDBStatFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 1893
End Sub
Sub AddGangFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    
    AddField Map, ctr, FLD_STRING, 20   'key name
    AddField Map, ctr, FLD_STRING, 20   'display name
    AddField Map, ctr, FLD_INTEGER, 4   'exp
    AddField Map, ctr, FLD_STRING, 30   'leader
    AddField Map, ctr, FLD_INTEGER, 2   'date created
    AddField Map, ctr, FLD_INTEGER, 2   'unknown
    AddField Map, ctr, FLD_INTEGER, 2   '#members
    AddField Map, ctr, FLD_INTEGER, 4   'unknown
    AddField Map, ctr, FLD_INTEGER, 4   'unknown
    AddField Map, ctr, FLD_INTEGER, 4   'rollover
    AddField Map, ctr, FLD_INTEGER, 4   'rolltimes
    AddField Map, ctr, FLD_STRING, 160   'unknown

End Sub
Sub AddBankKeyFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    AddField Map, ctr, FLD_STRING, 30   'bbs name
    AddField Map, ctr, FLD_BYTE, 1   'nothing
    AddField Map, ctr, FLD_BYTE, 1   'nothing
    AddField Map, ctr, FLD_INTEGER, 4   'Shop number
End Sub

Sub AddBankFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    AddField Map, ctr, FLD_STRING, 30   'bbs name
    AddField Map, ctr, FLD_INTEGER, 2   'nothing
    AddField Map, ctr, FLD_INTEGER, 4   'Shop number
    AddField Map, ctr, FLD_INTEGER, 4   'Cash amount
End Sub

Sub AddUserFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    AddField Map, ctr, FLD_STRING, 30   'bbs name
    AddField Map, ctr, FLD_STRING, 10   'firstname
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 18   'lastname
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_INTEGER, 4   'notexperience
    AddField Map, ctr, FLD_INTEGER, 2   'splcast 1
    AddField Map, ctr, FLD_INTEGER, 2   '2
    AddField Map, ctr, FLD_INTEGER, 2   '3
    AddField Map, ctr, FLD_INTEGER, 2   '4
    AddField Map, ctr, FLD_INTEGER, 2   '5
    AddField Map, ctr, FLD_INTEGER, 2   '6
    AddField Map, ctr, FLD_INTEGER, 2   '7
    AddField Map, ctr, FLD_INTEGER, 2   '8
    AddField Map, ctr, FLD_INTEGER, 2   '9
    AddField Map, ctr, FLD_INTEGER, 2   'splcast 10
    AddField Map, ctr, FLD_INTEGER, 2   'splval 1
    AddField Map, ctr, FLD_INTEGER, 2   '2
    AddField Map, ctr, FLD_INTEGER, 2   '3
    AddField Map, ctr, FLD_INTEGER, 2   '4
    AddField Map, ctr, FLD_INTEGER, 2   '5
    AddField Map, ctr, FLD_INTEGER, 2   '6
    AddField Map, ctr, FLD_INTEGER, 2   '7
    AddField Map, ctr, FLD_INTEGER, 2   '8
    AddField Map, ctr, FLD_INTEGER, 2   '9
    AddField Map, ctr, FLD_INTEGER, 2   'splval 10
    AddField Map, ctr, FLD_INTEGER, 2   'splround 1
    AddField Map, ctr, FLD_INTEGER, 2   '2
    AddField Map, ctr, FLD_INTEGER, 2   '3
    AddField Map, ctr, FLD_INTEGER, 2   '4
    AddField Map, ctr, FLD_INTEGER, 2   '5
    AddField Map, ctr, FLD_INTEGER, 2   '6
    AddField Map, ctr, FLD_INTEGER, 2   '7
    AddField Map, ctr, FLD_INTEGER, 2   '8
    AddField Map, ctr, FLD_INTEGER, 2   '9
    AddField Map, ctr, FLD_INTEGER, 2   'splround 10
    AddField Map, ctr, FLD_STRING, 20   'title
    AddField Map, ctr, FLD_INTEGER, 2   'Race
    AddField Map, ctr, FLD_INTEGER, 2   'Class
    AddField Map, ctr, FLD_INTEGER, 2   'Level
    AddField Map, ctr, FLD_INTEGER, 2   'Stat1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   'Stat12
    AddField Map, ctr, FLD_INTEGER, 2   'MaxHP
    AddField Map, ctr, FLD_INTEGER, 2   'CurrentHP
    AddField Map, ctr, FLD_INTEGER, 2   'max encum
    AddField Map, ctr, FLD_INTEGER, 2   'cur encum
    AddField Map, ctr, FLD_INTEGER, 2   'energy
    AddField Map, ctr, FLD_INTEGER, 2   'energy
    AddField Map, ctr, FLD_INTEGER, 2   'energy
    AddField Map, ctr, FLD_INTEGER, 2   '?
    AddField Map, ctr, FLD_INTEGER, 2   '?
    AddField Map, ctr, FLD_INTEGER, 2   'mr
    AddField Map, ctr, FLD_INTEGER, 2   'mr
    AddField Map, ctr, FLD_INTEGER, 4   'map
    AddField Map, ctr, FLD_INTEGER, 4   'room
    AddField Map, ctr, FLD_INTEGER, 2   'nothing
    AddField Map, ctr, FLD_INTEGER, 2   '?
    AddField Map, ctr, FLD_INTEGER, 2   '?
    AddField Map, ctr, FLD_INTEGER, 2   'nothing
    AddField Map, ctr, FLD_INTEGER, 1   '?
    AddField Map, ctr, FLD_INTEGER, 1   '?
    AddField Map, ctr, FLD_INTEGER, 2   'nothing
    AddField Map, ctr, FLD_INTEGER, 4 'Item 1
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4 'Item25
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4 'Item50
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4 'Item75
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4 'Item100
    AddField Map, ctr, FLD_INTEGER, 2 'ItemUses1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'ItemUses25
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'ItemUses50
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'ItemUses75
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'ItemUses100
    AddField Map, ctr, FLD_INTEGER, 4 'nothing
    AddField Map, ctr, FLD_INTEGER, 4 'keys 1
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4 'keys 25
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4 'keys 50
    AddField Map, ctr, FLD_INTEGER, 2 'key uses 1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'key uses 25
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'key uses 50
    AddField Map, ctr, FLD_INTEGER, 4 'unknown
    AddField Map, ctr, FLD_INTEGER, 4 'unknown
    AddField Map, ctr, FLD_INTEGER, 4 'unknown
    AddField Map, ctr, FLD_INTEGER, 4 'unknown
    AddField Map, ctr, FLD_INTEGER, 4 'exp bill
    AddField Map, ctr, FLD_INTEGER, 4 'exp mill
    AddField Map, ctr, FLD_INTEGER, 2 'nothing
    AddField Map, ctr, FLD_INTEGER, 2 'spell 1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'spell 25
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'spell 50
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'spell 75
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'spell 100
    AddField Map, ctr, FLD_INTEGER, 2 'evil points
    AddField Map, ctr, FLD_INTEGER, 4 'nothing
    AddField Map, ctr, FLD_INTEGER, 4 'nothing
    AddField Map, ctr, FLD_INTEGER, 4 'nothing
    AddField Map, ctr, FLD_INTEGER, 4 'last map 1
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4 '10
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4 'last map 20
    AddField Map, ctr, FLD_INTEGER, 4 'last room 1
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4 '10
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4 'last room 20
    AddField Map, ctr, FLD_INTEGER, 2 'nothing
    AddField Map, ctr, FLD_INTEGER, 2 'broadcast chan
    AddField Map, ctr, FLD_INTEGER, 4 'unknown
    AddField Map, ctr, FLD_INTEGER, 2 'percep
    AddField Map, ctr, FLD_INTEGER, 2 'stealth
    AddField Map, ctr, FLD_INTEGER, 2 'ma
    AddField Map, ctr, FLD_INTEGER, 2 'thievery
    AddField Map, ctr, FLD_INTEGER, 2 'max mana
    AddField Map, ctr, FLD_INTEGER, 2 'cur mana
    AddField Map, ctr, FLD_INTEGER, 2 'spell cast
    AddField Map, ctr, FLD_INTEGER, 2 'traps
    AddField Map, ctr, FLD_INTEGER, 2 'unknown
    AddField Map, ctr, FLD_INTEGER, 2 'pick locks
    AddField Map, ctr, FLD_INTEGER, 2 'tracking
    AddField Map, ctr, FLD_INTEGER, 2 'nothing
    AddField Map, ctr, FLD_INTEGER, 4 'runic
    AddField Map, ctr, FLD_INTEGER, 4 'plat
    AddField Map, ctr, FLD_INTEGER, 4 'gold
    AddField Map, ctr, FLD_INTEGER, 4 'silver
    AddField Map, ctr, FLD_INTEGER, 4 'copper
    AddField Map, ctr, FLD_INTEGER, 4 'weapon in hand
    AddField Map, ctr, FLD_INTEGER, 4 'nothing
    AddField Map, ctr, FLD_INTEGER, 4 'worn item 1
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4 '10
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4 'worn item 20
    AddField Map, ctr, FLD_INTEGER, 2 'unknown 1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 '10
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'unknown 20
    AddField Map, ctr, FLD_INTEGER, 2 'unknown
    AddField Map, ctr, FLD_INTEGER, 2 'lives remaining
    AddField Map, ctr, FLD_INTEGER, 2 'unknown 1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'unknown 10
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'unknown 16
    AddField Map, ctr, FLD_STRING, 19 'gang name
    AddField Map, ctr, FLD_INTEGER, 1 'after gang name
    AddField Map, ctr, FLD_INTEGER, 1 'unknown 1
    AddField Map, ctr, FLD_INTEGER, 1
    AddField Map, ctr, FLD_INTEGER, 1
    AddField Map, ctr, FLD_INTEGER, 1
    AddField Map, ctr, FLD_INTEGER, 1
    AddField Map, ctr, FLD_INTEGER, 1 'unknown 6
    AddField Map, ctr, FLD_INTEGER, 2 'cps
    AddField Map, ctr, FLD_STRING, 8 'suicide
    AddField Map, ctr, FLD_INTEGER, 2 'unknown 1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    
    
    'bEDITED As Byte
    'unknown12c As Byte
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    
    'AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 '10
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 '20
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 '30
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'unknown 39
    AddField Map, ctr, FLD_INTEGER, 2 'ability 1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 '10
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 '20
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'ability 30
    AddField Map, ctr, FLD_INTEGER, 2 'ability mod 1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 '10
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 '20
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'ability mod 30
    AddField Map, ctr, FLD_INTEGER, 2 'unknown13a
    AddField Map, ctr, FLD_INTEGER, 2 'unknown13b
    AddField Map, ctr, FLD_INTEGER, 2 'unknown13c
    AddField Map, ctr, FLD_INTEGER, 2 'unknown13d
    AddField Map, ctr, FLD_INTEGER, 2 'unknown13e
    AddField Map, ctr, FLD_INTEGER, 2 'unknown13f
    AddField Map, ctr, FLD_INTEGER, 2 'unknown13g
    AddField Map, ctr, FLD_INTEGER, 4 'charlife
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 '10
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 '15
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_BYTE, 1 'bitmasktest1
    AddField Map, ctr, FLD_BYTE, 1 'bitmasktest2
    AddField Map, ctr, FLD_BYTE, 1 'test1
    AddField Map, ctr, FLD_BYTE, 1 'test2
    'AddField map, ctr, FLD_INTEGER, 2 'test3
    AddField Map, ctr, FLD_INTEGER, 2 'unknown14 -1
    AddField Map, ctr, FLD_INTEGER, 4 'unknown15 -1
    AddField Map, ctr, FLD_INTEGER, 4 'unknown15 -2
    AddField Map, ctr, FLD_INTEGER, 4 'unknown15 -3
    AddField Map, ctr, FLD_INTEGER, 4 'unknown15 -4
End Sub

Sub AddTextblockFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_INTEGER, 4
    AddField Map, ctr, FLD_STRING, 2000
End Sub

Sub AddRaceFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    AddField Map, ctr, FLD_INTEGER, 2       '####    -   Number as Integer            2
    AddField Map, ctr, FLD_STRING, 29       'NNNN    -   Name as String * 29     29
    AddField Map, ctr, FLD_BYTE, 1          '  00    -   Nothing as Byte          1
    AddField Map, ctr, FLD_INTEGER, 2       '-INT    -   Min Int as Integer       2
    AddField Map, ctr, FLD_INTEGER, 2       '-Wil    -   Min Wil as Integer       2
    AddField Map, ctr, FLD_INTEGER, 2       '-Str    -   Min Str as Integer       2
    AddField Map, ctr, FLD_INTEGER, 2       '-Hea    -   Min Hea as Integer       2
    AddField Map, ctr, FLD_INTEGER, 2       '-Agl    -   Min Agl as Integer       2
    AddField Map, ctr, FLD_INTEGER, 2       '-Chm    -   Min Chm as Integer       2
    AddField Map, ctr, FLD_INTEGER, 2       'HP00    -   HP Bonus as Integer      2
    AddField Map, ctr, FLD_INTEGER, 4       '0000    -   Nothing as Long          4
    AddField Map, ctr, FLD_INTEGER, 2       'AAA1    -   AbilityA1 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA2    -   AbilityA2 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA3    -   AbilityA3 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA4    -   AbilityA4 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA5    -   AbilityA5 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA6    -   AbilityA6 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA7    -   AbilityA7 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA8    -   AbilityA8 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA9    -   AbilityA9 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AA10    -   AbilityA10 as Integer        2
    AddField Map, ctr, FLD_INTEGER, 2       'CP00    -   Starting CP as Integer       2
    AddField Map, ctr, FLD_INTEGER, 2       'BBB1    -   AbilityB1 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'BBB2    -   AbilityB2 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'BBB3    -   AbilityB3 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'BBB4    -   AbilityB4 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'BBB5    -   AbilityB5 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'BBB6    -   AbilityB6 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'BBB7    -   AbilityB7 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'BBB8    -   AbilityB8 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'BBB9    -   AbilityB9 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'BB10    -   AbilityB10 as Integer        2
    AddField Map, ctr, FLD_INTEGER, 4       '0020    -   Nothing as Long          4
    AddField Map, ctr, FLD_INTEGER, 2       '0000    -   Nothing as Integer       2
    AddField Map, ctr, FLD_INTEGER, 2       'EXPE    -   Exp Chart as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       '0000    -   Nothing as Integer       2
    AddField Map, ctr, FLD_INTEGER, 2       '+Int    -   Max Int as Integer       2
    AddField Map, ctr, FLD_INTEGER, 2       '+Wil    -   Max Wil as Integer       2
    AddField Map, ctr, FLD_INTEGER, 2       '+Str    -   Max Str as Integer       2
    AddField Map, ctr, FLD_INTEGER, 2       '+Hea    -   Max Hea as Integer       2
    AddField Map, ctr, FLD_INTEGER, 2       '+Agl    -   Max Agl as Integer       2
    AddField Map, ctr, FLD_INTEGER, 2       '+Chm    -   Max Chm as Integer       2
    AddField Map, ctr, FLD_INTEGER, 4       '0000    -   Nothing as Long          4
    AddField Map, ctr, FLD_INTEGER, 4       '0000    -   Nothing as Long          4
    AddField Map, ctr, FLD_INTEGER, 4       '0000    -   Nothing as Long          4
End Sub

Sub AddClassFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    AddField Map, ctr, FLD_INTEGER, 2 'CNumber
    AddField Map, ctr, FLD_STRING, 29 'CName
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_INTEGER, 2 'MinHP
    AddField Map, ctr, FLD_INTEGER, 2 'MaxHP
    AddField Map, ctr, FLD_INTEGER, 2 'Exp %
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 'AbilA1
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 'AbilA5
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 'AbilA10
    AddField Map, ctr, FLD_INTEGER, 2 'MagicType
    AddField Map, ctr, FLD_INTEGER, 2 'MagicLvl
    AddField Map, ctr, FLD_INTEGER, 2 'Weapon
    AddField Map, ctr, FLD_INTEGER, 2 'Armour
    AddField Map, ctr, FLD_INTEGER, 2 'Combat
    AddField Map, ctr, FLD_INTEGER, 2 'AbilB1
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 'AbilB5
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 'AbilB10
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 4 'Title
End Sub

Sub AddSpellFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    AddField Map, ctr, FLD_INTEGER, 2 'SNumber
    AddField Map, ctr, FLD_STRING, 29 'SName
    AddField Map, ctr, FLD_BYTE, 1    '
    AddField Map, ctr, FLD_STRING, 50 'SDescriptionA
    AddField Map, ctr, FLD_BYTE, 1    '
    AddField Map, ctr, FLD_STRING, 50 'SDescriptionB
    AddField Map, ctr, FLD_BYTE, 1    '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 4 'Cast MsgA
    AddField Map, ctr, FLD_INTEGER, 2 '1
    AddField Map, ctr, FLD_INTEGER, 2 '2
    AddField Map, ctr, FLD_INTEGER, 2 '3
    AddField Map, ctr, FLD_INTEGER, 2 '4
    AddField Map, ctr, FLD_INTEGER, 2 '5
    AddField Map, ctr, FLD_INTEGER, 2 '6
    AddField Map, ctr, FLD_INTEGER, 2 '7
    AddField Map, ctr, FLD_INTEGER, 2 '8
    AddField Map, ctr, FLD_INTEGER, 2 '9
    AddField Map, ctr, FLD_INTEGER, 2 '10
    AddField Map, ctr, FLD_INTEGER, 2 '11
    AddField Map, ctr, FLD_BYTE, 1    'SLevelCap
    AddField Map, ctr, FLD_BYTE, 1    '
    AddField Map, ctr, FLD_BYTE, 1    '
    AddField Map, ctr, FLD_BYTE, 1    '
    AddField Map, ctr, FLD_BYTE, 1    '
    AddField Map, ctr, FLD_BYTE, 1    '
    AddField Map, ctr, FLD_INTEGER, 2 'SAbilB1
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 'SabilB10
    AddField Map, ctr, FLD_INTEGER, 2 'SEnergy
    AddField Map, ctr, FLD_INTEGER, 2 'SLvl
    AddField Map, ctr, FLD_INTEGER, 2 'SMin
    AddField Map, ctr, FLD_INTEGER, 2 'SMax
    AddField Map, ctr, FLD_INTEGER, 2 'SStyle
    AddField Map, ctr, FLD_INTEGER, 2 '33
    AddField Map, ctr, FLD_INTEGER, 2 'SDifficulty
    AddField Map, ctr, FLD_INTEGER, 2 '35
    AddField Map, ctr, FLD_INTEGER, 2 'STarget
    AddField Map, ctr, FLD_INTEGER, 2 'SLength
    AddField Map, ctr, FLD_INTEGER, 2 'SElement
    AddField Map, ctr, FLD_INTEGER, 2 'UNDEFINED2
    AddField Map, ctr, FLD_INTEGER, 2 'ResistAbil
    AddField Map, ctr, FLD_INTEGER, 2 'STypeA
    AddField Map, ctr, FLD_INTEGER, 2 'SAbilA1
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 '
    AddField Map, ctr, FLD_INTEGER, 2 'SAbilA10
    AddField Map, ctr, FLD_INTEGER, 4 'SCastMsg
    'AddField map, ctr, FLD_INTEGER, 2 '53
    AddField Map, ctr, FLD_INTEGER, 2 'SMana
    AddField Map, ctr, FLD_BYTE, 1    'SLvlMod
    AddField Map, ctr, FLD_BYTE, 1    'SIncrease
    AddField Map, ctr, FLD_INTEGER, 2 'STypeB
    AddField Map, ctr, FLD_BYTE, 1    '
    AddField Map, ctr, FLD_BYTE, 1    '
    AddField Map, ctr, FLD_BYTE, 1    '
    AddField Map, ctr, FLD_BYTE, 1    '
    AddField Map, ctr, FLD_STRING, 5  'SShortName
    AddField Map, ctr, FLD_BYTE, 1    '
    AddField Map, ctr, FLD_INTEGER, 4 '
End Sub

Sub AddMonsterFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    AddField Map, ctr, FLD_INTEGER, 4   'Number
    AddField Map, ctr, FLD_STRING, 50
    AddField Map, ctr, FLD_STRING, 29   'Name
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_INTEGER, 2   'Group
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 4   'exp multi
    'AddField map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   'Index
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 4   'Something
    AddField Map, ctr, FLD_INTEGER, 4   'Weapon
    AddField Map, ctr, FLD_INTEGER, 2   'DR
    AddField Map, ctr, FLD_INTEGER, 2   'AC
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   'Follow%
    AddField Map, ctr, FLD_INTEGER, 2   'MR
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 4   'Experience
    'AddField map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   'Hitpoints
    AddField Map, ctr, FLD_INTEGER, 2   'Energy
    AddField Map, ctr, FLD_INTEGER, 2   'HPRegen
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityA1
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityA2
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityA3
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityA4
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityA5
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityA6
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityA7
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityA8
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityA9
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityA10
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityB1
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityB2
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityB3
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityB4
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityB5
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityB6
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityB7
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityB8
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityB9
    AddField Map, ctr, FLD_INTEGER, 2   'AbilityB10
    AddField Map, ctr, FLD_INTEGER, 2   'GameLimit
    AddField Map, ctr, FLD_INTEGER, 2   'Active
    AddField Map, ctr, FLD_INTEGER, 2   'Type
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1      'Undead
    AddField Map, ctr, FLD_INTEGER, 2   'Alignment
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   'Regen Time
    AddField Map, ctr, FLD_INTEGER, 2   'date killed
    AddField Map, ctr, FLD_INTEGER, 2   'time killed
    AddField Map, ctr, FLD_INTEGER, 4   'MoveMsg
    AddField Map, ctr, FLD_INTEGER, 4   'DeathMsg
    AddField Map, ctr, FLD_INTEGER, 4   'ItemNumber1
    AddField Map, ctr, FLD_INTEGER, 4   'ItemNumber2
    AddField Map, ctr, FLD_INTEGER, 4   'ItemNumber3
    AddField Map, ctr, FLD_INTEGER, 4   'ItemNumber4
    AddField Map, ctr, FLD_INTEGER, 4   'ItemNumber5
    AddField Map, ctr, FLD_INTEGER, 4   'ItemNumber6
    AddField Map, ctr, FLD_INTEGER, 4   'ItemNumber7
    AddField Map, ctr, FLD_INTEGER, 4   'ItemNumber8
    AddField Map, ctr, FLD_INTEGER, 4   'ItemNumber9
    AddField Map, ctr, FLD_INTEGER, 4   'ItemNumber10
    AddField Map, ctr, FLD_INTEGER, 2   'ItemUses1
    AddField Map, ctr, FLD_INTEGER, 2   'ItemUses2
    AddField Map, ctr, FLD_INTEGER, 2   'ItemUses3
    AddField Map, ctr, FLD_INTEGER, 2   'ItemUses4
    AddField Map, ctr, FLD_INTEGER, 2   'ItemUses5
    AddField Map, ctr, FLD_INTEGER, 2   'ItemUses6
    AddField Map, ctr, FLD_INTEGER, 2   'ItemUses7
    AddField Map, ctr, FLD_INTEGER, 2   'ItemUses8
    AddField Map, ctr, FLD_INTEGER, 2   'ItemUses9
    AddField Map, ctr, FLD_INTEGER, 2   'ItemUses10
    AddField Map, ctr, FLD_BYTE, 1      'ItemDropPer1
    AddField Map, ctr, FLD_BYTE, 1      'ItemDropPer2
    AddField Map, ctr, FLD_BYTE, 1      'ItemDropPer3
    AddField Map, ctr, FLD_BYTE, 1      'ItemDropPer4
    AddField Map, ctr, FLD_BYTE, 1      'ItemDropPer5
    AddField Map, ctr, FLD_BYTE, 1      'ItemDropPer6
    AddField Map, ctr, FLD_BYTE, 1      'ItemDropPer7
    AddField Map, ctr, FLD_BYTE, 1      'ItemDropPer8
    AddField Map, ctr, FLD_BYTE, 1      'ItemDropPer9
    AddField Map, ctr, FLD_BYTE, 1      'ItemDropPer10
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 4   'Runic
    AddField Map, ctr, FLD_INTEGER, 4   'Platinum
    AddField Map, ctr, FLD_INTEGER, 4   'Gold
    AddField Map, ctr, FLD_INTEGER, 4   'Silver
    AddField Map, ctr, FLD_INTEGER, 4   'Copper
    AddField Map, ctr, FLD_INTEGER, 4   'GreetTxt
    AddField Map, ctr, FLD_INTEGER, 2   'CharmLvL
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 4   'DescTxt
    AddField Map, ctr, FLD_BYTE, 1      'AttackType1     As Byte
    AddField Map, ctr, FLD_BYTE, 1      'AttackType2     As Byte
    AddField Map, ctr, FLD_BYTE, 1      'AttackType3     As Byte
    AddField Map, ctr, FLD_BYTE, 1      'AttackType4     As Byte
    AddField Map, ctr, FLD_BYTE, 1      'AttackType5     As Byte
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_INTEGER, 2   'AttackAccuSpell1    As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackAccuSpell2    As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackAccuSpell3    As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackAccuSpell4    As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackAccuSpell5    As Integer
    AddField Map, ctr, FLD_BYTE, 1      'AttackPer1          As Byte
    AddField Map, ctr, FLD_BYTE, 1      'AttackPer2          As Byte
    AddField Map, ctr, FLD_BYTE, 1      'AttackPer3          As Byte
    AddField Map, ctr, FLD_BYTE, 1      'AttackPer4          As Byte
    AddField Map, ctr, FLD_BYTE, 1      'AttackPer5          As Byte
    AddField Map, ctr, FLD_BYTE, 1      'Nothing17           As Byte
    AddField Map, ctr, FLD_INTEGER, 2   'AttackMinHCastPer1  As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackMinHCastPer2  As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackMinHCastPer3  As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackMinHCastPer4  As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackMinHCastPer5  As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackMaxHCastLvl1  As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackMaxHCastLvl2  As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackMaxHCastLvl3  As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackMaxHCastLvl4  As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackMaxHCastLvl5  As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'Nothing18           As Integer
    AddField Map, ctr, FLD_INTEGER, 4   'AttackHitMsg1       As Long
    AddField Map, ctr, FLD_INTEGER, 4   'AttackHitMsg2       As Long
    AddField Map, ctr, FLD_INTEGER, 4   'AttackHitMsg3       As Long
    AddField Map, ctr, FLD_INTEGER, 4   'AttackHitMsg4       As Long
    AddField Map, ctr, FLD_INTEGER, 4   'AttackHitMsg5       As Long
    AddField Map, ctr, FLD_INTEGER, 4   'AttackDodgeMsg1     As Long
    AddField Map, ctr, FLD_INTEGER, 4   'AttackDodgeMsg2     As Long
    AddField Map, ctr, FLD_INTEGER, 4   'AttackDodgeMsg3     As Long
    AddField Map, ctr, FLD_INTEGER, 4   'AttackDodgeMsg4     As Long
    AddField Map, ctr, FLD_INTEGER, 4   'AttackDodgeMsg5     As Long
    AddField Map, ctr, FLD_INTEGER, 4   'AttackMissMsg1      As Long
    AddField Map, ctr, FLD_INTEGER, 4   'AttackMissMsg2      As Long
    AddField Map, ctr, FLD_INTEGER, 4   'AttackMissMsg3      As Long
    AddField Map, ctr, FLD_INTEGER, 4   'AttackMissMsg4      As Long
    AddField Map, ctr, FLD_INTEGER, 4   'AttackMissMsg5      As Long
    AddField Map, ctr, FLD_INTEGER, 2   'AttackEnergy1       As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackEnergy2       As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackEnergy3       As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackEnergy4       As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackEnergy5       As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'Nothing19           As Integer
    AddField Map, ctr, FLD_INTEGER, 4   'TalkTxt             As Long
    AddField Map, ctr, FLD_INTEGER, 2   'Nothing20           As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'Nothing21           As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackHitSpell1     As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackHitSpell2     As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackHitSpell3     As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackHitSpell4     As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'AttackHitSpell5     As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'DeathSpellNumber    As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'Nothing23           As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'Nothing24           As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'Nothing25           As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'Nothing26           As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'Nothing27           As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'Nothing28           As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'Nothing29           As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'CreateSpellNumber   As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'SpellNumber1        As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'SpellNumber2        As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'SpellNumber3        As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'SpellNumber4        As Integer
    AddField Map, ctr, FLD_INTEGER, 2   'SpellNumber5        As Integer
    AddField Map, ctr, FLD_BYTE, 1      'SpellCastPer1       As Byte
    AddField Map, ctr, FLD_BYTE, 1      'SpellCastPer2       As Byte
    AddField Map, ctr, FLD_BYTE, 1      'SpellCastPer3       As Byte
    AddField Map, ctr, FLD_BYTE, 1      'SpellCastPer4       As Byte
    AddField Map, ctr, FLD_BYTE, 1      'SpellCastPer5       As Byte
    AddField Map, ctr, FLD_BYTE, 1      'SpellCastLvl1       As Byte
    AddField Map, ctr, FLD_BYTE, 1      'SpellCastLvl2       As Byte
    AddField Map, ctr, FLD_BYTE, 1      'SpellCastLvl3       As Byte
    AddField Map, ctr, FLD_BYTE, 1      'SpellCastLvl4       As Byte
    AddField Map, ctr, FLD_BYTE, 1      'SpellCastLvl5       As Byte
    AddField Map, ctr, FLD_STRING, 70   'Desc Line1
    AddField Map, ctr, FLD_INTEGER, 1
    AddField Map, ctr, FLD_STRING, 70   'Desc Line2
    AddField Map, ctr, FLD_INTEGER, 1
    AddField Map, ctr, FLD_STRING, 70   'Desc Line3
    AddField Map, ctr, FLD_INTEGER, 1
    AddField Map, ctr, FLD_STRING, 70   'Desc Line4
    AddField Map, ctr, FLD_INTEGER, 1
    AddField Map, ctr, FLD_BYTE, 1      'Gender
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_INTEGER, 2
End Sub

Sub AddItemFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    AddField Map, ctr, FLD_INTEGER, 4   'Item Number
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_STRING, 156
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 29   'Item Name
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 60   'desc1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 60   'desc2
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 60   'desc3
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 60   'desc4
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 60   'desc5
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 60   'desc6
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 60   'desc7
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 60   'desc8
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 60   'desc9
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_INTEGER, 2   'unknown6
    AddField Map, ctr, FLD_INTEGER, 2   'weight
    AddField Map, ctr, FLD_INTEGER, 2   'Type
    AddField Map, ctr, FLD_INTEGER, 2       'AAA1    -   AbilityA1 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA2    -   AbilityA2 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA3    -   AbilityA3 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA4    -   AbilityA4 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA5    -   AbilityA5 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA6    -   AbilityA6 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA7    -   AbilityA7 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA8    -   AbilityA8 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA9    -   AbilityA9 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AA10    -   AbilityA10 as Integer        2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA1    -   AbilityA1 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA2    -   AbilityA2 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA3    -   AbilityA3 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA4    -   AbilityA4 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA5    -   AbilityA5 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA6    -   AbilityA6 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA7    -   AbilityA7 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA8    -   AbilityA8 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AAA9    -   AbilityA9 as Integer         2
    AddField Map, ctr, FLD_INTEGER, 2       'AA10    -   AbilityA10 as Integer        2
    AddField Map, ctr, FLD_INTEGER, 2       'Uses
    AddField Map, ctr, FLD_INTEGER, 2       'nothing
    AddField Map, ctr, FLD_INTEGER, 2       'Cost
    AddField Map, ctr, FLD_INTEGER, 2       'class1
    AddField Map, ctr, FLD_INTEGER, 2       'class2
    AddField Map, ctr, FLD_INTEGER, 2       'class3
    AddField Map, ctr, FLD_INTEGER, 2       'class4
    AddField Map, ctr, FLD_INTEGER, 2       'class5
    AddField Map, ctr, FLD_INTEGER, 2       'class6
    AddField Map, ctr, FLD_INTEGER, 2       'class7
    AddField Map, ctr, FLD_INTEGER, 2       'class8
    AddField Map, ctr, FLD_INTEGER, 2       'class9
    AddField Map, ctr, FLD_INTEGER, 2       'class10
    AddField Map, ctr, FLD_INTEGER, 2       'unknown8
    AddField Map, ctr, FLD_INTEGER, 2       'nothing2
    AddField Map, ctr, FLD_INTEGER, 2       'nothing3
    AddField Map, ctr, FLD_INTEGER, 2       'minhit
    AddField Map, ctr, FLD_INTEGER, 2       'maxhit
    AddField Map, ctr, FLD_INTEGER, 2       'AC
    AddField Map, ctr, FLD_INTEGER, 4       'race1
    AddField Map, ctr, FLD_INTEGER, 4       'race2
    AddField Map, ctr, FLD_INTEGER, 4       'race3
    AddField Map, ctr, FLD_INTEGER, 4       'race4
    AddField Map, ctr, FLD_INTEGER, 4       'race5
    AddField Map, ctr, FLD_INTEGER, 4       'race6
    AddField Map, ctr, FLD_INTEGER, 4       'race7
    AddField Map, ctr, FLD_INTEGER, 4       'race8
    AddField Map, ctr, FLD_INTEGER, 4       'race9
    AddField Map, ctr, FLD_INTEGER, 4       'race10
    AddField Map, ctr, FLD_INTEGER, 2       'negate1
    AddField Map, ctr, FLD_INTEGER, 2       'negate2
    AddField Map, ctr, FLD_INTEGER, 2       'negate3
    AddField Map, ctr, FLD_INTEGER, 2       'negate4
    AddField Map, ctr, FLD_INTEGER, 2       'negate5
    AddField Map, ctr, FLD_INTEGER, 2       'negate6
    AddField Map, ctr, FLD_INTEGER, 2       'negate7
    AddField Map, ctr, FLD_INTEGER, 2       'negate8
    AddField Map, ctr, FLD_INTEGER, 2       'negate9
    AddField Map, ctr, FLD_INTEGER, 2       'negate10
    AddField Map, ctr, FLD_INTEGER, 2       'negate11
    AddField Map, ctr, FLD_INTEGER, 2       'negate12
    AddField Map, ctr, FLD_INTEGER, 2       'negate13
    AddField Map, ctr, FLD_INTEGER, 2       'negate14
    AddField Map, ctr, FLD_INTEGER, 2       'negate15
    AddField Map, ctr, FLD_INTEGER, 2       'negate16
    AddField Map, ctr, FLD_INTEGER, 2       'negate17
    AddField Map, ctr, FLD_INTEGER, 2       'negate18
    AddField Map, ctr, FLD_INTEGER, 2       'negate19
    AddField Map, ctr, FLD_INTEGER, 2       'negate20
    AddField Map, ctr, FLD_INTEGER, 2       'Weapon
    AddField Map, ctr, FLD_INTEGER, 2       'Armour
    AddField Map, ctr, FLD_INTEGER, 2       'Wornon
    AddField Map, ctr, FLD_INTEGER, 2       'Accuracy
    AddField Map, ctr, FLD_INTEGER, 2       'DR
    AddField Map, ctr, FLD_BYTE, 1      'Gettable
    AddField Map, ctr, FLD_BYTE, 1      'unknown12
    AddField Map, ctr, FLD_INTEGER, 2   'ReqStr
    AddField Map, ctr, FLD_INTEGER, 2       'unknown13a 1
    AddField Map, ctr, FLD_INTEGER, 2       '2
    AddField Map, ctr, FLD_INTEGER, 2       '3
    AddField Map, ctr, FLD_INTEGER, 2       '4
    AddField Map, ctr, FLD_INTEGER, 2       '5
    AddField Map, ctr, FLD_INTEGER, 2       '6
    AddField Map, ctr, FLD_INTEGER, 2       '7
    AddField Map, ctr, FLD_INTEGER, 4       'runic
    AddField Map, ctr, FLD_INTEGER, 4       'plat
    AddField Map, ctr, FLD_INTEGER, 4       'gold
    AddField Map, ctr, FLD_INTEGER, 4       'silver
    AddField Map, ctr, FLD_INTEGER, 4       'copper
    AddField Map, ctr, FLD_INTEGER, 2       'unknown13b 1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2       '4
    AddField Map, ctr, FLD_INTEGER, 2       '5
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2       '8
    AddField Map, ctr, FLD_INTEGER, 2       '9
    AddField Map, ctr, FLD_INTEGER, 2       '10
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2       '13
'    AddField map, ctr, FLD_INTEGER, 2                  -- took out when added cash drops
'    AddField map, ctr, FLD_INTEGER, 2
'    AddField map, ctr, FLD_INTEGER, 2       '16
'    AddField map, ctr, FLD_INTEGER, 2       '17
'    AddField map, ctr, FLD_INTEGER, 2       '18
    AddField Map, ctr, FLD_INTEGER, 2   'speed
    AddField Map, ctr, FLD_INTEGER, 2   'unknown14
    AddField Map, ctr, FLD_INTEGER, 2       'AbilB 1
    AddField Map, ctr, FLD_INTEGER, 2       '2
    AddField Map, ctr, FLD_INTEGER, 2       '3
    AddField Map, ctr, FLD_INTEGER, 2       '4
    AddField Map, ctr, FLD_INTEGER, 2       '5
    AddField Map, ctr, FLD_INTEGER, 2       '6
    AddField Map, ctr, FLD_INTEGER, 2       '7
    AddField Map, ctr, FLD_INTEGER, 2       '8
    AddField Map, ctr, FLD_INTEGER, 2       '9
    AddField Map, ctr, FLD_INTEGER, 2       '10
    AddField Map, ctr, FLD_INTEGER, 2       '11
    AddField Map, ctr, FLD_INTEGER, 2       '12
    AddField Map, ctr, FLD_INTEGER, 2       '13
    AddField Map, ctr, FLD_INTEGER, 2       '14
    AddField Map, ctr, FLD_INTEGER, 2       '15
    AddField Map, ctr, FLD_INTEGER, 2       '16
    AddField Map, ctr, FLD_INTEGER, 2       '17
    AddField Map, ctr, FLD_INTEGER, 2       '18
    AddField Map, ctr, FLD_INTEGER, 2       '19
    AddField Map, ctr, FLD_INTEGER, 2       '20
    AddField Map, ctr, FLD_INTEGER, 2   'unknown15
    AddField Map, ctr, FLD_INTEGER, 4   'Hit
    AddField Map, ctr, FLD_INTEGER, 4   'Miss
    AddField Map, ctr, FLD_INTEGER, 4   'read
    AddField Map, ctr, FLD_INTEGER, 4   'distruct
    AddField Map, ctr, FLD_INTEGER, 2   'unknown16 1
    AddField Map, ctr, FLD_INTEGER, 2   '2
    AddField Map, ctr, FLD_INTEGER, 2   '3
    AddField Map, ctr, FLD_INTEGER, 2   '4
    AddField Map, ctr, FLD_INTEGER, 2   '5
    AddField Map, ctr, FLD_INTEGER, 2   '6
    AddField Map, ctr, FLD_BYTE, 1      'NotDroppable
    AddField Map, ctr, FLD_BYTE, 1      'Costype
    AddField Map, ctr, FLD_BYTE, 1      'retain after uses
    AddField Map, ctr, FLD_BYTE, 1      'u18
    AddField Map, ctr, FLD_BYTE, 1      'destroy
    AddField Map, ctr, FLD_BYTE, 1      'u19
'    AddField map, ctr, FLD_BYTE, 1      '
'    AddField map, ctr, FLD_BYTE, 1      '
'    AddField map, ctr, FLD_BYTE, 1      '
'    AddField map, ctr, FLD_BYTE, 1      '
'    AddField map, ctr, FLD_BYTE, 1      '
'    AddField map, ctr, FLD_BYTE, 1      '
'    AddField map, ctr, FLD_BYTE, 1      '
'    AddField map, ctr, FLD_BYTE, 1      '
'    AddField map, ctr, FLD_BYTE, 1      '
'    AddField map, ctr, FLD_BYTE, 1      '
'    AddField map, ctr, FLD_BYTE, 1      '
'    AddField map, ctr, FLD_BYTE, 1      '
'    AddField map, ctr, FLD_BYTE, 1      '
End Sub

Sub AddShopFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    AddField Map, ctr, FLD_INTEGER, 4   'SNumber
    AddField Map, ctr, FLD_STRING, 39   'SName
    AddField Map, ctr, FLD_INTEGER, 2   '
    AddField Map, ctr, FLD_STRING, 52   'SDescA
    AddField Map, ctr, FLD_INTEGER, 1   '
    AddField Map, ctr, FLD_STRING, 52   'SDescB
    AddField Map, ctr, FLD_INTEGER, 1   '
    AddField Map, ctr, FLD_STRING, 52   'SDescC
    AddField Map, ctr, FLD_INTEGER, 1   '
    AddField Map, ctr, FLD_INTEGER, 2   'SType
    AddField Map, ctr, FLD_INTEGER, 2   'SMinLvl
    AddField Map, ctr, FLD_INTEGER, 2   'SMaxLvl
    AddField Map, ctr, FLD_INTEGER, 2   'SMarkUp
    AddField Map, ctr, FLD_INTEGER, 2   '
    AddField Map, ctr, FLD_INTEGER, 1   'ClassLimit
    AddField Map, ctr, FLD_INTEGER, 1   '
    AddField Map, ctr, FLD_INTEGER, 4   'Item 1
    AddField Map, ctr, FLD_INTEGER, 4   'Item 2
    AddField Map, ctr, FLD_INTEGER, 4   'Item 3
    AddField Map, ctr, FLD_INTEGER, 4   'Item 4
    AddField Map, ctr, FLD_INTEGER, 4   'Item 5
    AddField Map, ctr, FLD_INTEGER, 4   'Item 6
    AddField Map, ctr, FLD_INTEGER, 4   'Item 7
    AddField Map, ctr, FLD_INTEGER, 4   'Item 8
    AddField Map, ctr, FLD_INTEGER, 4   'Item 9
    AddField Map, ctr, FLD_INTEGER, 4   'Item 10
    AddField Map, ctr, FLD_INTEGER, 4   'Item 11
    AddField Map, ctr, FLD_INTEGER, 4   'Item 12
    AddField Map, ctr, FLD_INTEGER, 4   'Item 13
    AddField Map, ctr, FLD_INTEGER, 4   'Item 14
    AddField Map, ctr, FLD_INTEGER, 4   'Item 15
    AddField Map, ctr, FLD_INTEGER, 4   'Item 16
    AddField Map, ctr, FLD_INTEGER, 4   'Item 17
    AddField Map, ctr, FLD_INTEGER, 4   'Item 18
    AddField Map, ctr, FLD_INTEGER, 4   'Item 19
    AddField Map, ctr, FLD_INTEGER, 4   'Item 20
    AddField Map, ctr, FLD_INTEGER, 2   'Max 1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '5
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '10
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '15
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '20
    AddField Map, ctr, FLD_INTEGER, 2   'Normal1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '5
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '10
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '15
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '20
    AddField Map, ctr, FLD_INTEGER, 2   'RgnTimer1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '5
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '10
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '15
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '20
    AddField Map, ctr, FLD_INTEGER, 2   'Number1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '5
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '10
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '15
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2   '20
    AddField Map, ctr, FLD_BYTE, 1      'RgnPercentage1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1      '5
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1      '10
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1      '15
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1      '20
End Sub

Sub AddRoomFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    AddField Map, ctr, FLD_INTEGER, 4   'Map Number
    AddField Map, ctr, FLD_INTEGER, 4   'Room Number
    AddField Map, ctr, FLD_STRING, 253   'emptyspace
    AddField Map, ctr, FLD_STRING, 53   'Name
    AddField Map, ctr, FLD_STRING, 71   'Desc1
    AddField Map, ctr, FLD_STRING, 71   'Desc2
    AddField Map, ctr, FLD_STRING, 71   'Desc3
    AddField Map, ctr, FLD_STRING, 71   'Desc4
    AddField Map, ctr, FLD_STRING, 71   'Desc5
    AddField Map, ctr, FLD_STRING, 71   'Desc6
    AddField Map, ctr, FLD_STRING, 71   'Desc7
    AddField Map, ctr, FLD_STRING, 13   'AnsiMap
    AddField Map, ctr, FLD_INTEGER, 4   'RoomExit1
    AddField Map, ctr, FLD_INTEGER, 4   '2
    AddField Map, ctr, FLD_INTEGER, 4   '3
    AddField Map, ctr, FLD_INTEGER, 4   '4
    AddField Map, ctr, FLD_INTEGER, 4   '5
    AddField Map, ctr, FLD_INTEGER, 4   '6
    AddField Map, ctr, FLD_INTEGER, 4   '7
    AddField Map, ctr, FLD_INTEGER, 4   '8
    AddField Map, ctr, FLD_INTEGER, 4   '9
    AddField Map, ctr, FLD_INTEGER, 4   'RoomExit10
    AddField Map, ctr, FLD_INTEGER, 2   'Room Type 1
    AddField Map, ctr, FLD_INTEGER, 2   '2
    AddField Map, ctr, FLD_INTEGER, 2   '3
    AddField Map, ctr, FLD_INTEGER, 2   '4
    AddField Map, ctr, FLD_INTEGER, 2   '5
    AddField Map, ctr, FLD_INTEGER, 2   '6
    AddField Map, ctr, FLD_INTEGER, 2   '7
    AddField Map, ctr, FLD_INTEGER, 2   '8
    AddField Map, ctr, FLD_INTEGER, 2   '9
    AddField Map, ctr, FLD_INTEGER, 2   'Room Type 10
    AddField Map, ctr, FLD_INTEGER, 4   'Room Para1 1
    AddField Map, ctr, FLD_INTEGER, 4   '2
    AddField Map, ctr, FLD_INTEGER, 4   '3
    AddField Map, ctr, FLD_INTEGER, 4   '4
    AddField Map, ctr, FLD_INTEGER, 4   '5
    AddField Map, ctr, FLD_INTEGER, 4   '6
    AddField Map, ctr, FLD_INTEGER, 4   '7
    AddField Map, ctr, FLD_INTEGER, 4   '8
    AddField Map, ctr, FLD_INTEGER, 4   '9
    AddField Map, ctr, FLD_INTEGER, 4   'Room Para1 10
    AddField Map, ctr, FLD_INTEGER, 2   'Room Para2 1
    AddField Map, ctr, FLD_INTEGER, 2   '2
    AddField Map, ctr, FLD_INTEGER, 2   '3
    AddField Map, ctr, FLD_INTEGER, 2   '4
    AddField Map, ctr, FLD_INTEGER, 2   '5
    AddField Map, ctr, FLD_INTEGER, 2   '6
    AddField Map, ctr, FLD_INTEGER, 2   '7
    AddField Map, ctr, FLD_INTEGER, 2   '8
    AddField Map, ctr, FLD_INTEGER, 2   '9
    AddField Map, ctr, FLD_INTEGER, 2   'Room Para2 10
    AddField Map, ctr, FLD_INTEGER, 4   'Room Para3 1
    AddField Map, ctr, FLD_INTEGER, 4   '2
    AddField Map, ctr, FLD_INTEGER, 4   '3
    AddField Map, ctr, FLD_INTEGER, 4   '4
    AddField Map, ctr, FLD_INTEGER, 4   '5
    AddField Map, ctr, FLD_INTEGER, 4   '6
    AddField Map, ctr, FLD_INTEGER, 4   '7
    AddField Map, ctr, FLD_INTEGER, 4   '8
    AddField Map, ctr, FLD_INTEGER, 4   '9
    AddField Map, ctr, FLD_INTEGER, 4   'Room Para3 10
    AddField Map, ctr, FLD_INTEGER, 4   'Room Para4 1
    AddField Map, ctr, FLD_INTEGER, 4   '2
    AddField Map, ctr, FLD_INTEGER, 4   '3
    AddField Map, ctr, FLD_INTEGER, 4   '4
    AddField Map, ctr, FLD_INTEGER, 4   '5
    AddField Map, ctr, FLD_INTEGER, 4   '6
    AddField Map, ctr, FLD_INTEGER, 4   '7
    AddField Map, ctr, FLD_INTEGER, 4   '8
    AddField Map, ctr, FLD_INTEGER, 4   '9
    AddField Map, ctr, FLD_INTEGER, 4   'Room Para4 10
    
'    AddField map, ctr, FLD_INTEGER, 4   'current mons 1        '***I_THROUGH_N*** (comment for n)
'    AddField map, ctr, FLD_INTEGER, 4   '2                     '         ||
'    AddField map, ctr, FLD_INTEGER, 4   '3                     '         ||
'    AddField map, ctr, FLD_INTEGER, 4   '4                     '         ||
'    AddField map, ctr, FLD_INTEGER, 4   '5                     '         ||
'    AddField map, ctr, FLD_INTEGER, 4   '6                     '         ||
'    AddField map, ctr, FLD_INTEGER, 4   '7                     '         ||
'    AddField map, ctr, FLD_INTEGER, 4   '8                     '         ||
'    AddField map, ctr, FLD_INTEGER, 4   '9                     '         ||
'    AddField map, ctr, FLD_INTEGER, 4   '10                    '         ||
'    AddField map, ctr, FLD_INTEGER, 4   '11                    '         ||
'    AddField map, ctr, FLD_INTEGER, 4   '12                    '         ||
'    AddField map, ctr, FLD_INTEGER, 4   '13                    '         ||
'    AddField map, ctr, FLD_INTEGER, 4   '14                    '        \\//
'    AddField map, ctr, FLD_INTEGER, 4   'current mons 15       '***I_THROUGH_N*** (comment for n)

    AddField Map, ctr, FLD_INTEGER, 2   'current mons 1         '***I_THROUGH_N*** (UNcomment for n)
    AddField Map, ctr, FLD_INTEGER, 2   '2                      '         ||
    AddField Map, ctr, FLD_INTEGER, 2   '3                      '         ||
    AddField Map, ctr, FLD_INTEGER, 2   '4                      '         ||
    AddField Map, ctr, FLD_INTEGER, 2   '5                      '         ||
    AddField Map, ctr, FLD_INTEGER, 2   '6                      '         ||
    AddField Map, ctr, FLD_INTEGER, 2   '7                      '         ||
    AddField Map, ctr, FLD_INTEGER, 2   '8                      '         ||
    AddField Map, ctr, FLD_INTEGER, 2   '9                      '         ||
    AddField Map, ctr, FLD_INTEGER, 2   '10                     '         ||
    AddField Map, ctr, FLD_INTEGER, 2   '11                     '         ||
    AddField Map, ctr, FLD_INTEGER, 2   '12                     '         ||
    AddField Map, ctr, FLD_INTEGER, 2   '13                     '         ||
    AddField Map, ctr, FLD_INTEGER, 2   '14                     '        \\//
    AddField Map, ctr, FLD_INTEGER, 2   'current mons 15        '***I_THROUGH_N*** (UNcomment for n)
    
    AddField Map, ctr, FLD_INTEGER, 2   'Type
    'AddField map, ctr, FLD_INTEGER, 2   'new spot           '***I_THROUGH_N*** (comment for n)
    AddField Map, ctr, FLD_INTEGER, 4   'ShopNum
    AddField Map, ctr, FLD_INTEGER, 2   'nothingc1
    AddField Map, ctr, FLD_INTEGER, 2   'nothingc2
    AddField Map, ctr, FLD_INTEGER, 2   'nothingc3
    AddField Map, ctr, FLD_INTEGER, 2   'nothingc4
    AddField Map, ctr, FLD_INTEGER, 2   'nothingc5
    AddField Map, ctr, FLD_INTEGER, 2   'nothingc6
    AddField Map, ctr, FLD_INTEGER, 2   'nothingc7
    AddField Map, ctr, FLD_INTEGER, 2   'nothingc8
    AddField Map, ctr, FLD_INTEGER, 2   'nothingc9
    AddField Map, ctr, FLD_INTEGER, 2   'nothingc10
    AddField Map, ctr, FLD_INTEGER, 2   'nothingc11
    AddField Map, ctr, FLD_INTEGER, 2   'nothingc12
    AddField Map, ctr, FLD_INTEGER, 2   'nothingc13
    AddField Map, ctr, FLD_INTEGER, 2   'nothingc14
    AddField Map, ctr, FLD_INTEGER, 2   'nothingc15
    AddField Map, ctr, FLD_INTEGER, 2   'MinIndex
    AddField Map, ctr, FLD_INTEGER, 2   'MaxIndex
    AddField Map, ctr, FLD_INTEGER, 4   'ByNumber
    AddField Map, ctr, FLD_INTEGER, 2   '????
    AddField Map, ctr, FLD_INTEGER, 2   'Light
    AddField Map, ctr, FLD_INTEGER, 2   'GangHouseNumber
    AddField Map, ctr, FLD_INTEGER, 4   'Roomitem1
    AddField Map, ctr, FLD_INTEGER, 4   '2
    AddField Map, ctr, FLD_INTEGER, 4   '3
    AddField Map, ctr, FLD_INTEGER, 4   '4
    AddField Map, ctr, FLD_INTEGER, 4   '5
    AddField Map, ctr, FLD_INTEGER, 4   '6
    AddField Map, ctr, FLD_INTEGER, 4   '7
    AddField Map, ctr, FLD_INTEGER, 4   '8
    AddField Map, ctr, FLD_INTEGER, 4   '9
    AddField Map, ctr, FLD_INTEGER, 4   '10
    AddField Map, ctr, FLD_INTEGER, 4   '11
    AddField Map, ctr, FLD_INTEGER, 4   '12
    AddField Map, ctr, FLD_INTEGER, 4   '13
    AddField Map, ctr, FLD_INTEGER, 4   '14
    AddField Map, ctr, FLD_INTEGER, 4   '15
    AddField Map, ctr, FLD_INTEGER, 4   '16
    AddField Map, ctr, FLD_INTEGER, 4   'RoomItem17
    AddField Map, ctr, FLD_INTEGER, 2   'RoomItemUses1
    AddField Map, ctr, FLD_INTEGER, 2   '2
    AddField Map, ctr, FLD_INTEGER, 2   '3
    AddField Map, ctr, FLD_INTEGER, 2   '4
    AddField Map, ctr, FLD_INTEGER, 2   'RoomItemUses5
    AddField Map, ctr, FLD_INTEGER, 2   '6
    AddField Map, ctr, FLD_INTEGER, 2   '7
    AddField Map, ctr, FLD_INTEGER, 2   '8
    AddField Map, ctr, FLD_INTEGER, 2   '9
    AddField Map, ctr, FLD_INTEGER, 2   'RoomItemUses10
    AddField Map, ctr, FLD_INTEGER, 2   '11
    AddField Map, ctr, FLD_INTEGER, 2   '12
    AddField Map, ctr, FLD_INTEGER, 2   '13
    AddField Map, ctr, FLD_INTEGER, 2   '14
    AddField Map, ctr, FLD_INTEGER, 2   'RoomItemUses15
    AddField Map, ctr, FLD_INTEGER, 2   '16
    AddField Map, ctr, FLD_INTEGER, 2   '17
    AddField Map, ctr, FLD_INTEGER, 2   'nothing4
    AddField Map, ctr, FLD_INTEGER, 4   'InvisItem1
    AddField Map, ctr, FLD_INTEGER, 4   '2
    AddField Map, ctr, FLD_INTEGER, 4   '3
    AddField Map, ctr, FLD_INTEGER, 4   '4
    AddField Map, ctr, FLD_INTEGER, 4   '5
    AddField Map, ctr, FLD_INTEGER, 4   '6
    AddField Map, ctr, FLD_INTEGER, 4   '7
    AddField Map, ctr, FLD_INTEGER, 4   '8
    AddField Map, ctr, FLD_INTEGER, 4   '9
    AddField Map, ctr, FLD_INTEGER, 4   '10
    AddField Map, ctr, FLD_INTEGER, 4   '11
    AddField Map, ctr, FLD_INTEGER, 4   '12
    AddField Map, ctr, FLD_INTEGER, 4   '13
    AddField Map, ctr, FLD_INTEGER, 4   '14
    AddField Map, ctr, FLD_INTEGER, 4   'InvisItem15
    AddField Map, ctr, FLD_INTEGER, 2   'InvisItemUses1
    AddField Map, ctr, FLD_INTEGER, 2   '2
    AddField Map, ctr, FLD_INTEGER, 2   '3
    AddField Map, ctr, FLD_INTEGER, 2   '4
    AddField Map, ctr, FLD_INTEGER, 2   'InvisItemUses5
    AddField Map, ctr, FLD_INTEGER, 2   '6
    AddField Map, ctr, FLD_INTEGER, 2   '7
    AddField Map, ctr, FLD_INTEGER, 2   '8
    AddField Map, ctr, FLD_INTEGER, 2   '9
    AddField Map, ctr, FLD_INTEGER, 2   'InvisItemUses10
    AddField Map, ctr, FLD_INTEGER, 2   '11
    AddField Map, ctr, FLD_INTEGER, 2   '12
    AddField Map, ctr, FLD_INTEGER, 2   '13
    AddField Map, ctr, FLD_INTEGER, 2   '14
    AddField Map, ctr, FLD_INTEGER, 2   'InvisItemUses15
    AddField Map, ctr, FLD_INTEGER, 2   'nothing5
    AddField Map, ctr, FLD_INTEGER, 4   'Runic
    AddField Map, ctr, FLD_INTEGER, 4   'Plat
    AddField Map, ctr, FLD_INTEGER, 4   'Gold
    AddField Map, ctr, FLD_INTEGER, 4   'Silver
    AddField Map, ctr, FLD_INTEGER, 4   'Copper
    AddField Map, ctr, FLD_INTEGER, 4   'InvisRunic
    AddField Map, ctr, FLD_INTEGER, 4   'InvisPlat
    AddField Map, ctr, FLD_INTEGER, 4   'InvisGold
    AddField Map, ctr, FLD_INTEGER, 4   'InvisSilver
    AddField Map, ctr, FLD_INTEGER, 4   'InvisCopper
    AddField Map, ctr, FLD_INTEGER, 4   'MaxRegen
    AddField Map, ctr, FLD_INTEGER, 2   'MonsterType
    AddField Map, ctr, FLD_INTEGER, 2   'unknown69
    AddField Map, ctr, FLD_INTEGER, 4   'Attributes
    AddField Map, ctr, FLD_INTEGER, 4   'unknown1
    AddField Map, ctr, FLD_INTEGER, 4   'DeathRoom
    AddField Map, ctr, FLD_INTEGER, 4   'ExitRoom
    AddField Map, ctr, FLD_INTEGER, 2   'RoomItemQty1
    AddField Map, ctr, FLD_INTEGER, 2   '2
    AddField Map, ctr, FLD_INTEGER, 2   '3
    AddField Map, ctr, FLD_INTEGER, 2   '4
    AddField Map, ctr, FLD_INTEGER, 2   '5
    AddField Map, ctr, FLD_INTEGER, 2   '6
    AddField Map, ctr, FLD_INTEGER, 2   '7
    AddField Map, ctr, FLD_INTEGER, 2   '8
    AddField Map, ctr, FLD_INTEGER, 2   '9
    AddField Map, ctr, FLD_INTEGER, 2   '10
    AddField Map, ctr, FLD_INTEGER, 2   '11
    AddField Map, ctr, FLD_INTEGER, 2   '12
    AddField Map, ctr, FLD_INTEGER, 2   '13
    AddField Map, ctr, FLD_INTEGER, 2   '14
    AddField Map, ctr, FLD_INTEGER, 2   '15
    AddField Map, ctr, FLD_INTEGER, 2   '16
    AddField Map, ctr, FLD_INTEGER, 2   'RoomItemQty17
    AddField Map, ctr, FLD_INTEGER, 2   'InvisItemQty1
    AddField Map, ctr, FLD_INTEGER, 2   '2
    AddField Map, ctr, FLD_INTEGER, 2   '3
    AddField Map, ctr, FLD_INTEGER, 2   '4
    AddField Map, ctr, FLD_INTEGER, 2   '5
    AddField Map, ctr, FLD_INTEGER, 2   '6
    AddField Map, ctr, FLD_INTEGER, 2   '7
    AddField Map, ctr, FLD_INTEGER, 2   '8
    AddField Map, ctr, FLD_INTEGER, 2   '9
    AddField Map, ctr, FLD_INTEGER, 2   '10
    AddField Map, ctr, FLD_INTEGER, 2   '11
    AddField Map, ctr, FLD_INTEGER, 2   '12
    AddField Map, ctr, FLD_INTEGER, 2   '13
    AddField Map, ctr, FLD_INTEGER, 2   '14
    AddField Map, ctr, FLD_INTEGER, 2   'InvisItemQty15
    AddField Map, ctr, FLD_INTEGER, 4   'Command Text
    AddField Map, ctr, FLD_INTEGER, 4   'nothingR1
    AddField Map, ctr, FLD_INTEGER, 2   'Delay
    AddField Map, ctr, FLD_INTEGER, 2   'MaxArea
    AddField Map, ctr, FLD_INTEGER, 4   'NothingS
    AddField Map, ctr, FLD_INTEGER, 4   'ControlRoom
    AddField Map, ctr, FLD_INTEGER, 4   'PermNPC
    AddField Map, ctr, FLD_INTEGER, 4   'PlacedItems1
    AddField Map, ctr, FLD_INTEGER, 4   '2
    AddField Map, ctr, FLD_INTEGER, 4   '3
    AddField Map, ctr, FLD_INTEGER, 4   '4
    AddField Map, ctr, FLD_INTEGER, 4   '5
    AddField Map, ctr, FLD_INTEGER, 4   '6
    AddField Map, ctr, FLD_INTEGER, 4   '7
    AddField Map, ctr, FLD_INTEGER, 4   '8
    AddField Map, ctr, FLD_INTEGER, 4   '9
    AddField Map, ctr, FLD_INTEGER, 4   'PlacedItems10
    AddField Map, ctr, FLD_INTEGER, 4   'NothingT1
    AddField Map, ctr, FLD_INTEGER, 4   '2
    AddField Map, ctr, FLD_INTEGER, 4   'something1
    AddField Map, ctr, FLD_INTEGER, 4   'Spell
    AddField Map, ctr, FLD_INTEGER, 2   'unknown70
    AddField Map, ctr, FLD_BYTE, 1   'num of mons
    AddField Map, ctr, FLD_BYTE, 1   'unknown71
End Sub

Sub AddMessageFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    AddField Map, ctr, FLD_INTEGER, 4   'MNumber
    AddField Map, ctr, FLD_STRING, 74   'MLine1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_STRING, 74   'MLine2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_STRING, 74   'Mline3
End Sub
Sub AddActionFieldMap(Map() As FieldMap, ByRef ctr As Integer)
    AddField Map, ctr, FLD_STRING, 29   'name
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 74   '1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_STRING, 74   '2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_STRING, 74   '3
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_STRING, 74   '4
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_STRING, 74   '5
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_STRING, 74   '6
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_STRING, 74   '7
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_STRING, 74   '8
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_STRING, 74   '9
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_STRING, 74   '10
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_STRING, 74   '11
End Sub

Sub AddField(Map() As FieldMap, ByRef ctr As Integer, dataType As Long, length As Long)
  SetField Map(ctr), dataType, length
  ctr = ctr + 1
End Sub

