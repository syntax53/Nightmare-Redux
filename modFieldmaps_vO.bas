Attribute VB_Name = "modFieldmaps_vO"
Option Base 0
Option Explicit

'***I_THROUGH_N*** (search on this comment for all the vN changes)
'set this to true if this version works with v1.11h to v1.11n (set to true for n, false for o+)
'Public Const WorksWithN = True
Public Const WorksWithN = False

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
Public Type RaceRecType '                                         A     B       C       A=field len, B=Cumulative field len, C=field count
    Number As Integer       '####    -   Number as Integer        2     2       1
    Name As String * 29     'NNNN    -   Name as String * 29      29    31      2
    nothing1 As Byte        '  00    -   Nothing as Byte          1     32      3
    MinInt As Integer       '-INT    -   Min Int as Integer       2     34      4
    MinWil As Integer       '-Wil    -   Min Wil as Integer       2     36      5
    MinStr As Integer       '-Str    -   Min Str as Integer       2     38      6
    MinHea As Integer       '-Hea    -   Min Hea as Integer       2     40      7
    MinAgl As Integer       '-Agl    -   Min Agl as Integer       2     42      8
    MinChm As Integer       '-Chm    -   Min Chm as Integer       2     44      9
    HPBonus As Integer      'HP00    -   HP Bonus as Integer      2     46      10
    nothing2 As Long        '0000    -   Nothing as Long          4     50      11
    AbilityA(9) As Integer  'AAA1    -   AbilityA1 as Integer     20    70      21
    CP As Integer           'CP00    -   Starting CP as Integer   2     72      22
    AbilityB(9) As Integer  'BBB1    -   AbilityB1 as Integer     20    92      32
    nothing3 As Long        '0020    -   Nothing as Long          4     96      33
    nothing4 As Integer     '0000    -   Nothing as Integer       2     98      34
    ExpChart As Integer     'EXPE    -   Exp Chart as Integer     2     100     35
    nothing5 As Integer     '0000    -   Nothing as Integer       2     102     36
    MaxInt As Integer       '+Int    -   Max Int as Integer       2     104     37
    MaxWil As Integer       '+Wil    -   Max Wil as Integer       2     106     38
    MaxStr As Integer       '+Str    -   Max Str as Integer       2     108     39
    MaxHea As Integer       '+Hea    -   Max Hea as Integer       2     110     40
    MaxAgl As Integer       '+Agl    -   Max Agl as Integer       2     112     41
    MaxChm As Integer       '+Chm    -   Max Chm as Integer       2     114     42
    Nothing6 As Long        '0000    -   Nothing as Long          4     118     43
    nothing7 As Long        '0000    -   Nothing as Long          4     122     44
    nothing8 As Long        '0000    -   Nothing as Long          4     126     45
End Type
Const RaceDataBufSize = 126
Public RaceFldMap(0 To 44) As PALN32.FieldMap
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
Public Type ClassRecType        'A      B       C       A=field len, B=Cumulative field len, C=field count
    Number     As Integer       '2      2       1
    Name       As String * 29   '29     31      2
    AfterName  As Byte          '1      32      3
    MinHp      As Integer       '2      34      4
    MaxHP      As Integer       '2      36      5
    Exp        As Integer       '2      38      6
    nothing1   As Integer       '2      40      7
    nothing2   As Integer       '2      42      8
    nothing3   As Integer       '2      44      9
    AbilityA(9)   As Integer    '20     64      19
    MagicType  As Integer       '2      66      20
    MagicLvL   As Integer       '2      68      21
    Weapon     As Integer       '2      70      22
    Armour     As Integer       '2      72      23
    Combat     As Integer       '2      74      24
    AbilityB(9)   As Integer    '20     94      34
    nothing4   As Integer       '2      96      35
    nothing5   As Integer       '2      98      36
    Nothing6   As Integer       '2      100     37
    TitleText  As Long          '4      104     38
End Type
Const ClassDataBufSize = 156
Public ClassFldMap(0 To 37) As PALN32.FieldMap
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
Public Type SpellRecType        'A      B       C       A=field len, B=Cumulative field len, C=field count
    Number As Integer           '2      2       1
    Name As String * 29         '29     31      2
    AfterName  As Byte          '1      32      3
    DescA As String * 50        '50     82      4
    AfterDescA As Byte          '1      83      5
    DescB As String * 50        '50     133     6
    AfterDescB As Byte          '1      134     7
    N01 As Integer              '2      136     8
    CastMsgA As Long            '4      140     9
    N02(10) As Integer          '22     162     20
    LevelCap As Byte            '1      163     21
    N03 As Byte                 '1      164     22
    MsgStyle As Byte            '1      165     23
    N04(2) As Byte              '3      168     26
    AbilityB(9) As Integer      '20     188     36
    Energy As Integer           '2      190     37
    Level As Integer            '2      192     38
    Min As Integer              '2      194     39
    Max As Integer              '2      196     40
    SpellType As Integer        '2      198     41
    TypeOfResists As Integer    '2      200     42
    Difficulty As Integer       '2      202     43
    UNDEFINED01 As Integer      '2      204     44
    Target As Integer           '2      206     45
    duration As Integer         '2      208     46
    TypeOfAttack As Integer     '2      210     47
    UNDEFINED02 As Integer      '2      212     48
    ResistAbility As Integer    '2      214     49
    MageryA As Integer          '2      216     50
    AbilityA(9) As Integer      '20     236     60
    CastMsgB As Long            '4      240     61
    Mana As Integer             '2      242     62
    MaxIncrease As Byte         '1      243     63
    LVLSMaxIncr As Byte         '1      244     64
    MageryB As Integer          '2      246     65
    MinIncrease As Byte         '1      247     66    u3
    LVLSMinIncr As Byte         '1      248     67    u4
    DurIncrease As Byte         '1      249     68    u5
    LVLSDurIncr As Byte         '1      250     69    u6
    ShortName As String * 5     '5      255     70
    AfterShortName As Byte      '1      256     71
    N06 As Long                 '4      260     72
End Type
Const SpellDataBufSize = 260
Public SpellFldMap(0 To 71) As PALN32.FieldMap
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
Public Type MonsterRecType              'A      B       C       A=field len, B=Cumulative field len, C=field count
    Number          As Long             '4      4       1
    EmptySpace      As String * 50      '50     54      2
    Name            As String * 29      '29     83      3
    nothing1        As Byte             '1      84      4
    Group           As Integer          '2      86      5
    nothingXX1      As Integer          '2      88      6
    ExpMulti        As Long             '4      92      7
    Index           As Integer          '2      94      8
    nothingXX3      As Integer          '2      96      9
    Something2      As Long             '4      100     10
    WeaponNumber    As Long             '4      104     11
    DR              As Integer          '2      106     12
    AC              As Integer          '2      108     13
    Something3      As Integer          '2      110     14
    Follow          As Integer          '2      112     15
    MR              As Integer          '2      114     16
    BSDefence       As Integer          '2      116     17
    Experience      As Long             '4      120     18
    Hitpoints       As Integer          '2      122     19
    Energy          As Integer          '2      124     20
    HPRegen         As Integer          '2      126     21
    AbilityA(9)     As Integer          '20     146     31
    AbilityB(9)     As Integer          '20     166     41
    GameLimit       As Integer          '2      168     42
    Active          As Integer          '2      170     43
    Type            As Integer          '2      172     44
    nothing2        As Byte             '1      173     45
    Undead          As Byte             '1      174     46
    Alignment       As Integer          '2      176     47
    nothing3        As Integer          '2      178     48
    RegenTime       As Integer          '2      180     49
    DateKilled      As Integer          '2      182     50
    TimeKilled      As Integer          '2      184     51
    MoveMsg         As Long             '4      188     52
    DeathMsg        As Long             '4      192     53
    ItemNumber(9)   As Long             '40     232     63
    ItemUses(9)     As Integer          '20     252     73
    ItemDropPer(9)  As Byte             '10     262     83
    nothing9        As Integer          '2      264     84
    Runic           As Long             '4      268     85
    Platinum        As Long             '4      272     86
    Gold            As Long             '4      276     87
    Silver          As Long             '4      280     88
    Copper          As Long             '4      284     89
    GreetTxt        As Long             '4      288     90
    CharmLvL        As Integer          '2      290     91
    Nothing16       As Integer          '2      292     92
    DescTxt         As Long             '4      296     93
    AttackType(4)   As Byte             '5      301     98
    Nothing22       As Byte             '1      302     99
    AttackAccuSpell(4) As Integer       '10     312     104
    AttackPer(4)    As Byte             '5      317     109
    Nothing17       As Byte             '1      318     110
    AttackMinHCastPer(4) As Integer     '10     328     115
    AttackMaxHCastLvl(4) As Integer     '10     338     120
    Nothing18       As Integer          '2      340     121
    AttackHitMsg(4) As Long             '20     360     126
    AttackDodgeMsg(4) As Long           '20     380     131
    AttackMissMsg(4) As Long            '20     400     136
    AttackEnergy(4) As Integer          '10     410     141
    Nothing19       As Integer          '2      412     142
    TalkTxt         As Long             '4      416     143
    CharmRes        As Integer          '2      418     144
    Nothing21       As Integer          '2      420     145
    AttackHitSpell(4) As Integer        '10     430     150
    DeathSpellNumber As Integer         '2      432     151
    Nothing23       As Integer          '2      434     152
    Nothing24       As Integer          '2      436     153
    Nothing25       As Integer          '2      438     154
    Nothing26       As Integer          '2      440     155
    Nothing27       As Integer          '2      442     156
    Nothing28       As Integer          '2      444     157
    Nothing29       As Integer          '2      446     158
    CreateSpellNumber As Integer        '2      448     159
    SpellNumber(4)  As Integer          '10     458     164
    SpellCastPer(4) As Byte             '5      463     169
    SpellCastLvl(4) As Byte             '5      468     174
    DescLine1       As String * 70      '70     538     175
    nothing10       As Byte             '1      539     176
    DescLine2       As String * 70      '70     609     177
    Nothing11       As Byte             '1      610     178
    DescLine3       As String * 70      '70     680     179
    Nothing12       As Byte             '1      681     180
    DescLine4       As String * 70      '70     751     181
    Nothing13       As Byte             '1      752     182
    Gender          As Byte             '1      753     183
    Nothing14       As Byte             '1      754     184
    Nothing15       As Integer          '2      756     185  section total: buf 300, 300+456=756 .. fields 185
End Type
Const MonsterDataBufSize = 756
Public MonsterFldMap(0 To 184) As PALN32.FieldMap
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
Public Type ItemRecType                 'A      B       C       A=field len, B=Cumulative field len, C=field count
    Number          As Long             '4      4       1
    unknown1        As Integer          '2      6       2
    GameLimit       As Integer          '2      8       3
    unknown2        As Integer          '2      10      4
    unknown3        As Integer          '2      12      5
    unknown4        As Integer          '2      14      6
    unknown5        As Integer          '2      16      7
    EmptySpace1     As String * 156     '156    172     8
    EmptySpace2     As Byte             '1      173     9
    Name            As String * 29      '29     202     10
    nothing1        As Byte             '1      203     11
    Desc1           As String * 60      '60     263     12
    nothing2        As Byte             '1      264     13
    Desc2           As String * 60      '60     324     14
    nothing3        As Byte             '1      325     15
    Desc3           As String * 60      '60     385     16
    nothing4        As Byte             '1      386     17
    Desc4           As String * 60      '60     446     18
    nothing5        As Byte             '1      447     19
    Desc5           As String * 60      '60     507     20
    Nothing6        As Byte             '1      508     21
    Desc6           As String * 60      '60     568     22
    nothing7        As Byte             '1      569     23
    Desc7           As String * 60      '60     629     24  these last three lines aren't actually descriptions
    nothing8        As Byte             '1      630     25
    Desc8           As String * 60      '60     690     26
    nothing9        As Byte             '1      691     27
    Desc9           As String * 60      '60     751     28
    nothing10       As Byte             '1      752     29
    unknown6        As Integer          '2      754     30
    Weight          As Integer          '2      756     31  en
    Type            As Integer          '2      758     32  item type
    AbilityA(19)    As Integer          '40     798     52
    Uses            As Integer          '2      800     53  uses
    unknown7        As Integer          '2      802     54  unknown
    Cost            As Integer          '2      804     55  cost
    Class(9)        As Integer          '20     824     65
    unknown8        As Integer          '2      826     66
    unknown9        As Integer          '2      828     67
    unknown10       As Integer          '2      830     68
    Minhit          As Integer          '2      832     69
    Maxhit          As Integer          '2      834     70
    AC              As Integer          '2      836     71
    Race(9)         As Long             '40     876     81
    Negate(19)      As Integer          '40     916     101
    Weapon          As Integer          '2      918     102
    Armour          As Integer          '2      920     103
    WornOn          As Integer          '2      922     104
    Accuracy        As Integer          '2      924     105
    DR              As Integer          '2      926     106
    Gettable        As Byte             '1      927     107
    unknown12       As Byte             '1      928     108
    ReqStr          As Integer          '2      930     109
    unknown13a(6)   As Integer          '14     944     116
    OpenRunic       As Long             '4      948     117
    OpenPlatinum    As Long             '4      952     118
    OpenGold        As Long             '4      956     119
    OpenSilver      As Long             '4      960     120
    OpenCopper      As Long             '4      964     121
    unknown13b(12)  As Integer          '26     990     134
    Speed           As Integer          '2      992     135
    unknown14       As Integer          '2      994     136
    AbilityB(19)    As Integer          '40     1034    156
    unknown15       As Integer          '2      1036    157
    HitMsg          As Long             '4      1040    158
    MissMsg         As Long             '4      1044    159
    ReadTB          As Long             '4      1048    160
    DistructMsg     As Long             '4      1052    161
    unknown16(5)    As Integer          '12     1064    167
    NotDroppable    As Byte             '1      1065    168
    CostType        As Byte             '1      1066    169
    RetainAfterUses As Byte             '1      1067    170
    Robable         As Byte             '1      1068    171
    DestroyOnDeath  As Byte             '1      1069    172
    unknown19       As Byte             '1      1070    173
End Type
Const ItemDataBufSize = 1072 '1948 '1072
Public ItemFldMap(0 To 172) As PALN32.FieldMap '188
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
Public Type ShopRecType                  'A      B       C       A=field len, B=Cumulative field len, C=field count
    Number              As Long          '4      4       1
    Name                As String * 39   '39     43      2
    ShopAfterName       As Integer       '2      45      3
    ShopDescriptionA    As String * 52   '52     97      4
    ShopNothing1        As Byte          '1      98      5
    ShopDescriptionB    As String * 52   '52     150     6
    ShopNothing2        As Byte          '1      151     7
    ShopDescriptionC    As String * 52   '52     203     8
    ShopNothing3        As Byte          '1      204     9
    ShopType            As Integer       '2      206     10
    ShopMinLvL          As Integer       '2      208     11
    ShopMaxLvl          As Integer       '2      210     12
    ShopMarkUp          As Integer       '2      212     13
    ShopNothing4        As Integer       '2      214     14
    ShopClassLimit      As Byte          '1      215     15
    ShopNothingAA       As Byte          '1      216     16
    ShopItemNumber(19)  As Long          '80     296     36
    ShopMax(19)         As Integer       '40     336     56
    ShopNow(19)         As Integer       '40     376     76
    ShopRgnTime(19)     As Integer       '40     416     96
    ShopRgnNumber(19)   As Integer       '40     456     116
    ShopRgnPercentage(19) As Byte        '20     476     136
End Type
Const ShopDataBufSize = 504
Public ShopFldMap(0 To 135) As PALN32.FieldMap
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
Public Type RoomRecType                     'A     B      C       A=field len, B=Cumulative field len, C=field count
    MapNumber           As Long             '4     4      1
    RoomNumber          As Long             '4     8      2
    EmptySpace          As String * 253     '253   261    3
    Name                As String * 53      '53    314    4
    Desc(6)             As String * 71      '497   811    11
    AnsiMap             As String * 13      '13    824    12
    RoomExit(9)         As Long             '40    864    22
    RoomType(9)         As Integer          '20    884    32
    Para1(9)            As Long             '40    924    42
    Para2(9)            As Integer          '20    944    52
    Para3(9)            As Long             '40    984    62
    Para4(9)            As Long             '40    1024   72
    CurrentRoomMon(14)  As Long             '60    1084   87            '***I_THROUGH_N*** (comment for n)
    'CurrentRoomMon(14)  As Integer          '*30*  *1054*   87            '***I_THROUGH_N*** (UNcomment for n)  -30
    Type                As Integer          '2     1086   88
    NewSpot             As Integer          '2     1088   89            '***I_THROUGH_N*** (comment for n) -2 == -32 less
    ShopNum             As Long             '4     1092   90
    nothing1(14)        As Integer          '30    1122   105
    MinIndex            As Integer          '2     1124   106
    MaxIndex            As Integer          '2     1126   107
    ByNumber            As Long             '4     1130   108
    dontknow            As Integer          '2     1132   109
    Light               As Integer          '2     1134   110
    GangHouseNumber     As Integer          '2     1136   111
    RoomItems(16)       As Long             '68    1204   128
    RoomItemUses(16)    As Integer          '34    1238   145
    nothing4            As Integer          '2     1240   146
    InvisItems(14)      As Long             '60    1300   161
    InvisItemUses(14)   As Integer          '30    1330   176
    nothing5            As Integer          '2     1332   177
    Runic               As Long             '4     1336   178
    Platinum            As Long             '4     1340   179
    Gold                As Long             '4     1344   180
    Silver              As Long             '4     1348   181
    Copper              As Long             '4     1352   182
    InvisRunic          As Long             '4     1356   183
    InvisPlatinum       As Long             '4     1360   184
    InvisGold           As Long             '4     1364   185
    InvisSilver         As Long             '4     1368   186
    InvisCopper         As Long             '4     1372   187
    MaxRegen            As Long             '4     1376   188
    MonsterType         As Integer          '2     1378   189
    unknown69           As Integer          '2     1380   190
    Attributes          As Long             '4     1384   191
    nothing9            As Long             '4     1388   192
    DeathRoom           As Long             '4     1392   193
    ExitRoom            As Long             '4     1396   194
    RoomItemQty(16)     As Integer          '34    1430   211
    InvisItemQty(14)    As Integer          '30    1460   226
    CmdText             As Long             '4     1464   227
    nothing10           As Long             '4     1468   228
    Delay               As Integer          '2     1470   229
    MaxArea             As Integer          '2     1472   230
    Nothing11           As Long             '4     1476   231
    ControlRoom         As Long             '4     1480   232
    PermNPC             As Long             '4     1484   233
    PlacedItems(9)      As Long             '40    1524   243
    Nothing12(1)        As Long             '8     1532   245
    Something1          As Long             '4     1536   246
    Spell               As Long             '4     1540   247
    unknown70           As Integer          '2     1542   248
    NumMons             As Byte             '1     1543   249
    unknown71           As Byte             '1     1544   250
End Type

Const RoomDataBufSize = 1544 '***I_THROUGH_N*** (comment for n)
'Const RoomDataBufSize = 1512 '***I_THROUGH_N*** (UNcomment for n)

Public RoomFldMap(0 To 249) As PALN32.FieldMap     '***I_THROUGH_N*** (comment for n)
'Public RoomFldMap(0 To 248) As PALN32.FieldMap    '***I_THROUGH_N*** (UNcomment for n)
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
Public Type MessageRecType                  'A      B       C       A=field len, B=Cumulative field len, C=field count
    Number              As Long             '4      4       1
    MessageLine1        As String * 74      '74     78      2
    MessageNothing02    As Integer          '2      80      3
    MessageNothing03    As Integer          '2      82      4
    MessageNothing04    As Integer          '2      84      5
    MessageNothing05    As Integer          '2      86      6
    MessageLine2        As String * 74      '74     160     7
    MessageNothing06    As Integer          '2      162     8
    MessageNothing07    As Integer          '2      164     9
    MessageNothing08    As Integer          '2      166     10
    MessageNothing09    As Integer          '2      168     11
    MessageLine3        As String * 74      '74     242     12
End Type
Const MessageDataBufSize = 256
Public MessageFldMap(0 To 11) As PALN32.FieldMap
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
Public Type TextblockRecType             'A      B       C       A=field len, B=Cumulative field len, C=field count
    PartNum         As Integer           '2      2       1
    LeadIn(1 To 14) As Byte              '14     16      15
    Number          As Long              '4      20      16
    LinkTo          As Long              '4      24      17
    Data            As String * 2000     '2000   2024    18
End Type
Public Const TextblockMaxBufSize = 2024
Public TextblockFldMap(0 To 17) As PALN32.FieldMap
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
Public Type UserRecType                 'A      B      C       A=field len, B=Cumulative field len, C=field count
    BBSName As String * 30              '30     30     1
    FirstName As String * 10            '10     40     2
    AfterFirstName As Byte              '1      41     3
    LastName As String * 18             '18     59     4
    AfterLastName As Byte               '1      60     5
    NotExperience As Long               '4      64     6
    SpellCasted(9) As Integer           '20     84     16  (10 elements × 2 bytes each = 20 bytes)
    SpellValue(9) As Integer            '20     104    26  (10 elements × 2 bytes each = 20 bytes)
    SpellRoundsLeft(9) As Integer       '20     124    36  (10 elements × 2 bytes each = 20 bytes)
    Title As String * 20                '20     144    37
    Race As Integer                     '2      146    38
    Class As Integer                    '2      148    39
    Level As Integer                    '2      150    40
    Stat(11) As Integer                 '24     174    52  (12 elements × 2 bytes each = 24 bytes)
    MaxHP As Integer                    '2      176    53  - defaults to 20 (… ) hardcoded in dll @ data_482cc8
    CurrentHP As Integer                '2      178    54  - defaults to 20 (… ) hardcoded in dll @ data_482cc8
    MaxENC As Integer                   '2      180    55
    CurrentENC As Integer               '2      182    56
    Energy(2) As Integer                '6      188    59  - 0 NEVER USED??? energy start? - hardcoded 1000 (…)
    unknown1 As Integer                 '2      190    60  - NEVER USED??? hardcoded 125 (…)
    SomeSortOfFlag As Integer           '2      192    61
    MagicRes As Integer                 '2      194    62
    MagicRes2 As Integer                '2      196    63
    MapNumber As Long                   '4      200    64
    RoomNum As Long                     '4      204    65
    nothing2 As Integer                 '2      206    66
    unknown2(1) As Integer              '4      210    68  (2 elements × 2 bytes each = 4 bytes)
    nothing3 As Integer                 '2      212    69
    unknown3(1) As Byte                 '2      214    71  (2 elements × 1 byte each = 2 bytes)
    nothing4 As Integer                 '2      216    72
    Item(99) As Long                    '400    616    172 (100 elements × 4 bytes each = 400 bytes)
    ItemUses(99) As Integer             '200    816    272 (100 elements × 2 bytes each = 200 bytes)
    nothing5 As Long                    '4      820    273
    Key(49) As Long                     '200    1020   323 (50 elements × 4 bytes each = 200 bytes)
    KeyUses(49) As Integer              '100    1120   373 (50 elements × 2 bytes each = 100 bytes)
    unknown4(3) As Long                 '16     1136   377 (4 elements × 4 bytes each = 16 bytes)
    BillionsOfExperience As Long        '4      1140   378
    MillionsOfExperience As Long        '4      1144   379
    Nothing6 As Integer                 '2      1146   380
    Spell(99) As Integer                '200    1346   480 (100 elements × 2 bytes each = 200 bytes)
    EvilPoints As Integer               '2      1348   481
    nothing7(2) As Long                 '12     1360   484 (3 elements × 4 bytes each = 12 bytes)
    LastMap(19) As Long                 '80     1440   504 (20 elements × 4 bytes each = 80 bytes)
    LastRoom(19) As Long                '80     1520   524 (20 elements × 4 bytes each = 80 bytes)
    nothing8 As Integer                 '2      1522   525
    BroadcastChan As Integer            '2      1524   526
    unknown5 As Long                    '4      1528   527  may be a bunch of bytes, flag for hiding (0x5f6==1526): *(uint8_t*)((char*)arg2 + 0x5f6) = 1;
    Perception As Integer               '2      1530   528
    Stealth As Integer                  '2      1532   529
    MartialArts As Integer              '2      1534   530
    Thievery As Integer                 '2      1536   531
    MaxMana As Integer                  '2      1538   532
    CurrentMana As Integer              '2      1540   533
    SpellCasting As Integer             '2      1542   534
    Traps As Integer                    '2      1544   535
    unknown6 As Integer                 '2      1546   536
    Picklocks As Integer                '2      1548   537
    Tracking As Integer                 '2      1550   538
    nothing9 As Integer                 '2      1552   539
    Runic As Long                       '4      1556   540
    Platinum As Long                    '4      1560   541
    Gold As Long                        '4      1564   542
    Silver As Long                      '4      1568   543
    Copper As Long                      '4      1572   544
    WeaponHand As Long                  '4      1576   545
    nothing10 As Long                   '4      1580   546
    WornItem(19) As Long                '80     1660   566 (20 elements × 4 bytes each = 80 bytes)
    unknown7(19) As Integer             '40     1700   586 (20 elements × 2 bytes each = 40 bytes)
    unknown8 As Integer                 '2      1702   587
    LivesRemaining As Integer           '2      1704   588
    unknown9a(5) As Byte                '6      1710   594
    unknown9(12) As Integer             '26     1736   607 (13 elements × 2 bytes each = 26 bytes)
    GangName As String * 19             '19     1755   608
    AfterGangName As Byte               '1      1756   609
    unknown11(5) As Byte                '6      1762   615 (6 elements × 1 byte each = 6 bytes)
    CPRemaining As Integer              '2      1764   616
    SuicidePassword As String * 8       '8      1772   617
    unknown12a(3) As Integer            '8      1780   621
    SomeUserFlags(7) As Byte            '8      1788   629  flag sneaking (0x6f4==1780): *(uint16_t*)((char*)arg1 + 0x6f4) |= 4;
    bEDITED As Byte                     '1      1789   630
    unknown12c As Byte                  '1      1790   631
    unknown12d(18) As Integer           '38     1828   650  see bitfield notes ?
        'Bit 00 (mask 0x0001): Room Display Detail
        'Bit 01 (mask 0x0002): PVP Combat Notification/Interruption
        'Bit 02 (mask 0x0004): I think... 1 == CAN cast between round spell.  0 == already cast between round spell.
        'Bit 03 (mask 0x0008): Gift Acceptance
        'Bit 04 (mask 0x0010): Evil Warning/Interruption
        'Bit 05 (mask 0x0020): Telepath Blocking
        'Bit 06 (mask 0x0040): set just after "dark cloud"
        'Bit 07 (mask 0x0080):
        'Bit 08 (mask 0x0100): Response Verbosity
        'Bit 09 (mask 0x0200):
        'Bit 10 (mask 0x0400): Keep Setting on Death/Reroll
        'Bit 11 (mask 0x0800): Message Style (Technical vs. Fantasy)
        'Bit 12 (mask 0x1000):
        'Bit 13 (mask 0x2000):
        'Bit 14 (mask 0x3000):
        'Bit 15 (mask 0x8000): reference to having a character but not paid on the bbs
        'unknown12d(0) and unknown12d(1) referenced in dll as userrecord + 0x700
        'unknown12d(5) CURRENT ENCUMBRANCE AS A PERCENTAGE - 2025.01.12
        'unknown12d(6) TOTAL ACCURACY FROM ability 22 (auras and items) - 2025.03.01
        'unknown12d(7) TOTAL AC FROM ability 2
        'unknown12d(8) TOTAL MAX DAMAGE from ability 4
        'unknown12d(9) something referenced here for crits 2025.03.12
        'unknown12d(17) something related to a transition in evil points at one point?
    HitPointRolls As Byte               '1      1829   651
    unknown12e As Byte                  '1      1830   652
    unknown12f(9) As Integer            '20     1850   662 (10 elements × 2 bytes each = 20 bytes)
    Ability(29) As Integer              '60     1910   692 (30 elements × 2 bytes each = 60 bytes)
    AbilityModifier(29) As Integer      '60     1970   722 (30 elements × 2 bytes each = 60 bytes)
    unknown13a As Integer               '2      1972   723
    unknown13b As Integer               '2      1974   724  TOTAL CRITS FROM ability 58
    unknown13c As Integer               '2      1976   725  TOTAL DR FROM ability 7
    unknown13d As Integer               '2      1978   726  TOTAL AlterDRpercent FROM ability 99
    unknown13e As Integer               '2      1980   727  TOTAL SPEED FROM ability 87
    unknown13f As Integer               '2      1982   728  TOTAL AC/BLUR FROM ability 10
    unknown13g As Integer               '2      1984   729  TOTAL "DefenseModifier" FROM ability 104
    CharLife As Long                    '4      1988   730
    unknown13(8) As Integer             '18     2006   739  _energy_update_character sets unknown13(0) to 0 if ENCUM_PCT > 66; flags/bitmask in element 8
        'unknown13(1) gets set to a timestamp in ljngame_polling
        'unknown13(2) SOME BITMASK RELATED TO MOVEMENT IMPEDANCE OR STEALTH?
        'unknown13(4) this may be a 4-byte long
        'unknown13(8) flags/bitmask here 0x20 = class stealth, 0x40 = race stealth, 0x60 = race+class stealth
    Bitmask1 As Byte                    '1      2007   740
    Bitmask2 As Byte                    '1      2008   741
    TestFlag1 As Byte                   '1      2009   742
    TestFlag2 As Byte                   '1      2010   743
    unknown14 As Integer                '2      2012   744
    unknown15(3) As Long                '16     2028   748  (4 elements × 4 bytes each = 16 bytes)
End Type

Const UserDataBufSize = 2028
Public UserFldMap(0 To 747) As PALN32.FieldMap
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
Public Type BankrecType                'A      B      C       A=field len, B=Cumulative field len, C=field count
    BBSName     As String * 30         '30     30     1
    nothing1    As Integer             '2      32     2
    ShopNumber  As Long                '4      36     3
    Cash        As Long                '4      40     4
End Type
Const BankDatabufSize = 76
Public BankFldMap(0 To 3) As PALN32.FieldMap
Public Type BankDatabufType
    buf(1 To BankDatabufSize) As Byte
End Type
Public Type BankPosBlockType
    buf(1 To 128) As Byte
End Type

Public BankKey As BankKeyType
Public BankKeyDataBuf As BankKeyDataBufType
Public Type BankKeyType                'A      B      C       A=field len, B=Cumulative field len, C=field count
    BBSName     As String * 30         '30     30     1
    nothing1    As Byte                '1      31     2
    nothing2    As Byte                '1      32     3
    ShopNumber  As Long                '4      36     4
End Type
Const BankKeyDataBufSize = 36
Public BankKeyFldMap(0 To 3) As PALN32.FieldMap
Public Type BankKeyDataBufType
    buf(1 To BankKeyDataBufSize) As Byte
End Type

'====> ACTIONS
Public Actionrec As ActionRecType
Public ActionDatabuf As ActionDatabufType
Public ActionPosBlock As ActionPosBlockType
Public ActionKeyBuffer As String * 255
Public Type ActionRecType              'A      B      C       A=field len, B=Cumulative field len, C=field count
    Name                As String * 29 '29     29     1
    AfterName           As Byte        '1      30     2
    SingleToUser        As String * 74 '74     104    3
    nothing2(2)         As Integer     '6      110    6    (3 elements × 2 bytes)
    SingleToRoom        As String * 74 '74     184    7
    nothing3(2)         As Integer     '6      190    10   (3 elements × 2 bytes)
    UserToUser          As String * 74 '74     264    11
    nothing4(2)         As Integer     '6      270    14   (3 elements × 2 bytes)
    UserToOtherUser     As String * 74 '74     344    15
    nothing5(2)         As Integer     '6      350    18   (3 elements × 2 bytes)
    UserToRoom          As String * 74 '74     424    19
    Nothing6(2)         As Integer     '6      430    22   (3 elements × 2 bytes)
    MonsterToUser       As String * 74 '74     504    23
    nothing7(2)         As Integer     '6      510    26   (3 elements × 2 bytes)
    MonsterToRoom       As String * 74 '74     584    27
    nothing8(2)         As Integer     '6      590    30   (3 elements × 2 bytes)
    InventoryToUser     As String * 74 '74     664    31
    nothing9(2)         As Integer     '6      670    34   (3 elements × 2 bytes)
    InventoryToRoom     As String * 74 '74     744    35
    nothing10(2)        As Integer     '6      750    38   (3 elements × 2 bytes)
    FloorItemToUser     As String * 74 '74     824    39
    Nothing11(2)        As Integer     '6      830    42   (3 elements × 2 bytes)
    FloorItemToRoom     As String * 74 '74     904    43
    Nothing12(2)        As Integer     '6      910    46   (3 elements × 2 bytes)
    Offset              As Byte        '1      911    47
End Type
Const ActionDataBufSize = 1010
Public ActionFldMap(0 To 46) As PALN32.FieldMap
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
Public Type GangrecType                'A      B      C       A=field len, B=Cumulative field len, C=field count
    KeyName         As String * 20     '20     20     1
    DisplayName     As String * 20     '20     40     2
    Exp             As Long            '4      44     3
    Leader          As String * 30     '30     74     4
    DateCreated     As Integer         '2      76     5
    unknown1        As Integer         '2      78     6
    Members         As Integer         '2      80     7
    unknown2        As Long            '4      84     8
    unknown3        As Long            '4      88     9
    RollOver        As Long            '4      92     10
    RollTimes       As Long            '4      96     11
    unknown5        As String * 160    '160    256    12
End Type

Const GangDatabufSize = 256
Public GangFldMap(0 To 11) As PALN32.FieldMap
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
Public Type DBStatType                 'A      B      C       A=field len, B=Cumulative field len, C=field count
    RecLen          As Integer         '2      2      1
    PageSize        As Integer         '2      4      2
    nIndexes        As Integer         '2      6      3
    nRecords        As Long            '4      10     4
    FileFlags       As Integer         '2      12     5
    ReservedWord    As Integer         '2      14     6
    UnusedPages     As Integer         '2      16     7
    KeyPosition     As Integer         '2      18     8
    KeyLength       As Integer         '2      20     9
    KeyFlags        As Integer         '2      22     10
    nUniqueKeys     As Long            '4      26     11
    ExtendedDataType As Byte           '1      27     12
    TheRest         As String * 1893   '1893   1920   13
End Type
Public Const DBStatBufSize = 1920
Public DBStatFldMap(0 To 12) As PALN32.FieldMap
Public Type DBStatDatabufType
    buf(1 To DBStatBufSize) As Byte
End Type
Public DBStatDatabuf As DBStatDatabufType

Sub IntFieldMaps()

DLog "AddRaceFieldMap RaceFldMap..."
    AddRaceFieldMap RaceFldMap, 0
DLog "AddClassFieldMap ClassFldMap..."
    AddClassFieldMap ClassFldMap, 0
    AddSpellFieldMap SpellFldMap, 0
    AddMonsterFieldMap MonsterFldMap, 0
    AddItemFieldMap ItemFldMap, 0
    AddShopFieldMap ShopFldMap, 0
    AddRoomFieldMap RoomFldMap, 0
    AddMessageFieldMap MessageFldMap, 0
    AddTextblockFieldMap TextblockFldMap, 0
DLog "AddUserFieldMap UserFldMap..."
    AddUserFieldMap UserFldMap, 0
    AddActionFieldMap ActionFldMap, 0
    AddBankFieldMap BankFldMap, 0
    AddBankKeyFieldMap BankKeyFldMap, 0
    AddGangFieldMap GangFldMap, 0
    AddDBStatFieldMap DBStatFldMap, 0
End Sub

Sub AddRoomFilterFieldMap(Map() As PALN32.FieldMap, ByRef ctr As Long)
    AddField Map, ctr, FLD_INTEGER, 4 'map
    AddField Map, ctr, FLD_INTEGER, 4 'room
    AddField Map, ctr, FLD_INTEGER, 4 'value
End Sub
Sub AddDBStatFieldMap(Map() As PALN32.FieldMap, ByRef ctr As Long)
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
Sub AddGangFieldMap(Map() As PALN32.FieldMap, ByRef ctr As Long)
    
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
Sub AddBankKeyFieldMap(Map() As PALN32.FieldMap, ByRef ctr As Long)
    AddField Map, ctr, FLD_STRING, 30   'bbs name
    AddField Map, ctr, FLD_BYTE, 1   'nothing
    AddField Map, ctr, FLD_BYTE, 1   'nothing
    AddField Map, ctr, FLD_INTEGER, 4   'Shop number
End Sub

Sub AddBankFieldMap(Map() As PALN32.FieldMap, ByRef ctr As Long)
    AddField Map, ctr, FLD_STRING, 30   'bbs name
    AddField Map, ctr, FLD_INTEGER, 2   'nothing
    AddField Map, ctr, FLD_INTEGER, 4   'Shop number
    AddField Map, ctr, FLD_INTEGER, 4   'Cash amount
End Sub

Sub AddUserFieldMap(Map() As PALN32.FieldMap, ByRef ctr As Long)
    AddField Map, ctr, FLD_STRING, 30   'bbs name               'bytes 1-30
    AddField Map, ctr, FLD_STRING, 10   'firstname              'starts at byte 31-40
    AddField Map, ctr, FLD_BYTE, 1                              '41
    AddField Map, ctr, FLD_STRING, 18   'lastname               '42-59
    AddField Map, ctr, FLD_BYTE, 1                              '60
    AddField Map, ctr, FLD_INTEGER, 4   'notexperience          '61-64
    AddField Map, ctr, FLD_INTEGER, 2   'splcast 1              'starts at byte 65-66
    AddField Map, ctr, FLD_INTEGER, 2   '2                      '67-68
    AddField Map, ctr, FLD_INTEGER, 2   '3                      '69-70
    AddField Map, ctr, FLD_INTEGER, 2   '4                      '71-72
    AddField Map, ctr, FLD_INTEGER, 2   '5                      '73-74
    AddField Map, ctr, FLD_INTEGER, 2   '6                      '75-76
    AddField Map, ctr, FLD_INTEGER, 2   '7                      '77-78
    AddField Map, ctr, FLD_INTEGER, 2   '8                      '79-80
    AddField Map, ctr, FLD_INTEGER, 2   '9                      '81-82
    AddField Map, ctr, FLD_INTEGER, 2   'splcast 10             '83-84
    AddField Map, ctr, FLD_INTEGER, 2   'splval 1               '85-86
    AddField Map, ctr, FLD_INTEGER, 2   '2 88
    AddField Map, ctr, FLD_INTEGER, 2   '3 90
    AddField Map, ctr, FLD_INTEGER, 2   '4 92
    AddField Map, ctr, FLD_INTEGER, 2   '5 94
    AddField Map, ctr, FLD_INTEGER, 2   '6 96
    AddField Map, ctr, FLD_INTEGER, 2   '7 98
    AddField Map, ctr, FLD_INTEGER, 2   '8 100
    AddField Map, ctr, FLD_INTEGER, 2   '9 102
    AddField Map, ctr, FLD_INTEGER, 2   'splval 104
    AddField Map, ctr, FLD_INTEGER, 2   'splround 1             '105-106
    AddField Map, ctr, FLD_INTEGER, 2   '2
    AddField Map, ctr, FLD_INTEGER, 2   '3
    AddField Map, ctr, FLD_INTEGER, 2   '4
    AddField Map, ctr, FLD_INTEGER, 2   '5
    AddField Map, ctr, FLD_INTEGER, 2   '6
    AddField Map, ctr, FLD_INTEGER, 2   '7
    AddField Map, ctr, FLD_INTEGER, 2   '8
    AddField Map, ctr, FLD_INTEGER, 2   '9
    AddField Map, ctr, FLD_INTEGER, 2   'splround 10            '123-124
    AddField Map, ctr, FLD_STRING, 20   'title                  '125-145
    AddField Map, ctr, FLD_INTEGER, 2   'Race                   '147-148
    AddField Map, ctr, FLD_INTEGER, 2   'Class                  '149-150
    AddField Map, ctr, FLD_INTEGER, 2   'Level                  '151-152
    AddField Map, ctr, FLD_INTEGER, 2   'Stat1                  '153-154
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
    AddField Map, ctr, FLD_INTEGER, 2   'Stat12                 '177
    AddField Map, ctr, FLD_INTEGER, 2   'MaxHP                  '176
    AddField Map, ctr, FLD_INTEGER, 2   'CurrentHP              '178-179
    AddField Map, ctr, FLD_INTEGER, 2   'max encum              '180-181
    AddField Map, ctr, FLD_INTEGER, 2   'cur encum              '182-183
    AddField Map, ctr, FLD_INTEGER, 2   'energy                 '184-185
    AddField Map, ctr, FLD_INTEGER, 2   'energy                 '186-187
    AddField Map, ctr, FLD_INTEGER, 2   'energy                 '188-189
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
    AddField Map, ctr, FLD_BYTE, 1   '?
    AddField Map, ctr, FLD_BYTE, 1   '?
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
    AddField Map, ctr, FLD_INTEGER, 2 'unknown7 1
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
    AddField Map, ctr, FLD_INTEGER, 2 'unknown7 20
    AddField Map, ctr, FLD_INTEGER, 2 'unknown8
    AddField Map, ctr, FLD_INTEGER, 2 'lives remaining
    AddField Map, ctr, FLD_BYTE, 1 'unknown9A 1
    AddField Map, ctr, FLD_BYTE, 1 'unknown9A 2
    AddField Map, ctr, FLD_BYTE, 1 'unknown9A 3
    AddField Map, ctr, FLD_BYTE, 1 'unknown9A 4
    AddField Map, ctr, FLD_BYTE, 1 'unknown9A 5
    AddField Map, ctr, FLD_BYTE, 1 'unknown9A 6
    AddField Map, ctr, FLD_INTEGER, 2 'unknown9 1
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'unknown9 5
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'unknown9 10
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'unknown9 13
    AddField Map, ctr, FLD_STRING, 19 'gang name
    AddField Map, ctr, FLD_BYTE, 1 'after gang name
    AddField Map, ctr, FLD_BYTE, 1 'unknown 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1 'unknown 6
    AddField Map, ctr, FLD_INTEGER, 2 'cps
    AddField Map, ctr, FLD_STRING, 8 'suicide
    AddField Map, ctr, FLD_INTEGER, 2 'unknown12a 0
    AddField Map, ctr, FLD_INTEGER, 2 'unknown12a 1
    AddField Map, ctr, FLD_INTEGER, 2 'unknown12a 2
    AddField Map, ctr, FLD_INTEGER, 2 'unknown12a 3
    AddField Map, ctr, FLD_BYTE, 1 'SomeUserFlags 1
    AddField Map, ctr, FLD_BYTE, 1 'SomeUserFlags 2
    AddField Map, ctr, FLD_BYTE, 1 'SomeUserFlags 3
    AddField Map, ctr, FLD_BYTE, 1 'SomeUserFlags 4
    AddField Map, ctr, FLD_BYTE, 1 'SomeUserFlags 5
    AddField Map, ctr, FLD_BYTE, 1 'SomeUserFlags 6
    AddField Map, ctr, FLD_BYTE, 1 'SomeUserFlags 7
    AddField Map, ctr, FLD_BYTE, 1 'SomeUserFlags 8
    AddField Map, ctr, FLD_BYTE, 1 'bEDITED As Byte
    AddField Map, ctr, FLD_BYTE, 1 'unknown12c As Byte
    AddField Map, ctr, FLD_INTEGER, 2 'unknown12d(18)0
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'unknown12d(18)10
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2 'unknown12d(18)
    
    AddField Map, ctr, FLD_BYTE, 1 'HitPointRolls As Byte
    AddField Map, ctr, FLD_BYTE, 1 'unknown12e As Byte
    
    AddField Map, ctr, FLD_INTEGER, 2
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

Sub AddTextblockFieldMap(Map() As PALN32.FieldMap, ByRef ctr As Long)
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

Sub AddRaceFieldMap(ByRef Map() As PALN32.FieldMap, ByRef ctr As Long)
    DLog "race..." & ctr
    AddField Map, ctr, FLD_INTEGER, 2
    DLog "race..." & ctr
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

Sub AddClassFieldMap(Map() As PALN32.FieldMap, ByRef ctr As Long)
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

Sub AddSpellFieldMap(Map() As PALN32.FieldMap, ByRef ctr As Long)
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

Sub AddMonsterFieldMap(Map() As PALN32.FieldMap, ByRef ctr As Long)
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
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 70   'Desc Line2
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 70   'Desc Line3
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_STRING, 70   'Desc Line4
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_BYTE, 1      'Gender
    AddField Map, ctr, FLD_BYTE, 1
    AddField Map, ctr, FLD_INTEGER, 2
End Sub

Sub AddItemFieldMap(Map() As PALN32.FieldMap, ByRef ctr As Long)
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

Sub AddShopFieldMap(Map() As PALN32.FieldMap, ByRef ctr As Long)
    AddField Map, ctr, FLD_INTEGER, 4   'SNumber
    AddField Map, ctr, FLD_STRING, 39   'SName
    AddField Map, ctr, FLD_INTEGER, 2   '
    AddField Map, ctr, FLD_STRING, 52   'SDescA
    AddField Map, ctr, FLD_BYTE, 1   '
    AddField Map, ctr, FLD_STRING, 52   'SDescB
    AddField Map, ctr, FLD_BYTE, 1   '
    AddField Map, ctr, FLD_STRING, 52   'SDescC
    AddField Map, ctr, FLD_BYTE, 1   '
    AddField Map, ctr, FLD_INTEGER, 2   'SType
    AddField Map, ctr, FLD_INTEGER, 2   'SMinLvl
    AddField Map, ctr, FLD_INTEGER, 2   'SMaxLvl
    AddField Map, ctr, FLD_INTEGER, 2   'SMarkUp
    AddField Map, ctr, FLD_INTEGER, 2   '
    AddField Map, ctr, FLD_BYTE, 1   'ClassLimit
    AddField Map, ctr, FLD_BYTE, 1   '
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

Sub AddRoomFieldMap(Map() As PALN32.FieldMap, ByRef ctr As Long)
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
    
    AddField Map, ctr, FLD_INTEGER, 4   'current mons 1        '***I_THROUGH_N*** (comment for n)
    AddField Map, ctr, FLD_INTEGER, 4   '2                     '         ||
    AddField Map, ctr, FLD_INTEGER, 4   '3                     '         ||
    AddField Map, ctr, FLD_INTEGER, 4   '4                     '         ||
    AddField Map, ctr, FLD_INTEGER, 4   '5                     '         ||
    AddField Map, ctr, FLD_INTEGER, 4   '6                     '         ||
    AddField Map, ctr, FLD_INTEGER, 4   '7                     '         ||
    AddField Map, ctr, FLD_INTEGER, 4   '8                     '         ||
    AddField Map, ctr, FLD_INTEGER, 4   '9                     '         ||
    AddField Map, ctr, FLD_INTEGER, 4   '10                    '         ||
    AddField Map, ctr, FLD_INTEGER, 4   '11                    '         ||
    AddField Map, ctr, FLD_INTEGER, 4   '12                    '         ||
    AddField Map, ctr, FLD_INTEGER, 4   '13                    '         ||
    AddField Map, ctr, FLD_INTEGER, 4   '14                    '        \\//
    AddField Map, ctr, FLD_INTEGER, 4   'current mons 15       '***I_THROUGH_N*** (comment for n)

'    AddField Map, ctr, FLD_INTEGER, 2   'current mons 1         '***I_THROUGH_N*** (UNcomment for n)
'    AddField Map, ctr, FLD_INTEGER, 2   '2                      '         ||
'    AddField Map, ctr, FLD_INTEGER, 2   '3                      '         ||
'    AddField Map, ctr, FLD_INTEGER, 2   '4                      '         ||
'    AddField Map, ctr, FLD_INTEGER, 2   '5                      '         ||
'    AddField Map, ctr, FLD_INTEGER, 2   '6                      '         ||
'    AddField Map, ctr, FLD_INTEGER, 2   '7                      '         ||
'    AddField Map, ctr, FLD_INTEGER, 2   '8                      '         ||
'    AddField Map, ctr, FLD_INTEGER, 2   '9                      '         ||
'    AddField Map, ctr, FLD_INTEGER, 2   '10                     '         ||
'    AddField Map, ctr, FLD_INTEGER, 2   '11                     '         ||
'    AddField Map, ctr, FLD_INTEGER, 2   '12                     '         ||
'    AddField Map, ctr, FLD_INTEGER, 2   '13                     '         ||
'    AddField Map, ctr, FLD_INTEGER, 2   '14                     '        \\//
'    AddField Map, ctr, FLD_INTEGER, 2   'current mons 15        '***I_THROUGH_N*** (UNcomment for n)
    
    AddField Map, ctr, FLD_INTEGER, 2   'Type
    AddField Map, ctr, FLD_INTEGER, 2   'new spot           '***I_THROUGH_N*** (comment for n)
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

Sub AddMessageFieldMap(Map() As PALN32.FieldMap, ByRef ctr As Long)
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
Sub AddActionFieldMap(Map() As PALN32.FieldMap, ByRef ctr As Long)
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
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_INTEGER, 2
    AddField Map, ctr, FLD_BYTE, 1      'Offset
End Sub

Sub AddField(Map() As PALN32.FieldMap, ByRef ctr As Long, dataType As PALN32.FldDataType, length As Long)
DLog "AddField > SetField..."
  PALN32.AlignFuncs.SetField Map(ctr), dataType, CLng(length)
DLog "ctr + 1..."
  ctr = ctr + 1
End Sub
