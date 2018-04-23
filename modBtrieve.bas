Attribute VB_Name = "modBtrieve"
Option Base 0
Option Explicit

DefInt A-Z
Public Const BOPEN = 0
Public Const BCLOSE = 1
Public Const BINSERT = 2
Public Const BUPDATE = 3
Public Const BDELETE = 4
Public Const BGETEQUAL = 5
Public Const BGETNEXT = 6
Public Const BGETPREVIOUS = 7
Public Const BGETGREATER = 8
Public Const BGETGREATEROREQUAL = 9
Public Const BGETFIRST = 12
Public Const BGETLAST = 13
Public Const BCREATE = 14
Public Const BSTAT = 15
Public Const BBEGINTRANS = 19
Public Const BTRANSSEND = 20
Public Const BABORTTRANS = 21
Public Const BGETPOSITION = 22
Public Const BGETRECORD = 23
Public Const BSTOP = 25
Public Const BVERSION = 26
Public Const BRESET = 28
Public Const BGETNEXTEXTENDED = 36
Public Const BGETKEY = 50

Public Const KEY_BUF_LEN = 255

Public Const FIXED = 67

Rem  Key Flags
Public Const DUP = 1
Public Const MODIFIABLE = 2
Public Const BIN = 4
Public Const NUL = 8
Public Const SEGMENT = 16

Public Const SEQ = 32
Public Const DEC = 64
Public Const SUP = 128

Rem  Key Types
Public Const EXTTYPE = 256
Public Const MANUAL = 512
Public Const BSTRING = 0
Public Const BINTEGER = 1
Public Const BFLOAT = 2
Public Const BDATE = 3
Public Const BTIME = 4
Public Const BDECIMAL = 5
Public Const BNUMERIC = 8
Public Const BZSTRING = 11
Public Const BAUTOINC = 15

Public Const B_NO_ERROR = 0
Public Const B_END_OF_FILE = 9

Public Const VAR_RECS = 1
Public Const BLANK_TRUNC = 2
Public Const PRE_ALLOC = 4
Public Const DATA_COMP = 8
Public Const KEY_ONLY = 16
Public Const BALANCED_KEYS = 32
Public Const FREE_10 = 64
Public Const FREE_20 = 128
Public Const FREE_30 = 192
Public Const DUP_PTRS = 256
Public Const INCLUDE_SYSTEM_DATA = 512
Public Const NO_INCLUDE_SYSTEM_DATA = 4608
Public Const SPECIFY_KEY_NUMS = 1024
Public Const VATS_SUPPORT = 2048

Public Const FLD_STRING = 0
Public Const FLD_INTEGER = 1
Public Const FLD_IEEE = 2
Public Const FLD_DATE = 3
Public Const FLD_TIME = 4
Public Const FLD_MONEY = 6
Public Const FLD_LOGICAL = 7
Public Const FLD_BYTE = 19
Public Const FLD_UNICODE = 20
Public Const FLD_UNSIGNEDBINARY = 14

Declare Function BTRCALL Lib "wbtrv32" (ByVal OP As Integer, Pb As Any, DB As Any, DL As Long, ByRef Kb As Any, ByVal Kl As Integer, ByVal Kn As Integer) As Integer

Public Function BtrieveErrorCode(ByVal nStatus As Variant, Optional ByVal NoCRLF As Boolean) As String
Select Case nStatus
    Case 1
        BtrieveErrorCode = "Invalid Operation (1)"
    Case 2
        BtrieveErrorCode = "Disk I/O Error (2)"
    Case 3
        BtrieveErrorCode = "File Not Open (3)"
    Case 4
        BtrieveErrorCode = "Record Not Found (4)"
        If Not NoCRLF Then
            BtrieveErrorCode = BtrieveErrorCode & vbCrLf & vbCrLf & "The record you are trying to goto doesn't exist."
        End If
    Case 5
        BtrieveErrorCode = "Duplicate Record (5)"
        If Not NoCRLF Then
            BtrieveErrorCode = BtrieveErrorCode & vbCrLf & vbCrLf & "That record already exists."
        End If
    Case 6
        BtrieveErrorCode = "Invalid Record Number (6)"
    Case 7
        BtrieveErrorCode = "Different Record Number (7)"
    Case 8
        BtrieveErrorCode = "Invalid Positioning (8)"
    Case 9
        BtrieveErrorCode = "End-Of-File (9)"
        If Not NoCRLF Then
            BtrieveErrorCode = BtrieveErrorCode & vbCrLf & vbCrLf & "You are already at the beginning/end of the database, or the Database is empty."
        End If
    Case 10
        BtrieveErrorCode = "Modifiable Index Value Error (10)"
    Case 11
        BtrieveErrorCode = "Invalid Location (11)"
        If Not NoCRLF Then
            BtrieveErrorCode = BtrieveErrorCode & vbCrLf & vbCrLf & "Go to File --> Settings and set your Datfile Path and Call Letters"
        End If
    Case 12
        BtrieveErrorCode = "File Not Found (12)"
        If Not NoCRLF Then
            BtrieveErrorCode = BtrieveErrorCode & vbCrLf & vbCrLf & "Go to File --> Settings and set your Datfile Path and Call Letters"
        End If
    Case 13
        BtrieveErrorCode = "Extended File Error (13)"
    Case 14
        BtrieveErrorCode = "Pre-Image Open Error (14)"
    Case 15
        BtrieveErrorCode = "Pre-Image I/O Error  (15)"
    Case 17
        BtrieveErrorCode = "Close Error (17)"
    Case 18
        BtrieveErrorCode = "Disk Full (18)"
    Case 19
        BtrieveErrorCode = "Unrecoverable Error (19)"
    Case 20
        BtrieveErrorCode = "Record Manager Inactive (20)"
    Case 21
        BtrieveErrorCode = "Index Buffer Too Short (21)"
    Case 22
        BtrieveErrorCode = "Data Buffer Length (22)"
    Case 23
        BtrieveErrorCode = "Position Block Length (23)"
    Case 24
        BtrieveErrorCode = "Page Size Error (24)"
    Case 25
        BtrieveErrorCode = "Create I/O Error (25)"
    Case 26
        BtrieveErrorCode = "Number of Keys (26)"
    Case 27
        BtrieveErrorCode = "Invalid Key Position (27)"
    Case 28
        BtrieveErrorCode = "Invalid Record Length (28)"
    Case 29
        BtrieveErrorCode = "Invalid Record Length (29)"
    Case 30
        BtrieveErrorCode = "Not A Btrieve File (30)"
    Case 35
        BtrieveErrorCode = "Directory Error (35), Go to File --> Settings and set your Datfile Path"
    Case 36
        BtrieveErrorCode = "TransactiOn Error (36)"
    Case 37
        BtrieveErrorCode = "Transaction Is Active (37)"
    Case 38
        BtrieveErrorCode = "Transaction Control File I/O Error (38)"
    Case 39
        BtrieveErrorCode = "End/Abort TransactiOn Error (39)"
    Case 40
        BtrieveErrorCode = "Transaction Max Files (40)"
    Case 41
        BtrieveErrorCode = "Operation Not Allowed (41)"
    Case 43
        BtrieveErrorCode = "Invalid Record Access (43)"
    Case 44
        BtrieveErrorCode = "Null Index Path (44)"
    Case 46
        BtrieveErrorCode = "Access To File Denied (46)"
    Case 51
        BtrieveErrorCode = "Invalid Owner (51)"
    Case 52
        BtrieveErrorCode = "Error Writing Cache (52)"
    Case 54
        BtrieveErrorCode = "Variable Page Error During a Step Direct operation (54)"
    Case 55
        BtrieveErrorCode = "Autoincrement Error (55)"
    Case 58
        BtrieveErrorCode = "Compression Buffer Too Short (58)"
    Case 66
        BtrieveErrorCode = "Maximum Number of Open Databases Exceeded (66)"
    Case 67
        BtrieveErrorCode = "Rl Could Not Open SQL Data Dictionaries (67)"
    Case 68
        BtrieveErrorCode = "Rl Cascades Too Deeply (68)"
    Case 69
        BtrieveErrorCode = "Rl Cascade Error (69)"
    Case 71
        BtrieveErrorCode = "Rl Definitions Violation (71)"
    Case 72
        BtrieveErrorCode = "Rl Referenced File Could Not Be Opnend (72)"
    Case 73
        BtrieveErrorCode = "Rl Definition Out Of Sync (73)"
    Case 76
        BtrieveErrorCode = "Rl Referenced File Conflict (76)"
    Case 77
        BtrieveErrorCode = "Wait Error (77)"
    Case 78
        BtrieveErrorCode = "Deadlock Detected (78)"
    Case 79
        BtrieveErrorCode = "Programming Error (79)"
    Case 80
        BtrieveErrorCode = "Conflict (80)"
    Case 81
        BtrieveErrorCode = "Lock Error (81)"
    Case 82
        BtrieveErrorCode = "Lost Position (82)"
    Case 83
        BtrieveErrorCode = "Read Outside Transaction (83)"
    Case 84
        BtrieveErrorCode = "Record In Use (84)"
    Case 85
        BtrieveErrorCode = "File In Use (85)"
    Case 86
        BtrieveErrorCode = "File Table Full"
    Case 87
        BtrieveErrorCode = "Handle Table Full"
    Case 88
        BtrieveErrorCode = "Incompatible Mode Error"
    Case 90
        BtrieveErrorCode = "Redirected Device Table Full"
    Case 91
        BtrieveErrorCode = "Server Error"
    Case 92
        BtrieveErrorCode = "Transaction Table Full"
    Case 93
        BtrieveErrorCode = "Incompatible Lock Type"
    Case 94
        BtrieveErrorCode = "PermissiOn Error"
    Case 95
        BtrieveErrorCode = "Session No Longer Valid"
    Case 96
        BtrieveErrorCode = "Communications Environment Error"
    Case 97
        BtrieveErrorCode = "Data Message Too Small"
    Case 98
        BtrieveErrorCode = "Internal TransactiOn Error"
    Case 100
        BtrieveErrorCode = "No Cache Buffers Available"
    Case 101
        BtrieveErrorCode = "No OS Memory Availabl"
    Case 102
        BtrieveErrorCode = "Not Enough Stack space"
    Case 1001
        BtrieveErrorCode = "Lock Option Out Of Range"
    Case 1002
        BtrieveErrorCode = "Memory AllocatiOn Error"
    Case 1003
        BtrieveErrorCode = "Memory Option Too Small"
    Case 1004
        BtrieveErrorCode = "Page Size Option Out Of Range"
    Case 1005
        BtrieveErrorCode = "Invalid Pre-Image Drive Option"
    Case 1007
        BtrieveErrorCode = "Files Option Out of Range"
    Case 1008
        BtrieveErrorCode = "Invalid Initialization Option"
    Case 1009
        BtrieveErrorCode = "Invalid Transaction File Open"
    Case 1011
        BtrieveErrorCode = "Compression Buffer Out Of Range"
    Case 1013
        BtrieveErrorCode = "Task Table Full"
    Case 1014
        BtrieveErrorCode = "Stop Warning"
    Case 1015
        BtrieveErrorCode = "Invalid Pointer"
    Case 1016
        BtrieveErrorCode = "Already Initialized"
    Case 2001
        BtrieveErrorCode = "Insufficient Memory"
    Case 2003
        BtrieveErrorCode = "No Local Access Allowed"
    Case 2006
        BtrieveErrorCode = "No Available SPX Connection"
    Case 2007
        BtrieveErrorCode = "Invalid Parameter"
    Case Else
        BtrieveErrorCode = "Unknown BTRIEVE Error, #" & nStatus
End Select
End Function
