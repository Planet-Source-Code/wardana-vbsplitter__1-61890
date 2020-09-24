VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   Caption         =   "VBSplitter"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4320
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "New"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   5400
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   720
      TabIndex        =   15
      Top             =   2040
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Number of split"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   5400
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Output :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   4095
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3360
         Top             =   1920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Location"
         Height          =   375
         Left            =   3120
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2895
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00C0FFFF&
         Height          =   1620
         ItemData        =   "Form1.frx":1CCA
         Left            =   120
         List            =   "Form1.frx":1CCC
         TabIndex        =   10
         Top             =   960
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Input :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton Command2 
         Caption         =   "GO"
         Height          =   615
         Left            =   3120
         TabIndex        =   7
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "Form1.frx":1CCE
         Left            =   1680
         List            =   "Form1.frx":1CE4
         TabIndex        =   6
         Text            =   "100"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Location"
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   120
         ScrollBars      =   1  'Horizontal
         TabIndex        =   4
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "3"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Size of split"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "File Name:"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "0%"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************
'*                                                     *
'* Created by Wardana from Indonesia, in Bogor,        *
'* July,25 2005                                        *
'*                                                     *
'*******************************************************

'*******************************************************
'*                                                     *
'* This code is from Makino who created FileSplitter,  *
'* and I'm interesting to his so  I added some featur- *
'* es to the code, e.g. we can choose the splitting    *
'* method depend on the number of file-splits,  and    *
'* there is a batch file to join the splits            *
'*                                                     *
'*******************************************************

Option Explicit
Const GENERIC_WRITE = &H40000000
Const GENERIC_READ = &H80000000
Const FILE_ATTRIBUTE_NORMAL = &H80
Const CREATE_ALWAYS = 2
Const OPEN_ALWAYS = 4
Const INVALID_HANDLE_VALUE = -1
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long

Public Function SplitFiles(ByVal inputFilename As String, newFileSizeBytes As Long) As Boolean
    On Error Resume Next
    Dim fReadHandle As Long
    Dim fWriteHandle As Long
    Dim fSuccess As Long
    Dim lBytesWritten As Long
    Dim lBytesRead As Long
    Dim ReadBuffer() As Byte
    Dim TotalCount As Long
    Dim Count As Integer
    Dim tambah As String
    Dim ukur As String
    Dim ukur2 As String
    Dim lok1 As String
    Dim ukur3 As String
    
    Me.MousePointer = vbHourglass
    Command1.Enabled = False
    ProgressBar1.Max = FileLen(inputFilename)
    ProgressBar1.Value = 0
    Count = 1
    ReDim ReadBuffer(0 To newFileSizeBytes)
    TotalCount = (FileLen(inputFilename) \ UBound(ReadBuffer)) + 1
    fReadHandle = CreateFile(inputFilename, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    
    If fReadHandle <> INVALID_HANDLE_VALUE Then
        fSuccess = ReadFile(fReadHandle, ReadBuffer(0), UBound(ReadBuffer), lBytesRead, 0)
        ProgressBar1.Value = ProgressBar1.Value + lBytesRead
        ProgressBar1.Refresh
        tambah = ""
        Do While lBytesRead > 0
            If Dir(inputFilename & "." & Count) <> "" Then
                Kill inputFilename & "." & Count
            End If
            fWriteHandle = CreateFile(inputFilename & "." & Count, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
            If fWriteHandle <> INVALID_HANDLE_VALUE Then
                fSuccess = WriteFile(fWriteHandle, ReadBuffer(0), lBytesRead, lBytesWritten, 0)
                If fSuccess <> 0 Then
                    fSuccess = FlushFileBuffers(fWriteHandle)
                    fSuccess = CloseHandle(fWriteHandle)
                Else
                    MsgBox ("There is a mistake")
                    SplitFiles = False
                    Exit Function
                End If
            Else
                MsgBox ("There is a mistake")
                SplitFiles = False
                Exit Function
            End If
            fSuccess = ReadFile(fReadHandle, ReadBuffer(0), UBound(ReadBuffer), lBytesRead, 0)
            ProgressBar1.Value = ProgressBar1.Value + lBytesRead
            ProgressBar1.Refresh
            Label2.Caption = Left(Str(ProgressBar1.Value / ProgressBar1.Max * 100), 6) & "%"
            Label2.Refresh
            Count = Count + 1
            List1.AddItem CommonDialog1.FileTitle & "." & Count - 1
        tambah = tambah & Chr(34) & CommonDialog1.FileTitle & "." & Count - 1 & Chr(34) & " + "
        If Text2.Text <> "" Then
            FileCopy CommonDialog1.FileName & "." & Count - 1, Text2.Text & "\" & CommonDialog1.FileTitle & "." & Count - 1
            Kill CommonDialog1.FileName & "." & Count - 1
        End If
        Loop
        fSuccess = CloseHandle(fReadHandle)
    Else
        MsgBox ("There is a mistake")
        SplitFiles = False
        Exit Function
    End If
    tambah = "copy /B " & tambah
    ukur = Len(tambah)
    If Text2.Text <> "" Then
        ukur2 = Text2.Text & "\"
    Else
        ukur2 = Left(Text1.Text, Len(Text1.Text) - Len(CommonDialog1.FileTitle))
    End If
    ukur = Left(tambah, Val(ukur - 3))
    ukur3 = Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4)
    
    'Make a batch file to join any splits of file
    Open ukur2 & "Join_" & ukur3 & ".bat" For Output As 1
        Print #1, "@Echo off"
        Print #1, "Echo."
        Print #1, "Echo Making by VBSplitter on : " & Date
        Print #1, "Echo ------------------------------------------------------"
        Print #1, "Echo."
        Print #1, "Echo To join all of this splits of " & CommonDialog1.FileTitle
        Print #1, "Echo."
        Print #1, "Echo Press any key to process"
        Print #1, "Pause>nul"
        Print #1, ukur & " " & Chr(34) & CommonDialog1.FileTitle & Chr(34)
    Close #1
    Text2.Text = "Same location with the origin file"
    'Show information by List1 about the batch file
    List1.AddItem "Join_" & ukur3 & ".bat"
    List1.AddItem ""
    List1.AddItem "Use Join_" & ukur3 & ".bat to make an origin file"

    ProgressBar1.Value = ProgressBar1.Max
    Call txtFileName_Change
    Me.MousePointer = vbDefault
    SplitFiles = True
    Command1.Enabled = True
End Function

Private Sub Command1_Click()
    On Error Resume Next

    Dim lok1 As String
    CommonDialog1.ShowOpen
    lok1 = CommonDialog1.FileName
    Text1.Text = lok1
    Label1.Caption = "File Name : " & CommonDialog1.FileTitle
    List1.Clear
End Sub

Private Sub txtFileName_Change()
    If Len(Text1.Text) = 0 Then
        Command1.Enabled = True
    Else
        Command1.Enabled = False
    End If
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Dim ukur As String
    Command1.Enabled = True
    If Option1(1).Value = True Then
        If Text3.Text <> "" Then Text3.Text = "3"
        ukur = FileLen(Text1.Text) / (Val(Text3.Text) * 1024)
        Combo1.Enabled = False
    Else
        Combo1.Enabled = True
        ukur = Combo1.Text
    End If
    List1.Clear
    If Text1.Text <> "" And Dir(Text1.Text) <> "" Then
        Call SplitFiles(Text1.Text, CLng(ukur) * CLng(1024))
    Else
        Call txtFileName_Change
        Call MsgBox("Can't find the file", vbOKOnly + vbCritical, "Salah membaca file")
    End If
End Sub

Private Sub Command3_Click()
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat("Choose the folder", "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    Text2.Text = sPath
End Sub

Private Sub Command4_Click()
    End
End Sub

Private Sub Command5_Click()
    Text1.Text = ""
    Text2.Text = ""
    List1.Clear
    Label1.Caption = "File Name:"
    Label2.Caption = "0%"
    ProgressBar1.Value = 0
End Sub

Private Sub Option1_Click(Index As Integer)
    On Error Resume Next
    Command1.Enabled = True
    If Index = 0 Then
        Option1(1).Value = False
        Option1(0).Value = True
        Text3.Enabled = False
        Combo1.Enabled = True
    Else
        Option1(1).Value = True
        Option1(0).Value = False
        Text3.Enabled = True
        Combo1.Enabled = False
    End If
End Sub


