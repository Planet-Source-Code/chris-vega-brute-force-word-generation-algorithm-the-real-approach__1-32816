VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmBrute 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Brute Force Cracking, An Algorithm!"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   Icon            =   "frmBrute.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComCtl2.UpDown ud 
      Height          =   285
      Left            =   1755
      TabIndex        =   13
      Top             =   720
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   327681
      Value           =   1
      BuddyControl    =   "txtLen"
      BuddyDispid     =   196609
      OrigLeft        =   1665
      OrigTop         =   720
      OrigRight       =   1905
      OrigBottom      =   1005
      Max             =   255
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtLen 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   720
      Width           =   465
   End
   Begin VB.Timer tmrx 
      Interval        =   1000
      Left            =   4320
      Top             =   405
   End
   Begin VB.TextBox txtBruteWord 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   32
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   405
      Width           =   3480
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   285
      Left            =   2385
      TabIndex        =   8
      Top             =   2835
      Width           =   2085
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "&Start"
      Height          =   285
      Left            =   270
      TabIndex        =   7
      Top             =   2835
      Width           =   2085
   End
   Begin VB.TextBox txtCustomSet 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   405
      MaxLength       =   255
      TabIndex        =   6
      Text            =   "`~!@#$%^&*()_-=+"
      Top             =   2385
      Width           =   4335
   End
   Begin VB.OptionButton opts 
      Caption         =   "Custom Character Set"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   135
      TabIndex        =   5
      Top             =   2115
      Width           =   4560
   End
   Begin VB.OptionButton opts 
      Caption         =   "Use Numbers and Letters (a-z, A-Z, 0-9)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   135
      TabIndex        =   4
      Top             =   1800
      Value           =   -1  'True
      Width           =   4560
   End
   Begin VB.OptionButton opts 
      Caption         =   "Use Numbers Only (0-9)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   135
      TabIndex        =   3
      Top             =   1485
      Width           =   4560
   End
   Begin VB.OptionButton opts 
      Caption         =   "Use Letters Only (a-z, A-Z)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   135
      TabIndex        =   2
      Top             =   1170
      Width           =   4560
   End
   Begin VB.TextBox txtWordToCrack 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1260
      MaxLength       =   255
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   90
      Width           =   3480
   End
   Begin VB.Label lblPercent 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4605
      TabIndex        =   14
      Top             =   765
      Width           =   45
   End
   Begin VB.Label lbls 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Length"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   11
      Top             =   765
      Width           =   1080
   End
   Begin VB.Label lbls 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brute Word"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   10
      Top             =   450
      Width           =   810
   End
   Begin VB.Label lbls 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Word to Crack"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   1035
   End
End
Attribute VB_Name = "frmBrute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' BruteForce Cracking Algorithm
' Presented by : gewdlooking crackers of philippines
'
' Original BruteForce Word Generation Routine
' Created  : February 1998 as a Win32 Console Application
' Language : Win32 Assembly (TASM32)
' Author   : Chris Vega [gwapo@models.com]
'
' Recreated Using Visual Basic by:
'       Chris Vega and Destro Ex [webmaster@win32asm.8m.com]
'    +- Added Counters
'    +- Added User Interface
'    +- Added Cancel Procedures
'    +- Removed Dictionary File Attack Support
'    +- Quick Exit Animation
'
' Copyright 1998-2001 by Chris Vega [gwapo@models.com]
' No Rights Reserved, Use Without Permission

Private charSet As String
Private IsStarting As Boolean
Private lCustom As String

Private pChar() As Long
Private pCount As Long

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdStartStop_Click()
    Dim cx_ As Control
    IsStarting = Not IsStarting
    If IsStarting Then
        cmdStartStop.Caption = "&Stop"
        For Each cx_ In Controls
            cx_.Enabled = False
        Next
        cmdStartStop.Enabled = True
        txtBruteWord.Enabled = True
        lblPercent.Enabled = True
        doBruteForce
    Else
        cmdStartStop.Caption = "&Start"
        For Each cx_ In Controls
            cx_.Enabled = True
        Next
        lblPercent.Enabled = False
        lblPercent.Caption = ""
        txtBruteWord.Enabled = False
        If Not opts(3).Value Then _
           txtCustomSet.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Show
    opts_Click 2
    IsStarting = False
    txtWordToCrack = "win32"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsStarting Then
        Cancel = True
    Else
        Dim iw As Long
        MsgBox Caption & _
               vbCrLf & vbCrLf & _
               LoadResString(777) & _
               vbCrLf & vbCrLf & _
               LoadResString(888), _
               vbInformation, _
               LoadResString(666)
        While Height > 405
            Height = Height - 25
        Wend
        While Width > 870
            Width = Width - 25
        Wend
        While Left > (Width * -1)
            Left = Left - 15
        Wend
        End
    End If
End Sub

Private Sub opts_Click(Index As Integer)
    Select Case Index
        Case 0
            charSet = LoadResString(101) & _
                      LoadResString(102)
            txtCustomSet.Enabled = False
        Case 1
            charSet = LoadResString(103)
            txtCustomSet.Enabled = False
        Case 2
            charSet = LoadResString(101) & _
                      LoadResString(102) & _
                      LoadResString(103)
            txtCustomSet.Enabled = False
        Case 3
            charSet = gset(txtCustomSet)
            With txtCustomSet
                .Enabled = True
                .SelStart = 0
                .SelLength = 255
                .SetFocus
            End With
    End Select
End Sub

Private Function gset(txt_) As String
    Dim xChar As String * 1
    
    gset = ""
    
    While txt_ & "?" <> "?"
        xChar = Left(txt_, 1)
        gset = gset & xChar
        txt_ = Replace(txt_, _
                       xChar, _
                       Space(0))
    Wend
End Function

Public Sub doBruteForce()
    Dim cSet() As Byte
    Dim cCount As Long
    
    Dim i As Long, i_ As Byte
    Dim j As Double, k As Double
    Dim xStr As String
    
    If charSet & "?" = "?" Then
        ' =====================================================
        ' On Custom Set = NULL, Halt Brute Forcing!
        ' =====================================================
        MsgBox LoadResString(104), _
               vbExclamation, _
               LoadResString(666)
        cmdStartStop_Click
    Else
        ' =====================================================
        ' Build Character Set Array (Sorted)
        ' =====================================================
        For cCount = 1 To Len(charSet)
            ReDim Preserve cSet(cCount - 1) As Byte
            cSet(cCount - 1) = Asc(Mid(charSet, cCount, 1))
        Next
        cCount = cCount - 1
        For i = 0 To UBound(cSet) - 1
            For j = i To UBound(cSet)
                If cSet(i) > cSet(j) Then
                    i_ = cSet(i)
                    cSet(i) = cSet(j)
                    cSet(j) = i_
                End If
            Next
        Next
        
        ' =====================================================
        ' Initialize BruteForcing as:
        '    -+ Starting Counter
        '    -+ Total Number of Combinations
        '    -+ Others
        ' =====================================================
        txtBruteWord = ""
        k = 1
        For i = 1 To Len(txtWordToCrack)
            k = k * Len(charSet)
        Next
        j = 1
        For i = 1 To ud.Value - 1
            j = j * Len(charSet)
        Next
        
        pCount = ud.Value - 1
        ReDim pChar(pCount) As Long
        For i = 0 To pCount: pChar(i) = 0: Next
        If opts(3).Value Then
            charSet = ""
            For i = 0 To UBound(cSet)
                charSet = charSet & Chr(cSet(i))
            Next
            txtCustomSet = charSet
        End If
        
        ' =====================================================
        ' Do the BruteForcing!
        ' =====================================================
        Do While True
            If IsStarting Then
                ' =====================================================
                ' Test Current Generated String
                ' =====================================================
                xStr = ""
                For i = 0 To pCount
                    xStr = xStr & _
                           Chr(cSet(pChar(i)))
                Next
                txtBruteWord = xStr
                ' =====================================================
                ' If it equals then trigger success and exit
                ' =====================================================
                If txtBruteWord = txtWordToCrack Then
                    MsgBox LoadResString(999) & xStr, _
                           vbInformation, _
                           LoadResString(666)
                    cmdStartStop_Click
                    Exit Sub
                End If
                DoEvents
                ' =====================================================
                ' Generate next String with accordance to maximus
                ' =====================================================
                IncPointers cCount, _
                            Len(txtWordToCrack) - 1
            Else
                ' =====================================================
                ' If Cancelled by User then Exit
                ' =====================================================
                MsgBox LoadResString(997), _
                       vbExclamation, _
                       LoadResString(666)
                Exit Sub
            End If

            ' =====================================================
            ' Progress Counters and value-maximus check-out
            ' =====================================================
            If pChar(0) = cCount Then Exit Do
            j = j + 1
            lblPercent.Caption = "Word : " & _
                                 FormatNumber(j, 0) & _
                                 "/" & _
                                 FormatNumber(k, 0)
        Loop

        ' =====================================================
        ' If value-maximus has been reached and no matched
        ' found then blame user for not giving enough
        ' character set (lolz)
        ' =====================================================
        MsgBox LoadResString(998), _
               vbExclamation, _
               LoadResString(666)
        cmdStartStop_Click
    End If
End Sub

Private Sub IncPointers(maxIn_ As Long, _
                        last_P As Long)
    Dim i As Long
    ' =====================================================
    ' Build a dynamic If Statement to test length-In
    ' digitizing maximal, array bounding
    ' =====================================================
    For i = UBound(pChar) To 0 Step -1
        pChar(i) = pChar(i) + 1
        If ((pChar(i) >= maxIn_) And _
            (i - 1 >= 0)) Then _
            pChar(i) = 0 _
        Else _
            Exit For
    Next
    
    ' =====================================================
    ' If not yet digitized maximus increase length of
    ' BruteForced Word Length and Reset all to 1^nth
    ' =====================================================
    If ((pChar(0) >= maxIn_) And _
        (pCount < last_P)) Then
        pCount = pCount + 1
        ReDim pChar(pCount) As Long
        For i = 0 To pCount: pChar(i) = 0: Next
    End If
End Sub

Private Sub txtCustomSet_Change()
    charSet = gset(txtCustomSet)
End Sub

Private Sub txtLen_GotFocus()
    ud.SetFocus
End Sub

Private Sub txtWordToCrack_Change()
    ud.Max = Len(txtWordToCrack)
    If ud.Max = 0 Then ud.Min = 0 Else ud.Min = 1
End Sub
