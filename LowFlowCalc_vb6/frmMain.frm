VERSION 5.00
Object = "{F8774D51-DD61-11D2-B202-00104B55E536}#19.0#0"; "SevenQ10.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{60D9A323-259E-11D2-BD29-444553540000}#7.0#0"; "shCoolButton.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Low Flow Calculator"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   FillColor       =   &H8000000F&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "test"
      Height          =   615
      Left            =   3480
      TabIndex        =   48
      Top             =   120
      Width           =   495
   End
   Begin shCoolButtons.shCoolButton cmdCompute 
      Height          =   615
      Left            =   2040
      TabIndex        =   45
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      BackColor       =   -2147483633
      BeginProperty Font1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1        =   "Compute"
      Caption2        =   "Low Flow"
      Picture0        =   "frmMain.frx":0E42
      Picture1        =   "frmMain.frx":1294
      Picture2        =   "frmMain.frx":2DE6
      Picture3        =   "frmMain.frx":4938
      ButtonType      =   1
   End
   Begin shCoolButtons.shCoolButton cmdLoad 
      Height          =   615
      Left            =   120
      TabIndex        =   44
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      BackColor       =   -2147483633
      BeginProperty Font1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1        =   "Load"
      Caption2        =   "Data"
      Picture0        =   "frmMain.frx":648A
      Picture1        =   "frmMain.frx":68DC
      Picture2        =   "frmMain.frx":842E
      Picture3        =   "frmMain.frx":9F80
      ButtonType      =   1
   End
   Begin TMDLSevenQ10.SevenQ10 sq10 
      Left            =   7560
      Top             =   4080
      _ExtentX        =   873
      _ExtentY        =   767
   End
   Begin VB.Frame Frame5 
      Height          =   1335
      Left            =   4680
      TabIndex        =   39
      Top             =   3960
      Width           =   3735
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version 1.0.x"
         Height          =   195
         Left            =   360
         TabIndex        =   43
         Top             =   600
         Width           =   915
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2880
         Picture         =   "frmMain.frx":BAD2
         Stretch         =   -1  'True
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Daniel P. Ames"
         Height          =   195
         Left            =   360
         TabIndex        =   42
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Copyright 1999-2002"
         Height          =   195
         Left            =   360
         TabIndex        =   41
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Low Flow Calculator"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   345
         TabIndex        =   40
         Top             =   210
         Width           =   3255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parameters"
      Height          =   2895
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   4455
      Begin VB.OptionButton optFull 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   50
         Top             =   1680
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optFull 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   49
         Top             =   1680
         Width           =   600
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   255
         Left            =   2160
         TabIndex        =   36
         Top             =   600
         Width           =   2175
         Begin VB.OptionButton optFunct 
            Caption         =   "Lognormal"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   38
            Top             =   0
            Width           =   1320
         End
         Begin VB.OptionButton optFunct 
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.TextBox txtAve 
         Height          =   285
         Left            =   3240
         TabIndex        =   32
         Text            =   "7"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtRet 
         Height          =   285
         Left            =   3240
         TabIndex        =   31
         Text            =   "10"
         Top             =   2400
         Width           =   975
      End
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Index           =   1
         ItemData        =   "frmMain.frx":C914
         Left            =   2760
         List            =   "frmMain.frx":C93F
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":C9A6
         Left            =   2760
         List            =   "frmMain.frx":C9D1
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optBoot 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   27
         Top             =   240
         Width           =   600
      End
      Begin VB.OptionButton optBoot 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Probability function:"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Averaging period (days):"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   34
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Return period (years):"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   33
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Only use full periods:"
         Height          =   195
         Left            =   360
         TabIndex        =   30
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Use bootstrapping:"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Starting month:"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Ending month:"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Set"
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   8295
      Begin VB.TextBox txtEnd 
         Height          =   285
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtStart 
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtMean 
         Height          =   285
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtMin 
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtMax 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtCount 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtFilename 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label13 
         Caption         =   "End date:"
         Height          =   255
         Left            =   6240
         TabIndex        =   25
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Start date:"
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Mean:"
         Height          =   255
         Left            =   6240
         TabIndex        =   17
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Minimum value:"
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Maximum value:"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "Number of data points:"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "File name:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Results"
      Height          =   1455
      Left            =   4680
      TabIndex        =   0
      Top             =   2400
      Width           =   3735
      Begin VB.TextBox txtConf95 
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   795
      End
      Begin VB.TextBox txtConf5 
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   795
      End
      Begin VB.TextBox txtExpect 
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "95th percentile confidence limit:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "5th percentile confidence limit:"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Expected flow:"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   2000
      End
   End
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   480
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin shCoolButtons.shCoolButton cmdHelp 
      Height          =   615
      Left            =   6600
      TabIndex        =   46
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      BackColor       =   -2147483633
      BeginProperty Font1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1        =   "Help/"
      Caption2        =   "Information"
      Picture0        =   "frmMain.frx":CA38
      Picture1        =   "frmMain.frx":CD52
      Picture2        =   "frmMain.frx":E8A4
      Picture3        =   "frmMain.frx":103F6
      ButtonType      =   1
   End
   Begin VB.Label lblRunning 
      Caption         =   "Running... This may take several minutes..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4080
      TabIndex        =   47
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'dpa 7/17/02
'This is the main form that is the GUI for the LowFlowCalculator.

Option Explicit
Private d() As Date, v() As Single
Private DataLoaded As Boolean


Private Sub cmdCompute_Click()
    If DataLoaded = True Then
        lblRunning.Visible = True
        lblRunning.Refresh
        ComputeFlow
        lblRunning.Visible = False
    End If
End Sub

Private Sub ComputeFlow()
    'dpa 7/17/02
    'sets up the inputs and parameters for the 7q10 calculator,
    'runs it and displays outputs
    Dim s As Long, t As Single, result As Single, i As Long
    Dim Tries As Integer
    sq10.Clear
    If optFunct(0).Value = True Then
        For i = 0 To UBound(d)
            sq10.AddPoint d(i), v(i)
        Next i
    Else
        For i = 0 To UBound(d)
            sq10.AddPoint d(i), Log(v(i))
        Next i
    End If

    frmMain.MousePointer = 11
    
    With sq10
        If optBoot(0).Value = True Then
            .Bootstrap = True
        Else
            .Bootstrap = False
        End If
        .MonthStart = cmbMonth(0).ItemData(cmbMonth(0).ListIndex)
        .MonthEnd = cmbMonth(1).ItemData(cmbMonth(1).ListIndex)
        '.StartYear = cmbYear.List(cmbYear.ListIndex)
        s = txtAve.Text
        t = txtRet.Text
        Tries = 0
TRYAGAIN:
        Tries = Tries + 1
        If Tries > 10 Then
            MsgBox "Can not converge on a valid solution.", vbExclamation + vbOKOnly, "No solution"
            GoTo ERRORHANDLER
        End If
        result = sq10.Get7Q10(s, t)
        If optFunct(0).Value = True Then
            If result > Val(txtMean.Text) Then GoTo TRYAGAIN
            txtExpect.Text = Round(result, 2)
            txtConf5.Text = Round(sq10.Conf5, 2)
            txtConf95.Text = Round(sq10.Conf95, 2)
        Else
            If result > 10 Then GoTo TRYAGAIN
            If 10 ^ result > Val(txtMean) Then GoTo TRYAGAIN
            txtExpect.Text = Round(10 ^ result, 2)
            txtConf5.Text = Round(10 ^ sq10.Conf5, 2)
            txtConf95.Text = Round(10 ^ sq10.Conf95, 2)
        End If
        If optBoot(1).Value = True Then
            txtConf5.Text = ""
            txtConf95.Text = ""
        End If
    End With
ERRORHANDLER:
    frmMain.MousePointer = 0
    
End Sub


Private Sub cmdHelp_Click()
    'dpa 7/17/02
    'show a help file with information about this thing
    Dim HelpFile As String
    HelpFile = App.Path & "\help\LowFlowCalc.chm"
    If OpenTools.OpenFileByExtentsion(HelpFile) = False Then
        MsgBox "Failed to open Help File " & HelpFile, vbCritical + vbOKOnly, "Help file not found"
    End If

End Sub

Private Sub cmdLoad_Click()
    'dpa 7/17/02
    'handles the dialog box display for loading a data file
    Dim FN As String
    cdlMain.CancelError = True
    cdlMain.DialogTitle = "Load Streamflow Data"
    On Error GoTo ERRORHANDLER
    cdlMain.ShowOpen
    On Error GoTo 0
    FN = cdlMain.FileName
    frmMain.Refresh
    LoadData FN
ERRORHANDLER:
    
End Sub
Private Sub LoadData(FileName As String)
    'dpa 7/17/02
    'actually loads the data from the selected file if possible
    Dim L As String, a() As String, i As Long
    Dim vMax As Single, vMin As Single, vMean As Single, vCount As Long
    On Error GoTo ERRORHANDLER
    frmMain.MousePointer = 11
   ' cmbYear.Clear
    txtExpect.Text = ""
    txtConf5.Text = ""
    txtConf95 = ""
    
    Open FileName For Input As #1
    txtFilename.Text = FileName
    i = -1
    Do Until EOF(1)
        i = i + 1
        If i > 31999 Then
            MsgBox "This version of Low Flow Calculator only accepts time series with up to 32,000 points.  The first 32,000 points of the selected data set will be used.", vbInformation + vbOKOnly, "Low Flow Calculator"
            Exit Do
        End If
        Line Input #1, L
        a = Split(L, vbTab)
        ReDim Preserve d(i)
        ReDim Preserve v(i)
        d(i) = CDate(a(0))
        v(i) = CSng(a(1))
    Loop
    Close #1
    vMax = v(0)
    vMin = v(0)
    vMean = 0
    vCount = UBound(d) + 1
   ' cmbYear.AddItem Year(d(0))
    For i = 0 To vCount - 1
        If v(i) > vMax Then vMax = v(i)
        If v(i) < vMin Then vMin = v(i)
        vMean = vMean + v(i)
'        If Year(d(i)) > cmbYear.List(cmbYear.ListCount - 1) Then
'            cmbYear.AddItem Year(d(i))
'        End If
    Next i
    vMean = vMean / vCount
    txtMean.Text = Round(vMean, 2)
    txtCount.Text = vCount
    txtMin.Text = vMin
    txtMax.Text = vMax
    txtStart.Text = d(0)
    txtEnd.Text = d(vCount - 1)
    'cmbYear.ListIndex = 0
    cmbMonth(1).ListIndex = 11
    DataLoaded = True
    frmMain.MousePointer = 0
    Exit Sub
    
ERRORHANDLER:
    Close #1
    frmMain.MousePointer = 0
    txtFilename.Text = ""
    ErrorTime
End Sub
Private Sub ErrorTime()
    'dpa 7/17/02
    'shows an error message if the data they tried to input is in the wrong format
    MsgBox "The input data is in the wrong format.  It must be a tab delimited text file with dates in the first column and values in the second column.", vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdTest_Click()
    If DataLoaded = True Then
        lblRunning.Visible = True
        lblRunning.Refresh
        ComputeFlow2
        lblRunning.Visible = False
    End If
End Sub

Private Sub Form_Load()
    'dpa 7/17/02
    'Sets up the gui and randomizes
    Dim i As Integer
    Randomize
    cmbMonth(0).ListIndex = 0
    cmbMonth(1).ListIndex = 0
    lblTitle.Caption = "Low Flow Calculator"
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblRunning.Visible = False
    lblRunning.Caption = "Running... This may take several minutes..."
    DataLoaded = False
End Sub

Private Sub optBoot_Click(Index As Integer)
    'dpa 7/17/02
    'enables and disables the confidence limits text boxes if they choose bootstrap
    If optBoot(0).Value = True Then
        txtConf5.Enabled = True
        txtConf95.Enabled = True
        txtConf5.BackColor = vbWhite
        txtConf95.BackColor = vbWhite
    Else
        txtConf5.Enabled = False
        txtConf95.Enabled = False
        txtConf5.BackColor = &H8000000F
        txtConf95.BackColor = &H8000000F
    End If
End Sub

Private Sub txtAve_Validate(Cancel As Boolean)
    'dpa 7/17/02
    'assures that the input data is correct format
    If Val(txtAve.Text) <= 0 Then
        MsgBox "Plese only enter positive integer values.", vbCritical + vbOKOnly, "Error"
        Cancel = True
    Else
        txtAve.Text = Int(Val(txtAve.Text))
    End If
End Sub
Private Sub txtRet_Validate(Cancel As Boolean)
    'dpa 7/17/02
    'assures that the input data is correct format
    If Val(txtRet.Text) <= 0 Then
        MsgBox "Plese only enter positive integer values.", vbCritical + vbOKOnly, "Error"
        Cancel = True
    Else
        txtRet.Text = Int(Val(txtRet.Text))
    End If
End Sub

Private Sub ComputeFlow2()
    'dpa 9/24/02
    'sets up the inputs and parameters for the new 7q10 calculator,
    'runs it and displays outputs
    Dim s As Integer, t As Integer, result As Single, i As Long
    Dim Tries As Integer, FullPeriod As Boolean
    Dim LowFlow As New clsLowFlow
    LowFlow.Clear
    If optFunct(0).Value = True Then
        LowFlow.LoadData v, d(0)
'    Else
'        For i = 0 To UBound(d)
'            sq10.AddPoint d(i), Log(v(i))
'        Next i
    End If

    frmMain.MousePointer = 11
    
    With LowFlow
'        If optBoot(0).Value = True Then
'            .Bootstrap = True
'        Else
'            .Bootstrap = False
'        End If
        .StartMonth = cmbMonth(0).ItemData(cmbMonth(0).ListIndex)
        .EndMonth = cmbMonth(1).ItemData(cmbMonth(1).ListIndex)
        If optFull(0).Value = True Then
            FullPeriod = True
        Else
            FullPeriod = False
        End If
        '.StartYear = cmbYear.List(cmbYear.ListIndex)
        s = txtAve.Text
        t = txtRet.Text
        Tries = 0
TRYAGAIN:
        Tries = Tries + 1
        If Tries > 10 Then
            MsgBox "Can not converge on a valid solution.", vbExclamation + vbOKOnly, "No solution"
            GoTo ERRORHANDLER
        End If
        result = LowFlow.Compute(s, t, FullPeriod)
'        If optFunct(0).Value = True Then
'            If result > Val(txtMean.Text) Then GoTo TRYAGAIN
'            txtExpect.Text = Round(result, 2)
'            txtConf5.Text = Round(sq10.Conf5, 2)
'            txtConf95.Text = Round(sq10.Conf95, 2)
'        Else
'            If result > 10 Then GoTo TRYAGAIN
'            If 10 ^ result > Val(txtMean) Then GoTo TRYAGAIN
'            txtExpect.Text = Round(10 ^ result, 2)
'            txtConf5.Text = Round(10 ^ sq10.Conf5, 2)
'            txtConf95.Text = Round(10 ^ sq10.Conf95, 2)
'        End If
'        If optBoot(1).Value = True Then
'            txtConf5.Text = ""
'            txtConf95.Text = ""
'        End If
    End With
ERRORHANDLER:
    frmMain.MousePointer = 0
    
End Sub

