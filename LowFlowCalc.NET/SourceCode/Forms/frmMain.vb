
Public Class frmMain
    Inherits System.Windows.Forms.Form
    Dim m_Filename As String
    Dim m_Dates() As Date
    Dim m_Values() As Double
    Dim m_StartDate As Date
    Dim m_EndDate As Date
    Dim m_Count As Integer
    Dim m_Max As Single
    Dim m_Min As Single
    Dim m_Iterations As Integer = 1000
    'background color
    Dim m_Red As Integer = 128
    Dim m_Green As Integer = 128
    Dim m_Blue As Integer = 255
    Dim m_ColorDir As Integer = 1

    'declare these here so we can capture their values during multiple runs
    Dim Conf5_Boot As Double, Conf95_Boot As Double
    Dim Conf5_Emp As Double, Conf95_Emp As Double
    Dim m_Skew As Double
    Dim m_XT As Double



#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLoadData As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents tlbMain As System.Windows.Forms.ToolBar
    Friend WithEvents btnOpen As System.Windows.Forms.ToolBarButton
    Friend WithEvents ilsMain As System.Windows.Forms.ImageList
    Friend WithEvents btnRun As System.Windows.Forms.ToolBarButton
    Friend WithEvents Separator1 As System.Windows.Forms.ToolBarButton
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents Separator2 As System.Windows.Forms.ToolBarButton
    Friend WithEvents btnConfig As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuParams As System.Windows.Forms.MenuItem
    Friend WithEvents Separator0 As System.Windows.Forms.ToolBarButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents TimeSeriesPlot As AxMSChart20Lib.AxMSChart
    Friend WithEvents mnuHelpWelcome As System.Windows.Forms.MenuItem
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents txtReport As System.Windows.Forms.TextBox
    Friend WithEvents mnuContents As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLP3 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.mnuLoadData = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.mnuExit = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.mnuParams = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.mnuHelpWelcome = New System.Windows.Forms.MenuItem
        Me.mnuContents = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.mnuLP3 = New System.Windows.Forms.MenuItem
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.tlbMain = New System.Windows.Forms.ToolBar
        Me.Separator0 = New System.Windows.Forms.ToolBarButton
        Me.btnOpen = New System.Windows.Forms.ToolBarButton
        Me.Separator1 = New System.Windows.Forms.ToolBarButton
        Me.btnConfig = New System.Windows.Forms.ToolBarButton
        Me.Separator2 = New System.Windows.Forms.ToolBarButton
        Me.btnRun = New System.Windows.Forms.ToolBarButton
        Me.ilsMain = New System.Windows.Forms.ImageList(Me.components)
        Me.lblStatus = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TimeSeriesPlot = New AxMSChart20Lib.AxMSChart
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtReport = New System.Windows.Forms.TextBox
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.GroupBox1.SuspendLayout()
        CType(Me.TimeSeriesPlot, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem5, Me.MenuItem6, Me.mnuLP3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuLoadData, Me.MenuItem3, Me.mnuExit})
        Me.MenuItem1.Text = "File"
        '
        'mnuLoadData
        '
        Me.mnuLoadData.Index = 0
        Me.mnuLoadData.Text = "&Load Data File"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "-"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 2
        Me.mnuExit.Text = "Exit"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 1
        Me.MenuItem5.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuParams})
        Me.MenuItem5.Text = "&Options"
        '
        'mnuParams
        '
        Me.mnuParams.Index = 0
        Me.mnuParams.Text = "&Change Parameters"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 2
        Me.MenuItem6.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuHelpWelcome, Me.mnuContents, Me.MenuItem10, Me.MenuItem11})
        Me.MenuItem6.Text = "&Help"
        '
        'mnuHelpWelcome
        '
        Me.mnuHelpWelcome.Index = 0
        Me.mnuHelpWelcome.Text = "&Welcome Screen"
        '
        'mnuContents
        '
        Me.mnuContents.Index = 1
        Me.mnuContents.Text = "&Contents"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 2
        Me.MenuItem10.Text = "-"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 3
        Me.MenuItem11.Text = "&About Low Flow Calculator"
        '
        'mnuLP3
        '
        Me.mnuLP3.Index = 3
        Me.mnuLP3.Text = "Test"
        '
        'Timer1
        '
        Me.Timer1.Interval = 250
        '
        'Timer2
        '
        Me.Timer2.Interval = 5000
        '
        'tlbMain
        '
        Me.tlbMain.Appearance = System.Windows.Forms.ToolBarAppearance.Flat
        Me.tlbMain.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.Separator0, Me.btnOpen, Me.Separator1, Me.btnConfig, Me.Separator2, Me.btnRun})
        Me.tlbMain.ButtonSize = New System.Drawing.Size(32, 32)
        Me.tlbMain.DropDownArrows = True
        Me.tlbMain.ImageList = Me.ilsMain
        Me.tlbMain.Location = New System.Drawing.Point(0, 0)
        Me.tlbMain.Name = "tlbMain"
        Me.tlbMain.ShowToolTips = True
        Me.tlbMain.Size = New System.Drawing.Size(632, 28)
        Me.tlbMain.TabIndex = 19
        '
        'Separator0
        '
        Me.Separator0.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'btnOpen
        '
        Me.btnOpen.ImageIndex = 4
        Me.btnOpen.Tag = "open"
        Me.btnOpen.ToolTipText = "Load Data File"
        '
        'Separator1
        '
        Me.Separator1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'btnConfig
        '
        Me.btnConfig.ImageIndex = 2
        Me.btnConfig.Tag = "config"
        Me.btnConfig.ToolTipText = "Analysis Parameters"
        '
        'Separator2
        '
        Me.Separator2.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'btnRun
        '
        Me.btnRun.ImageIndex = 3
        Me.btnRun.Tag = "compute"
        Me.btnRun.ToolTipText = "Compute 7Q10"
        '
        'ilsMain
        '
        Me.ilsMain.ImageSize = New System.Drawing.Size(16, 16)
        Me.ilsMain.ImageStream = CType(resources.GetObject("ilsMain.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ilsMain.TransparentColor = System.Drawing.Color.Transparent
        '
        'lblStatus
        '
        Me.lblStatus.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblStatus.Location = New System.Drawing.Point(440, 8)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(176, 16)
        Me.lblStatus.TabIndex = 21
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TimeSeriesPlot)
        Me.GroupBox1.Controls.Add(Me.PictureBox1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.Location = New System.Drawing.Point(0, 28)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(632, 208)
        Me.GroupBox1.TabIndex = 23
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Time Series Plot"
        '
        'TimeSeriesPlot
        '
        Me.TimeSeriesPlot.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TimeSeriesPlot.ContainingControl = Me
        Me.TimeSeriesPlot.DataSource = Nothing
        Me.TimeSeriesPlot.Location = New System.Drawing.Point(13, 18)
        Me.TimeSeriesPlot.Name = "TimeSeriesPlot"
        Me.TimeSeriesPlot.OcxState = CType(resources.GetObject("TimeSeriesPlot.OcxState"), System.Windows.Forms.AxHost.State)
        Me.TimeSeriesPlot.Size = New System.Drawing.Size(603, 174)
        Me.TimeSeriesPlot.TabIndex = 10
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PictureBox1.Location = New System.Drawing.Point(11, 16)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(613, 184)
        Me.PictureBox1.TabIndex = 9
        Me.PictureBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtReport)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 244)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(632, 202)
        Me.GroupBox2.TabIndex = 26
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Report"
        '
        'txtReport
        '
        Me.txtReport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtReport.BackColor = System.Drawing.Color.White
        Me.txtReport.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReport.Location = New System.Drawing.Point(10, 16)
        Me.txtReport.Multiline = True
        Me.txtReport.Name = "txtReport"
        Me.txtReport.ReadOnly = True
        Me.txtReport.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtReport.Size = New System.Drawing.Size(613, 181)
        Me.txtReport.TabIndex = 17
        Me.txtReport.Text = ""
        '
        'Splitter1
        '
        Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Splitter1.Location = New System.Drawing.Point(0, 236)
        Me.Splitter1.MinExtra = 100
        Me.Splitter1.MinSize = 100
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(632, 8)
        Me.Splitter1.TabIndex = 27
        Me.Splitter1.TabStop = False
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(632, 446)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.tlbMain)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.MinimumSize = New System.Drawing.Size(640, 480)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Low Flow Calculator"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.TimeSeriesPlot, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub GetData()
        'show dialog, get file name, plot data and update summary
        Dim OpenDlg As New Windows.Forms.OpenFileDialog()
        Dim FileName As String
        Try
            btnOpen.Enabled = False
            lblStatus.Text = "Select a data file to load."
            OpenDlg.ShowDialog()
            If OpenDlg.FileName = "" Then
                btnOpen.Enabled = True
                lblStatus.Text = ""
                Exit Sub
            End If
            If System.IO.File.Exists(OpenDlg.FileName) = False Then
                MsgBox(OpenDlg.FileName & " is not a valid file name.", MsgBoxStyle.Exclamation, "Low Flow Calculator")
                btnOpen.Enabled = True
                lblStatus.Text = ""
                Exit Sub
            End If
            Me.Refresh()
            SetupData(OpenDlg.FileName)
        Catch ex As System.Exception
            Cursor = System.Windows.Forms.Cursors.Default
            MsgBox("Error opening " & FileName & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Low Flow Calculator")
        End Try

    End Sub

    Public Sub SetupData(ByVal FileName As String)
        Try
            Cursor = System.Windows.Forms.Cursors.WaitCursor
            If LoadData(FileName) = True Then
                Me.Text = "Low Flow Calculator - " & FileName
                m_Filename = FileName
                UpdateSummary()
                lblStatus.Text = "Drawing plot ..."
                lblStatus.Refresh()
                'PlotData()
                btnRun.Enabled = True
            End If
            lblStatus.Text = ""
            Cursor = System.Windows.Forms.Cursors.Default
        Catch ex As System.Exception
            Cursor = System.Windows.Forms.Cursors.Default
            MsgBox("Error opening " & FileName & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Low Flow Calculator")
        End Try
        btnOpen.Enabled = True
    End Sub

    Private Sub Reinitialize()
        Me.Text = "Low Flow Calculator"
        TimeSeriesPlot.RowCount = 0
        TimeSeriesPlot.Refresh()
        txtReport.Text = ""
        txtReport.Refresh()
        Erase m_Dates
        Erase m_Values
    End Sub

    Private Function LoadData(ByVal filename As String) As Boolean
        'Read in the file of data must be either tab delimited or comma delimited, one
        'day per line
        Dim OneLine As String
        Dim a() As String
        Dim i As Integer
        Try
            Reinitialize()
            FileOpen(1, filename, OpenMode.Input)
            i = -1
            ' lblLoading.Visible = True
            OneLine = LineInput(1)
            'if comma delimited then ...
            If InStr(OneLine, ",") > 0 Then
                a = Split(OneLine, ",")
                i += 1
                ReDim Preserve m_Dates(i)
                ReDim Preserve m_Values(i)
                m_Dates(i) = CDate(a(0))
                m_Values(i) = Val(a(1))
                m_Max = m_Values(i)
                m_Min = m_Values(i)
                Do Until EOF(1)
                    OneLine = LineInput(1)
                    a = Split(OneLine, ",")
                    i += 1
                    If i > m_Dates.GetUpperBound(0) Then
                        ReDim Preserve m_Dates(i + 100)
                        ReDim Preserve m_Values(i + 100)
                        lblStatus.Text = "Loading " & i
                        lblStatus.Refresh()
                    End If
                    m_Dates(i) = CDate(a(0))
                    m_Values(i) = Val(a(1))
                    If m_Values(i) > m_Max Then m_Max = m_Values(i)
                    If m_Values(i) < m_Min Then m_Min = m_Values(i)
                Loop
                FileClose(1)
            ElseIf InStr(OneLine, vbTab) > 0 Then
                a = Split(OneLine, vbTab)
                i += 1
                ReDim Preserve m_Dates(i)
                ReDim Preserve m_Values(i)
                m_Dates(i) = CDate(a(0))
                m_Values(i) = Val(a(1))
                m_Max = m_Values(i)
                m_Min = m_Values(i)
                Do Until EOF(1)
                    OneLine = LineInput(1)
                    a = Split(OneLine, vbTab)
                    i += 1
                    If i > m_Dates.GetUpperBound(0) Then
                        ReDim Preserve m_Dates(i + 100)
                        ReDim Preserve m_Values(i + 100)
                        lblStatus.Text = "Loading " & i
                        lblStatus.Refresh()
                    End If
                    m_Dates(i) = CDate(a(0))
                    m_Values(i) = Val(a(1))
                    If m_Values(i) > m_Max Then m_Max = m_Values(i)
                    If m_Values(i) < m_Min Then m_Min = m_Values(i)
                Loop
                FileClose(1)
            Else
                MsgBox("Data must be either tab or comma delimited date and value pairs, one pair per line.", MsgBoxStyle.Exclamation, "Low Flow Calculator")
                FileClose(1)
                Return False
            End If
            m_Count = i + 1
            m_StartDate = m_Dates(0)
            m_EndDate = m_Dates(i)
            ReDim Preserve m_Dates(i)
            ReDim Preserve m_Values(i)
            FileClose(1)
            Return True
        Catch ex As System.Exception
            MsgBox("Failed to load data from " & filename & vbCrLf & "Please ensure that the file contains ASCII formatted data as described in the help file." & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Low Flow Calculator")
            lblStatus.Text = ""
            FileClose(1)
            Return False
        End Try
    End Function

    Private Function PlotData()
        'Draw a simple plot of the data on the mschart
        Dim i As Integer
        Dim numPoints As Integer
        Dim start_time As DateTime
        Dim stop_time As DateTime
        Dim elapsed_time As TimeSpan
        Dim ComputeTime As String

        Try
            btnOpen.Text = "Plotting Data"
            tlbMain.Refresh()
            numPoints = m_Dates.GetUpperBound(0) + 1
            TimeSeriesPlot.RowCount = numPoints
            TimeSeriesPlot.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).ValueScale.Auto = False
            TimeSeriesPlot.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).CategoryScale.DivisionsPerLabel = Int(numPoints / 10)
            TimeSeriesPlot.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).CategoryScale.DivisionsPerTick = Int(numPoints / 10)
            TimeSeriesPlot.RowLabelIndex = 1
            start_time = Now
            'TimeSeriesPlot.Visible = False
            For i = 0 To numPoints - 1
                TimeSeriesPlot.DataGrid.SetData(i + 1, 1, m_Values(i), 0)
                TimeSeriesPlot.DataGrid.RowLabel(i + 1, 1) = m_Dates(i)
                'TimeSeriesPlot.Row = i + 1
                'TimeSeriesPlot.Data = m_Values(i)
                'TimeSeriesPlot.RowLabel = m_Dates(i)
            Next
            'TimeSeriesPlot.Visible = True
            stop_time = Now
            elapsed_time = stop_time.Subtract(start_time)
            ComputeTime = elapsed_time.TotalSeconds.ToString("0.000")
            'MsgBox("drawn in " & ComputeTime & " seconds.")
        Catch ex As System.Exception
            MsgBox("Failed to plot data." & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Low Flow Calculator")
        End Try
        btnOpen.Text = "Load Data File"
    End Function

    Private Sub UpdateSummary()
        'Create a summary of the input data
        Dim a As String
        Dim Transpired As Integer
        a += "Data file: " & vbTab & m_Filename & vbCrLf
        a += "Start date:" & vbTab & m_StartDate & vbCrLf
        a += "End date:  " & vbTab & m_EndDate & vbCrLf
        a += "Count:     " & vbTab & m_Count & vbCrLf
        a += "Maximum:   " & vbTab & m_Max & vbCrLf
        a += "Minimum:   " & vbTab & m_Min & vbCrLf
        Transpired = m_EndDate.Subtract(m_StartDate).Days + 1
        a += "Missing:   " & vbTab & Transpired - m_Count

        txtReport.AppendText("*** Report date and time = " & Now & " ***" & vbCrLf & vbCrLf)
        txtReport.AppendText("Input Data Summary:" & vbCrLf)
        txtReport.AppendText("----------------------------------------" & vbCrLf)
        txtReport.AppendText(a & vbCrLf & vbCrLf)
        txtReport.SelectionStart = Len(txtReport.Text) - 1
        txtReport.ScrollToCaret()


        If m_Count < 365 Then
            MsgBox("This file has less than one year of data which will result in unreliable results.", MsgBoxStyle.Information, "Low Flow Calculator")
        End If
        If Transpired - m_Count > Transpired / 4 Then
            MsgBox("This file is missing more than one fourth of the data for the specified time period.  Results will be computed for available data, but may be unreliable.", MsgBoxStyle.Information, "Low Flow Calculator")
        End If
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'clear out the plot and set the caption
        Dim Params As String
        Dim FileNum As Integer = FreeFile()
        Dim a() As String
        Randomize(Now.Millisecond)

        Try
            TimeSeriesPlot.RowCount = 0
            Me.Text = "Low Flow Calculator"
            btnRun.Enabled = False
            Timer1.Enabled = False
            Randomize()
            FileOpen(FileNum, System.Windows.Forms.Application.StartupPath & "\params.txt", OpenMode.Input)
            Params = LineInput(FileNum)
            FileClose(FileNum)
            a = Split(Params, ",")
            m_Version = a(0)
            m_AvePeriod = a(1)
            m_RetPeriod = a(2)
            m_ExType = a(3)
            m_ExtremeTypeString = Mid(a(4), 2, Len(a(4)) - 2) ' get rid of the quotes

            m_StartMonth = a(5)
            m_EndMonth = a(6)
            If InStr(a(7), "TRUE") > 0 Then
                m_ShowWelcome = True
            Else
                m_ShowWelcome = False
            End If
            If m_ShowWelcome = True Then
                ShowWelcome()
            End If
        Catch ex As System.Exception
            'if something doesn't work here, it is probably no big deal
            MsgBox("The following error occurred when loading the main form: " & ex.Message, MsgBoxStyle.Exclamation, "Low Flow Calculator")

        End Try

        'try this to compute the 7Q10 and conf bounds directly from year lows
        'Dim z(14) As Single
        'a = Split("6.73,6.72,5.41,5.76,5.85,5.76,6.00,5.62,6.53,6.26,5.95,6.80,6.04,6.64,6.28", ",")
        'Dim i As Integer, conf5 As Double, conf95 As Double
        'For i = 0 To 14
        '    z(i) = Val(a(i))
        'Next
        'GetConf(z, 0.1, conf5, conf95)
        'MsgBox(Math.Exp(conf5))
        'MsgBox(Math.Exp(conf95))



    End Sub

    Private Sub mnuLoadData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLoadData.Click
        'get the file and load it
        GetData()
    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        'Exit the program
        Me.Close()
    End Sub

    Private Sub MenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem11.Click
        Dim f As New frmDan()
        f.ShowDialog()
    End Sub

    Private Sub tlbMain_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles tlbMain.ButtonClick
        Select Case e.Button.Tag
            Case "open"
                'go get the data file and load it.
                GetData()
            Case "compute"
                'Run...
                Compute()
            Case "config"
                'show config form
                setParams()
        End Select
    End Sub

    Private Sub Compute()
        'Main entry point for doing the analysis.
        Dim b_Result As Boolean, bResult As Boolean
        Dim i As Integer, j As Integer
        Dim tempDates() As Date, tempValues() As Double
        Dim start_time As DateTime, stop_time As DateTime, elapsed_time As TimeSpan, ComputeTime As String
        Dim doDropAnalysis As Boolean = True 'Temporarily hard coded for the paper. 
        Dim AllYearExtremes() As Double
        'Try
        'Update progress indicator
        lblStatus.Text = "Computing " & m_AvePeriod & "Q" & m_RetPeriod & " " & m_ExtremeTypeString & " flow..."
        lblStatus.Refresh()

        'Initialize timer
        start_time = Now

        'Start report output
        InitialHeaderOutput()

        'Remove dates outside of the desired month range
        RemoveUnusedDates(tempDates, tempValues)

        'Compute extreme for each year
        AllYearExtremes = GetAllYearExtremes(m_ExType, tempDates, tempValues, m_AvePeriod)

        'Get logarithm of the single flow per year
        TakeLogs(AllYearExtremes)

        'Get the correct value for m_Prob
        If m_ExType = ExtremeTypes.LowFlow Then
            m_Prob = 1 / m_RetPeriod
        Else
            m_Prob = 1 - 1 / m_RetPeriod
        End If

        'Run analysis for correct distribution type
        Select Case m_DistType
            Case modMain.DistTypes.LogNormal
                LogNormalAnalysis(AllYearExtremes)
            Case modMain.DistTypes.LogPearsonTypeIII
                'PT3(AllYearExtremes)
                dropAnalysis(AllYearExtremes)
        End Select

        'Compute run time
        stop_time = Now
        elapsed_time = stop_time.Subtract(start_time)
        ComputeTime = elapsed_time.TotalSeconds.ToString("0.000")

        'Finialize report
        txtReport.AppendText(vbCrLf & vbCrLf & "Computation time = " & ComputeTime & " seconds.")

        'move to end of txtreport
        txtReport.SelectionStart = Len(txtReport.Text) - 1
        txtReport.ScrollToCaret()
        lblStatus.Text = ""

        'Catch ex As Exception
        '    txtReport.AppendText(vbCrLf & vbCrLf & "Computation failed.  Please check the input data set to ensure that it has at least 3 years of daily data with no negative values. Also, this may occur if the selected averaging period (in days) is greater than the number of daily observations in a year of the data. Any averaging period greater than 365 will have this problem.")
        '    txtReport.AppendText(vbCrLf & "Error message: " & ex.Message)
        '    txtReport.SelectionStart = Len(txtReport.Text) - 1
        '    txtReport.ScrollToCaret()
        '    lblStatus.Text = ""
        'End Try
    End Sub

    Public Sub LogNormalAnalysis(ByVal X() As Double)
        'Do lognormal analysis.
        Dim StdDev As Double, Mean As Double, Prob As Double
        Dim Result As Double
        Dim Conf95 As Double, Conf5 As Double

        Prob = m_Prob
        Mean = GetMean(X)
        StdDev = GetStdDev(X)
        Result = GetInvNorm(Prob, Mean, StdDev)
        GetConf(X, Prob, Conf5, Conf95)
        Result = Math.Round(Math.Exp(Result), 2)
        Conf5 = Math.Round(Math.Exp(Conf5), 2)
        Conf95 = Math.Round(Math.Exp(Conf95), 2)
        txtReport.AppendText(vbCrLf)
        txtReport.AppendText("Log Normal Distribution Parameters" & vbCrLf)
        txtReport.AppendText("Mean(Log(Q)) is" & vbTab & vbTab & Math.Round(Mean, 2) & vbCrLf)
        txtReport.AppendText("StdDev(Log(Q)) is" & vbTab & vbTab & Math.Round(StdDev, 2) & vbCrLf)
        If Result < 0 Then Result = 0
        If Conf5 < 0 Then Conf5 = 0
        If Conf95 < 0 Then Conf95 = 0
        txtReport.AppendText(vbCrLf & m_AvePeriod & "Q" & m_RetPeriod & " " & m_ExtremeTypeString & " flow = " & Result)
        txtReport.AppendText(vbCrLf & "Bootstrap method estimate of 5% confidence limit = " & Conf5)
        txtReport.AppendText(vbCrLf & "Bootstrap method estimate of 95% confidence limit = " & Conf95)
    End Sub

    Private Sub PT3(ByVal X() As Double)
        'Computes method of moments and maximum likelihood estimates for T year events and standard errors for Pearson type 3
        'distribution
        'N = number of annual maximum events
        'X = series of events

        Dim M1 As Double, M2 As Double, M3 As Double, K As Double
        Dim C1 As Double, C2 As Double, C3 As Double
        'Dim SND(6) As Double
        'Dim XT(6) As Double, SX(6) As Double
        Dim XTLog As Double, SXLog As Double
        Dim XT As Double, SX As Double
        Dim A As Double, B As Double, C As Double, D As Double, StDev As Double
        Dim i As Integer, j As Integer, XN As Integer
        Dim Skew As Double, Beta As Double, Alpha As Double, Gamma As Double
        Dim T As Double, T1 As Double, T2 As Double, T3 As Double, T4 As Double, T5 As Double, T6 As Double, T7 As Double, T8 As Double, T9 As Double, T10 As Double
        Dim Slope As Double, Delta As Double, Prob As Double

        Prob = m_Prob

        XN = X.GetUpperBound(0) + 1
        A = 0.0
        B = 0.0
        C = 0.0
        For i = 0 To XN - 1
            A += X(i)
            B += X(i) ^ 2
            C += X(i) ^ 3
        Next i
        M1 = A / XN
        M2 = (B / XN) - (A / XN) ^ 2
        M3 = (C / XN) + 2 * M1 ^ 3 - 3 * M1 * (B / XN)
        C1 = (Math.Sqrt(XN * (XN - 1))) / (XN - 2)
        C2 = 1.0 + 8.5 / XN
        C3 = XN / (XN - 1)

        'Faster, less accurate skew:
        Skew = M3 / (M2 ^ 1.5)
        Skew = Skew * C1 * C2

        'Slower but more accurate skew
        'StDev = M2 ^ 0.5
        'D = 0
        'For i = 0 To XN - 1
        '    D += ((X(i) - M1) / StDev) ^ 3
        'Next
        'Skew = XN / ((XN - 1) * (XN - 2)) * D
        m_Skew = Skew 'store it globally for capturing later...

        M2 = M2 * C3
        Beta = (2 / Skew) ^ 2
        Alpha = (M2 ^ 0.5) / (Beta ^ 0.5)
        Gamma = M1 - (M2 ^ 0.5) * (Beta ^ 0.5)

        txtReport.AppendText(vbCrLf)
        txtReport.AppendText("Log-Pearson Type III Distribution Parameters Using the Method of Moments" & vbCrLf)
        txtReport.AppendText("Mean is:                 " & Math.Round(M1, 2) & vbCrLf)
        txtReport.AppendText("Variance is:             " & Math.Round(M2, 2) & vbCrLf)
        txtReport.AppendText("Coefficient of Skew is:  " & Math.Round(Skew, 2) & vbCrLf)
        txtReport.AppendText("Parameter Alpha is:      " & Math.Round(Alpha, 2) & vbCrLf)
        txtReport.AppendText("Parameter Beta is:       " & Math.Round(Beta, 2) & vbCrLf)
        txtReport.AppendText("Parameter Gamma is:      " & Math.Round(Gamma, 2) & vbCrLf)

        T = Math.Abs(GetStandardNormalDeviate(Prob))
        T1 = T
        T2 = (T ^ 2 - 1) / 6
        T3 = 2 * (T ^ 3 - 6 * T) / 6 ^ 3
        T4 = (T ^ 2 - 1) / 6 ^ 3
        T5 = T / 6 ^ 4
        T6 = 2 / 6 ^ 6
        K = T1 + T2 * Skew + T3 * Skew ^ 2 - T4 * Skew ^ 3 + T5 * Skew ^ 4 - T6 * Skew ^ 5
        'At this point we have K, M1, and M2, needed for the estimate, XT.
        'The following stuff is used for a standard error estimate, SX.
        Slope = T2 + T3 * 2 * Skew - T4 * 3 * Skew ^ 2 + T5 * 4 * Skew ^ 3 - T6 * 5 * Skew ^ 4
        T7 = (1 + 0.75 * Skew ^ 2) * (0.5 * K ^ 2)
        T8 = K * Skew
        T9 = 6 * (1 + 0.25 * Skew ^ 2) * Slope
        T10 = Slope * (1 + 1.25 * Skew ^ 2) + (Skew * K / 2)
        Delta = 1 + T7 + T8 + T9 * T10
        'In the following lines, keep everything in log, until it is reported.
        If m_ExType = ExtremeTypes.HighFlow Then
            XTLog = M1 + K * Math.Sqrt(M2)
        Else 'for low flow do the other direction
            XTLog = M1 - K * Math.Sqrt(M2)
        End If
        SXLog = Math.Sqrt(M2 * Delta / XN)
        XT = Math.Exp(XTLog)
        m_XT = XT
        Conf5_Emp = Math.Exp(XTLog - GetStandardNormalDeviate(0.05) * SXLog)
        Conf95_Emp = Math.Exp(XTLog + GetStandardNormalDeviate(0.05) * SXLog)

        'now do bootstrapping to get the confidence limits.
        'This function computes the upper and lower confidence limits on the flow estimate using normal distribution.
        Dim StdDev As Double, Mean As Double
        Dim Realization(999) As Double
        Dim ResampledArray() As Double
        Dim Ki As Integer
        ReDim ResampledArray(XN - 1)
        For j = 0 To 999

            'Resample the original data into a new array
            For i = 0 To XN - 1
                Ki = Int(Rnd() * XN)
                ResampledArray(i) = X(Ki)
            Next i
            'now I have a new array to do fun on...
            '
            A = 0.0
            B = 0.0
            C = 0.0
            For i = 0 To XN - 1
                A += ResampledArray(i)
                B += ResampledArray(i) ^ 2
                C += ResampledArray(i) ^ 3
            Next i
            M1 = A / XN
            M2 = (B / XN) - (A / XN) ^ 2
            M3 = (C / XN) + 2 * M1 ^ 3 - 3 * M1 * (B / XN)
            Skew = M3 / (M2 ^ 1.5)
            C1 = (Math.Sqrt(XN * (XN - 1))) / (XN - 2)
            C2 = 1.0 + 8.5 / XN
            C3 = XN / (XN - 1)
            Skew = Skew * C1 * C2
            M2 = M2 * C3
            Beta = (2 / Skew) ^ 2
            Alpha = (M2 ^ 0.5) / (Beta ^ 0.5)
            Gamma = M1 - (M2 ^ 0.5) * (Beta ^ 0.5)
            'T1 - T6 don't change, but skew does, so we have to recompute K now
            K = T1 + T2 * Skew + T3 * Skew ^ 2 - T4 * Skew ^ 3 + T5 * Skew ^ 4 - T6 * Skew ^ 5
            'don't need the stuff used to compute the standard error...
            If m_ExType = ExtremeTypes.HighFlow Then
                Realization(j) = Math.Exp(M1 + K * Math.Sqrt(M2))
            Else 'for low flow do the other direction
                Realization(j) = Math.Exp(M1 - K * Math.Sqrt(M2))
            End If

        Next
        Array.Sort(Realization)
        Conf5_Boot = Realization(50)
        Conf95_Boot = Realization(949)

        'Report results
        txtReport.AppendText(vbCrLf & m_AvePeriod & "Q" & m_RetPeriod & " " & m_ExtremeTypeString & " flow using Log-Pearson Type III Distribution is " & Math.Round(XT, 2) & vbCrLf)
        ' txtReport.AppendText("Standard error on this estimate is " & Math.Round(SX1, 2) & vbCrLf)
        ' txtReport.AppendText(vbCrLf & "Confidence Limits" & vbTab & "Empirical Method" & vbTab & "Bootstrap Method")
        'txtReport.AppendText(vbCrLf & "Lower 5%         " & vbTab & Math.Round(Conf5_Emp, 1) & "    " & vbTab & Math.Round(Conf5_Boot, 1))
        'txtReport.AppendText(vbCrLf & "Upper 95%        " & vbTab & Math.Round(Conf95_Emp, 1) & "    " & vbTab & Math.Round(Conf95_Boot, 1))

    End Sub

    Public Sub dropAnalysis(ByVal InArray() As Double)

        Dim i As Integer, j As Integer, k As Integer, z As Integer
        Dim smallArray() As Double
        Dim N As Integer = UBound(InArray, 1) + 1
        Dim tempArray() As Double
        Dim ConfArray(,) As Double
        Dim tempString As String
        Dim NumSims As Integer = 1
        Dim c(3) As Integer
        Dim p(3) As Double

        ReDim tempArray(UBound(InArray, 1))
        ReDim ConfArray(NumSims * (N - 10) + NumSims - 1, 8)   'create space to run the drop analysis 10 times and store all the results.

        MsgBox(N)

        'Disable the output window
        txtReport.Visible = False

        'Store a copy of the input array.
        For j = 0 To UBound(InArray, 1)
            tempArray(j) = InArray(j)
        Next j
        z = 0   'counter for the master array that collects results of all runs.
        For k = 0 To NumSims - 1 'number of times to run the drop analysis.
            lblStatus.Text = "Running simulation #" & k + 1
            lblStatus.Refresh()
            For i = 0 To N - 10
                txtReport.AppendText("simulation #" & k & vbTab & "drop #" & i)
                ReDim smallArray(N - 1 - i)
                'Reset the input array. It gets damaged in the RandomDrop function.
                For j = 0 To UBound(InArray, 1)
                    InArray(j) = tempArray(j)
                Next j
                RandomDrop(InArray, smallArray, i)
                PT3(smallArray)                    'run the pearson type III analysis. 

                'now log the computed value

                ConfArray(z, 0) = Math.Round(m_XT, 2)
                ConfArray(z, 1) = Math.Round(Conf5_Emp, 2)
                ConfArray(z, 2) = Math.Round(Conf95_Emp, 2)
                ConfArray(z, 3) = Math.Round(Conf5_Boot, 2)
                ConfArray(z, 4) = Math.Round(Conf95_Boot, 2)
                'Boot CI - Emp CI
                ConfArray(z, 5) = Math.Round((Conf95_Boot - Conf5_Boot) - (Conf95_Emp - Conf5_Emp), 2)
                'Qest-7Q10
                ConfArray(z, 6) = Math.Round(m_XT - ConfArray(0, 0), 2)
                ConfArray(z, 7) = i
                ConfArray(z, 8) = Math.Round(m_Skew, 3)
                z += 1
            Next i
        Next k

        FileOpen(1, "C:\Dan\Papers\ASCE 7Q10 Confidence Limits\More Data\output.dat", OpenMode.Output)
        Print(1, "Dropped" & vbTab & "Q   " & vbTab & "5% emp" & vbTab & "95% emp" & vbTab & "5% boot" & vbTab & "95% boot" & vbTab & "boot-emp" & vbTab & "error" & vbTab & "Skew")
        For z = 0 To ConfArray.GetUpperBound(0)
            tempString = vbCrLf & ConfArray(z, 7) & vbTab & ConfArray(z, 0) & vbTab & ConfArray(z, 1) & vbTab & ConfArray(z, 2) & vbTab & ConfArray(z, 3) & vbTab & ConfArray(z, 4) & vbTab & ConfArray(z, 5) & vbTab & ConfArray(z, 6) & vbTab & ConfArray(z, 8)
            Print(1, tempString)
        Next
        FileClose(1)

        'Enable the output window
        txtReport.Text = ""
        txtReport.Visible = True
        txtReport.Refresh()

        'now compute the probability of wider confidence bands using bootstrap given an overestimated 7Q10
        For z = 0 To ConfArray.GetUpperBound(0)
            If ConfArray(z, 6) > 0 Then 'this means we over estimated the 7Q10
                If ConfArray(z, 5) > 0 Then 'this means that the boot ci is greater than the emp ci
                    c(0) += 1
                Else
                    c(1) += 1
                End If
            Else
                If ConfArray(z, 5) > 0 Then
                    c(2) += 1
                Else
                    c(3) += 1
                End If
            End If
        Next z
        p(0) = c(0) / (c(0) + c(1))
        p(1) = c(1) / (c(0) + c(1))
        p(2) = c(2) / (c(2) + c(3))
        p(3) = c(3) / (c(2) + c(3))
        txtReport.AppendText(vbCrLf & "P(Wider Bootstrap | Overestimate) = " & p(0))
        txtReport.AppendText(vbCrLf & "P(Wider Empirical | Overestimate) = " & p(1))
        txtReport.AppendText(vbCrLf & "P(Wider Bootstrap | Underestimate) = " & p(2))
        txtReport.AppendText(vbCrLf & "P(Wider Empirical | Underestimate) = " & p(3))

    End Sub

    Public Function RandomDrop(ByVal InArray() As Double, ByRef OutArray() As Double, ByVal NumberToDrop As Integer)
        'randomly select values from inarray to drop.
        Dim i As Integer, Index As Integer, Temp() As Double
        Dim j As Integer, N As Integer

        txtReport.AppendText(vbCrLf & "Drop indices: ")
        N = UBound(InArray, 1) + 1
        For i = 1 To NumberToDrop
TRYAGAIN:
            Index = Int(Rnd() * (N))
            ' MsgBox("trying to drop index, " & Index)
            If InArray(Index) <> 99999 Then
                InArray(Index) = 99999
                'txtReport.AppendText(Index & " ")
            Else
                GoTo TRYAGAIN
            End If
        Next

        'txtReport.AppendText(vbCrLf & "Censored data: ")
        'For j = 0 To UBound(InArray, 1)
        '    If InArray(j) = 99999 Then
        '        txtReport.AppendText("* ")
        '    Else
        '        txtReport.AppendText(Math.Round(InArray(j), 0) & " ")
        '    End If
        'Next

        ReDim OutArray(N - 1 - NumberToDrop)
        j = 0
        For i = 0 To N - 1
            If InArray(i) <> 99999 Then
                OutArray(j) = InArray(i)
                j = j + 1
            End If
        Next
    End Function

    Public Function TakeLogs(ByRef InArray As Object) As Boolean
        'This function takes a numeric array and returns the logs of all of the data. 
        'If it succeeds, it returns True.
        'changed to use natural logs based on Kite(1988 p.55)
        'math.log is the natural base e logarithm
        Dim i As Integer
        For i = 0 To UBound(InArray)
            InArray(i) = Math.Log(InArray(i))
        Next i
    End Function

    Public Function GetConf(ByVal InArray As Object, ByVal Probability As Double, ByRef Conf5 As Double, ByRef Conf95 As Double)
        'This function computes the upper and lower confidence limits on the flow estimate using normal distribution.
        Dim i As Integer, j As Integer, k As Integer, n As Integer
        Dim StdDev As Double, Mean As Double
        Dim Realization(999) As Double
        Dim ResampledArray() As Double
        n = UBound(InArray) + 1
        ReDim ResampledArray(n - 1)
        For j = 0 To 999
            For i = 0 To n - 1
                k = Int(Rnd() * n)
                ResampledArray(i) = InArray(k)
            Next i
            Mean = GetMean(ResampledArray)
            StdDev = GetStdDev(ResampledArray)
            Realization(j) = GetInvNorm(Probability, Mean, StdDev)
        Next
        Array.Sort(Realization)
        Conf5 = Realization(50)
        Conf95 = Realization(949)
    End Function

    Private Function GetAllYearExtremes(ByVal ExtremeType As modMain.ExtremeTypes, ByVal D As Object, ByVal V As Object, ByVal M As Integer) As Object
        'This function computes the year extremes for the full data set.
        Dim i As Integer, j As Integer, k As Integer
        Dim Y As Integer
        Dim OneYearD() As Date, OneYearV() As Double, AllExtremes() As Double
        Y = Year(D(0))
        j = -1
        k = -1
        For i = 0 To UBound(D)
            'step through each date, D
            If Year(D(i)) = Y Then 'same year
                j = j + 1
                ReDim Preserve OneYearD(j)
                ReDim Preserve OneYearV(j)
                OneYearV(j) = V(i)
                OneYearD(j) = D(i)
            Else
                'just changed years so take the data from the last year and get extreme value
                'make sure there is some data in year data
                k = k + 1
                ReDim Preserve AllExtremes(k)
                AllExtremes(k) = GetYearExtreme(ExtremeType, OneYearD, OneYearV, M)
                'reset the variables for the next year
                Y = Year(D(i))
                txtReport.AppendText(Y - 1 & vbTab & Math.Round(AllExtremes(k), 2) & vbTab & Math.Round(Math.Log(AllExtremes(k)), 2) & vbCrLf)
                i = i - 1
                Erase OneYearV
                j = -1
            End If
        Next
        'done checking data, so do the last year if a last year was filled up.
        k = k + 1
        ReDim Preserve AllExtremes(k)
        AllExtremes(k) = GetYearExtreme(ExtremeType, OneYearD, OneYearV, M)
        Y = Year(D(i - 1))
        txtReport.AppendText(Y & vbTab & Math.Round(AllExtremes(k), 2) & vbTab & Math.Round(Math.Log(AllExtremes(k)), 2) & vbCrLf)
        txtReport.AppendText(vbCrLf)
        Return AllExtremes
    End Function

    Private Function GetYearExtreme(ByVal ExtremeType As modMain.ExtremeTypes, ByVal OneYearD As Object, ByVal OneYearV As Object, ByVal M As Integer) As Double
        'This function computes the year extreme for a single year.
        Dim i As Integer, j As Integer, k As Integer
        Dim AvgV() As Double, Consecutive As Boolean, ExtremeVal As Double
        Dim OneYearDCopy() As Date
        OneYearDCopy = OneYearD
        k = -1
        'get the array of m-day averages
        For i = 0 To UBound(OneYearD) - M
            Consecutive = True
            For j = 1 To M
                If OneYearD(i + j) <> OneYearDCopy(i).AddDays(j) Then 'check for m consecutive dates
                    Consecutive = False
                End If
                'Check for flows of value 0.  Only used for m_ZerosMethod = 1
                If m_ZerosMethod = 1 Then
                    If OneYearV(i + j) = 0 Then
                        Consecutive = False
                        'this way, if there is a zero in any of the coming m days, we just don't compute an average for that period
                    End If
                End If
            Next
            If Consecutive = True Then
                k = k + 1
                ReDim Preserve AvgV(k)
                For j = 1 To M
                    AvgV(k) += OneYearV(i + j)
                Next
                AvgV(k) = AvgV(k) / M
            End If
        Next
        'get the extreme m-day average
        ExtremeVal = AvgV(0)
        For i = 0 To UBound(AvgV)
            If ExtremeType = ExtremeTypes.HighFlow Then
                If AvgV(i) > ExtremeVal Then
                    ExtremeVal = AvgV(i)
                End If
            Else
                If AvgV(i) < ExtremeVal Then
                    If m_ZerosMethod = 2 Then
                        'check to make sure we don't return zero as minimum
                        If AvgV(i) > 0 Then
                            ExtremeVal = AvgV(i)
                        End If
                    Else
                        ExtremeVal = AvgV(i)
                    End If
                End If
            End If
        Next
        Return ExtremeVal
        'AvgV.Sort(AvgV)
        'If ExtremeType = ExtremeTypes.HighFlow Then
        '    Return AvgV(AvgV.GetUpperBound(0))
        'Else
        '    Return AvgV(0)
        'End If
    End Function

    Private Sub mnuParams_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuParams.Click
        setParams()
    End Sub

    Private Sub setParams()
        Dim f As New frmParams
        f.ShowDialog()
        btnRun.ToolTipText = "Compute " & m_AvePeriod & "Q" & m_RetPeriod & " " & m_ExtremeTypeString & " flow"
    End Sub

    Private Sub mnuHelpWelcome_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpWelcome.Click
        ShowWelcome()
    End Sub

    Private Sub ShowWelcome()
        Dim f As New frmWelcome
        f.ShowDialog()

        Me.Refresh()
        If f.ReturnAction = "open" Then
            GetData()
        ElseIf f.ReturnAction > "" Then 'returns the filename of the sample data to open
            SetupData(f.ReturnAction)
        End If
    End Sub

    Private Sub frmMain_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        'On closing, write out the list of parameters
        Dim FileNum As Integer = FreeFile()
        'Try
        FileOpen(FileNum, System.Windows.Forms.Application.StartupPath & "\params.txt", OpenMode.Output)
        Write(FileNum, m_Version, m_AvePeriod, m_RetPeriod, m_ExType, m_ExtremeTypeString, m_StartMonth, m_EndMonth, m_ShowWelcome)
        FileClose(FileNum)
        ' Catch ex As System.Exception
        'MsgBox("The following error occurred when unloading the main form: " & ex.Message, MsgBoxStyle.Exclamation, "Low Flow Calculator")
        'End Try
    End Sub

    Private Sub mnuContents_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuContents.Click
        Dim p As New Process
        p.Start(System.Windows.Forms.Application.StartupPath & "\LowFlowCalc.chm")
    End Sub

    Private Sub mnuLP3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLP3.Click
        txtReport.AppendText(vbCrLf)
        txtReport.AppendText("for P=0.001, Z=" & GetStandardNormalDeviate(0.001) & vbCrLf)
        txtReport.AppendText("for P=0.05, Z=" & GetStandardNormalDeviate(0.05) & vbCrLf)
        txtReport.AppendText("for P=0.10 Z=" & GetStandardNormalDeviate(0.1) & vbCrLf)
        txtReport.AppendText("for P=0.50, Z=" & GetStandardNormalDeviate(0.5) & vbCrLf)
        txtReport.AppendText("for P=0.90, Z=" & GetStandardNormalDeviate(0.9) & vbCrLf)
        txtReport.AppendText("for P=0.999, Z=" & GetStandardNormalDeviate(0.999) & vbCrLf)

    End Sub

    Private Sub InitialHeaderOutput()
        'This is the header output that is first displayed in the report at the beginning of a new run.
        txtReport.AppendText(vbCrLf & vbCrLf & "Results Summary:" & vbCrLf)
        txtReport.AppendText("----------------------------------------" & vbCrLf)
        If m_ZerosMethod = 0 Then txtReport.AppendText("No method selected for handling streamflows of value zero." & vbCrLf)
        If m_ZerosMethod = 1 Then txtReport.AppendText("Streamflows of value zero are dropped before computing " & m_AvePeriod & "-day averages." & vbCrLf)
        If m_ZerosMethod = 2 Then txtReport.AppendText(m_AvePeriod & "-day average flows of value zero are not allowed to be the year average." & vbCrLf)
        txtReport.AppendText("The following " & m_AvePeriod & "-day average " & m_ExtremeTypeString & " flows are based on data from " & MonthName(m_StartMonth) & " through " & MonthName(m_EndMonth) & " each year." & vbCrLf & vbCrLf)
        txtReport.AppendText("Year" & vbTab & "Q" & vbTab & "Log(Q)" & vbCrLf)
        txtReport.AppendText("_____________________________________________________________" & vbCrLf)
    End Sub

    Private Sub RemoveUnusedDates(ByRef tempDates() As Date, ByRef tempValues() As Double)
        'remove the dates that aren't part of the month range we are interested in
        Dim i As Integer, j As Integer
        j = 0
        ReDim tempDates(0)
        ReDim tempValues(0)
        j = -1
        For i = 0 To m_Dates.GetUpperBound(0)
            If Month(m_Dates(i)) <= m_EndMonth And Month(m_Dates(i)) >= m_StartMonth Then
                'include it
                j += 1
                If j > tempDates.GetUpperBound(0) Then
                    ReDim Preserve tempDates(j + 100)
                    ReDim Preserve tempValues(j + 100)
                End If
                tempDates(j) = m_Dates(i)
                tempValues(j) = m_Values(i)
            End If
        Next
        ReDim Preserve tempDates(j)
        ReDim Preserve tempValues(j)
    End Sub
End Class
