Public Class frmWelcome
    Inherits System.Windows.Forms.Form
    Public ReturnAction As String

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
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents chkShowWelcome As System.Windows.Forms.CheckBox
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents picGlobe As System.Windows.Forms.PictureBox
    Friend WithEvents picGlobeSpin As System.Windows.Forms.PictureBox
    Friend WithEvents picGlobeStill As System.Windows.Forms.PictureBox
    Friend WithEvents picFolder1 As System.Windows.Forms.PictureBox
    Friend WithEvents picFolderStill As System.Windows.Forms.PictureBox
    Friend WithEvents picFolderAnim As System.Windows.Forms.PictureBox
    Friend WithEvents picFolder0 As System.Windows.Forms.PictureBox
    Friend WithEvents picHelp As System.Windows.Forms.PictureBox
    Friend WithEvents picHelpSpin As System.Windows.Forms.PictureBox
    Friend WithEvents picHelpStill As System.Windows.Forms.PictureBox
    Friend WithEvents lnkFindData As System.Windows.Forms.LinkLabel
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents lnkHelp As System.Windows.Forms.LinkLabel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmWelcome))
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.chkShowWelcome = New System.Windows.Forms.CheckBox()
        Me.picFolder1 = New System.Windows.Forms.PictureBox()
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel()
        Me.lnkHelp = New System.Windows.Forms.LinkLabel()
        Me.picHelp = New System.Windows.Forms.PictureBox()
        Me.lnkFindData = New System.Windows.Forms.LinkLabel()
        Me.picGlobe = New System.Windows.Forms.PictureBox()
        Me.picGlobeSpin = New System.Windows.Forms.PictureBox()
        Me.picGlobeStill = New System.Windows.Forms.PictureBox()
        Me.picFolderStill = New System.Windows.Forms.PictureBox()
        Me.picFolderAnim = New System.Windows.Forms.PictureBox()
        Me.picFolder0 = New System.Windows.Forms.PictureBox()
        Me.picHelpSpin = New System.Windows.Forms.PictureBox()
        Me.picHelpStill = New System.Windows.Forms.PictureBox()
        Me.SuspendLayout()
        '
        'picLogo
        '
        Me.picLogo.BackColor = System.Drawing.Color.White
        Me.picLogo.Image = CType(resources.GetObject("picLogo.Image"), System.Drawing.Bitmap)
        Me.picLogo.Location = New System.Drawing.Point(4, 5)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(192, 312)
        Me.picLogo.TabIndex = 0
        Me.picLogo.TabStop = False
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Location = New System.Drawing.Point(204, 40)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(144, 16)
        Me.LinkLabel1.TabIndex = 2
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Load data from a file"
        Me.LinkLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkShowWelcome
        '
        Me.chkShowWelcome.Checked = True
        Me.chkShowWelcome.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowWelcome.Location = New System.Drawing.Point(208, 288)
        Me.chkShowWelcome.Name = "chkShowWelcome"
        Me.chkShowWelcome.Size = New System.Drawing.Size(208, 24)
        Me.chkShowWelcome.TabIndex = 4
        Me.chkShowWelcome.Text = "Show this screen at start-up"
        '
        'picFolder1
        '
        Me.picFolder1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.picFolder1.Image = CType(resources.GetObject("picFolder1.Image"), System.Drawing.Bitmap)
        Me.picFolder1.Location = New System.Drawing.Point(356, 89)
        Me.picFolder1.Name = "picFolder1"
        Me.picFolder1.Size = New System.Drawing.Size(48, 40)
        Me.picFolder1.TabIndex = 5
        Me.picFolder1.TabStop = False
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Location = New System.Drawing.Point(204, 101)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(144, 16)
        Me.LinkLabel2.TabIndex = 6
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Load a sample data set"
        Me.LinkLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lnkHelp
        '
        Me.lnkHelp.Location = New System.Drawing.Point(204, 162)
        Me.lnkHelp.Name = "lnkHelp"
        Me.lnkHelp.Size = New System.Drawing.Size(144, 16)
        Me.lnkHelp.TabIndex = 8
        Me.lnkHelp.TabStop = True
        Me.lnkHelp.Text = "View help file"
        Me.lnkHelp.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'picHelp
        '
        Me.picHelp.Cursor = System.Windows.Forms.Cursors.Hand
        Me.picHelp.Image = CType(resources.GetObject("picHelp.Image"), System.Drawing.Bitmap)
        Me.picHelp.Location = New System.Drawing.Point(356, 150)
        Me.picHelp.Name = "picHelp"
        Me.picHelp.Size = New System.Drawing.Size(32, 40)
        Me.picHelp.TabIndex = 9
        Me.picHelp.TabStop = False
        '
        'lnkFindData
        '
        Me.lnkFindData.Location = New System.Drawing.Point(204, 223)
        Me.lnkFindData.Name = "lnkFindData"
        Me.lnkFindData.Size = New System.Drawing.Size(144, 16)
        Me.lnkFindData.TabIndex = 10
        Me.lnkFindData.TabStop = True
        Me.lnkFindData.Text = "Find streamflow data online"
        Me.lnkFindData.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'picGlobe
        '
        Me.picGlobe.Cursor = System.Windows.Forms.Cursors.Hand
        Me.picGlobe.Image = CType(resources.GetObject("picGlobe.Image"), System.Drawing.Bitmap)
        Me.picGlobe.Location = New System.Drawing.Point(357, 211)
        Me.picGlobe.Name = "picGlobe"
        Me.picGlobe.Size = New System.Drawing.Size(48, 48)
        Me.picGlobe.TabIndex = 11
        Me.picGlobe.TabStop = False
        '
        'picGlobeSpin
        '
        Me.picGlobeSpin.Image = CType(resources.GetObject("picGlobeSpin.Image"), System.Drawing.Bitmap)
        Me.picGlobeSpin.Location = New System.Drawing.Point(216, 240)
        Me.picGlobeSpin.Name = "picGlobeSpin"
        Me.picGlobeSpin.Size = New System.Drawing.Size(48, 48)
        Me.picGlobeSpin.TabIndex = 12
        Me.picGlobeSpin.TabStop = False
        Me.picGlobeSpin.Visible = False
        '
        'picGlobeStill
        '
        Me.picGlobeStill.Image = CType(resources.GetObject("picGlobeStill.Image"), System.Drawing.Bitmap)
        Me.picGlobeStill.Location = New System.Drawing.Point(216, 240)
        Me.picGlobeStill.Name = "picGlobeStill"
        Me.picGlobeStill.Size = New System.Drawing.Size(48, 48)
        Me.picGlobeStill.TabIndex = 13
        Me.picGlobeStill.TabStop = False
        Me.picGlobeStill.Visible = False
        '
        'picFolderStill
        '
        Me.picFolderStill.Image = CType(resources.GetObject("picFolderStill.Image"), System.Drawing.Bitmap)
        Me.picFolderStill.Location = New System.Drawing.Point(208, 128)
        Me.picFolderStill.Name = "picFolderStill"
        Me.picFolderStill.Size = New System.Drawing.Size(48, 40)
        Me.picFolderStill.TabIndex = 14
        Me.picFolderStill.TabStop = False
        Me.picFolderStill.Visible = False
        '
        'picFolderAnim
        '
        Me.picFolderAnim.Image = CType(resources.GetObject("picFolderAnim.Image"), System.Drawing.Bitmap)
        Me.picFolderAnim.Location = New System.Drawing.Point(208, 128)
        Me.picFolderAnim.Name = "picFolderAnim"
        Me.picFolderAnim.Size = New System.Drawing.Size(48, 40)
        Me.picFolderAnim.TabIndex = 15
        Me.picFolderAnim.TabStop = False
        Me.picFolderAnim.Visible = False
        '
        'picFolder0
        '
        Me.picFolder0.Cursor = System.Windows.Forms.Cursors.Hand
        Me.picFolder0.Image = CType(resources.GetObject("picFolder0.Image"), System.Drawing.Bitmap)
        Me.picFolder0.Location = New System.Drawing.Point(356, 28)
        Me.picFolder0.Name = "picFolder0"
        Me.picFolder0.Size = New System.Drawing.Size(48, 40)
        Me.picFolder0.TabIndex = 16
        Me.picFolder0.TabStop = False
        '
        'picHelpSpin
        '
        Me.picHelpSpin.Image = CType(resources.GetObject("picHelpSpin.Image"), System.Drawing.Bitmap)
        Me.picHelpSpin.Location = New System.Drawing.Point(224, 176)
        Me.picHelpSpin.Name = "picHelpSpin"
        Me.picHelpSpin.Size = New System.Drawing.Size(32, 40)
        Me.picHelpSpin.TabIndex = 17
        Me.picHelpSpin.TabStop = False
        Me.picHelpSpin.Visible = False
        '
        'picHelpStill
        '
        Me.picHelpStill.Image = CType(resources.GetObject("picHelpStill.Image"), System.Drawing.Bitmap)
        Me.picHelpStill.Location = New System.Drawing.Point(232, 176)
        Me.picHelpStill.Name = "picHelpStill"
        Me.picHelpStill.Size = New System.Drawing.Size(32, 40)
        Me.picHelpStill.TabIndex = 18
        Me.picHelpStill.TabStop = False
        Me.picHelpStill.Visible = False
        '
        'frmWelcome
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(424, 318)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.picHelpStill, Me.picHelpSpin, Me.picFolder0, Me.picFolderStill, Me.picGlobeStill, Me.picGlobeSpin, Me.picGlobe, Me.picHelp, Me.lnkFindData, Me.lnkHelp, Me.picFolder1, Me.LinkLabel2, Me.chkShowWelcome, Me.LinkLabel1, Me.picLogo, Me.picFolderAnim})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmWelcome"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Welcome to Low Flow Calculator"
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub picGlobe_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles picGlobe.MouseLeave
        picGlobe.Image = picGlobeStill.Image
    End Sub


    Private Sub picGlobe_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles picGlobe.MouseEnter
        picGlobe.Image = picGlobeSpin.Image
    End Sub

    Private Sub picFolder0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picFolder0.Click
        openfile()
    End Sub

    Private Sub picFolder0_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles picFolder0.MouseEnter
        picFolder0.Image = picFolderAnim.Image
    End Sub

    Private Sub picFolder0_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles picFolder0.MouseLeave
        picFolder0.Image = picFolderStill.Image
    End Sub

    Private Sub picFolder1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles picFolder1.MouseEnter
        picFolder1.Image = picFolderAnim.Image
    End Sub

    Private Sub picFolder1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles picFolder1.MouseLeave
        picFolder1.Image = picFolderStill.Image
    End Sub

    Private Sub chkShowWelcome_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowWelcome.CheckedChanged
        If Me.Visible = True Then
            m_ShowWelcome = chkShowWelcome.Checked
        End If
    End Sub

    Private Sub frmWelcome_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        chkShowWelcome.Checked = m_ShowWelcome
    End Sub

    Private Sub picHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picHelp.Click
        Dim p As New Process()
        p.Start(System.Windows.Forms.Application.StartupPath & "\LowFlowCalc.chm")
    End Sub

    Private Sub picHelp_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles picHelp.MouseEnter
        picHelp.Image = picHelpSpin.Image
    End Sub

    Private Sub picHelp_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles picHelp.MouseLeave
        picHelp.Image = picHelpStill.Image
    End Sub

    Private Sub lnkFindData_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkFindData.LinkClicked
        FindData()
    End Sub

    Private Sub FindData()
        Dim p As New Process()
        p.Start("http://www.google.com/search?q=streamflow%20data")
    End Sub

    Private Sub picGlobe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picGlobe.Click
        FindData()
    End Sub

    Private Sub picFolder1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picFolder1.Click
        LoadSample()
    End Sub

    Private Sub LoadSample()
        Dim FileName As String = System.Windows.Forms.Application.StartupPath & "\data\sample.txt"
        ReturnAction = FileName
        Me.Close()
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        LoadSample()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        OpenFile()
    End Sub
    Private Sub OpenFile()
        ReturnAction = "open"
        Me.Close()
    End Sub

    Private Sub picLogo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picLogo.Click

    End Sub

    Private Sub picLogo_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picLogo.MouseDown
        Dim p As New Process()
        Try
            If e.X > (32 - picLogo.Left) And e.X < (32 - picLogo.Left + 128) And e.Y > (40 - picLogo.Top) And e.Y < (40 - picLogo.Top + 16) Then 'we are in the URL area
                p.Start("http://www.hydromap.com/")
            End If
        Catch
        End Try
    End Sub


    Private Sub lnkHelp_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkHelp.LinkClicked
        Dim p As New Process()
        p.Start(System.Windows.Forms.Application.StartupPath & "\LowFlowCalc.chm")
    End Sub
End Class
