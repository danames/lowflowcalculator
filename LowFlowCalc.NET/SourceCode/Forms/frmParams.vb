Public Class frmParams
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents llReturnPeriod As System.Windows.Forms.LinkLabel
    Friend WithEvents llAveragingPeriod As System.Windows.Forms.LinkLabel
    Friend WithEvents txtReturnPeriod As System.Windows.Forms.TextBox
    Friend WithEvents txtAveragingPeriod As System.Windows.Forms.TextBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents lnkHighLow As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbStartMonth As System.Windows.Forms.ComboBox
    Friend WithEvents cmbEndMonth As System.Windows.Forms.ComboBox
    Friend WithEvents lnkMonths As System.Windows.Forms.LinkLabel
    Friend WithEvents cmbZeros As System.Windows.Forms.ComboBox
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents lnkProbDist As System.Windows.Forms.LinkLabel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents radLow As System.Windows.Forms.RadioButton
    Friend WithEvents radHigh As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents radLogNorm As System.Windows.Forms.RadioButton
    Friend WithEvents radPT3 As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmParams))
        Me.llReturnPeriod = New System.Windows.Forms.LinkLabel
        Me.llAveragingPeriod = New System.Windows.Forms.LinkLabel
        Me.txtReturnPeriod = New System.Windows.Forms.TextBox
        Me.txtAveragingPeriod = New System.Windows.Forms.TextBox
        Me.lnkHighLow = New System.Windows.Forms.LinkLabel
        Me.btnOK = New System.Windows.Forms.Button
        Me.lnkMonths = New System.Windows.Forms.LinkLabel
        Me.cmbStartMonth = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmbEndMonth = New System.Windows.Forms.ComboBox
        Me.cmbZeros = New System.Windows.Forms.ComboBox
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.lnkProbDist = New System.Windows.Forms.LinkLabel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.radLow = New System.Windows.Forms.RadioButton
        Me.radHigh = New System.Windows.Forms.RadioButton
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.radLogNorm = New System.Windows.Forms.RadioButton
        Me.radPT3 = New System.Windows.Forms.RadioButton
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'llReturnPeriod
        '
        Me.llReturnPeriod.AutoSize = True
        Me.llReturnPeriod.Location = New System.Drawing.Point(21, 39)
        Me.llReturnPeriod.Name = "llReturnPeriod"
        Me.llReturnPeriod.Size = New System.Drawing.Size(115, 16)
        Me.llReturnPeriod.TabIndex = 20
        Me.llReturnPeriod.TabStop = True
        Me.llReturnPeriod.Text = "Return Period (years):"
        Me.llReturnPeriod.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'llAveragingPeriod
        '
        Me.llAveragingPeriod.AutoSize = True
        Me.llAveragingPeriod.Location = New System.Drawing.Point(7, 7)
        Me.llAveragingPeriod.Name = "llAveragingPeriod"
        Me.llAveragingPeriod.Size = New System.Drawing.Size(129, 16)
        Me.llAveragingPeriod.TabIndex = 19
        Me.llAveragingPeriod.TabStop = True
        Me.llAveragingPeriod.Text = "Averaging Period (days):"
        Me.llAveragingPeriod.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtReturnPeriod
        '
        Me.txtReturnPeriod.Location = New System.Drawing.Point(142, 38)
        Me.txtReturnPeriod.Name = "txtReturnPeriod"
        Me.txtReturnPeriod.Size = New System.Drawing.Size(107, 20)
        Me.txtReturnPeriod.TabIndex = 18
        Me.txtReturnPeriod.Text = "10"
        '
        'txtAveragingPeriod
        '
        Me.txtAveragingPeriod.Location = New System.Drawing.Point(142, 5)
        Me.txtAveragingPeriod.Name = "txtAveragingPeriod"
        Me.txtAveragingPeriod.Size = New System.Drawing.Size(106, 20)
        Me.txtAveragingPeriod.TabIndex = 17
        Me.txtAveragingPeriod.Text = "7"
        '
        'lnkHighLow
        '
        Me.lnkHighLow.AutoSize = True
        Me.lnkHighLow.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lnkHighLow.Location = New System.Drawing.Point(43, 71)
        Me.lnkHighLow.Name = "lnkHighLow"
        Me.lnkHighLow.Size = New System.Drawing.Size(93, 16)
        Me.lnkHighLow.TabIndex = 22
        Me.lnkHighLow.TabStop = True
        Me.lnkHighLow.Text = "Flow to Compute:"
        Me.lnkHighLow.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(270, 237)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(88, 24)
        Me.btnOK.TabIndex = 25
        Me.btnOK.Text = "OK"
        '
        'lnkMonths
        '
        Me.lnkMonths.AutoSize = True
        Me.lnkMonths.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lnkMonths.Location = New System.Drawing.Point(39, 133)
        Me.lnkMonths.Name = "lnkMonths"
        Me.lnkMonths.Size = New System.Drawing.Size(97, 16)
        Me.lnkMonths.TabIndex = 26
        Me.lnkMonths.TabStop = True
        Me.lnkMonths.Text = "Months to Include:"
        Me.lnkMonths.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbStartMonth
        '
        Me.cmbStartMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbStartMonth.Items.AddRange(New Object() {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"})
        Me.cmbStartMonth.Location = New System.Drawing.Point(142, 128)
        Me.cmbStartMonth.Name = "cmbStartMonth"
        Me.cmbStartMonth.Size = New System.Drawing.Size(93, 21)
        Me.cmbStartMonth.TabIndex = 27
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(240, 132)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(14, 16)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "to"
        '
        'cmbEndMonth
        '
        Me.cmbEndMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbEndMonth.Items.AddRange(New Object() {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"})
        Me.cmbEndMonth.Location = New System.Drawing.Point(261, 128)
        Me.cmbEndMonth.Name = "cmbEndMonth"
        Me.cmbEndMonth.Size = New System.Drawing.Size(93, 21)
        Me.cmbEndMonth.TabIndex = 29
        '
        'cmbZeros
        '
        Me.cmbZeros.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbZeros.Items.AddRange(New Object() {"No Action", "Drop zeros before averaging days", "Drop zeros after averaging days"})
        Me.cmbZeros.Location = New System.Drawing.Point(142, 157)
        Me.cmbZeros.Name = "cmbZeros"
        Me.cmbZeros.Size = New System.Drawing.Size(211, 21)
        Me.cmbZeros.TabIndex = 30
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(0, 161)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(138, 16)
        Me.LinkLabel1.TabIndex = 31
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Method for handling zeros:"
        Me.LinkLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lnkProbDist
        '
        Me.lnkProbDist.AutoSize = True
        Me.lnkProbDist.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lnkProbDist.Location = New System.Drawing.Point(18, 189)
        Me.lnkProbDist.Name = "lnkProbDist"
        Me.lnkProbDist.Size = New System.Drawing.Size(120, 16)
        Me.lnkProbDist.TabIndex = 32
        Me.lnkProbDist.TabStop = True
        Me.lnkProbDist.Text = "Probability Distribution:"
        Me.lnkProbDist.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.radLow)
        Me.GroupBox1.Controls.Add(Me.radHigh)
        Me.GroupBox1.Location = New System.Drawing.Point(140, 63)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(217, 52)
        Me.GroupBox1.TabIndex = 35
        Me.GroupBox1.TabStop = False
        '
        'radLow
        '
        Me.radLow.Checked = True
        Me.radLow.Location = New System.Drawing.Point(8, 11)
        Me.radLow.Name = "radLow"
        Me.radLow.Size = New System.Drawing.Size(176, 16)
        Me.radLow.TabIndex = 26
        Me.radLow.TabStop = True
        Me.radLow.Text = "Low flow (i.e. drought events)"
        '
        'radHigh
        '
        Me.radHigh.Location = New System.Drawing.Point(8, 31)
        Me.radHigh.Name = "radHigh"
        Me.radHigh.Size = New System.Drawing.Size(176, 16)
        Me.radHigh.TabIndex = 25
        Me.radHigh.Text = "High flow (i.e. flood events)"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.radLogNorm)
        Me.GroupBox2.Controls.Add(Me.radPT3)
        Me.GroupBox2.Location = New System.Drawing.Point(148, 182)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(209, 48)
        Me.GroupBox2.TabIndex = 36
        Me.GroupBox2.TabStop = False
        '
        'radLogNorm
        '
        Me.radLogNorm.Location = New System.Drawing.Point(7, 30)
        Me.radLogNorm.Name = "radLogNorm"
        Me.radLogNorm.Size = New System.Drawing.Size(176, 16)
        Me.radLogNorm.TabIndex = 36
        Me.radLogNorm.Text = "Log Normal"
        '
        'radPT3
        '
        Me.radPT3.Checked = True
        Me.radPT3.Location = New System.Drawing.Point(7, 10)
        Me.radPT3.Name = "radPT3"
        Me.radPT3.Size = New System.Drawing.Size(176, 16)
        Me.radPT3.TabIndex = 35
        Me.radPT3.TabStop = True
        Me.radPT3.Text = "Log Pearson Type III"
        '
        'frmParams
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(365, 268)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lnkProbDist)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.cmbZeros)
        Me.Controls.Add(Me.cmbEndMonth)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbStartMonth)
        Me.Controls.Add(Me.lnkMonths)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.lnkHighLow)
        Me.Controls.Add(Me.llReturnPeriod)
        Me.Controls.Add(Me.llAveragingPeriod)
        Me.Controls.Add(Me.txtReturnPeriod)
        Me.Controls.Add(Me.txtAveragingPeriod)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmParams"
        Me.Text = "Parameters"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub txtReturnPeriod_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtReturnPeriod.Validating
        txtReturnPeriod.Text = Math.Min(Math.Abs(Val(txtReturnPeriod.Text)), 1000000)
        If txtReturnPeriod.Text = 0 Then
            txtReturnPeriod.Text = 1
        End If
    End Sub

    Private Sub txtAveragingPeriod_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtAveragingPeriod.Validating
        txtAveragingPeriod.Text = Math.Min(Math.Abs(Val(txtAveragingPeriod.Text)), 365)
        If txtAveragingPeriod.Text = 0 Then
            txtAveragingPeriod.Text = 1
        End If
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        'Save the settings to global variables and close the form.
        m_AvePeriod = txtAveragingPeriod.Text
        m_RetPeriod = txtReturnPeriod.Text
        If radHigh.Checked = True Then
            m_ExType = ExtremeTypes.HighFlow
            m_ExtremeTypeString = "high"
        Else
            m_ExType = ExtremeTypes.LowFlow
            m_ExtremeTypeString = "low"
        End If
        If radPT3.Checked = True Then
            m_DistType = DistTypes.LogPearsonTypeIII
            m_DistTypeString = "Log Pearson Type III"
        Else
            m_DistType = DistTypes.LogNormal
            m_DistTypeString = "Log Normal"
        End If
        m_ZerosMethod = cmbZeros.SelectedIndex
        Me.Close()
    End Sub

    Private Sub llAveragingPeriod_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llAveragingPeriod.LinkClicked
        MsgBox("This is the averaging period in days.  For a 7Q10 low flow calculation this value is 7.", MsgBoxStyle.Information, "Low Flow Calculator")
    End Sub

    Private Sub llReturnPeriod_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llReturnPeriod.LinkClicked
        MsgBox("This is the return period in years. For a 7Q10 low flow calculation this value is 10, For a 100-year flood this value would be 100.", MsgBoxStyle.Information, "Low Flow Calculator")
    End Sub

    Private Sub frmParams_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Set the form contents to the global variables values.
        txtReturnPeriod.Text = m_RetPeriod
        txtAveragingPeriod.Text = m_AvePeriod
        cmbEndMonth.SelectedIndex = m_EndMonth - 1
        cmbStartMonth.SelectedIndex = m_StartMonth - 1
        If m_ExType = ExtremeTypes.HighFlow Then
            radHigh.Checked = True
            radLow.Checked = False
        Else
            radHigh.Checked = False
            radLow.Checked = True
        End If
        If m_DistType = DistTypes.LogPearsonTypeIII Then
            radPT3.Checked = True
            radLogNorm.Checked = False
        Else
            radPT3.Checked = False
            radLogNorm.Checked = True
        End If
        cmbZeros.SelectedIndex = m_ZerosMethod
    End Sub

    Private Sub lnkHighLow_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkHighLow.LinkClicked
        MsgBox("Choose 'Low' to compute the drought with the specified averaging and return periods, choose 'High' to compute the flood with the specified averaging and return periods.", MsgBoxStyle.Information, "Low Flow Calculator")
    End Sub

    Private Sub lnkMonths_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkMonths.LinkClicked
        MsgBox("Only data falling within the selected months will be used in the analysis.", MsgBoxStyle.Information, "Low Flow Calculator")
    End Sub

    Private Sub cmbStartMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbStartMonth.SelectedIndexChanged
        m_StartMonth = cmbStartMonth.SelectedIndex + 1
        If cmbEndMonth.SelectedIndex < cmbStartMonth.SelectedIndex Then
            cmbEndMonth.SelectedIndex = cmbStartMonth.SelectedIndex
        End If
    End Sub

    Private Sub cmbEndMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbEndMonth.SelectedIndexChanged
        m_EndMonth = cmbEndMonth.SelectedIndex + 1
        If cmbStartMonth.SelectedIndex > cmbEndMonth.SelectedIndex Then
            cmbStartMonth.SelectedIndex = cmbEndMonth.SelectedIndex
        End If

    End Sub

    Private Sub lnkProbDist_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkProbDist.LinkClicked
        MsgBox("Select the probability distribution to use in the analysis.", MsgBoxStyle.Information, "Low Flow Calculator")
    End Sub
End Class
