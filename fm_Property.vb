Public Class fmProperty
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents tbName As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cbDefault As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents tbPath As System.Windows.Forms.TextBox
    Friend WithEvents cbImage As System.Windows.Forms.CheckBox
    Friend WithEvents cbControls As System.Windows.Forms.CheckBox
    Friend WithEvents cbSelection As System.Windows.Forms.CheckBox
    Friend WithEvents cbTables As System.Windows.Forms.CheckBox
    Friend WithEvents cbAnchors As System.Windows.Forms.CheckBox
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents ofdPAth As System.Windows.Forms.OpenFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.GroupBox1 = New System.Windows.Forms.GroupBox
		Me.tbName = New System.Windows.Forms.TextBox
		Me.GroupBox2 = New System.Windows.Forms.GroupBox
		Me.cbAnchors = New System.Windows.Forms.CheckBox
		Me.cbSelection = New System.Windows.Forms.CheckBox
		Me.cbTables = New System.Windows.Forms.CheckBox
		Me.cbControls = New System.Windows.Forms.CheckBox
		Me.cbImage = New System.Windows.Forms.CheckBox
		Me.cbDefault = New System.Windows.Forms.CheckBox
		Me.GroupBox3 = New System.Windows.Forms.GroupBox
		Me.btnBrowse = New System.Windows.Forms.Button
		Me.tbPath = New System.Windows.Forms.TextBox
		Me.btnOK = New System.Windows.Forms.Button
		Me.btnCancel = New System.Windows.Forms.Button
		Me.ofdPAth = New System.Windows.Forms.OpenFileDialog
		Me.GroupBox1.SuspendLayout()
		Me.GroupBox2.SuspendLayout()
		Me.GroupBox3.SuspendLayout()
		Me.SuspendLayout()
		'
		'GroupBox1
		'
		Me.GroupBox1.Controls.Add(Me.tbName)
		Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
		Me.GroupBox1.Name = "GroupBox1"
		Me.GroupBox1.Size = New System.Drawing.Size(440, 48)
		Me.GroupBox1.TabIndex = 0
		Me.GroupBox1.TabStop = False
		Me.GroupBox1.Text = "Название"
		'
		'tbName
		'
		Me.tbName.Location = New System.Drawing.Point(8, 16)
		Me.tbName.Name = "tbName"
		Me.tbName.Size = New System.Drawing.Size(424, 20)
		Me.tbName.TabIndex = 0
		'
		'GroupBox2
		'
		Me.GroupBox2.Controls.Add(Me.cbAnchors)
		Me.GroupBox2.Controls.Add(Me.cbSelection)
		Me.GroupBox2.Controls.Add(Me.cbTables)
		Me.GroupBox2.Controls.Add(Me.cbControls)
		Me.GroupBox2.Controls.Add(Me.cbImage)
		Me.GroupBox2.Controls.Add(Me.cbDefault)
		Me.GroupBox2.Location = New System.Drawing.Point(8, 120)
		Me.GroupBox2.Name = "GroupBox2"
		Me.GroupBox2.Size = New System.Drawing.Size(440, 96)
		Me.GroupBox2.TabIndex = 1
		Me.GroupBox2.TabStop = False
		Me.GroupBox2.Text = "Применять..."
		'
		'cbAnchors
		'
		Me.cbAnchors.Location = New System.Drawing.Point(224, 64)
		Me.cbAnchors.Name = "cbAnchors"
		Me.cbAnchors.Size = New System.Drawing.Size(208, 24)
		Me.cbAnchors.TabIndex = 5
		Me.cbAnchors.Tag = "32"
		Me.cbAnchors.Text = "... для ссылок"
		'
		'cbSelection
		'
		Me.cbSelection.Location = New System.Drawing.Point(224, 40)
		Me.cbSelection.Name = "cbSelection"
		Me.cbSelection.Size = New System.Drawing.Size(208, 24)
		Me.cbSelection.TabIndex = 4
		Me.cbSelection.Tag = "16"
		Me.cbSelection.Text = "... для выделения"
		'
		'cbTables
		'
		Me.cbTables.Location = New System.Drawing.Point(224, 16)
		Me.cbTables.Name = "cbTables"
		Me.cbTables.Size = New System.Drawing.Size(208, 24)
		Me.cbTables.TabIndex = 3
		Me.cbTables.Tag = "8"
		Me.cbTables.Text = "... для таблиц"
		'
		'cbControls
		'
		Me.cbControls.Location = New System.Drawing.Point(8, 64)
		Me.cbControls.Name = "cbControls"
		Me.cbControls.Size = New System.Drawing.Size(208, 24)
		Me.cbControls.TabIndex = 2
		Me.cbControls.Tag = "4"
		Me.cbControls.Text = "... для ""контролов"""
		'
		'cbImage
		'
		Me.cbImage.Location = New System.Drawing.Point(8, 40)
		Me.cbImage.Name = "cbImage"
		Me.cbImage.Size = New System.Drawing.Size(208, 24)
		Me.cbImage.TabIndex = 1
		Me.cbImage.Tag = "2"
		Me.cbImage.Text = "... для картинок"
		'
		'cbDefault
		'
		Me.cbDefault.Location = New System.Drawing.Point(8, 16)
		Me.cbDefault.Name = "cbDefault"
		Me.cbDefault.Size = New System.Drawing.Size(208, 24)
		Me.cbDefault.TabIndex = 0
		Me.cbDefault.Tag = "1"
		Me.cbDefault.Text = "... по умолчанию"
		'
		'GroupBox3
		'
		Me.GroupBox3.Controls.Add(Me.btnBrowse)
		Me.GroupBox3.Controls.Add(Me.tbPath)
		Me.GroupBox3.Location = New System.Drawing.Point(8, 64)
		Me.GroupBox3.Name = "GroupBox3"
		Me.GroupBox3.Size = New System.Drawing.Size(440, 48)
		Me.GroupBox3.TabIndex = 2
		Me.GroupBox3.TabStop = False
		Me.GroupBox3.Text = "Путь"
		'
		'btnBrowse
		'
		Me.btnBrowse.Location = New System.Drawing.Point(400, 16)
		Me.btnBrowse.Name = "btnBrowse"
		Me.btnBrowse.Size = New System.Drawing.Size(24, 23)
		Me.btnBrowse.TabIndex = 1
		Me.btnBrowse.Text = "..."
		'
		'tbPath
		'
		Me.tbPath.Location = New System.Drawing.Point(8, 16)
		Me.tbPath.Name = "tbPath"
		Me.tbPath.Size = New System.Drawing.Size(384, 20)
		Me.tbPath.TabIndex = 0
		'
		'btnOK
		'
		Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
		Me.btnOK.Location = New System.Drawing.Point(128, 224)
		Me.btnOK.Name = "btnOK"
		Me.btnOK.Size = New System.Drawing.Size(75, 23)
		Me.btnOK.TabIndex = 3
		Me.btnOK.Text = "OK"
		'
		'btnCancel
		'
		Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.btnCancel.Location = New System.Drawing.Point(224, 224)
		Me.btnCancel.Name = "btnCancel"
		Me.btnCancel.Size = New System.Drawing.Size(75, 23)
		Me.btnCancel.TabIndex = 4
		Me.btnCancel.Text = "Отмена"
		'
		'ofdPAth
		'
		Me.ofdPAth.ShowReadOnly = True
		Me.ofdPAth.SupportMultiDottedExtensions = True
		Me.ofdPAth.Title = "Выберите файл"
		'
		'fmProperty
		'
		Me.AcceptButton = Me.btnOK
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.CancelButton = Me.btnCancel
		Me.ClientSize = New System.Drawing.Size(456, 255)
		Me.Controls.Add(Me.btnCancel)
		Me.Controls.Add(Me.btnOK)
		Me.Controls.Add(Me.GroupBox3)
		Me.Controls.Add(Me.GroupBox2)
		Me.Controls.Add(Me.GroupBox1)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "fmProperty"
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Text = "Свойства элемента"
		Me.GroupBox1.ResumeLayout(False)
		Me.GroupBox1.PerformLayout()
		Me.GroupBox2.ResumeLayout(False)
		Me.GroupBox3.ResumeLayout(False)
		Me.GroupBox3.PerformLayout()
		Me.ResumeLayout(False)

	End Sub

#End Region
    Public MenuItem As IEMenuItem

    Private Sub fmProperty_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not MenuItem Is Nothing Then
            tbName.Text = MenuItem.Name
            tbPath.Text = MenuItem.Path
            cbDefault.Checked = (MenuItem.Context And 1) = 1
            cbImage.Checked = (MenuItem.Context And 2) = 2
            cbControls.Checked = (MenuItem.Context And 4) = 4
            cbTables.Checked = (MenuItem.Context And 8) = 8
            cbSelection.Checked = (MenuItem.Context And 16) = 16
            cbAnchors.Checked = (MenuItem.Context And 32) = 32
        End If
    End Sub

    Private Sub tbName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbName.TextChanged
        MenuItem.Name = tbName.Text
    End Sub

    Private Sub MyClick_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAnchors.Click, cbControls.Click, cbDefault.Click, cbImage.Click, cbSelection.Click, cbTables.Click
        Dim cb As CheckBox
        cb = CType(sender, CheckBox)
        If cb.Checked Then
            MenuItem.Context = (MenuItem.Context Or CInt(cb.Tag))
        Else
            MenuItem.Context = (MenuItem.Context And Not CInt(cb.Tag))
        End If
    End Sub

    Private Sub tbPath_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPath.TextChanged
        MenuItem.Path = tbPath.Text
    End Sub

	Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
		If (ofdPAth.ShowDialog() = Windows.Forms.DialogResult.OK) Then
			tbPath.Text = ofdPAth.FileName
		End If
	End Sub
End Class
