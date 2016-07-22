Option Strict Off
Imports Microsoft.Win32

Public Class fmMain
    Inherits System.Windows.Forms.Form
    Const sContexts As String = "Contexts"
    Const sRegPath As String = "Software\Microsoft\Internet Explorer\MenuExt"
	Const AppName As String = "IEContextMenuManager"
	Friend WithEvents btnNew As System.Windows.Forms.Button
	Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
	Friend WithEvents miDelete As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
	Friend WithEvents miRefresh As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents btnEdit As System.Windows.Forms.Button
	Friend WithEvents btnDelete As System.Windows.Forms.Button
	Friend WithEvents miChange As System.Windows.Forms.ToolStripMenuItem
	Private bModified As Boolean = False
	Private arDeleted As New ArrayList()

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
	Friend WithEvents ProjectVersion1 As azLib.WindowsControls.ProjectVersion
	Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
	Friend WithEvents lvMenuItems As System.Windows.Forms.ListView
	Friend WithEvents chName As System.Windows.Forms.ColumnHeader
	Friend WithEvents chPath As System.Windows.Forms.ColumnHeader
	Friend WithEvents chContext As System.Windows.Forms.ColumnHeader
	Friend WithEvents btnOk As System.Windows.Forms.Button
	Friend WithEvents btnCancel As System.Windows.Forms.Button
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container
		Me.ProjectVersion1 = New azLib.WindowsControls.ProjectVersion(Me.components)
		Me.GroupBox1 = New System.Windows.Forms.GroupBox
		Me.lvMenuItems = New System.Windows.Forms.ListView
		Me.chName = New System.Windows.Forms.ColumnHeader
		Me.chPath = New System.Windows.Forms.ColumnHeader
		Me.chContext = New System.Windows.Forms.ColumnHeader
		Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
		Me.miChange = New System.Windows.Forms.ToolStripMenuItem
		Me.miDelete = New System.Windows.Forms.ToolStripMenuItem
		Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
		Me.miRefresh = New System.Windows.Forms.ToolStripMenuItem
		Me.btnOk = New System.Windows.Forms.Button
		Me.btnCancel = New System.Windows.Forms.Button
		Me.btnNew = New System.Windows.Forms.Button
		Me.btnEdit = New System.Windows.Forms.Button
		Me.btnDelete = New System.Windows.Forms.Button
		Me.GroupBox1.SuspendLayout()
		Me.ContextMenuStrip1.SuspendLayout()
		Me.SuspendLayout()
		'
		'ProjectVersion1
		'
		Me.ProjectVersion1.AddVersionToTitle = True
		Me.ProjectVersion1.ConnectedForm = Me
		'
		'GroupBox1
		'
		Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
					Or System.Windows.Forms.AnchorStyles.Left) _
					Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.GroupBox1.Controls.Add(Me.lvMenuItems)
		Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
		Me.GroupBox1.Name = "GroupBox1"
		Me.GroupBox1.Size = New System.Drawing.Size(552, 320)
		Me.GroupBox1.TabIndex = 0
		Me.GroupBox1.TabStop = False
		Me.GroupBox1.Text = "Установленные расширения меню"
		'
		'lvMenuItems
		'
		Me.lvMenuItems.Alignment = System.Windows.Forms.ListViewAlignment.[Default]
		Me.lvMenuItems.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.chName, Me.chPath, Me.chContext})
		Me.lvMenuItems.ContextMenuStrip = Me.ContextMenuStrip1
		Me.lvMenuItems.Dock = System.Windows.Forms.DockStyle.Fill
		Me.lvMenuItems.FullRowSelect = True
		Me.lvMenuItems.GridLines = True
		Me.lvMenuItems.HideSelection = False
		Me.lvMenuItems.LabelEdit = True
		Me.lvMenuItems.Location = New System.Drawing.Point(3, 16)
		Me.lvMenuItems.MultiSelect = False
		Me.lvMenuItems.Name = "lvMenuItems"
		Me.lvMenuItems.Size = New System.Drawing.Size(546, 301)
		Me.lvMenuItems.TabIndex = 0
		Me.lvMenuItems.UseCompatibleStateImageBehavior = False
		Me.lvMenuItems.View = System.Windows.Forms.View.Details
		'
		'chName
		'
		Me.chName.Text = "Название"
		Me.chName.Width = 110
		'
		'chPath
		'
		Me.chPath.Text = "Путь"
		'
		'chContext
		'
		Me.chContext.Text = "Контекст"
		'
		'ContextMenuStrip1
		'
		Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.miChange, Me.miDelete, Me.ToolStripSeparator1, Me.miRefresh})
		Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
		Me.ContextMenuStrip1.ShowImageMargin = False
		Me.ContextMenuStrip1.Size = New System.Drawing.Size(104, 76)
		'
		'miChange
		'
		Me.miChange.Enabled = False
		Me.miChange.Name = "miChange"
		Me.miChange.Size = New System.Drawing.Size(103, 22)
		Me.miChange.Text = "Изменить"
		'
		'miDelete
		'
		Me.miDelete.Enabled = False
		Me.miDelete.Name = "miDelete"
		Me.miDelete.Size = New System.Drawing.Size(103, 22)
		Me.miDelete.Text = "Удалить"
		'
		'ToolStripSeparator1
		'
		Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
		Me.ToolStripSeparator1.Size = New System.Drawing.Size(100, 6)
		'
		'miRefresh
		'
		Me.miRefresh.Name = "miRefresh"
		Me.miRefresh.Size = New System.Drawing.Size(103, 22)
		Me.miRefresh.Text = "Обновить"
		'
		'btnOk
		'
		Me.btnOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK
		Me.btnOk.Location = New System.Drawing.Point(392, 336)
		Me.btnOk.Name = "btnOk"
		Me.btnOk.Size = New System.Drawing.Size(75, 23)
		Me.btnOk.TabIndex = 1
		Me.btnOk.Text = "OK"
		'
		'btnCancel
		'
		Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.btnCancel.Location = New System.Drawing.Point(480, 336)
		Me.btnCancel.Name = "btnCancel"
		Me.btnCancel.Size = New System.Drawing.Size(75, 23)
		Me.btnCancel.TabIndex = 2
		Me.btnCancel.Text = "Отмена"
		'
		'btnNew
		'
		Me.btnNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.btnNew.DialogResult = System.Windows.Forms.DialogResult.OK
		Me.btnNew.Location = New System.Drawing.Point(8, 336)
		Me.btnNew.Name = "btnNew"
		Me.btnNew.Size = New System.Drawing.Size(75, 23)
		Me.btnNew.TabIndex = 1
		Me.btnNew.Text = "Новое"
		'
		'btnEdit
		'
		Me.btnEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.btnEdit.DialogResult = System.Windows.Forms.DialogResult.OK
		Me.btnEdit.Enabled = False
		Me.btnEdit.Location = New System.Drawing.Point(89, 336)
		Me.btnEdit.Name = "btnEdit"
		Me.btnEdit.Size = New System.Drawing.Size(75, 23)
		Me.btnEdit.TabIndex = 1
		Me.btnEdit.Text = "Изменить"
		'
		'btnDelete
		'
		Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.btnDelete.DialogResult = System.Windows.Forms.DialogResult.OK
		Me.btnDelete.Enabled = False
		Me.btnDelete.Location = New System.Drawing.Point(170, 336)
		Me.btnDelete.Name = "btnDelete"
		Me.btnDelete.Size = New System.Drawing.Size(75, 23)
		Me.btnDelete.TabIndex = 1
		Me.btnDelete.Text = "Удалить"
		'
		'fmMain
		'
		Me.AcceptButton = Me.btnOk
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.CancelButton = Me.btnCancel
		Me.ClientSize = New System.Drawing.Size(568, 373)
		Me.Controls.Add(Me.btnCancel)
		Me.Controls.Add(Me.btnNew)
		Me.Controls.Add(Me.btnEdit)
		Me.Controls.Add(Me.btnDelete)
		Me.Controls.Add(Me.btnOk)
		Me.Controls.Add(Me.GroupBox1)
		Me.MinimumSize = New System.Drawing.Size(248, 160)
		Me.Name = "fmMain"
		Me.Text = "IE Context Menu Manager"
		Me.GroupBox1.ResumeLayout(False)
		Me.ContextMenuStrip1.ResumeLayout(False)
		Me.ResumeLayout(False)

	End Sub

#End Region




	Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
		bModified = False
		Close()
	End Sub

	Private Sub fmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		Dim s As New Settings
		s = Settings.Load(AppName, s.GetType)
		If (s.FormWidth > 0) Then
			Me.SetBounds(s.FormLeft, s.FormTop, s.FormWidth, s.FormHeight)
			Me.WindowState = s.WindowState
			Me.lvMenuItems.Columns(0).Width = s.Column1Width
			Me.lvMenuItems.Columns(1).Width = s.Column2Width
		End If


		LoadItems()

	End Sub

	Private Sub LoadItems()
		lvMenuItems.BeginUpdate()
		Try
			Dim reg As RegistryKey
			reg = Registry.CurrentUser.OpenSubKey(sRegPath)
			Try
				Dim FakeArray() As Integer = {0}
				lvMenuItems.Items.Clear()
				For Each Menu As String In reg.GetSubKeyNames()
					Dim mreg As RegistryKey = reg.OpenSubKey(Menu)
					Dim li As ListViewItem = lvMenuItems.Items.Add(Menu)
                    li.SubItems.Add(mreg.GetValue("", "").ToString())
                    Dim value As Integer = 0
                    Try
                        Dim kind As RegistryValueKind = mreg.GetValueKind(sContexts)

                        If kind = RegistryValueKind.Binary Then
                            Dim ARR() As Byte = mreg.GetValue(sContexts, Nothing)
                            If ARR.Length > 0 Then
                                value = ARR(0)
                            End If
                        Else
                            Dim Contexts As Object = mreg.GetValue(sContexts, Nothing)
                            value = CType(Contexts, Integer)
                        End If
                    Catch
                    Finally
                        li.SubItems.Add(value.ToString())
                        li.SubItems.Add(Menu) ' Old name
                    End Try
                Next
			Finally
				reg = Nothing
			End Try
		Finally
			lvMenuItems.EndUpdate()
			bModified = False
		End Try
	End Sub

	Private Sub SaveItems()
		Dim reg As RegistryKey
		reg = Registry.CurrentUser.OpenSubKey(sRegPath, RegistryKeyPermissionCheck.ReadWriteSubTree)
		Try
			Dim li As ListViewItem
			Dim newkey As Boolean
			Dim Context() As Byte = {0}
			For Each li In lvMenuItems.Items
				Dim name, oldName As String
				name = li.SubItems(0).Text
				oldName = li.SubItems(3).Text

				Dim mreg As RegistryKey = reg.OpenSubKey(name, True)

				newkey = False
				If mreg Is Nothing Then
					mreg = reg.CreateSubKey(name, RegistryKeyPermissionCheck.ReadWriteSubTree)
					newkey = True
				End If
				mreg.SetValue("", li.SubItems(1).Text)
				Context(0) = CInt(li.SubItems(2).Text)
				If Not newkey Then
					Try
						mreg.DeleteValue(sContexts)
					Catch ex As System.ArgumentException

					End Try
				End If
				mreg.SetValue(sContexts, Context, RegistryValueKind.Binary)

				If (Not (oldName = name)) And (Not (oldName = String.Empty)) Then
					reg.DeleteSubKey(oldName, False)
				End If
				li.SubItems(3).Text = name
			Next

			For Each Name As String In arDeleted
				reg.DeleteSubKey(Name, False)
			Next

		Finally
			reg = Nothing
			bModified = False
		End Try
	End Sub

	Private Sub SaveSettings()
		Dim s As New Settings(AppName)
		s.Column1Width = Me.lvMenuItems.Columns(0).Width
		s.Column2Width = Me.lvMenuItems.Columns(1).Width
		s.FormHeight = Me.Height
		s.FormLeft = Me.Left
		s.FormTop = Me.Top
		s.FormWidth = Me.Width
		s.WindowState = Me.WindowState
		s.Save()
	End Sub

	Private Sub fmMain_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
		SaveSettings()
	End Sub

	Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
		If bModified Then SaveItems()
	End Sub

	Private Sub fmMain_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing

		If bModified Then
			Dim res As MsgBoxResult = MsgBox("Элементы были изменены. Сохранить изменения?", MsgBoxStyle.YesNoCancel)
			If res = MsgBoxResult.Cancel Then
				e.Cancel = True
			ElseIf res = MsgBoxResult.Yes Then
				SaveItems()
			End If
		End If
	End Sub


	Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
		Dim Prop As fmProperty
		Prop = New fmProperty
		Dim NewName As String = "Новое"
		Dim BaseName As String = NewName
		Dim Li As ListViewItem
		Dim i As Integer = 1

		Do
			Li = lvMenuItems.FindItemWithText(NewName)
			If Li Is Nothing Then Exit Do
			NewName = BaseName & i.ToString(" #")
			i = i + 1
		Loop While True

		Prop.MenuItem = New IEMenuItem(NewName, "", 0)
		If Prop.ShowDialog() = Windows.Forms.DialogResult.OK Then
			With Prop.MenuItem
				Li = lvMenuItems.Items.Add(.Name)
				'Li.SubItems(0).Text = .Name
				Li.SubItems.Add(.Path)
				Li.SubItems.Add(.Context.ToString())
				Li.SubItems.Add(String.Empty)
				bModified = True
			End With
		End If
		Prop = Nothing
	End Sub

	Private Sub lvMenuItems_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvMenuItems.MouseDoubleClick
		If e.Button = Windows.Forms.MouseButtons.Left Then EditItem()
	End Sub

	Private Sub EditItem()
		Dim Prop As fmProperty
		Prop = New fmProperty
		Dim LVItem As ListViewItem
		LVItem = lvMenuItems.FocusedItem
		Prop.MenuItem = New IEMenuItem(LVItem.Text, LVItem.SubItems(1).Text, Integer.Parse(LVItem.SubItems(2).Text))
		If Prop.ShowDialog() = Windows.Forms.DialogResult.OK Then
			LVItem = lvMenuItems.FocusedItem
			With Prop.MenuItem
				LVItem.SubItems(0).Text = .Name
				LVItem.SubItems(1).Text = .Path
				LVItem.SubItems(2).Text = .Context.ToString()
				bModified = True
			End With
		End If
		Prop = Nothing
	End Sub

	Private Sub lvMenuItems_ItemSelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles lvMenuItems.ItemSelectionChanged
		Dim sel As Boolean = e.IsSelected
		miChange.Enabled = sel
		miDelete.Enabled = sel
		btnEdit.Enabled = sel
		btnDelete.Enabled = sel
	End Sub

	Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
		DeleteItem()
	End Sub

	Private Sub DeleteItem()
		Dim li As ListViewItem = lvMenuItems.SelectedItems(0)
		If Not li Is Nothing Then
			Dim res As MsgBoxResult = MsgBox("Вы действительно хотите удалить это элемент?", MsgBoxStyle.YesNo)
			If res = MsgBoxResult.Yes Then
				Dim name As String = li.SubItems(3).Text
				If Not name = String.Empty Then
					arDeleted.Add(name)
				End If
				li.Remove()
				bModified = True
			End If
		End If
	End Sub

	Private Sub miRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miRefresh.Click
		LoadItems()
	End Sub

	Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
		EditItem()
	End Sub

	Private Sub miChange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miChange.Click
		EditItem()
	End Sub

	Private Sub miDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miDelete.Click
		DeleteItem()
	End Sub

	Private Sub btnOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
		SaveItems()
		bModified = False
		Close()
	End Sub
End Class
