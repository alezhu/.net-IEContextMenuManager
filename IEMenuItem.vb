Public Class IEMenuItem
	Sub New(ByVal Name As String, ByVal Path As String, ByVal Context As Integer)
		Me.Name = Name
		Me.Path = Path
		Me.Context = Context
	End Sub
	Public Name, Path As String
	Public Context As Integer
End Class
