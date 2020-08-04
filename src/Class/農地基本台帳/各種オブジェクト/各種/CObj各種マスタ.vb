
'

Public Class CObj各種マスタ : Inherits CObj各種
    Public Sub New(ByVal sKey As String)
        MyBase.New(sKey)
        Me.SetRow(App農地基本台帳.DataMaster.Rows.Find({Me.Key.ID, Me.Key.DataClass}))

    End Sub
End Class
