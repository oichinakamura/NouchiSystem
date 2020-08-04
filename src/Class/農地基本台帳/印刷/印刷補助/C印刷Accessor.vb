Public MustInherit Class C印刷Accessor
    Inherits HimTools2012.clsAccessor

    Public Overrides Sub Execute()

    End Sub


    Protected Property n置換え桁数 As Integer = 0

    Protected Function TargetStr2(ByVal sName As String, ByVal nNumber As Integer, Optional nFigureLength As Integer = 0) As String
        If nFigureLength = 0 Then
            Return "{" & sName & Microsoft.VisualBasic.Strings.Right("00000" & nNumber.ToString, n置換え桁数) & "}"
        Else
            Return "{" & sName & Microsoft.VisualBasic.Strings.Right("00000" & nNumber.ToString, nFigureLength) & "}"
        End If

    End Function
End Class
