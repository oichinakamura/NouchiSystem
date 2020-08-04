Imports System.ComponentModel


Public Class CPrint農地法第3_3_1
    Public Sub New(ByVal n受け手ID As Integer, ByVal s農地List As String, ByVal s事由 As String, ByVal 受付日 As DateTime)
        Stop
        'Dim p農地 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect( "SELECT * FROM [D:___] WHERE \\\")
        'Dim p個人 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect( "SELECT * FROM [D:___] WHERE \\\")


        ' Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportPath(SysAD.市町村.市町村名) & "\" & sFile)
        ' HimTools2012.TextAdapter.SaveTextFile(sDesktopFolder & "\" & Replace(sFile, ".", IIf(sCount.Length > 0, "(", "") & sCount & "."),sxml)
    End Sub

End Class
Public Class imputPrint農地法第3_3_1
    Inherits InputObjectParam

    <Category("届出内容")> <HimTools2012.controls.PropertyGridSupport.PropertyOrderAttribute(0)>
    Public Property 受付日 As DateTime

    <Category("届出内容")> <HimTools2012.controls.PropertyGridSupport.PropertyOrderAttribute(1)>
    Public Property 事由 As String = ""

    Public Overrides Function AddRecord() As Long
        Return 0
    End Function


    Public Overrides Function CheckValues() As Boolean
        If 事由.Length = 0 Then
            Return False
        ElseIf 受付日.Year = 1 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Overrides Sub SetUpdateRow(ByRef pUpdateRow As HimTools2012.Data.UpdateRow)

    End Sub
End Class
