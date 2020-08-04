

Public Class CTBL土地履歴
    Inherits HimTools2012.Data.DataTableWith

    Public Sub New(ByRef DSet As DataSet, ByRef pTable As DataTable)
        MyBase.New(pTable, sLRDB)
        With Me
            DataInitAfter(Me.Body.DataSet)
        End With
    End Sub

    Public Sub 履歴消去(ByVal sWhere As String)
        
        Dim pView As New DataView(Me.Body, sWhere, "", DataViewRowState.CurrentRows)
        Dim pList As New List(Of Integer)
        For Each pRow As DataRowView In pView
            pList.Add(pRow.Item("ID"))
        Next

        For Each nID As Integer In pList
            Try
                Me.Rows.Remove(Me.Rows.Find(nID))
            Catch ex As Exception

            End Try
        Next
    End Sub

End Class
