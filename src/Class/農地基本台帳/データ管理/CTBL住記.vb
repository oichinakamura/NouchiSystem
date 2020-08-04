

Public Class CTBL住記
    Inherits HimTools2012.Data.DataTableWith

    Public Sub New(DSet As DataSet, pTable As DataTable)
        MyBase.New(pTable, sLRDB)
    End Sub


    Public Overrides Sub MergePlus(pTable As System.Data.DataTable, Optional preserveChanges As Boolean = False, Optional pAction As System.Data.MissingSchemaAction = System.Data.MissingSchemaAction.Add)
        MyBase.MergePlus(pTable, preserveChanges, pAction)
    End Sub

    Dim mvarCollection As New Dictionary(Of Integer, String)

 
    Public Overloads Overrides Function GetDataView(sWhere As String, sFilter As String, sOrderBy As String) As System.Data.DataView
        Return Nothing
    End Function
End Class
