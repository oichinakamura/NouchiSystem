
Public Class CTBL固定情報
    Inherits HimTools2012.Data.DataTableEx

    Public Sub New(ByRef DSet As DataSet)
        MyBase.New("M_固定情報")
    End Sub
    Public Overrides Sub InitTable()
        With Me
            .Columns.Add(New DataColumn("大字名", GetType(String), "Parent(R固大字).名称"))
            .Columns.Add(New DataColumn("小字名", GetType(String), "Parent(R固小字).名称"))
            .Columns.Add(New DataColumn("登記地目名", GetType(String), "Parent(R固登地目).名称"))
            .Columns.Add(New DataColumn("現況地目名", GetType(String), "Parent(R固現地目).名称"))
            DataInitAfter(Me.DataSet, "nID")
        End With

        MyBase.InitTable()
    End Sub
End Class
