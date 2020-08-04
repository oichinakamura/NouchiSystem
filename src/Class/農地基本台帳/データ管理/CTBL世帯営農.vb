

Public Class CTBL世帯営農
    Inherits HimTools2012.Data.DataTableWith

    Public Sub New(DSet As DataSet, ByRef pTable As DataTable)
        MyBase.New(pTable, sLRDB)
        Dim bReLoad As Boolean = False

        With Me.Body
            .Columns.Add(New DataColumn("Key", GetType(String), "'営農情報.' + [ID]"))
            .Columns.Add(New DataColumn("アイコン", GetType(String), "'営農情報'"))

            'DSet.Relations.Add(New DataRelation("営農情報", App農地基本台帳.TBL世帯.Columns("ID"), .Columns("ID"), False))

            '.Columns.Add(New DataColumn("世帯営農氏名", GetType(String), "Parent(営農情報).世帯主氏名"))
            '.Columns.Add(New DataColumn("世帯営農住所", GetType(String), "Parent(営農情報).住所"))
        End With

        DataInitAfter(DSet)
    End Sub


End Class
