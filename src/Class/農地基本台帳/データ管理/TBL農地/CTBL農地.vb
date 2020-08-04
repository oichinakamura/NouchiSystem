
Public Class CTBL農地
    Inherits CTBL農地共通

    Public Sub New(ByRef DSet As DataSet, ByRef pTable As DataTable)
        MyBase.New(DSet, pTable)


        With Me.Body
            .Columns.Add(New DataColumn("登記簿地目名", GetType(String), "Parent(R登記地目).名称"))
            .Columns.Add(New DataColumn("現況地目名", GetType(String), "Parent(R現況地目).名称"))
            .Columns.Add(New DataColumn("農委地目名", GetType(String), "Parent(R農委地目).名称"))

            DSet.Relations.Add(New DataRelation("R適用法令", DSet.Tables("V_適用法令").Columns("ID"), .Columns("小作地適用法"), False))
            .Columns.Add(New DataColumn("適用法令", GetType(String), "IIF(自小作別=0,Null,Parent(R適用法令).名称)"))

            DSet.Relations.Add(New DataRelation("小作形態", DSet.Tables("V_小作形態").Columns("ID"), .Columns("小作形態"), False))
            .Columns.Add(New DataColumn("小作形態種別", GetType(String), "IIF(自小作別=0,Null,Parent(小作形態).名称)"))

            .Columns.Add(New DataColumn("農地状況名", GetType(String), "Parent(R農地状況).名称"))
            '.Columns.Add(New DataColumn("農振法区分名", GetType(String), "IIF(農振法区分=1,'農用地',IIF(農振法区分=2,'農振地',IIF(農振法区分=3,'農振外','-')))"))
        End With

        DSet.Relations.Add("R所有者", App農地基本台帳.TBL個人.Columns("ID"), Columns("所有者ID"), False)
        DSet.Relations.Add("管理者", App農地基本台帳.TBL個人.Columns("ID"), Columns("管理者ID"), False)
        DSet.Relations.Add("名義人", App農地基本台帳.TBL個人.Columns("ID"), Columns("登記名義人ID"), False)
        DSet.Relations.Add("相続人", App農地基本台帳.TBL個人.Columns("ID"), Columns("推測耕作者ID"), False)
        DSet.Relations.Add("借受人", App農地基本台帳.TBL個人.Columns("ID"), Columns("借受人ID"), False)
        DSet.Relations.Add("経由農業生産法人", App農地基本台帳.TBL個人.Columns("ID"), Columns("経由農業生産法人ID"), False)

        Me.Columns.Add(New DataColumn("所有者郵便番号", GetType(String), "Parent(R所有者).郵便番号"))
        Me.Columns.Add(New DataColumn("所有者氏名", GetType(String), "Parent(R所有者).氏名"))
        Me.Columns.Add(New DataColumn("所有者住所", GetType(String), "Parent(R所有者).住所"))
        Me.Columns.Add(New DataColumn("所有者住民区分", GetType(String), "Parent(R所有者).住民区分"))
        Me.Columns.Add(New DataColumn("管理者氏名", GetType(String), "Parent(管理者).氏名"))
        Me.Columns.Add(New DataColumn("管理者住所", GetType(String), "Parent(管理者).住所"))

        Me.Columns.Add(New DataColumn("名義人氏名", GetType(String), "Parent(名義人).氏名"))
        Me.Columns.Add(New DataColumn("相続者名", GetType(String), "Parent(相続人).氏名"))

        Me.Columns.Add(New DataColumn("借受人郵便番号", GetType(String), "IIF(自小作別=0,'',Parent(借受人).郵便番号)"))
        Me.Columns.Add(New DataColumn("借受人氏名", GetType(String), "IIF(自小作別=0,'',Parent(借受人).氏名)"))
        Me.Columns.Add(New DataColumn("借受人住所", GetType(String), "IIF(自小作別=0,'',Parent(借受人).住所)"))
        Me.Columns.Add(New DataColumn("借受人個人法人の別", GetType(String), "IIF(Parent(借受人).性別=3,'法人','個人')"))
        Me.Columns.Add(New DataColumn("認定農家区分", GetType(Integer), "IIF(自小作別 = 0, 0, Parent(借受人).農業改善計画認定)"))

        'Me.Columns.Add(New DataColumn("経由農業生産法人ID", GetType(Decimal), "IIF(自小作別 = 0, 0, 経由農業生産法人ID)"))
        Me.Columns.Add(New DataColumn("経由農業生産法人名", GetType(String), "IIF(自小作別 = 0, '', IIF(経由農業生産法人ID=0,'',Parent(経由農業生産法人).氏名))"))
        Me.Columns.Add(New DataColumn("小作料表示", GetType(String), "IIF(自小作別>0 AND 小作形態=1, 小作料 + 小作料単位, '')"))

        Dim bChk As TableCheckAndUpdate.CheckResult = TableCheckAndUpdate.CheckResult.NoUpdate
        Dim pTBL As DataTable
        Do
            bChk = TableCheckAndUpdate.CheckResult.NoUpdate
            pTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID]=0")
            Dim nCount As Integer = 0
            Do Until SysAD.DB(sLRDB).ResultMessage.Length = 0 OrElse SysAD.DB(sLRDB).ResultMessage = "OK"
                pTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID]=0")
                nCount += 1
                If nCount > 3 Then
                    Stop
                End If
            Loop
            With New 農地InfoUpdate("D:農地Info", pTBL)
                bChk = .CheckStart

                If bChk = TableCheckAndUpdate.CheckResult.CompleteUpdate Then
                    CompTable.Comp("D:農地Info", "D_転用農地")
                    CompTable.Comp("D:農地Info", "D_削除農地")
                End If
            End With
        Loop Until bChk = TableCheckAndUpdate.CheckResult.NoUpdate
        Me.MergePlus(pTBL)


        DataInitAfter(DSet)
        'Catch ex As Exception
        '    MsgBox("CTBL農地(" & sError & "):" & ex.Message)
        'End Try
    End Sub
End Class


Public Class CTBL転用農地
    Inherits CTBL農地共通


    Public Sub New(ByRef DSet As DataSet, ByRef pTable As DataTable)
        MyBase.New(DSet, pTable)

        CompTable.Comp("D:農地Info", "D_転用農地")

        With Me
         
            DSet.Relations.Add(New DataRelation("転用登記地目", DSet.Tables("V_地目").Columns("ID"), .Columns("登記簿地目"), False))
            .Columns.Add(New DataColumn("登記簿地目名", GetType(String), "Parent(転用登記地目).名称"))

            DSet.Relations.Add(New DataRelation("転用現況地目", DSet.Tables("V_現況地目").Columns("ID"), .Columns("現況地目"), False))
            .Columns.Add(New DataColumn("現況地目名", GetType(String), "Parent(転用現況地目).名称"))

            DSet.Relations.Add(New DataRelation("転用農委地目", DSet.Tables("V_農委地目").Columns("ID"), .Columns("農委地目ID"), False))
            .Columns.Add(New DataColumn("農委地目名", GetType(String), "Parent(転用農委地目).名称"))
            .Columns.Add(New DataColumn("転用済み", GetType(Boolean), "True"))
        End With

        DSet.Relations.Add("転用所有者", App農地基本台帳.TBL個人.Columns("ID"), Me.Columns("所有者ID"), False)
        DSet.Relations.Add("転用管理者", App農地基本台帳.TBL個人.Columns("ID"), Me.Columns("管理者ID"), False)
        DSet.Relations.Add("転用名義人", App農地基本台帳.TBL個人.Columns("ID"), Me.Columns("登記名義人ID"), False)
        DSet.Relations.Add("転用借受人", App農地基本台帳.TBL個人.Columns("ID"), Me.Columns("借受人ID"), False)
        DSet.Relations.Add("転用経由農業生産法人", App農地基本台帳.TBL個人.Columns("ID"), Me.Columns("経由農業生産法人ID"), False)

        Me.Columns.Add(New DataColumn("所有者氏名", GetType(String), "Parent(転用所有者).氏名"))
        Me.Columns.Add(New DataColumn("所有者住所", GetType(String), "Parent(転用所有者).住所"))
        Me.Columns.Add(New DataColumn("管理者氏名", GetType(String), "Parent(転用管理者).氏名"))
        Me.Columns.Add(New DataColumn("管理者住所", GetType(String), "Parent(転用管理者).住所"))

        Me.Columns.Add(New DataColumn("名義人氏名", GetType(String), "Parent(転用名義人).氏名"))
        Me.Columns.Add(New DataColumn("借受人氏名", GetType(String), "IIF(自小作別=0,'',Parent(転用借受人).氏名)"))
        Me.Columns.Add(New DataColumn("借受人住所", GetType(String), "IIF(自小作別=0,'',Parent(転用借受人).住所)"))
        Me.Columns.Add(New DataColumn("経由農業生産法人名", GetType(String), "IIF(自小作別=0,'',IIF(経由農業生産法人ID=0,'',Parent(転用経由農業生産法人).氏名))"))

        DataInitAfter(DSet)
    End Sub

End Class

Public Class CTBL削除農地
    Inherits CTBL農地共通

    Public Sub New(ByRef DSet As DataSet, ByRef pTable As DataTable)
        MyBase.New(DSet, pTable)

        CompTable.Comp("D:農地Info", "D_削除農地")

        With Me
           

            DSet.Relations.Add(New DataRelation("削除登記地目", DSet.Tables("V_地目").Columns("ID"), .Columns("登記簿地目"), False))
            .Columns.Add(New DataColumn("登記簿地目名", GetType(String), "Parent(削除登記地目).名称"))

            DSet.Relations.Add(New DataRelation("削除現況地目", DSet.Tables("V_現況地目").Columns("ID"), .Columns("現況地目"), False))
            .Columns.Add(New DataColumn("現況地目名", GetType(String), "Parent(削除現況地目).名称"))
        End With

        DSet.Relations.Add("削除所有者", App農地基本台帳.TBL個人.Columns("ID"), Me.Columns("所有者ID"), False)
        DSet.Relations.Add("削除管理者", App農地基本台帳.TBL個人.Columns("ID"), Me.Columns("管理者ID"), False)
        DSet.Relations.Add("削除名義人", App農地基本台帳.TBL個人.Columns("ID"), Me.Columns("登記名義人ID"), False)
        DSet.Relations.Add("削除借受人", App農地基本台帳.TBL個人.Columns("ID"), Me.Columns("借受人ID"), False)
        DSet.Relations.Add("削除経由農業生産法人", App農地基本台帳.TBL個人.Columns("ID"), Me.Columns("経由農業生産法人ID"), False)

        Me.Columns.Add(New DataColumn("所有者氏名", GetType(String), "Parent(削除所有者).氏名"))
        Me.Columns.Add(New DataColumn("所有者住所", GetType(String), "Parent(削除所有者).住所"))
        Me.Columns.Add(New DataColumn("管理者氏名", GetType(String), "Parent(削除管理者).氏名"))
        Me.Columns.Add(New DataColumn("管理者住所", GetType(String), "Parent(削除管理者).住所"))

        Me.Columns.Add(New DataColumn("名義人氏名", GetType(String), "Parent(削除名義人).氏名"))
        Me.Columns.Add(New DataColumn("借受人氏名", GetType(String), "IIF(自小作別=0,'',Parent(削除借受人).氏名)"))
        Me.Columns.Add(New DataColumn("借受人住所", GetType(String), "IIF(自小作別=0,'',Parent(削除借受人).住所)"))

        Me.Columns.Add(New DataColumn("経由農業生産法人名", GetType(String), "IIF(自小作別=0,'',IIF(経由農業生産法人ID=0,'',Parent(削除経由農業生産法人).氏名))"))

        DataInitAfter(DSet)
    End Sub
End Class

Public Class CTBL筆情報
    Inherits HimTools2012.Data.DataTableWith

    Public Sub New(ByRef DSet As DataSet, ByRef pTable As DataTable)
        MyBase.New(pTable, s地図情報)

    End Sub
End Class


Public Class 農地InfoUpdate
    Inherits TableCheckAndUpdate

    Public Sub New(ByVal sTableName As String, ByRef pTarget As DataTable)
        MyBase.New(SysAD.DB(sLRDB), sTableName, pTarget, SysAD.DB(sLRDB).UpdateLog)
    End Sub


    Public Overrides Function CheckStart() As TableCheckAndUpdate.CheckResult
        Dim bCheck As CheckResult = Check農地基本修正20150619(bCheck)
        bCheck = Check農地基本削除20100101(bCheck)
        bCheck = Check生産緑地法20150317(bCheck)
        bCheck = Check各種交付金補助金20150129(bCheck)
        bCheck = Check納税猶予の適用状況20150304(bCheck)
        bCheck = Check農地中間管理権と利用配分計画等20150316(bCheck)
        bCheck = Check農地中間管理機構等との協議等20150129(bCheck)
        bCheck = Check特定作業受委託20150305(bCheck)
        bCheck = Check所有共有耕作者20150317(bCheck)
        bCheck = Check利用状況調査20150123(bCheck)
        bCheck = Check利用意向調査20150123(bCheck)
        bCheck = Check登記名義関連H270901(bCheck)
        bCheck = Check不要項目の削除20200116(bCheck)
        Return bCheck
    End Function


    ''' <summary>
    ''' 農地の基本部分
    ''' </summary>
    ''' <param name="bCheck"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Check農地基本修正20150619(bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then

            MakeTBLFieldModifySQL("耕作放棄解消区分", "LONG")
            MakeTBLFieldModifySQL("耕作放棄解消年月日", "DATETIME")
            MakeTBLFieldModifySQL("解除条件付きの農地の貸借", "BIT")
            MakeTBLFieldModifySQL("固定照合", "LONG")
            MakeTBLFieldModifySQL("固定異動日", "DATETIME")
            MakeTBLFieldModifySQL("固定照合日", "DATETIME")
            MakeTBLFieldModifySQL("本番", "LONG")
            MakeTBLFieldModifySQL("元ID", "DECIMAL(18,0)")
            MakeTBLFieldDropSQL("行政区ID")

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    ''' <summary>
    ''' 2014/6農地法改正による管理項目の追加
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Check利用状況調査20150123(bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            MakeTBLFieldModifySQL("利用状況調査日", "DATETIME")
            MakeTBLFieldModifySQL("利用状況調査農地法", "LONG")
            MakeTBLFieldModifySQL("利用状況調査荒廃", "LONG")
            MakeTBLFieldModifySQL("利用状況調査転用", "LONG")

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    Public Function Check利用意向調査20150123(bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            MakeTBLFieldModifySQL("利用意向調査日", "DATETIME")
            MakeTBLFieldModifySQL("利用意向根拠条項", "LONG")
            MakeTBLFieldModifySQL("利用意向意思表明日", "DATETIME")
            MakeTBLFieldModifySQL("利用意向意向内容区分", "LONG")
            MakeTBLFieldModifySQL("利用意向権利関係調査区分", "LONG")
            MakeTBLFieldModifySQL("利用意向権利関係調査記録", "TEXT")
            MakeTBLFieldModifySQL("利用意向公示年月日", "DATETIME")
            MakeTBLFieldModifySQL("利用意向通知年月日", "DATETIME")

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    Public Function Check生産緑地法20150317(ByVal bCheck As CheckResult) As CheckResult

        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            MakeTBLFieldModifySQL("耕地番号", "LONG")
            MakeTBLFieldModifySQL("本地面積", "LONG")
            MakeTBLFieldModifySQL("農振法区分", "LONG")
            MakeTBLFieldModifySQL("都市計画法区分", "LONG")
            MakeTBLFieldModifySQL("生産緑地法種別", "LONG")
            MakeTBLFieldModifySQL("生産緑地法指定日", "DATETIME")

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    Public Function Check所有共有耕作者20150317(ByVal bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            MakeTBLFieldModifySQL("所有者農地意向", "LONG")
            MakeTBLFieldModifySQL("農地法第52公表同意", "LONG")
            MakeTBLFieldModifySQL("共有地区分", "LONG")
            MakeTBLFieldModifySQL("共有者ID", "LONG")
            MakeTBLFieldModifySQL("共有者持分割合分子", "LONG")
            MakeTBLFieldModifySQL("共有者持分割合分母", "LONG")
            MakeTBLFieldModifySQL("耕作者整理番号", "LONG")
            MakeTBLFieldDropSQL("10アール賃借料")
            MakeTBLFieldModifySQL("10a賃借料", "CURRENCY")
            MakeTBLFieldModifySQL("分属管理番号", "DECIMAL")

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    Public Function Check特定作業受委託20150305(ByVal bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            MakeTBLFieldModifySQL("特定作業者ID", "LONG")
            MakeTBLFieldModifySQL("特定作業者名", "TEXT")
            MakeTBLFieldModifySQL("特定作業者住所", "TEXT")
            MakeTBLFieldDropSQL("特定作業作目")
            MakeTBLFieldModifySQL("特定作業作目種別", "TEXT")
            MakeTBLFieldModifySQL("特定作業内容", "TEXT")

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    Public Function Check農地中間管理機構等との協議等20150129(bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            MakeTBLFieldModifySQL("農地法35の1通知日", "DATETIME")
            MakeTBLFieldModifySQL("農地法35の2通知日", "DATETIME")
            MakeTBLFieldModifySQL("農地法35の3通知日", "DATETIME")
            MakeTBLFieldModifySQL("勧告年月日", "DATETIME")
            MakeTBLFieldModifySQL("勧告内容", "LONG")
            MakeTBLFieldModifySQL("中間管理勧告日", "DATETIME")
            MakeTBLFieldModifySQL("再生利用困難農地", "LONG")
            MakeTBLFieldModifySQL("農地法40裁定公告日", "DATETIME")
            MakeTBLFieldModifySQL("農地法43裁定公告日", "DATETIME")
            MakeTBLFieldModifySQL("農地法44の1裁定公告日", "DATETIME")
            MakeTBLFieldModifySQL("農地法44の3裁定公告日", "DATETIME")

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    Public Function Check農地中間管理権と利用配分計画等20150316(ByVal bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            MakeTBLFieldModifySQL("中間管理権取得日", "DATETIME")
            MakeTBLFieldModifySQL("意見回答日", "DATETIME")
            MakeTBLFieldModifySQL("知事公告日", "DATETIME")
            MakeTBLFieldModifySQL("認可通知日", "DATETIME")
            MakeTBLFieldModifySQL("権利設定内容", "LONG")
            MakeTBLFieldModifySQL("利用配分設定期間", "DATETIME")
            MakeTBLFieldModifySQL("利用配分計画設定期間年", "LONG")
            MakeTBLFieldModifySQL("利用配分計画設定期間月", "LONG")
            MakeTBLFieldModifySQL("利用配分計画始期日", "DATETIME")
            MakeTBLFieldModifySQL("利用配分計画終期日", "DATETIME")
            MakeTBLFieldModifySQL("利用配分計画借賃額", "LONG")
            MakeTBLFieldDropSQL("利用配分計画10アール賃借料")
            MakeTBLFieldModifySQL("利用配分計画10a賃借料", "CURRENCY")
            MakeTBLFieldModifySQL("貸借契約解除年月日", "DATETIME")

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    Public Function Check納税猶予の適用状況20150304(ByVal bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            MakeTBLFieldModifySQL("納税猶予種別", "LONG")
            MakeTBLFieldModifySQL("納税猶予相続日", "DATETIME")
            MakeTBLFieldModifySQL("納税猶予適用日", "DATETIME")
            MakeTBLFieldModifySQL("納税猶予継続日", "DATETIME")

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    Public Function Check各種交付金補助金20150129(bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            MakeTBLFieldModifySQL("環境保全交付金", "LONG")
            MakeTBLFieldModifySQL("農地維持交付金", "LONG")
            MakeTBLFieldModifySQL("資源向上交付金", "LONG")
            MakeTBLFieldModifySQL("中山間直接支払", "LONG")
            MakeTBLFieldModifySQL("特定処分対象農地等", "LONG")

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    Public Function Check登記名義関連H270901(bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            If MakeTBLFieldModifySQL("登記名義人住所", "VARCHAR(255)") Or
                MakeTBLFieldModifySQL("登記名義人氏名", "VARCHAR(255)") Then

                Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
            Else
                Return bCheck
            End If
        Else
            Return bCheck
        End If
    End Function

    Public Function Check農地基本削除20100101(ByVal bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            MakeTBLFieldDropSQL("旧所有世帯ID")
            MakeTBLFieldDropSQL("旧所有者ID")
            MakeTBLFieldDropSQL("旧借受世帯ID")
            MakeTBLFieldDropSQL("町名ID")

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    Public Function Check不要項目の削除20200116(bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then

            MakeTBLFieldDropSQL("固定照合日")
            MakeTBLFieldDropSQL("固定異動日")

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

End Class
