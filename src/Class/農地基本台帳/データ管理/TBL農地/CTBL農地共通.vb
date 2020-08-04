
Public MustInherit Class CTBL農地共通
    Inherits HimTools2012.Data.DataTableWith

    Public Sub New(ByRef DSet As DataSet, ByRef pTable As DataTable)
        MyBase.New(pTable, sLRDB)

    End Sub

    Public Function GetDefaultTable() As DataTable
        Dim mvarTBL As New DataTable("D:農地Info")
        mvarTBL.Columns.Add("ID", GetType(Integer))
        mvarTBL.Columns.Add("所有者ID", App農地基本台帳.TBL個人.Columns("ID").DataType)
        mvarTBL.Columns.Add("管理者ID", App農地基本台帳.TBL個人.Columns("ID").DataType)
        mvarTBL.Columns.Add("借受人ID", App農地基本台帳.TBL個人.Columns("ID").DataType)
        mvarTBL.Columns.Add("登記名義人ID", App農地基本台帳.TBL個人.Columns("ID").DataType)
        mvarTBL.Columns.Add("経由農業生産法人ID", App農地基本台帳.TBL個人.Columns("ID").DataType)

        Return mvarTBL
    End Function

    Protected Sub CheckDBType(mvarTBL As DataTable, ByVal sSQL As String)
        Do
            Dim sRet As String = ""
            Try
                mvarTBL.Merge(SysAD.DB(sLRDB).GetTableBySqlSelect(sSQL), False, MissingSchemaAction.AddWithKey)

                Exit Do
            Catch ex As Exception
                Dim sErr As String = ex.Message
                If InStr(sErr, "は競合するプロパティがあります : DataType プロパティの不一致") > 0 Then
                    Dim sTG As String = HimTools2012.StringF.Mid(sErr, InStr(sErr, "<target>.") + 9)
                    sTG = HimTools2012.StringF.Left(sTG, InStr(sTG, " ") - 1)

                    Select Case mvarTBL.Columns(sTG).DataType.ToString
                        Case "System.Decimal" : sRet = SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [{0}] ALTER COLUMN [" & sTG & "] DECIMAL", Me.Body.TableName)
                        Case "System.Int32" : sRet = SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [{0}] ALTER COLUMN [" & sTG & "] LONG", Me.Body.TableName)
                        Case Else
                            Stop
                    End Select

                    If sRet.ToUpper = "OK" OrElse sRet = "" Then

                    ElseIf InStr(sRet, "現在ほかのユーザーまたはプロセスで使用されている") > 0 Then
                        If MsgBox("データベースに不正な型情報が見つかりました。修正の為、他のPCの「農地台帳システム」及び旧台帳システムを終了してください。継続する場合は「はい」、中断する場合は「いいえ」を押してください。", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                            End
                        End If
                    Else
                        Stop
                    End If

                End If
            End Try
        Loop
    End Sub

    Public Function MinID() As Decimal
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Min(U_農地.ID) AS MinID FROM (Select ID FROM [D:農地Info] UNION Select ID FROM [D_転用農地]  UNION Select ID FROM [D_削除農地])  AS U_農地;")
        Dim n As Decimal = pTBL.Rows(0).Item("MinID")
        Return IIf(n >= 0, -1, n)
    End Function

    Public Function Update(ByVal mvarUpdateRow As HimTools2012.Data.UpdateRow, bAddNew As Boolean) As Boolean
        If bAddNew Then
            Stop

            Return True
        Else
            Return SysAD.DB(sLRDB).UpdateRecord(Me.Body, mvarUpdateRow)
        End If
    End Function

    Public Overrides Sub MergePlus(pTable As System.Data.DataTable, Optional preserveChanges As Boolean = False, Optional pAction As System.Data.MissingSchemaAction = System.Data.MissingSchemaAction.Add)
        SyncLock Me
            MyBase.MergePlus(pTable, preserveChanges, pAction)


            With New C農地関連情報取り込み(Me.Body)
                .Dialog.StartProc(pTable.Rows.Count > 10, False)
                If .Dialog._objException Is Nothing = False Then
                    If .Dialog._objException.Message = "Cancel" Then
                        MsgBox("処理を中止しました。　", , "処理中止")
                    Else

                    End If
                End If
            End With
        End SyncLock
    End Sub
End Class


Public Class C農地関連情報取り込み
    Inherits HimTools2012.clsAccessor
    Private mvarBody As DataTable

    Public Sub New(pBody As DataTable)
        MyBase.New()
        mvarBody = pBody
    End Sub

    Public Overrides Sub Execute()
        Dim pDic As New List(Of String)

        Message = "所有者名取得.."
        Dim p所有者 = From q In mvarBody Where Not IsDBNull(q.Item("所有者ID")) AndAlso Not q.Item("所有者ID") = 0 AndAlso IsDBNull(q.Item("所有者氏名")) = True
        For Each pRow As DataRow In p所有者
            If Not pDic.Contains(pRow.Item("所有者ID").ToString) Then
                pDic.Add(pRow.Item("所有者ID").ToString)
            End If
        Next

        Message = "管理者名取得.."
        Dim p管理者 = From q In mvarBody Where Not IsDBNull(q.Item("管理者ID")) AndAlso Not q.Item("管理者ID") = 0 AndAlso IsDBNull(q.Item("管理者氏名")) = True
        For Each pRow As DataRow In p管理者
            If Not pDic.Contains(pRow.Item("管理者ID").ToString) Then
                pDic.Add(pRow.Item("管理者ID").ToString)
            End If
        Next

        Message = "借受者名取得.."
        Dim p借受人 = From q In mvarBody Where Not IsDBNull(q.Item("借受人ID")) AndAlso Not q.Item("借受人ID") = 0 AndAlso IsDBNull(q.Item("借受人氏名")) = True
        For Each pRow As DataRow In p借受人
            If Not pDic.Contains(pRow.Item("借受人ID").ToString) Then
                pDic.Add(pRow.Item("借受人ID").ToString)
            End If
        Next

        Message = "経由法人名取得.."
        Dim p経由法人 = From q In mvarBody Where Not IsDBNull(q.Item("経由農業生産法人ID")) AndAlso Not q.Item("経由農業生産法人ID") = 0 AndAlso IsDBNull(q.Item("経由農業生産法人名")) = True
        For Each pRow As DataRow In p経由法人
            If Not pDic.Contains(pRow.Item("経由農業生産法人ID").ToString) Then
                pDic.Add(pRow.Item("経由農業生産法人ID").ToString)
            End If
        Next


        If pDic.Count > 0 Then
            Value = 0
            Maximum = pDic.Count
            Message = String.Format("データベースより関係者名取得..({0}/{1})", Value, Maximum)

            Dim sB As New System.Text.StringBuilder
            Dim sC As String = ""
            For i As Integer = 0 To pDic.Count - 1
                sB.Append(sC & pDic.Item(i))


                If (i Mod 64) = 63 AndAlso sB.Length > 0 Then
                    Message = String.Format("データベースより関係者名取得..({0}/{1})", i, pDic.Count)
                    Value = i

                    Dim pAddK As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE ID In (" & sB.ToString & ")")
                    App農地基本台帳.TBL個人.MergePlus(pAddK)
                    sB.Clear()
                    sC = ""
                Else
                    sC = ","
                End If
            Next

            If sB.Length > 0 Then
                Dim pAddK As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE ID In (" & sB.ToString & ")")
                App農地基本台帳.TBL個人.MergePlus(pAddK)
            End If
        End If
    End Sub

    Private Function Get検索文字(ByVal St As String) As String
        St = Replace(St, " ", "")
        Return St
    End Function
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class

Public Class CompTable
    Public Shared Sub Comp(ByVal sTBL1 As String, ByVal sTBL2 As String)
        Dim pTBLA As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [" & sTBL1 & "] WHERE ID=0")
        Dim pTBLB As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [" & sTBL2 & "] WHERE ID=0")

        For Each pCol As DataColumn In pTBLA.Columns
            If Not pTBLB.Columns.Contains(pCol.ColumnName) Then
                Select Case pCol.DataType.FullName
                    Case "System.Int32", "System.Int16"
                        SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [" & sTBL2 & "] ADD [" & pCol.ColumnName & "] LONG")
                    Case "System.Double", "System.Single"
                        SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [" & sTBL2 & "] ADD [" & pCol.ColumnName & "] DOUBLE")
                    Case "System.DateTime"
                        SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [" & sTBL2 & "] ADD [" & pCol.ColumnName & "] DATETIME")
                    Case "System.String"
                        SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [" & sTBL2 & "] ADD [" & pCol.ColumnName & "] VARCHAR(255);")
                    Case "System.Boolean"
                        SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [" & sTBL2 & "] ADD [" & pCol.ColumnName & "] BIT")
                    Case "System.Decimal"
                        SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [" & sTBL2 & "] ADD [" & pCol.ColumnName & "] DECIMAL")
                    Case Else
                        MsgBox("管理項目が異なります。久永情報マネジメントに連絡をお願いします。")
                        End
                End Select
            End If
        Next
        For Each pCol As DataColumn In pTBLB.Columns
            If Not pTBLA.Columns.Contains(pCol.ColumnName) Then
                Select Case pCol.DataType.FullName
                    Case "System.Int32", "System.DateTime"
                        SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [" & sTBL2 & "] DROP [" & pCol.ColumnName & "]")
                    Case "System.Decimal"
                        SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [" & sTBL2 & "] DROP [" & pCol.ColumnName & "]")
                    Case "System.Boolean"
                        SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [" & sTBL2 & "] DROP [" & pCol.ColumnName & "]")
                    Case "System.String"
                        SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [" & sTBL2 & "] DROP [" & pCol.ColumnName & "]")
                    Case Else
                        MsgBox("管理項目が異なります。久永情報マネジメントに連絡をお願いします。")
                        End
                End Select
            End If
        Next
    End Sub
End Class