Imports HimTools2012
Imports HimTools2012.TargetSystem
Imports HimTools2012.Excel.XMLSS2003

Module mod農地基本台帳

    Public Sub 耕作面積証明印刷(ByVal sKey As String)
        Dim p耕作証明書 As New CPrint耕作証明書(sKey)
    End Sub


    Public Sub 耕作多筆証明印刷(ByVal sKey As String)
        Dim objAcc As New CPrint耕作多筆証明願(sKey)
    End Sub

    Public Sub 基本台帳印刷(ByVal sKey As String, ByVal nViewMode As ExcelViewMode, ByVal pPrintMode As 印刷Mode)
        Dim sMode As String = IIf(pPrintMode = 0, "旧", "")
        Dim sFileName As String = SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & sMode & "農地基本台帳様式.xml"
        If Not IO.File.Exists(sFileName) Then
            sFileName = SysAD.SystemInfo.ApplicationDirectory & "\" & sMode & "農地基本台帳様式.xml"
        End If

        If IO.File.Exists(sFileName) Then
            Dim nSID As Long = CommonFunc.GetKeyCode(sKey)
            Dim nPID As Long = 0
            'Try
            Select Case CommonFunc.GetKeyHead(sKey)
                Case "農家"
                    Dim pView As DataView = New DataView(App農地基本台帳.TBL個人.Body, String.Format("[世帯ID]={0}", nSID), "", DataViewRowState.CurrentRows)
                    If pView.Count > 0 Then
                        For Each pRow As DataRowView In pView
                            Dim p個人 As CObj個人 = ObjectMan.GetObject("個人." + pRow.Item("ID").ToString)
                            sub農地の関連補正(p個人)
                        Next
                    End If
                Case "個人"
                    Dim p個人 As CObj個人 = ObjectMan.GetObject(sKey)
                    sub農地の関連補正(p個人)

                    nSID = p個人.世帯ID
                    nPID = p個人.ID
                Case Else
                    MsgBox("指定されたオブジェクトから基本台帳は出力できません。")
            End Select

            ' 処理中ダイアログ表示
            Dim objAcc As New CPrint基本台帳(New CXMLSS2003(TextAdapter.LoadTextFile(sFileName)), nSID, nPID, pPrintMode, nViewMode)

            With objAcc
                .Dialog.StartProc(True, True)

                If .Dialog._objException Is Nothing = False Then
                    If .Dialog._objException.Message = "Cancel" Then
                        MsgBox("処理を中止しました。　", , "処理中止")
                    Else
                        'Throw objDlg._objException
                    End If
                Else
                    .SaveAndOpen(nViewMode) '# デスクトップにファイルを保存する処理示
                End If
            End With
            'Catch ex As Exception

            'End Try

        Else
            MsgBox("ファイルがありません")
        End If
    End Sub

    Public Sub 農地台帳印刷(ByVal sKey As String, ByVal nViewMode As ExcelViewMode, ByVal PrintType As String)
        Dim sFileName As String = My.Application.Info.DirectoryPath & "\農地台帳" & PrintType & "様式.xml"
        If IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\農地台帳" & PrintType & "様式.xml") Then
            sFileName = SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\農地台帳" & PrintType & "様式.xml"
        End If

        If IO.File.Exists(sFileName) Then
            Dim nID As Integer = CommonFunc.GetKeyCode(sKey)

            Dim objAcc As New CPrint農地台帳(New CXMLSS2003(TextAdapter.LoadTextFile(sFileName)), nID, nViewMode)
            With objAcc
                .Dialog.StartProc(True, True)

                If .Dialog._objException Is Nothing = False Then
                    If .Dialog._objException.Message = "Cancel" Then
                        MsgBox("処理を中止しました。　", , "処理中止")
                    Else
                    End If
                Else
                    .SaveAndOpen(nViewMode, PrintType)
                End If
            End With
        Else
            MsgBox("ファイルがありません")
        End If
    End Sub

    Private Sub sub農地の関連補正(ByVal p個人 As CObj個人)
        If p個人.世帯ID = 0 Then
            'MsgBox("世帯番号が設定されていません", MsgBoxStyle.Critical)
        Else
            Dim p関連農地Tbl As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE ([D:農地Info].[所有者ID]={0} AND [D:農地Info].所有世帯ID=0) OR ([D:農地Info].[管理者ID]={0} AND [D:農地Info].管理世帯ID=0) OR ([D:農地Info].[借受人ID]={0} AND [D:農地Info].借受世帯ID=0);", p個人.ID)

            If p関連農地Tbl.Rows.Count > 0 Then
                App農地基本台帳.TBL農地.MergePlus(p関連農地Tbl)
                Dim pView As New DataView(App農地基本台帳.TBL農地.Body, String.Format("([所有者ID]={0} AND [所有世帯ID]=0) OR ([管理者ID]={0} AND [管理世帯ID]=0) OR ([借受人ID]={0} AND [借受世帯ID]=0)", p個人.ID), "", DataViewRowState.CurrentRows)

                If MsgBox("不整合農地があります。農地と世帯の関連を修復しますか？", vbOKCancel) = vbOK Then
                    For Each pRow As DataRowView In pView
                        Dim sSQL As New System.Text.StringBuilder
                        If Not IsDBNull(pRow.Item("所有者ID")) AndAlso pRow.Item("所有者ID") = p個人.ID Then
                            sSQL.Append(IIf(sSQL.Length > 0, ",", "") & "[所有世帯ID]=" & p個人.世帯ID)
                            pRow.Item("所有世帯ID") = p個人.世帯ID
                        End If
                        If Not IsDBNull(pRow.Item("管理者ID")) AndAlso pRow.Item("管理者ID") = p個人.ID Then
                            sSQL.Append(IIf(sSQL.Length > 0, ",", "") & "[管理世帯ID]=" & p個人.世帯ID)
                            pRow.Item("管理世帯ID") = p個人.世帯ID
                        End If
                        If Not IsDBNull(pRow.Item("借受人ID")) AndAlso pRow.Item("借受人ID") = p個人.ID Then
                            sSQL.Append(IIf(sSQL.Length > 0, ",", "") & "[借受世帯ID]=" & p個人.世帯ID)
                            pRow.Item("借受世帯ID") = p個人.世帯ID
                        End If

                        If sSQL.Length > 0 Then
                            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET {0} WHERE [ID]={1}", sSQL.ToString, pRow.Item("ID"))
                            Make農地履歴(pRow.Item("ID"), Now, Now, 0, enum法令.職権異動, "世帯間の関連付け設定", p個人.ID)
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    Public Sub OpenSQLList(ByVal sTitle As String, ByVal sSQL As String, Optional s合計列() As String = Nothing)
        Dim pList As SQLList
        If Not SysAD.page農家世帯.TabPageContainKey(sSQL) Then
            pList = New SQLList(sTitle, sSQL, True, s合計列)
            pList.Name = sSQL
            SysAD.page農家世帯.中央Tab.AddPage(pList)
            pList.SplitterDistance = 100
        Else
            pList = SysAD.page農家世帯.GetItem(sSQL)
        End If
    End Sub

    Public Sub OpenTableFilterList(ByVal sKey As String, ByVal sTitle As String, ByRef pTable As DataTable, ByVal RowFilter As String, ByVal Sort As String, Optional s合計列() As String = Nothing)
        Dim pList As CTableFilterList
        If Not SysAD.page農家世帯.TabPageContainKey(sKey) Then
            pList = New CTableFilterList(sKey, sTitle, pTable, RowFilter, Sort, True, s合計列)
            pList.Name = sKey
            SysAD.page農家世帯.中央Tab.AddPage(pList)
            pList.SplitterDistance = 100
        Else
            pList = SysAD.page農家世帯.GetItem(sKey)
        End If
    End Sub

    Public Sub 諮問意見書作成()
        Dim pDlg As New dlg諮問意見書選択

        pDlg.Load諮問意見書資料()
        If pDlg.ShowDialog() = DialogResult.OK Then
        End If
    End Sub

    Public Sub 総会資料作成()
        Dim pDlg As New dlg総会選択

        pDlg.Load総会資料()
        If pDlg.ShowDialog() = DialogResult.OK Then
        End If
    End Sub

    Private Function Open申請Wnd(bOpenWindow As Boolean, ID As Integer) As Boolean
        If bOpenWindow Then
            Dim pObj申請 As CTargetObjWithView = ObjectMan.GetObject("申請." & ID)
            If pObj申請 IsNot Nothing Then
                pObj申請.OpenDataViewNext(SysAD.page農家世帯.DataViewCollection)
            Else
                Return False
            End If
        End If
        Return True
    End Function
    Public Function Open農地(ByVal nID As Long) As Boolean
        Dim pRow As DataRow = App農地基本台帳.TBL農地.FindRowByID(nID)

        If pRow Is Nothing Then
            MsgBox("指定された農地の記録がありませんでした｡", vbInformation)
        Else
            Dim pObj As CObj農地 = ObjectMan.GetObjectDB("農地", pRow, GetType(CObj農地), True)
            pObj.OpenDataViewNext(SysAD.page農家世帯.DataViewCollection)
        End If
        Return True
    End Function

    Public Function Open世帯(ByVal nID As System.Int64, ByVal sMessErroe As String) As Boolean
        Dim sKey As String = "農家." & nID
        Dim pObj As CObj農家 = Nothing

        If IsDBNull(nID) OrElse nID = 0 Then
            MsgBox("世帯番号が正しく記録されていません。", vbInformation)
            Return False
        Else
            pObj = ObjectMan.GetObject(sKey)
            If pObj IsNot Nothing Then
                pObj.OpenDataViewNext(SysAD.page農家世帯.DataViewCollection)
                ObjectMan.AddObject(pObj)
            Else
                MsgBox(sMessErroe, MsgBoxStyle.Critical)
            End If
            Return True
        End If

    End Function

    Public Function Open個人(ByVal nID As Int64, ByVal sMessErroe As String) As Boolean
        Dim sKey As String = "個人." & nID
        Dim pObj As CObj個人 = Nothing

        If Not ObjectMan.ObjectCollection.ContainsKey(sKey) Then
            Dim pRow個人 As DataRow = App農地基本台帳.TBL個人.FindRowByID(nID)
            If pRow個人 Is Nothing Then
                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人INFO] WHERE [ID]=" & nID)
                If pTBL.Rows.Count > 0 Then
                    App農地基本台帳.TBL個人.MergePlus(pTBL)
                Else
                    MsgBox(sMessErroe, vbCritical)
                    Return False
                End If
            End If
            pObj = ObjectMan.GetObject(sKey)
            ObjectMan.ObjectCollection.Add(sKey, pObj)
        Else
            pObj = ObjectMan.GetObject(sKey)
        End If

        If pObj.DataViewPage Is Nothing Then
            CType(ObjectMan.GetObject(sKey), CTargetObjWithView).OpenDataViewNext(SysAD.page農家世帯.DataViewCollection)
        Else
            pObj.DataViewPage.Active()
        End If
        Return False
    End Function
End Module

