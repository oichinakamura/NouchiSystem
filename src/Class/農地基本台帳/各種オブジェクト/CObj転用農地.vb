Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CObj転用農地 : Inherits CTargetObjWithView農地台帳

    Public Sub New(ByVal pRow As DataRow, ByVal bAddNew As Boolean)
        MyBase.New(pRow, bAddNew, New HimTools2012.TargetSystem.DataKey("転用農地", pRow.Item("ID")), "D_転用農地")
    End Sub

    Public Overrides Function CanDropKeyHead(ByVal sKey As String, ByVal sOption As String) As Boolean
        If Not SysAD.IsClickOnceDeployed Then
            Debug.Print(sKey)
            Stop
        End If
        Return False
    End Function


    Public Overrides Function DoCommand(ByVal sCommand As String, ByVal ParamArray sParams() As String) As Object
        Select Case sCommand
            Case "農地への復活" : Sub転用農地の復活(Me)
            Case "農地履歴" : SysAD.page農家世帯.土地履歴リスト.検索開始("[LID]=" & Me.ID, "[LID]=" & Me.ID)
            Case "関連申請" : List関連申請(Me.ID)
            Case "４条申請" : Return New C申請データ作成("転用農地法4条の受付", Me.Key.KeyValue, Nothing)
            Case "５条所有権申請" : Return New C申請データ作成("転用を伴う所有権移転(5条)の申請受付", Me.Key.KeyValue, Nothing)
            Case "５条貸借申請" : Return New C申請データ作成("転用を伴う貸借権設定(5条)の申請受付", Me.Key.KeyValue, Nothing)
            Case Else
                If Not SysAD.IsClickOnceDeployed Then
                    Stop
                End If
        End Select
        Return Nothing
    End Function

    Public Overrides Sub DropDown(ByVal sSourceList As String, Optional ByVal sOption As String = "")

    End Sub


    Public Overrides Function GetContextMenu(Optional ByVal pMenu As HimTools2012.controls.MenuItemEx = Nothing, Optional nDips As Integer = 1, Optional sParam() As String = Nothing) As HimTools2012.controls.MenuPlus
        Dim pMenuf As New HimTools2012.controls.ContextMenuEx(AddressOf ClickMenu)
        Dim bEdit As Boolean = (SysAD.SystemInfo.ユーザー.n権利 > 0)
        pMenuf.AddMenu("農地への復活", , AddressOf ClickMenu)
        pMenuf.AddMenu("農地履歴", , AddressOf ClickMenu)

        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("Select * From [D_土地系図] Where [自ID]=" & Me.ID)
        If pTBL.Rows.Count > 0 Then
            With pMenuf.AddMenu("異動前農地")
                For Each pRow As DataRow In pTBL.Rows
                    .AddSubMenu(pRow.Item("元土地所在"), 2, ObjectMan.GetObject("削除農地." & pRow.Item("元ID")))

                    'Dim p元TBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("Select * From [D_土地系図] Where [自ID]=" & pRow.Item("元ID"))
                    'If p元TBL.Rows.Count > 0 Then
                    '    With .AddMenu("異動前農地")
                    '        For Each p元Row As DataRow In p元TBL.Rows
                    '            .AddSubMenu(p元Row.Item("元土地所在"), 2, ObjectMan.GetObject("削除農地." & p元Row.Item("元ID")))
                    '        Next
                    '    End With
                    'End If
                Next

            End With
        End If

        pMenuf.AddMenu("-", , AddressOf ClickMenu)
        pMenuf.AddMenu("関連申請", , AddressOf ClickMenu)
        With pMenuf.AddMenu("転用")
            .AddMenu("４条申請", , AddressOf ClickMenu, , bEdit)
            .AddMenu("５条所有権申請", , AddressOf ClickMenu, , bEdit)
            .AddMenu("５条貸借申請", , AddressOf ClickMenu, , bEdit)
        End With
        Return pMenuf
    End Function

    Public Overrides Function GetProperty(ByVal sParam As String) As Object
        Try
            Select Case sParam
                Case "管理者ID" : Return Me.GetItem("管理者ID", 0)
                Case "土地所在"
                    Return Me.土地所在
                Case Else
                    Return Me.Row.Body.Item(sParam)
            End Select
            Return ""
        Catch ex As Exception
        End Try
        Return ""
    End Function

    Public ReadOnly Property 土地所在() As String
        Get
            If MyBase.GetStringValue("所在").Length > 0 Then
                Return MyBase.GetStringValue("所在") & MyBase.GetStringValue("地番")
            Else
                Dim sB As New System.Text.StringBuilder
                If Not IsDBNull(Me.Row.Body("大字ID")) AndAlso Me.Row.Body("大字ID") <> 0 Then
                    Dim p大字() As DataRowView = SysAD.MasterView("大字").FindRows(Me.Row.Body("大字ID"))
                    If p大字 IsNot Nothing Then
                        sB.Append(p大字(0).Item("名称"))
                    End If
                End If

                If Not Me.Row.IsZero("小字ID") Then
                    Dim p小字() As DataRowView = SysAD.MasterView("小字").FindRows(Me.Row.Body("小字ID"))
                    If p小字 IsNot Nothing AndAlso p小字(0).Item("名称").ToString.Length > 0 AndAlso Replace(p小字(0).Item("名称").ToString, "-", "").ToString.Length > 0 Then
                        sB.Append("字" & p小字(0).Item("名称"))
                    End If
                End If

                Return sB.ToString & MyBase.GetStringValue("地番")
            End If

            Return MyBase.GetStringValue("所在")
        End Get

    End Property

    Public Sub Sub転用農地の復活(ByVal pDV As HimTools2012.TargetSystem.CTargetObjWithView)
        If MsgBox("転用済み農地を通常農地に戻しますか?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

            Dim sSQL As String = SysAD.DB(sLRDB).GetTranceTable2Table(
                New DataView(App農地基本台帳.TBL転用農地.Body, "[ID]=" & pDV.ID, "", DataViewRowState.CurrentRows),
                SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID]=0"),
                "D:農地Info",
                New String() {"Key", "アイコン"}, , True
            )

            If Len(sSQL) Then
                Dim sRet As String = SysAD.DB(sLRDB).ExecuteSQL(sSQL)
                'If sRet.Length = 0 OrElse sRet.StartsWith("OK") Then
                'Make農地履歴(pDV.DVProperty.ID, Now(), Now(), 0, "農地への復活", 0, 0, 0, "")

                SysAD.DB(sLRDB).ExecuteSQL("DELETE D_転用農地.ID FROM D_転用農地 WHERE D_転用農地.ID=" & pDV.ID)
                App農地基本台帳.TBL転用農地.Rows.Remove(pDV.Row.Body)
                MsgBox("処理しました。", vbInformation)
                'Else
                '    MsgBox("追加に失敗しました。" & sRet)
                'End If
            End If
        End If
    End Sub

    Public Function CopyObject(Optional ByVal NewID As Long? = Nothing) As HimTools2012.TargetSystem.CTargetObjectBase
        Dim sKeyIP As String = ""
        Dim adrList As System.Net.IPAddress() = SysAD.IPAddressList()
        If adrList.Length > 0 Then
            For Each padr As System.Net.IPAddress In adrList
                sKeyIP = Replace(padr.ToString, ".", "")
                If sKeyIP.IndexOf(":") = -1 AndAlso sKeyIP.Length > 4 Then
                    Exit For
                End If
            Next
        End If

        If NewID Is Nothing Then
            Dim pTBLMin As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Min([D:農地Info].ID) AS IDMin FROM [D:農地Info];")
            If pTBLMin.Rows(0).Item("IDMin") = 0 Then
                NewID = -1
            Else
                NewID = CLng(pTBLMin.Rows(0).Item("IDMin") - 1)
            End If
        End If

        Try
            SysAD.DB(sLRDB).ExecuteSQL("SELECT * INTO [農地追加{0}] FROM [D_転用農地] WHERE [ID]={1};", sKeyIP, Me.ID)
            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [農地追加{0}] SET [ID]={2} WHERE [ID]={1}", sKeyIP, Me.ID, NewID)
            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_転用農地] SELECT * FROM 農地追加{0}", sKeyIP)
            SysAD.DB(sLRDB).ExecuteSQL("DROP TABLE [農地追加{0}]", sKeyIP)
        Catch ex As Exception
            Return Nothing
        End Try

        Return ObjectMan.GetObject("転用農地." & NewID)
    End Function

    Public Overrides Function InitDataViewNext(ByRef pDB As HimTools2012.TargetSystem.CDataViewCollection, Optional InterfaceName As String = "") As Boolean
        If Me.DataViewPage Is Nothing Then
            Me.DataViewPage = New DataViewCtrl.DataViewNext転用農地(Me)
        End If
        Return True
    End Function
    Public Overrides ReadOnly Property DataTableWithUpdateList As HimTools2012.Data.DataTableWithUpdateList
        Get
            Return App農地基本台帳.TBL転用農地
        End Get
    End Property

    Public Overrides Sub ValueChange(ByVal sField As String, ByVal pValue As Object)
        App農地基本台帳.TBL転用農地.AddUpdateListwithDataViewPage(Me, sField, pValue)
    End Sub
End Class

Namespace DataViewCtrl
    Public Class DataViewNext転用農地
        Inherits CDataViewPanel農地台帳

        Public Sub New(pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
            MyBase.New(pTarget, Nothing, SysAD.page農家世帯.DataViewCollection, True, True)
            SetButtons()
            Dim nID As Integer = pTarget.ID
            Dim nHeight As Integer = Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsReadOnly, pTarget, "ID", , 80), "ID").Height
            Panel.AddCtrl(New ComboBoxPlus(App農地基本台帳.GetMasterView("大字"), "名称", "ID",
                   Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Numeric, emRO.IsCanEdit, pTarget, "大字ID", , 60), "大　字")
               ))

            Panel.AddCtrl(New TextBoxPlus(TextBoxMode.Free, emRO.IsCanEdit, pTarget, "地番", , 100), "地  番")


        End Sub


    End Class
End Namespace