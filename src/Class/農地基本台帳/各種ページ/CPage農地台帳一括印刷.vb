Imports HimTools2012.CommonFunc

Public Class CPage農地台帳一括印刷
    Inherits HimTools2012.controls.CTabPageWithToolStrip


    Private WithEvents mvarSp As HimTools2012.controls.SplitContainerEX
    Private WithEvents mvarTr As TreeView
    Private WithEvents mvarSpV As HimTools2012.controls.SplitContainerEX
    Private WithEvents mvarGr As HimTools2012.controls.DataGridViewWithDataView

    Private WithEvents mvarAllSelect As ToolStripButton
    Private WithEvents mvarClrSelect As ToolStripButton
    Private WithEvents mvarPrint As ToolStripButton
    Private WithEvents mvarCheckList As ToolStripButton

    Private mvarSK As DataTable
    Private WithEvents mvarXMLLayout As HimTools2012.controls.XMLLayout
    Private mvarDK As DataTable


    Public Sub New()
        MyBase.New(True, True, "農地台帳一括印刷", "農地台帳一括印刷")

        mvarXMLLayout = New HimTools2012.controls.XMLLayout(SysAD.DB(sLRDB), ObjectMan, Me)
        With mvarXMLLayout
            .StartLayout(My.Resources.Resource1.基本画面, "農地台帳一括印刷")
            mvarSp = .Controls("農地台帳一括印刷")
            mvarTr = .Controls("TreeView")
            mvarSpV = .Controls("農地台帳一括印刷")
            mvarGr = .Controls("mvarGr")
            mvarGr.Createエクセル出力Ctrl(.Controls("TX1"))
            'mvarAllSelect = .Controls("全選択")
            'mvarClrSelect = .Controls("選択解除")
            'mvarPrint = .Controls("印刷")
            'mvarCheckList = .Controls("チェック者リスト")
        End With

        mvarGr.AllowUserToAddRows = False

        mvarSK = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT M_BASICALL.ID, M_BASICALL.名称 FROM V_農地 INNER JOIN ([D:個人Info] INNER JOIN M_BASICALL ON [D:個人Info].行政区ID = M_BASICALL.ID) ON V_農地.耕作者ID = [D:個人Info].ID WHERE (M_BASICALL.ID<>0) AND (((M_BASICALL.Class)=""行政区"") AND (([D:個人Info].住民区分)=0)) GROUP BY M_BASICALL.ID, M_BASICALL.名称;")
        mvarDK = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT M_BASICALL.ID AS 集落ID, M_BASICALL.名称, [D:個人Info].世帯ID AS ID, [D:個人Info_1].氏名 AS 世帯主氏名, [D:個人Info].氏名 AS 耕作者名, [D:個人Info].[フリガナ], V_住民区分.名称 AS 住民区分, V_農地.[自小作別], V_農地.[所有世帯ID], V_農地.[借受世帯ID], V_農地.[登記簿面積], [D:個人Info].住所, V_農地.[地番] FROM ([D:個人Info] AS [D:個人Info_1] INNER JOIN ([D:世帯Info] INNER JOIN (V_農地 INNER JOIN ([D:個人Info] INNER JOIN M_BASICALL ON [D:個人Info].行政区ID = M_BASICALL.ID) ON V_農地.耕作者ID = [D:個人Info].ID) ON [D:世帯Info].ID = [D:個人Info].世帯ID) ON [D:個人Info_1].ID = [D:世帯Info].世帯主ID) INNER JOIN V_住民区分 ON [D:個人Info_1].住民区分 = V_住民区分.ID WHERE (((M_BASICALL.ID)<>0) AND ((M_BASICALL.Class)='行政区') AND (([D:個人Info].住民区分)=0));")
        mvarTr.Nodes.Clear()
        mvarTr.CheckBoxes = True
        Dim pParent As TreeNode = mvarTr.Nodes.Add("集落")
        pParent.Name = "All"

        For Each pRow As DataRow In mvarSK.Rows
            Dim pCNode As New TreeNode(pRow.Item("名称"))
            pCNode.Name = "集落." & pRow.Item("ID")
            pCNode.Tag = New NodeProperty()

            pParent.Nodes.Add(pCNode)
        Next
    End Sub

    Private Sub mvarXMLLayout_ClickButton(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarXMLLayout.ClickButton
        Select Case sender.name
            Case "全選択"
                'bSystemCheck = True
                SubCheck(mvarTr.Nodes, True)
                bSystemCheck = True
                'bSystemCheck = False
            Case "選択解除"
                'bSystemCheck = True
                SubCheck(mvarTr.Nodes, False)
                bSystemCheck = True
                'bSystemCheck = False
            Case "チェック者リスト"
                CheckList()
            Case "印刷"
                基本台帳印刷()
            Case Else
                Stop
        End Select
    End Sub

    Private mvarList As DataTable
    Private Sub CheckList()
        Try
            Dim pPNode As TreeNode = mvarTr.Nodes(0)

            If mvarList Is Nothing Then
                mvarList = New DataTable
                mvarList.Columns.Add(New DataColumn("ID", GetType(Long)))
                mvarList.Columns.Add(New DataColumn("集落", GetType(String)))
                mvarList.Columns.Add(New DataColumn("氏名", GetType(String)))
                mvarList.Columns.Add(New DataColumn("フリガナ", GetType(String)))
                mvarList.Columns.Add(New DataColumn("住民区分", GetType(String)))

                mvarList.Columns.Add(New DataColumn("住所", GetType(String)))
                mvarList.Columns.Add(New DataColumn("自作地地積", GetType(Decimal)))
                mvarList.Columns.Add(New DataColumn("借入地地積", GetType(Decimal)))
                'mvarList.Columns.Add(New DataColumn("借入地地積", GetType(Decimal), "[自作地地積]+[借入地地積]"))
                mvarList.Columns.Add(New DataColumn("貸付地地積", GetType(Decimal)))
                mvarList.PrimaryKey = {mvarList.Columns("ID")}
            Else
                mvarList.Rows.Clear()
            End If

            For Each pNode As TreeNode In pPNode.Nodes
                If pNode.Name.StartsWith("集落") Then
                    'If pNode.Checked Then
                    '    Dim pView As New DataView(mvarDK, "[集落ID]=" & GetKeyCode(pNode.Name), "", DataViewRowState.CurrentRows)
                    '    For Each pRowV As DataRowView In pView
                    '        Dim pRow As DataRow = mvarList.Rows.Find(pRowV.Item("ID"))
                    '        If pRow Is Nothing Then
                    '            mvarList.Rows.Add({pRowV.Item("ID"), pNode.Text, pRowV.Item("世帯主氏名"), pRowV.Item("フリガナ"), pRowV.Item("住民区分"), pRowV.Item("住所"),
                    '                               IIf(pRowV.Item("自小作別") = 0, pRowV.Item("登記簿面積"), 0),
                    '                               IIf(pRowV.Item("自小作別") > 0 AndAlso pRowV.Item("ID") = pRowV.Item("借受世帯ID"), pRowV.Item("登記簿面積"), 0),
                    '                               IIf(pRowV.Item("自小作別") > 0 AndAlso Not pRowV.Item("所有世帯ID") = pRowV.Item("借受世帯ID"), pRowV.Item("登記簿面積"), 0)})
                    '        Else
                    '            pRow.Item("自作地地積") += IIf(pRowV.Item("自小作別") = 0, pRowV.Item("登記簿面積"), 0)
                    '            pRow.Item("借入地地積") += IIf(pRowV.Item("自小作別") > 0 AndAlso pRowV.Item("ID") = pRowV.Item("借受世帯ID"), pRowV.Item("登記簿面積"), 0)
                    '            pRow.Item("貸付地地積") += IIf(pRowV.Item("自小作別") > 0 AndAlso Not pRowV.Item("所有世帯ID") = pRowV.Item("借受世帯ID"), pRowV.Item("登記簿面積"), 0)
                    '        End If
                    '    Next
                    'Else
                    For Each pKNode As TreeNode In pNode.Nodes
                        If pKNode.Checked Then
                            Dim nID As Integer = GetKeyCode(pKNode.Name)
                            Dim pRow As DataRow = App農地基本台帳.TBL世帯.FindRowByID(nID)
                            Dim pRowP As DataRow = App農地基本台帳.TBL個人.FindRowByID(pRow.Item("世帯主ID"))
                            If pKNode.Text <> pRowP.Item("氏名") OrElse pRowP.Item("氏名") <> pRow.Item("世帯主氏名") Then
                                Stop
                            End If
                            Dim pRowL As DataRow = mvarList.Rows.Add({nID, pNode.Text, pRow.Item("世帯主氏名"), pRow.Item("フリガナ"), pRowP.Item("住民区分名"), pRow.Item("住所"), 0, 0, 0})

                            Dim pView As New DataView(mvarDK, "[ID]=" & nID & "OR [所有世帯ID]=" & nID, "", DataViewRowState.CurrentRows)
                            For Each pRowV As DataRowView In pView
                                pRowL.Item("自作地地積") += IIf(Val(pRowV.Item("自小作別").ToString) = 0, Val(pRowV.Item("登記簿面積").ToString), 0)
                                pRowL.Item("借入地地積") += IIF(Val(pRowV.Item("自小作別").ToString) <> 0 AndAlso nID = Val(pRowV.Item("借受世帯ID").ToString), Val(pRowV.Item("登記簿面積").ToString), 0)
                                pRowL.Item("貸付地地積") += IIF(Val(pRowV.Item("自小作別").ToString) <> 0 AndAlso nID = Val(pRowV.Item("所有世帯ID").ToString) AndAlso nID <> Val(pRowV.Item("借受世帯ID").ToString), Val(pRowV.Item("登記簿面積").ToString), 0)
                            Next

                        End If
                    Next
                    'End If
                End If
            Next
            mvarGr.SetDataView(mvarList, "[自作地地積] > 0 Or [借入地地積] > 0 Or [貸付地地積] > 0", "")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Private Sub subCheckList()

    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property
    Protected Class NodeProperty
        Public FirstExpand As Boolean = True
    End Class

    Private Sub mvarTr_AfterCheck(sender As Object, e As System.Windows.Forms.TreeViewEventArgs) Handles mvarTr.AfterCheck
        If bSystemCheck = False Then
            Select Case GetKeyHead(e.Node.Name)
                Case "集落"
                    If e.Node.Nodes.Count > 0 Then
                        For Each pNode As TreeNode In e.Node.Nodes
                            pNode.Checked = e.Node.Checked
                        Next
                    End If
                Case "農家"
                    If e.Node.Checked = False Then
                        If Not e.Node.Parent.Checked = False Then
                            bSystemCheck = True
                            e.Node.Parent.Checked = False
                            bSystemCheck = False
                        End If
                    End If
            End Select
        End If
    End Sub

    Private Sub mvarTr_BeforeCheck(sender As Object, e As System.Windows.Forms.TreeViewCancelEventArgs) Handles mvarTr.BeforeCheck
        If bSystemCheck = False AndAlso e.Node.Nodes.Count = 0 AndAlso GetKeyHead(e.Node.Name) = "集落" Then
            MClick(e.Node)
        End If
    End Sub

    Private Sub mvarTr_NodeMouseClick(sender As Object, e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles mvarTr.NodeMouseClick
        MClick(e.Node)
    End Sub


    Private Sub MClick(pNode As TreeNode)
        If bSystemCheck = False AndAlso pNode.Tag IsNot Nothing Then
            Select Case GetKeyHead(pNode.Name)
                Case "集落"
                    With CType(pNode.Tag, NodeProperty)
                        If .FirstExpand Then
                            Dim nID As Integer = GetKeyCode(pNode.Name)
                            App農地基本台帳.TBL世帯.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:世帯Info].* FROM [D:世帯Info] INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID WHERE ((([D:個人Info].行政区ID)={0}));", nID))
                            Dim pView As New DataView(App農地基本台帳.TBL世帯.Body, "[世帯主行政区ID]=" & nID, "フリガナ", DataViewRowState.CurrentRows)

                            For Each pRow As DataRowView In pView
                                With CType(pNode.Nodes.Add(pRow.Item("世帯主氏名")), TreeNode)
                                    .Name = "農家." & pRow.Item("ID")
                                    .Tag = New NodeProperty()
                                End With
                            Next
                            .FirstExpand = False
                        End If
                    End With
                Case "農家"
                    With CType(pNode.Tag, NodeProperty)
                        If .FirstExpand Then

                        End If
                    End With
            End Select
        End If

    End Sub

    Private bSystemCheck As Boolean = False

    Private Sub SubCheck(ByVal pPNode As System.Windows.Forms.TreeNodeCollection, ByVal b As Boolean)
        For Each pNode As TreeNode In pPNode
            MClick(pNode)

            If pNode.Nodes.Count > 0 Then
                SubCheck(pNode.Nodes, b)
            End If
            pNode.Checked = b
        Next
    End Sub

    Private Sub mvarPrint_Click(sender As Object, e As System.EventArgs) Handles mvarPrint.Click
        基本台帳印刷()
    End Sub

    Public Sub 基本台帳印刷()
        Dim sMode As String = "旧"

        Dim sFileName As String = SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & sMode & "農地基本台帳様式.xml"
        If Not IO.File.Exists(sFileName) Then
            sFileName = SysAD.SystemInfo.ApplicationDirectory & "\" & sMode & "農地基本台帳様式.xml"
        End If

        If IO.File.Exists(sFileName) Then
            Using pExcel As New HimTools2012.Excel.Automation.ExcelAutomation
                Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(sFileName)

                Dim pPNode As TreeNode = mvarTr.Nodes(0)
                Do
                    For Each pNode As TreeNode In pPNode.Nodes
                        If pNode.Name.StartsWith("集落") Then
                            If pNode.Checked AndAlso CType(pNode.Tag, NodeProperty).FirstExpand Then
                                MClick(pNode)
                            End If
                            For Each pNNode As TreeNode In pNode.Nodes
                                If pNNode.Checked Then
                                    Dim nID As Integer = GetKeyCode(pNNode.Name)

                                    Try
                                        Dim objAcc As New CPrint基本台帳(New HimTools2012.Excel.XMLSS2003.CXMLSS2003(sXML), nID, 0, 印刷Mode.簡易印刷, ExcelViewMode.AutoPrint)

                                        With objAcc
                                            .Dialog.StartProc(True, True)

                                            If .Dialog._objException Is Nothing = False Then
                                                If .Dialog._objException.Message = "Cancel" Then
                                                    MsgBox("処理を中止しました。　", , "処理中止")
                                                    Exit Do
                                                Else
                                                    'Throw objDlg._objException
                                                End If
                                            ElseIf .HasLand Then
                                                Dim sDir As String = SysAD.OutputFolder & "\基本台帳.xml"
                                                HimTools2012.TextAdapter.SaveTextFile(sDir, .XMLSS.OutPut(True))

                                                pExcel.PrintBook(sDir)

                                                SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:世帯Info] SET [農地との関連]=True WHERE [ID]=" & nID)
                                            Else
                                                pNNode.Text = pNNode.Text & "(関連する農地が無いため印刷なし)"
                                                SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:世帯Info] SET [農地との関連]=False WHERE [ID]=" & nID)
                                            End If

                                        End With
                                        pNNode.Checked = False
                                    Catch ex As Exception

                                    End Try

                                End If
                            Next
                        End If
                    Next
                    Exit Do
                Loop
            End Using
        Else
            MsgBox("ファイルがありません")
        End If
    End Sub
End Class


