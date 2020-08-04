
Imports System.Reflection
Imports System.Globalization



'Public Class DataGridListViewTabPage
'    Inherits CTabPageWithDataGridView

'    Private WithEvents btnScrollLock As ToolStripCheckBox

'    Public Sub New(ByVal sText As String, ByVal sName As String, pGrid As DataGridViewWithDataView)
'        MyBase.New(True, sName, sText, ObjectMan, False, pGrid)

'        btnScrollLock = New ToolStripCheckBox("スクロールをロック", False)
'        Me.ToolStrip.ItemAdd("Excel", btnScrollLock)
'    End Sub

'    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
'        Get
'            Return HimTools2012.controls.CloseMode.CloseOK
'        End Get
'    End Property

'    Private Sub btnScrollLock_Click(sender As Object, e As System.EventArgs) Handles btnScrollLock.Click
'        If btnScrollLock.Checked AndAlso mvarGrid.CurrentCell IsNot Nothing Then
'            mvarGrid.Columns(mvarGrid.CurrentCell.ColumnIndex).Frozen = True
'        Else
'            For Each pCol As DataGridViewColumn In mvarGrid.Columns
'                If pCol.Frozen Then
'                    pCol.Frozen = False
'                End If
'            Next
'        End If
'    End Sub
'End Class

Public Class Cカスタムリスト
    Inherits CNList農地台帳

    Public Sub New(ByVal sText As String, ByVal sName As String, ByVal sSQL As String, Optional ByVal sHeaderText As String = "")
        MyBase.New(sText, sName, True)
        Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(sSQL)

        If sHeaderText.Length > 0 Then
            GView.AutoGenerateColumns = False
            Dim Ar() As String = Split(sHeaderText, ";")
            For n As Integer = 0 To UBound(Ar) Step 2
                Dim pColumn As New DataGridViewTextBoxColumn

                Dim sFN As String = Ar(n)
                Dim sFNBK As String

                Do
                    sFNBK = sFN
                    If sFN.StartsWith("@") Then
                        sFN = sFN.Substring(1)
                    End If
                Loop Until sFN = sFNBK

                Select Case pTable.Columns(sFN).DataType.FullName
                    Case "System.String"
                    Case "System.Int32"
                        pColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    Case "System.Double", "System.Single"
                        pColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    Case Else
                        If Not SysAD.IsClickOnceDeployed Then
                            Stop
                        End If
                End Select

                pColumn.DataPropertyName = sFN
                pColumn.Name = sFN
                pColumn.HeaderText = Ar(n + 1)

                GView.Columns.Add(pColumn)
            Next
            Dim pColumnKey As New DataGridViewTextBoxColumn
            pColumnKey.Name = "Key"
            pColumnKey.DataPropertyName = "Key"
            pColumnKey.Visible = False
            GView.Columns.Add(pColumnKey)

        Else
            GView.AutoGenerateColumns = True
        End If
        GView.SetDataView(pTable, "", "")
    End Sub



    Public Overrides Sub 検索開始(ByVal sWhere As String, ByVal sViewWhere As String, Optional sOrderBy As String = "", Optional sColumnStyle As String = "")

    End Sub


    Public Overrides Property IconKey As String
        Get
            Return "List"
        End Get
        Set(value As String)

        End Set
    End Property
End Class

Public Class SQLList
    Inherits CNList農地台帳

    Public Sub New(ByVal sTitle As String, ByVal sSQL As String, ByVal bCloseable As Boolean, Optional s合計列() As String = Nothing)
        MyBase.New(sTitle, sSQL, bCloseable)
        Try
            Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(sSQL)
            pTable.PrimaryKey = New DataColumn() {pTable.Columns("Key")}
            GView.AutoGenerateColumns = True
            GView.SetDataView(pTable, "", "")

            For Each pCol As DataGridViewColumn In GView.Columns
                Select Case pCol.DataPropertyName
                    Case "ID" : pCol.Visible = False
                    Case "Key" : pCol.Visible = False
                    Case "アイコン" : pCol.Visible = False
                End Select
            Next
            If s合計列 IsNot Nothing AndAlso s合計列.Length > 0 Then
                Dim pNewRow As DataRow = pTable.NewRow
                For Each s列 As String In s合計列
                    If pTable.Columns.Contains(s列) Then
                        pNewRow.Item(s列) = 0
                        For Each pRow As DataRow In pTable.Rows
                            pNewRow.Item(s列) += Val(pRow.Item(s列).ToString)
                        Next
                    End If
                Next
                pTable.Rows.Add(pNewRow)
            End If
        Catch ex As Exception
            Stop
        End Try

    End Sub
    Public Overrides Property IconKey As String
        Get
            Return "List"
        End Get
        Set(value As String)

        End Set
    End Property
    Public Overrides Sub 検索開始(ByVal sWhere As String, ByVal sViewWhere As String, Optional sOrderBy As String = "", Optional sColumnStyle As String = "")
    End Sub
End Class

Public Class CTableFilterList
    Inherits CNList農地台帳
    Private mvarParentObj As HimTools2012.TargetSystem.CTargetObjectBase

    Public Sub New(ByVal sKey As String, ByVal sTitle As String, ByRef pTable As DataTable, sRowFilter As String, ByVal Sort As String, ByVal bCloseable As Boolean, Optional s合計列() As String = Nothing, Optional ByRef pObj As HimTools2012.TargetSystem.CTargetObjectBase = Nothing)
        MyBase.New(sTitle, sKey, bCloseable)
        Try
            GView.AutoGenerateColumns = True
            GView.SetDataView(pTable, sRowFilter, Sort)
            mvarParentObj = pObj

            For Each pCol As DataGridViewColumn In GView.Columns
                Select Case pCol.DataPropertyName
                    Case "Key" : pCol.Visible = False
                    Case "アイコン" : pCol.Visible = False
                End Select
            Next
            If s合計列 IsNot Nothing AndAlso s合計列.Length > 0 Then
                Dim pNewRow As DataRow = pTable.NewRow
                For Each s列 As String In s合計列
                    If pTable.Columns.Contains(s列) Then
                        pNewRow.Item(s列) = 0
                        For Each pRow As DataRow In pTable.Rows
                            pNewRow.Item(s列) += Val(pRow.Item(s列).ToString)
                        Next
                    End If
                Next
                pTable.Rows.Add(pNewRow)
            End If
        Catch ex As Exception
            Stop
        End Try

    End Sub
    Public Overrides  Property IconKey As String
        Get
            Return "List"
        End Get
        Set(value As String)

        End Set
    End Property

    Public Overrides Sub 検索開始(ByVal sWhere As String, ByVal sViewWhere As String, Optional sOrderBy As String = "", Optional sColumnStyle As String = "")
    End Sub
    '
End Class

Public Class CObj各種List
    Inherits CNList農地台帳

    Public Sub New(ByVal sTitle As String, sKey As String, ByVal sTable As String, sWhere As String, ByVal bCloseable As Boolean)
        MyBase.New(sTitle, sKey, bCloseable)
        Try
            GView.AutoGenerateColumns = True
            GView.SetDataView(App農地基本台帳.DSet.Tables(sTable), sWhere, "")

            For Each pCol As DataGridViewColumn In GView.Columns
                Select Case pCol.DataPropertyName
                    Case "ID" : pCol.Visible = False
                    Case "Key" : pCol.Visible = False
                    Case "アイコン" : pCol.Visible = False
                End Select
            Next
        Catch ex As Exception
            Stop
        End Try

    End Sub

    Public Overrides Sub 検索開始(ByVal sWhere As String, ByVal sViewWhere As String, Optional sOrderBy As String = "", Optional sColumnStyle As String = "")

    End Sub

    Public Overrides Property IconKey As String
        Get
            Return "List"
        End Get
        Set(value As String)

        End Set
    End Property
End Class

