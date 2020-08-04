Imports System.ComponentModel
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class C申請農地一覧TBL
    Inherits HimTools2012.Data.DataTableEx

    Private mvar転用List As Boolean
    Private mvarWhere As String
    Private mvarParentGrid As GridViewNext
    Private mvarColumnList As New List(Of String)
    Private mvarDBFields As New List(Of String)
    Private mvarView As HimTools2012.TargetSystem.CDataViewPanel

    Public Sub New(ByVal sWhere As String, pView As HimTools2012.TargetSystem.CDataViewPanel, ByVal b転用List As Boolean, ParamArray 追加パラメータ() As String)
        MyBase.New("申請農地一覧")

        mvarView = pView
        mvar転用List = b転用List
        mvarWhere = sWhere
        mvarDBFields.AddRange(New String() {"Key", "ID", "自小作", "大字", "小字", "地番", "登記簿面積", "実面積", "一部現況", "登記簿地目名", "現況地目名", "借受人氏名", "所有者氏名"})
        If 追加パラメータ.Length > 0 Then
            mvarColumnList.AddRange(追加パラメータ)
        End If
    End Sub

    Public Sub DoStart(ByRef ParentGrid As GridViewNext)
        Dim pFieldNames As New List(Of String)
        Me.Columns.Add("Key", GetType(String))
        Me.PrimaryKey = {Me.Columns("Key")}
        mvarParentGrid = ParentGrid
        mvarParentGrid.Enabled = False

        Dim mvar農地View As DataView
        App農地基本台帳.TBL農地.FindRowBySQL(mvarWhere)
        mvar農地View = New DataView(App農地基本台帳.TBL農地.Body, mvarWhere, "", DataViewRowState.CurrentRows)

        Dim pTable As DataTable = mvar農地View.ToTable("農地", False, mvarDBFields.ToArray())
        Me.MergePlus(pTable)
        For Each sField As String In mvarColumnList
            Dim Ar As String() = Split(sField, ":")
            pFieldNames.Add(Ar(0))
            Select Case Ar.Length
                Case 1
                    Me.Columns.Add(Ar(0), GetType(String))
                Case 2
                    If InStr(Ar(1), ".") = 0 Then
                        Ar(1) = "System." & Ar(1)
                    End If
                    Dim masterType As Type = Type.GetType(Ar(1))
                    Me.Columns.Add(Ar(0), masterType)
                Case 3
                    If InStr(Ar(1), ".") = 0 Then
                        Ar(1) = "System." & Ar(1)
                    End If
                    Dim masterType As Type = Type.GetType(Ar(1))
                    Me.Columns.Add(Ar(0), masterType, Ar(2))
            End Select
        Next

        If mvar転用List Then
            App農地基本台帳.TBL転用農地.FindRowBySQL(mvarWhere)
            Dim mvar転用農地View As DataView = New DataView(App農地基本台帳.TBL転用農地.Body, mvarWhere, "", DataViewRowState.CurrentRows)
            Me.MergePlus(mvar転用農地View.ToTable("農地", False, mvarDBFields.ToArray()))
        End If
        For Each pCol As DataColumn In Me.Columns
            pCol.ReadOnly = False
        Next

        mvarParentGrid.DataSource = Me
        mvarParentGrid.Columns("KEY").Visible = False
        mvarParentGrid.AutoGenerateColumns = False
        mvarParentGrid.Enabled = True
        pFieldNames.Add("Key")
        mvarParentGrid.ParamFields = pFieldNames.ToArray

        For Each pCol As DataGridViewColumn In mvarParentGrid.Columns
            If mvarDBFields.Contains(pCol.DataPropertyName) Then
                pCol.ReadOnly = True
            End If
        Next
        mvarParentGrid.AddBind(mvarView, "パラメータリスト", pFieldNames.ToArray())



    End Sub

End Class
