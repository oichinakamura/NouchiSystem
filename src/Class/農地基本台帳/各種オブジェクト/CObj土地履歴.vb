Imports HimTools2012.controls
Imports HimTools2012.controls.DVCtrlCommonBase

Public Class CObj土地履歴 : Inherits CTargetObjWithView農地台帳

    Public Sub New(ByVal pRow As DataRow, ByVal bAddNew As Boolean)
        MyBase.New(pRow, bAddNew, New HimTools2012.TargetSystem.DataKey("土地履歴", pRow.Item("ID")), "D_土地履歴")

    End Sub
    Public Overrides Function InitDataViewNext(ByRef pDB As HimTools2012.TargetSystem.CDataViewCollection, Optional InterfaceName As String = "") As Boolean
        If Me.DataViewPage Is Nothing Then
            Me.DataViewPage = New DataViewNext農地台帳(Me, "D_土地履歴", "D_土地履歴", App農地基本台帳.TBL土地履歴, pDB)
        End If
        Return True
    End Function

    Public Overrides Function GetContextMenu(Optional ByVal pMenu As HimTools2012.controls.MenuItemEx = Nothing, Optional nDips As Integer = 1, Optional sParam() As String = Nothing) As HimTools2012.controls.MenuPlus
        Dim pMenuf As New HimTools2012.controls.ContextMenuEx(AddressOf ClickMenu)
        pMenuf.AddMenuByText({"開く", "土地を呼ぶ", "-"}, AddressOf ClickMenu, True)

        pMenuf.AddMenu("削除")

        SetDVMenu(pMenuf, pMenu)
        Return pMenuf
    End Function



    Public Overrides Function DoCommand(ByVal sCommand As String, ByVal ParamArray sParams() As String) As Object
        Select Case sCommand
            Case "開く" : Me.OpenDataViewNext(SysAD.page農家世帯.DataViewCollection)
            Case "更新" : Me.SaveMyself()
            Case "土地を呼ぶ" : Open農地(Val(Me.GetItem("LID").ToString))
            Case "削除" : Me.DoCommand("閉じる")
                If MsgBox("削除してもよろしいですか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    SysAD.DB(sLRDB).ExecuteSQL("DELETE D_土地履歴.ID FROM D_土地履歴 WHERE [ID]={0};", Me.ID)
                    Dim pRow As DataRow = App農地基本台帳.TBL土地履歴.FindRowByID(Me.ID)
                    If pRow IsNot Nothing Then
                        App農地基本台帳.TBL土地履歴.Rows.Remove(pRow)
                    End If
                End If
            Case Else
                Return MyBase.DoCommand(sCommand, sParams)
        End Select

        Return Nothing
    End Function

    Public Overrides Function ToString() As String
        Return "土地履歴"
    End Function
    Public Overrides ReadOnly Property DataTableWithUpdateList As HimTools2012.Data.DataTableWithUpdateList
        Get
            Return App農地基本台帳.TBL土地履歴
        End Get
    End Property


    Public Overrides Function CanDropKeyHead(ByVal sKey As String, ByVal sOption As String) As Boolean
        Return False
    End Function

    Public Overrides Sub DropDown(ByVal sSourceList As String, Optional ByVal sOption As String = "")

    End Sub

    Public Overrides Function GetProperty(ByVal sParam As String) As Object
        Return Nothing
    End Function

    Public Overrides Sub ValueChange(ByVal sField As String, ByVal pValue As Object)
        App農地基本台帳.TBL土地履歴.AddUpdateListwithDataViewPage(Me, sField, pValue)
    End Sub
End Class

'    Case "土地履歴-農地を呼ぶ" : mvarPDW.DataviewCol.Add(ADApp.ObjectMan.GetObject("農地." & DVProperty.Rs.Value("LID")))
'            '    Case "土地履歴" : St = "開く;農地を呼ぶ;" & n & "削除;"
'        Case "土地履歴-申請"
'            Stop
'            SysAD.DB(sLRDB).Execute("UPDATE [D_土地履歴] SET [申請ID]=" & FncNet.GetKeyCode(sSourceList) & " WHERE [ID]=" & DVProperty.ID)
'            mvarPDW.SQLListview.Refresh()

Public Class SaveInterface
    Dim xDocument As System.Xml.XmlDocument
    Dim xRoot As System.Xml.XmlElement
    Dim sNameSpace As String = ""

    Public Sub New(ByVal sName As String)
        sNameSpace = sName
        xDocument = New System.Xml.XmlDocument  'XMLドキュメント作成   

        Dim xEncode As System.Text.Encoding = System.Text.Encoding.GetEncoding(65001I) 'エンコード65001がUTF8   
        Dim xDeclaration As System.Xml.XmlDeclaration = xDocument.CreateXmlDeclaration("1.0", xEncode.BodyName, Nothing) 'ヘッダ(？)部作成   
        xRoot = xDocument.CreateElement("DataViewNext")   'ルート部作成   
        xRoot.SetAttribute("Name", sName)

        xDocument.AppendChild(xDeclaration)   'ドキュメントにヘッダを追加   
        xDocument.AppendChild(xRoot)    'ドキュメントにルートを追加   
    End Sub
    Public Function AddGroupBoxPlus(ByVal sName As String) As Xml.XmlElement
        Dim newNode As System.Xml.XmlElement = xDocument.CreateElement("GroupBoxPlus")
        newNode.SetAttribute("Name", sName)

        xRoot.AppendChild(newNode)
        Return newNode
    End Function

    Public Function AddTextBoxPlus(ByVal sKey As String,
                    pTextMode As TextBoxMode,
                    Optional bReadOnly As emRO = Nothing,
                    Optional ByVal WithLabel As String = "",
                    Optional ByVal nWidth As Integer = -1,
                    Optional ByVal nHeight As Integer = -1,
                    Optional ByVal bCanDragDrop As Boolean = Nothing,
                    Optional ByRef pPanel As System.Xml.XmlElement = Nothing) As System.Xml.XmlElement
        Dim newNode As System.Xml.XmlElement = xDocument.CreateElement("TextBoxPlus")

        newNode.SetAttribute("TextMode", pTextMode.ToString)

        If Not IsNothing(bReadOnly) Then
            newNode.SetAttribute("ReadOnly", bReadOnly.ToString)
        End If
        If Not IsNothing(bCanDragDrop) Then
            newNode.SetAttribute("CanDragDrop", bCanDragDrop.ToString)
        End If
        If nWidth > -1 Then
            newNode.SetAttribute("Width", nWidth)
        End If
        If nHeight > -1 Then
            newNode.SetAttribute("Height", nHeight)
        End If
        If WithLabel.Length > 0 Then
            newNode.SetAttribute("WithLabel", WithLabel)
        End If
        If pPanel IsNot Nothing Then
            pPanel.AppendChild(newNode)
        Else
            xRoot.AppendChild(newNode)
        End If
        Return newNode
    End Function

    Public Sub Save(ByVal sName As String)
        Dim sPath As String = My.Application.Info.DirectoryPath.ToLower.Replace("\bin\debug", "") & "\Resources\" & sName & ".xml"

        xDocument.Save(sPath)
    End Sub
End Class
