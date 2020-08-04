
Public Class CDataViewPanel農地台帳
    Inherits HimTools2012.TargetSystem.CDataViewPanel

    Public Sub New(ByRef pTarget As HimTools2012.TargetSystem.CTargetObjWithView, ByRef BaseTable As HimTools2012.Data.DataTableWithUpdateList, ByRef pParent As HimTools2012.TargetSystem.CDataViewCollection, ByVal bClose As Boolean, ByVal toolbarVisible As Boolean)
        MyBase.New(pTarget, BaseTable, pParent, bClose, toolbarVisible)

    End Sub


    Protected Overrides Function CreateCustomControl(ByVal pNode As Xml.XmlNode) As System.Windows.Forms.Control
        Select Case pNode.Name
            Case ""
        End Select

        Stop
        Return Nothing
    End Function

    Public Function b新項目保存(ByVal pTarget As HimTools2012.TargetSystem.CTargetObjWithView, ByVal pFieldName As String)
        Dim bReadOnly As Boolean = False
        If Not pTarget.Row.Body.Table.Columns.Contains(pFieldName) Then
            bReadOnly = True
        End If

        Return bReadOnly
    End Function
End Class
