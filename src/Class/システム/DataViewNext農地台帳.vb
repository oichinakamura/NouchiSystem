

Public Class DataViewNext農地台帳
    Inherits CDataViewPanel農地台帳

    Public Sub New(ByVal pTarget As CTargetObjWithView農地台帳, ByVal sTableName As String, ByRef mvarInterfaceName As String, ByRef pTBL As HimTools2012.Data.DataTableWithUpdateList, ByRef pDataViewCollection As HimTools2012.TargetSystem.CDataViewCollection)
        MyBase.New(pTarget, pTBL, pDataViewCollection, True, True)
        LoadXML(My.Resources.Resource1._Interface, mvarInterfaceName, App農地基本台帳.DSet)
    End Sub
End Class
