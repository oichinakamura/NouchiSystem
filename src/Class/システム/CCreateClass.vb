Imports System.CodeDom
Imports System.CodeDom.Compiler
Imports System.Reflection



Public Class CCreateClass
    Private codigo As New System.Text.StringBuilder("""")

    Public Function ItemClassCodeDom(ByRef pTBL As DataTable) As CodeNamespace
        Dim nsNameSpace As New CodeNamespace("DataProperty")

        nsNameSpace.Imports.Add(New CodeNamespaceImport("System"))
        nsNameSpace.Imports.Add(New CodeNamespaceImport("System.Data"))
        nsNameSpace.Imports.Add(New CodeNamespaceImport("CommonTools"))

        Dim dataClass As New CodeTypeDeclaration("CDataRow")
        dataClass.BaseTypes.Add(New CodeTypeReference("CommonTools.CTargetBase"))

        nsNameSpace.Types.Add(dataClass)


        Dim methodNew As New CodeConstructor()
        methodNew.Attributes = MemberAttributes.Public
        methodNew.Parameters.Add(New CodeParameterDeclarationExpression("System.Object", "pRow"))

        methodNew.BaseConstructorArgs.Add(New CodeVariableReferenceExpression("pRow"))
        '        methodNew.BaseConstructorArgs.Add(New CodeVariableReferenceExpression("False"))

        'Dim pRowM = New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "mvarRow")
        'Dim pRowP = New CodeArgumentReferenceExpression("pRow")

        'methodNew.Statements.Add(New CodeAssignStatement(pRowM, pRowP))
        dataClass.Members.Add(methodNew)

        For Each pCol As DataColumn In pTBL.Columns
            Dim pProperty As New CodeMemberProperty
            Dim sName As String = pCol.ColumnName


            If IsNumeric(sName.Substring(0, 1)) Then
                sName = "n" & sName
            End If

            pProperty.Name = sName
            pProperty.Type = New CodeTypeReference(pCol.DataType.FullName)
            pProperty.Attributes = MemberAttributes.Public

            Select Case pCol.DataType.FullName
                Case "System.Integer", "System.Int32"
                    pProperty.GetStatements.Add(New CodeMethodReturnStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "GetIntegerValue(""" & pCol.ColumnName & """)")))
                    pProperty.SetStatements.Add(New CodeAssignStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "mvarRow.Item(""" & pCol.ColumnName & """)"), New CodePropertySetValueReferenceExpression()))
                Case "System.Decimal"
                    pProperty.GetStatements.Add(New CodeMethodReturnStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "GetDecimalValue(""" & pCol.ColumnName & """)")))
                    pProperty.SetStatements.Add(New CodeAssignStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "mvarRow.Item(""" & pCol.ColumnName & """)"), New CodePropertySetValueReferenceExpression()))
                Case "System.Double"
                    pProperty.GetStatements.Add(New CodeMethodReturnStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "GetDoubleValue(""" & pCol.ColumnName & """)")))
                    pProperty.SetStatements.Add(New CodeAssignStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "mvarRow.Item(""" & pCol.ColumnName & """)"), New CodePropertySetValueReferenceExpression()))
                Case "System.Date", "System.DateTime"
                    pProperty.GetStatements.Add(New CodeMethodReturnStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "GetDateValue(""" & pCol.ColumnName & """)")))
                    pProperty.SetStatements.Add(New CodeAssignStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "mvarRow.Item(""" & pCol.ColumnName & """)"), New CodePropertySetValueReferenceExpression()))
                Case "System.String"
                    pProperty.GetStatements.Add(New CodeMethodReturnStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "GetStringValue(""" & pCol.ColumnName & """)")))
                    pProperty.SetStatements.Add(New CodeAssignStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "mvarRow.Item(""" & pCol.ColumnName & """)"), New CodePropertySetValueReferenceExpression()))
                Case Else
                    pProperty.GetStatements.Add(New CodeMethodReturnStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "mvarRow.Item(""" & pCol.ColumnName & """)")))
                    pProperty.SetStatements.Add(New CodeAssignStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "mvarRow.Item(""" & pCol.ColumnName & """)"), New CodePropertySetValueReferenceExpression()))
            End Select
            dataClass.Members.Add(pProperty)
        Next


        Return nsNameSpace
    End Function



    Public Function GenerateCode(codeNamespace As CodeNamespace) As String
        Dim compilerOptions = New CodeGeneratorOptions()
        compilerOptions.IndentString = "    "
        compilerOptions.BracingStyle = "C"

        Dim codeText As New System.Text.StringBuilder()
        Using codeWriter = New System.IO.StringWriter(codeText)
            CodeDomProvider.CreateProvider("VB").GenerateCodeFromNamespace(codeNamespace, codeWriter, compilerOptions)
        End Using
        Return codeText.ToString()
    End Function

    Public Function GetCode(ByRef pRow As DataRow) As String
        Dim ItemNamespace As CodeNamespace = ItemClassCodeDom(pRow.Table)
        Return GenerateCode(ItemNamespace)
    End Function

    Public Function CompileAssembly(codeNamespace As CodeNamespace, ByRef RText As RichTextBox) As Assembly
        Dim codeCompileUnit As New CodeCompileUnit()
        codeCompileUnit.Namespaces.Add(codeNamespace)
        codeCompileUnit.ReferencedAssemblies.Add(My.Application.Info.DirectoryPath & "\CommonTools.DLL")
        codeCompileUnit.ReferencedAssemblies.Add("System.Data.DLL")

        Dim pOptions As New CompilerParameters
        pOptions.GenerateExecutable = False
        pOptions.GenerateInMemory = True

        Dim pCompilerResults As CompilerResults = CodeDomProvider.CreateProvider("VB").CompileAssemblyFromDom(pOptions, codeCompileUnit)

        If pCompilerResults.Output.Count > 0 Then
            Dim sB As New System.Text.StringBuilder
            For Each sT As String In pCompilerResults.Output
                sB.AppendLine(sT)
            Next
            RText.Text = sB.ToString
            Return Nothing
        Else
            Dim pAsm As Assembly = pCompilerResults.CompiledAssembly
            Return pAsm
        End If
    End Function


    Public Function GetObject(ByVal sClassName As String, ByRef pRow As DataRow, ByRef RText As RichTextBox) As Object
        Dim ItemNamespace As CodeNamespace = ItemClassCodeDom(pRow.Table)

        Dim itemAssembly As System.Reflection.Assembly = CompileAssembly(ItemNamespace, RText)

        If itemAssembly IsNot Nothing Then

            For Each itemType As System.Type In itemAssembly.GetTypes
                If itemType.Name = sClassName Then
                    Return Activator.CreateInstance(itemType, {pRow})
                End If
            Next
            Return Nothing
        Else
            Return Nothing
        End If
    End Function
End Class
