Imports System.IO
Imports System.CodeDom

Module Module1
    Dim FileArray As List(Of String) = New List(Of String)()
    Dim searchRoot As String
    Dim ext As String
    Dim subStr As String

    Public Function GetAllSuitableFilePaths(searchRoot, ext, subStr) As String()
        Dim currentPath As String
        Dim directory As Object
        Dim sufix As String
        Dim FolderArray As New List(Of String)()

        If Right(searchRoot, 1) <> "\" Then
            searchRoot = searchRoot & "\"
        End If
        If searchRoot = "" Then
            Console.WriteLine("Hodnota adresáře nemůže být prázdná")
            Exit Function
        ElseIf System.IO.Directory.Exists(searchRoot) = False Then
            Console.WriteLine("Uvedený adresář neexistuje")
            Exit Function
        ElseIf ext = "" Then
            Console.WriteLine("Hodnota přípony nemůže být prázdná")
            Exit Function
        ElseIf Left(ext, 1) <> "." Then
            ext = ("." & ext)
        End If

        currentPath = Dir(searchRoot, vbDirectory)

        Do Until currentPath = vbNullString
            sufix = Right(currentPath, (Len(currentPath) - (InStrRev(currentPath, ".") - 1)))
            If (GetAttr(searchRoot & currentPath) And vbDirectory) = vbDirectory Then
                FolderArray.Add(currentPath)
            ElseIf (sufix = ext) And ((GetAttr(searchRoot & currentPath) And vbDirectory) <> vbDirectory) And (Left(currentPath, InStrRev(currentPath, ".") - 1).Contains(subStr)) Then
                FileArray.Add(currentPath)
            End If
            currentPath = Dir()
        Loop

        For Each directory In FolderArray
            GetAllSuitableFilePaths(searchRoot & directory & "\", ext, subStr)
        Next directory

        Return FileArray.ToArray()

    End Function

    Sub Main()
        searchRoot = "C:\temp\"
        ext = ".txt"
        subStr = "test"

        GetAllSuitableFilePaths(searchRoot, ext, subStr)

    End Sub

End Module
