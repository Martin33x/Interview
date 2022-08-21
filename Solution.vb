Imports System.IO
Module Module1

    Dim searchRoot As String
    Dim ext As String
    Dim subStr As String


    Public Function GetAllSuitableFilePaths(searchRoot, ext, subStr) As String()

        If searchRoot = "" Then
            Console.WriteLine("Hodnota adresáře nemůže být prázdná")
            Exit Function
        ElseIf System.IO.Directory.Exists(searchRoot) = False Then
            Console.WriteLine("Uvedený adresář neexistuje")
            Exit Function
        ElseIf ext = "" Then
            Console.WriteLine("Hodnota přípony nemůže být prázdná")
            Exit Function
        End If

        Dim Directory As New IO.DirectoryInfo(searchRoot)
        Dim FileArray As IO.FileInfo() = Directory.GetFiles("*" & subStr & "*" & ext, SearchOption.AllDirectories)
        Dim File As IO.FileInfo


        For Each File In FileArray
            Console.WriteLine("File Full Name: {0}", File.FullName)
        Next
    End Function

    Sub Main()

        Console.WriteLine("Zadejte název adresáře k prohledání:")
        Dim searchRoot = Console.ReadLine
        Console.WriteLine("Zadejte příponu souboru k vyhledání:")
        Dim ext = Console.ReadLine
        Console.WriteLine("(VOLITELNĚ) Zadejte řetezec obsažený v názvu souboru k vyhledání:")
        Dim subStr = Console.ReadLine

        GetAllSuitableFilePaths(searchRoot, ext, subStr)
    End Sub

End Module



