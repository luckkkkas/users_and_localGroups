Imports System.DirectoryServices.AccountManagement
Imports System.Security

Module Module1

    Sub Main(args As String())

        Dim hostName As String = Environment.MachineName
        Dim caminho As String = "C:\users\"
        Dim caminhoLogTxt As String = "c:\Windows\Temp\Usuarios_e_Grupos_locais.txt"

        If My.Computer.FileSystem.FileExists(caminhoLogTxt) Then
            Try
                My.Computer.FileSystem.DeleteFile(caminhoLogTxt)
            Catch ex As Exception
                Console.WriteLine("Erro ao deletar" & ex.Message)
            End Try
        Else
            Console.WriteLine("Obtendo informações...")

        End If

        If My.Computer.FileSystem.DirectoryExists(caminho) Then
            Dim logInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(caminho)
            Dim pastas As System.IO.DirectoryInfo() = logInfo.GetDirectories()

            Using file As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(caminhoLogTxt, True)

                For Each pasta As System.IO.DirectoryInfo In pastas

                    'GetGroups(pasta.Name)
                    Dim pastaTrim As String = pasta.LastWriteTime
                    Dim lastTrim As String = Left$(pastaTrim, 10)
                    Dim user As String = pasta.Name
                    Dim UserGroupAll As List(Of GroupPrincipal) = GetUserGroups(user)
                    Dim groupConcat As New List(Of String)()

                    If UserGroupAll IsNot Nothing Then
                        For Each group1 In UserGroupAll
                            groupConcat.Add(group1.Name)
                        Next
                    Else
                        groupConcat.Add("Usuário sem grupo")
                    End If

                    Dim allgroups = String.Join(";", groupConcat)
                    file.WriteLine(hostName & "," & pasta.Name & "," & lastTrim & "," & allgroups)
                Next
            End Using
        End If
    End Sub

    Function GetUserGroups(userName As String) As List(Of GroupPrincipal)
        Dim groups As List(Of GroupPrincipal) = GetAllGroups()
        Dim userGroups As New List(Of GroupPrincipal)
        Dim usuarioSemGrupo As List(Of GroupPrincipal) = Nothing

        For Each groupTab As GroupPrincipal In groups

            Dim members = groupTab.GetMembers()

            For Each member As Principal In members

                If TypeOf member Is UserPrincipal Then
                    Dim userMember As UserPrincipal = CType(member, UserPrincipal)

                    If userMember.SamAccountName = userName Then
                        userGroups.Add(groupTab)
                    End If

                End If
            Next

        Next
        If userGroups.Count = 0 Then
            Return usuarioSemGrupo
        End If
        Return userGroups
    End Function

    Function GetAllGroups() As List(Of GroupPrincipal)
        Dim groupsList As New List(Of GroupPrincipal)()
        Dim ctx As New PrincipalContext(ContextType.Machine, Environment.MachineName)

        ' Obtém todos os grupos no contexto
        Dim groupPrincipal As New GroupPrincipal(ctx)
        Dim groupSearcher As New PrincipalSearcher(groupPrincipal)

        For Each principal As Principal In groupSearcher.FindAll()
            If TypeOf principal Is GroupPrincipal Then
                Dim group As GroupPrincipal = CType(principal, GroupPrincipal)
                groupsList.Add(group) ' Ou use group.Name para obter o nome
            End If
        Next

        Return groupsList
    End Function
End Module