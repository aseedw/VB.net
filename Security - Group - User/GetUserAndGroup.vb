'Source : https://social.msdn.microsoft.com/Forums/vstudio/en-US/f3c56180-8e8a-4ecf-9709-94e2c30ff706/how-to-check-if-users-sid-is-in-local-administrator-group?forum=vbgeneral
'Title : [RESOLVED] Convert sid string to user name
'Requirement : ref - System.DirectoryServices.AccountManagement
'              UI - ListBox1 

Imports System.DirectoryServices.AccountManagement ' Add ref, Assemblies, Framework, System.DirectoryServices.AccountManagement

  Private Sub Button_Click(sender As Object, e As EventArgs) Handles Button.Click
        ListBox1.Items.Clear()
        Dim MachineContext As PrincipalContext = New PrincipalContext(ContextType.Machine)
        Dim userPrincipal As UserPrincipal = New UserPrincipal(MachineContext)
        Dim userPrincipalSearcher As PrincipalSearcher = New PrincipalSearcher(userPrincipal)

        Dim UPSresults As PrincipalSearchResult(Of Principal) = userPrincipalSearcher.FindAll
        For Each Item As Principal In UPSresults
            ListBox1.Items.Add(Item.Name)
            For Each GP In Item.GetGroups
                ListBox1.Items.Add("      Group = " & GP.Name)
            Next
            ListBox1.Items.Add(" ")
        Next
        userPrincipalSearcher.Dispose()
        userPrincipal.Dispose()
        MachineContext.Dispose()
    End Sub
