'Source : http://www.vbforums.com/showthread.php?616267-Creating-shared-folders-and-specifying-share-permissions
'Title : Creating shared folders and specifying share permissions
'Author : chris128

'Create a list that will hold our permissions
Dim PermissionsList As New List(Of SharePermissionEntry)

'Create a new permission entry for the Everyone group and specify that we want to allow them Read access
Dim PermEveryone As New SharePermissionEntry(String.Empty, "Everyone", SharedFolder.SharePermissions.Read, True)
'Create a new permission entry for the currently logged on user and specify that we want to allow them Full Control
Dim PermUser As New SharePermissionEntry(Environment.UserDomainName, Environment.UserName, SharedFolder.SharePermissions.FullControl, True)
 
'Add the two entries declared above to our list
PermissionsList.Add(PermUser)
PermissionsList.Add(PermEveryone)
 
'Share the folder as "Test Share" and pass in the desired permissions list
Dim Result As SharedFolder.NET_API_STATUS = _
SharedFolder.ShareExistingFolder("Test Share", "This is a test share", "C:\SomeFolder", PermissionsList)
 
'Show the result
If Result = SharedFolder.NET_API_STATUS.NERR_Success Then
    MessageBox.Show("Share created successfully!")
Else
    MessageBox.Show("Share was not created as the following error was returned: " & Result.ToString)
End If
