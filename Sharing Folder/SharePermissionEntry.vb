'Source : http://www.vbforums.com/showthread.php?616267-Creating-shared-folders-and-specifying-share-permissions
'Author : chris128

Imports System.Runtime.InteropServices

Public Class SharedFolder

#Region "Constants"

    Private Const STYPE_DISKTREE As UInteger = 0
    Private Const SECURITY_DESCRIPTOR_REVISION As UInteger = 1
    Private Const NO_INHERITANCE As UInteger = 0
    Private Const ERROR_NONE_MAPPED As UInteger = 1332

#End Region

#Region "Enums"

    Public Enum NET_API_STATUS As Integer
        NERR_Success = 0
        ERROR_ACCESS_DENIED = 5
        ERROR_INVALID_PARAMETER = 87
        ERROR_INVALID_NAME = 123
        ERROR_INVALID_LEVEL = 124
        NERR_UnknownDevDir = 2116
        NERR_RedirectedPath = 2117
        NERR_DuplicateShare = 2118
        NERR_BufTooSmall = 2123
    End Enum

    Public Enum SharePermissions As Integer
        Read = 1
        FullControl = 2
    End Enum

    Private Enum ACCESS_MASK As UInteger
        GENERIC_ALL = 268435456
        GENERIC_READ = 2147483648
        GENERIC_WRITE = 1073741824
        GENERIC_EXECUTE = 536870912
        STANDARD_RIGHTS_READ = 131072
    End Enum

    Private Enum ACCESS_MODE As UInteger
        NOT_USED_ACCESS = 0
        GRANT_ACCESS = 1
        SET_ACCESS = 2
        DENY_ACCESS = 3
        REVOKE_ACCESS = 4
        SET_AUDIT_SUCCESS = 5
        SET_AUDIT_FAILURE = 6
    End Enum

    Private Enum MULTIPLE_TRUSTEE_OPERATION As UInteger
        NO_MULTIPLE_TRUSTEE = 0
        TRUSTEE_IS_IMPERSONATE = 1
    End Enum

    Private Enum TRUSTEE_FORM As UInteger
        TRUSTEE_IS_SID = 0
        TRUSTEE_IS_NAME = 1
        TRUSTEE_BAD_FORM = 2
        TRUSTEE_IS_OBJECTS_AND_SID = 3
        TRUSTEE_IS_OBJECTS_AND_NAME = 4
    End Enum

    Private Enum TRUSTEE_TYPE As UInteger
        TRUSTEE_IS_UNKNOWN = 0
        TRUSTEE_IS_USER = 1
        TRUSTEE_IS_GROUP = 2
        TRUSTEE_IS_DOMAIN = 3
        TRUSTEE_IS_ALIAS = 4
        TRUSTEE_IS_WELL_KNOWN_GROUP = 5
        TRUSTEE_IS_DELETED = 6
        TRUSTEE_IS_INVALID = 7
        TRUSTEE_IS_COMPUTER = 8
    End Enum

#End Region

#Region "Structures"

    <StructLayoutAttribute(LayoutKind.Sequential)> _
    Private Structure SHARE_INFO_502
        <MarshalAsAttribute(UnmanagedType.LPWStr)> _
        Public shi502_netname As String
        Public shi502_type As UInteger
        <MarshalAsAttribute(UnmanagedType.LPWStr)> _
        Public shi502_remark As String
        Public shi502_permissions As Integer
        Public shi502_max_uses As Integer
        Public shi502_current_uses As Integer
        <MarshalAsAttribute(UnmanagedType.LPWStr)> _
        Public shi502_path As String
        <MarshalAsAttribute(UnmanagedType.LPWStr)> _
        Public shi502_passwd As String
        Public shi502_reserved As Integer
        Public shi502_security_descriptor As IntPtr
    End Structure

    <StructLayoutAttribute(LayoutKind.Sequential)> _
    Private Structure SECURITY_DESCRIPTOR
        Public Revision As Byte
        Public Sbz1 As Byte
        Public Control As UShort
        Public Owner As IntPtr
        Public Group As IntPtr
        Public Sacl As IntPtr
        Public Dacl As IntPtr
    End Structure

    <StructLayoutAttribute(LayoutKind.Sequential)> _
    Private Structure EXPLICIT_ACCESS
        Public grfAccessPermissions As UInteger
        Public grfAccessMode As ACCESS_MODE
        Public grfInheritance As UInteger
        Public Trustee As TRUSTEE
    End Structure

    <StructLayoutAttribute(LayoutKind.Sequential)> _
    Private Structure TRUSTEE
        Public pMultipleTrustee As UInteger
        Public MultipleTrusteeOperation As MULTIPLE_TRUSTEE_OPERATION
        Public TrusteeForm As TRUSTEE_FORM
        Public TrusteeType As TRUSTEE_TYPE
        <MarshalAsAttribute(UnmanagedType.LPTStr)> _
        Public ptstrName As String
    End Structure

#End Region

#Region "Native Methods"

    <DllImportAttribute("netapi32.dll", EntryPoint:="NetShareAdd", CharSet:=CharSet.Unicode)> _
    Private Shared Function NetShareAdd(<InAttribute(), MarshalAsAttribute(UnmanagedType.LPWStr)> ByVal servername As String, ByVal level As UInteger, <InAttribute()> ByRef buf As SHARE_INFO_502, <OutAttribute()> ByRef parm_err As Integer) As NET_API_STATUS
    End Function

    <DllImportAttribute("advapi32.dll", EntryPoint:="InitializeSecurityDescriptor")> _
    Private Shared Function InitializeSecurityDescriptor(ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal dwRevision As UInteger) As UInteger
    End Function

    <DllImportAttribute("advapi32.dll", EntryPoint:="SetEntriesInAclW")> _
    Private Shared Function SetEntriesInAcl(ByVal cCountOfExplicitEntries As Integer, <InAttribute()> ByRef pListOfExplicitEntries As EXPLICIT_ACCESS, <InAttribute()> ByVal OldAcl As System.IntPtr, ByRef NewAcl As System.IntPtr) As UInteger
    End Function

    <DllImportAttribute("advapi32.dll", EntryPoint:="SetSecurityDescriptorDacl")> _
    Private Shared Function SetSecurityDescriptorDacl(ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, <MarshalAsAttribute(UnmanagedType.Bool)> ByVal bDaclPresent As Boolean, <InAttribute()> ByVal pDacl As System.IntPtr, <MarshalAsAttribute(UnmanagedType.Bool)> ByVal bDaclDefaulted As Boolean) As UInteger
    End Function

    <DllImportAttribute("advapi32.dll", EntryPoint:="IsValidSecurityDescriptor")> _
    Private Shared Function IsValidSecurityDesctiptor(ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR) As UInteger
    End Function

#End Region

#Region "Managed Methods"

    ''' <summary>
    ''' Shares an existing folder on the local computer or on a remote computer
    ''' </summary>
    ''' <param name="ShareName">The name for the share</param>
    ''' <param name="ShareComment">An optional comment/description for the share</param>
    ''' <param name="LocalPath">The local path to the folder to be shared. If creating share on a remote computer then 
    ''' the path must be local to the remote computer. Do not use UNC paths</param>
    ''' <param name="SharePermissions">The share permissions to assign to this share</param>
    ''' <param name="ComputerName">OPTIONAL - the remote computer name to create the share on</param>
    Public Shared Function ShareExistingFolder(ByVal ShareName As String, ByVal ShareComment As String, ByVal LocalPath As String, ByVal SharePermissions As List(Of SharePermissionEntry), Optional ByVal ComputerName As String = Nothing) As NET_API_STATUS
        'Argument validation
        If String.IsNullOrEmpty(ShareName) OrElse String.IsNullOrEmpty(LocalPath) OrElse SharePermissions Is Nothing OrElse SharePermissions.Count = 0 Then
            Throw New ArgumentException("Invalid argument specified - ShareName, LocalPath and SharePermissions arguments must not be empty")
        End If

        'Create array of explicit access rules, one for each user specified in the SharePermissions argument
        Dim ExplicitAccessRule(SharePermissions.Count - 1) As EXPLICIT_ACCESS
        'This pointer will hold the full ACL (access control list) once the loop below has completed
        Dim AclPtr As IntPtr

        'Loop through each entry in our list of explicit access rules, build each one and add it to the ACL
        For i As Integer = 0 To ExplicitAccessRule.Length - 1
            'Build the user or group name
            Dim FullAccountName As String = String.Empty
            If Not String.IsNullOrEmpty(SharePermissions(i).DomainName) Then
                FullAccountName = SharePermissions(i).DomainName & "\"
            End If
            FullAccountName &= SharePermissions(i).UserOrGroupName
            'Create a TRUSTEE structure and populate it with the user account details
            Dim Account As New TRUSTEE
            With Account
                .MultipleTrusteeOperation = MULTIPLE_TRUSTEE_OPERATION.NO_MULTIPLE_TRUSTEE
                .pMultipleTrustee = 0
                .TrusteeForm = TRUSTEE_FORM.TRUSTEE_IS_NAME
                .ptstrName = FullAccountName
                .TrusteeType = TRUSTEE_TYPE.TRUSTEE_IS_UNKNOWN
            End With
            'Populate the explicit access rule for this user/permission
            With ExplicitAccessRule(i)
                'Set this to an Allow or Deny entry based on what was specified in the AllowOrDeny property
                If SharePermissions(i).AllowOrDeny Then
                    .grfAccessMode = ACCESS_MODE.GRANT_ACCESS
                Else
                    .grfAccessMode = ACCESS_MODE.DENY_ACCESS
                End If
                'Build the access mask for the share permission specified for this user
                If SharePermissions(i).Permission = SharedFolder.SharePermissions.Read Then
                    .grfAccessPermissions = ACCESS_MASK.GENERIC_READ Or ACCESS_MASK.STANDARD_RIGHTS_READ Or ACCESS_MASK.GENERIC_EXECUTE
                ElseIf SharePermissions(i).Permission = SharedFolder.SharePermissions.FullControl Then
                    .grfAccessPermissions = ACCESS_MASK.GENERIC_ALL
                End If
                'Not relevant for share permissions so just set to NO_INHERITANCE
                .grfInheritance = NO_INHERITANCE
                'Set the Trustee to the TRUSTEE structure we created earlier in the loop
                .Trustee = Account
            End With
            'Add this explicit access rule to the ACL
            Dim SetEntriesResult As UInteger = SetEntriesInAcl(1, ExplicitAccessRule(i), AclPtr, AclPtr)
            'Check the result of the SetEntriesInAcl API call
            If SetEntriesResult = ERROR_NONE_MAPPED Then
                Throw New ApplicationException("The account " & FullAccountName & " could not be mapped to a security identifier (SID). " & _
                                               "Check that the account name is correct and that the domain where the account is held is contactable. The share has not been created")
            ElseIf SetEntriesResult <> 0 Then
                Throw New ApplicationException("The account " & FullAccountName & " could not be added to the ACL as the follow error was encountered: " & SetEntriesResult & ". The share has not been created")
            End If
        Next

        'Create a SECURITY_DESCRIPTOR structure and set the Revision number
        Dim SecDesc As New SECURITY_DESCRIPTOR
        SecDesc.Revision = SECURITY_DESCRIPTOR_REVISION
        'Initialise the SECURITY_DESCRIPTOR instance - returns 0 if an error was encountered
        Dim DecriptorInitResult As UInteger = InitializeSecurityDescriptor(SecDesc, SECURITY_DESCRIPTOR_REVISION)
        If DecriptorInitResult = 0 Then
            Throw New ApplicationException("An error was encountered during the call to the InitializeSecurityDescriptor API. The share has not been created." & _
                                           "You may be able to get more information on the error by checking the Err.LastDllError property")
        End If
        'Add the ACL to the SECURITY_DESCRIPTOR
        Dim SetSecurityResult As UInteger = SetSecurityDescriptorDacl(SecDesc, True, AclPtr, False)
        If SetSecurityResult = 0 Then
            Throw New ApplicationException("An error was encountered during the call to the SetSecurityDescriptorDacl API. The share has not been created." & _
                                            "You may be able to get more information on the error by checking the Err.LastDllError property")
        End If
        'Check to make sure the SECURITY_DESCRIPTOR is valid
        If IsValidSecurityDesctiptor(SecDesc) = 0 Then
            Throw New ApplicationException("No errors were reported from previous API calls but the security descriptor is not valid. The share has not been created.")
        End If
        'Create a pointer for the SECURITY_DESCRIPTOR so that we can pass this in to the SHARE_INFO_502 structure
        Dim SecDescPtr As IntPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(SecDesc))
        Marshal.StructureToPtr(SecDesc, SecDescPtr, False)
        'Create and populate the SHARE_INFO_502 structure that specifies all of the share settings
        Dim ShareInfo As New SHARE_INFO_502
        With ShareInfo
            .shi502_netname = ShareName
            .shi502_type = STYPE_DISKTREE
            .shi502_remark = ShareComment
            .shi502_permissions = 0
            .shi502_max_uses = -1
            .shi502_current_uses = 0
            .shi502_path = LocalPath
            .shi502_passwd = Nothing
            .shi502_reserved = 0
            .shi502_security_descriptor = SecDescPtr
        End With
        'Call the NetShareAdd API to create the share
        Dim Result As NET_API_STATUS = NetShareAdd(ComputerName, 502, ShareInfo, 0)
        'Clean up and return the result of NetShareAdd
        Marshal.FreeCoTaskMem(SecDescPtr)
        Return Result
    End Function

#End Region

End Class
