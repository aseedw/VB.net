********************************************************************
Source : http://www.vbforums.com/showthread.php?616267-Creating-shared-folders-and-specifying-share-permissions
Title : Creating shared folders and specifying share permissions
Arthor : chris128
********************************************************************

Also note that these APIs do not seem to work on a 64 bit OS if your program is set to target 64 bit CPU - if you set your project to target x86 then it works fine on 64 bit and 32 bit OS but if you set to AnyCPU then it will only work correctly on a 32 bit OS.

EDIT: New version of SharedFolder class uploaded that now allows you to create a share on a remote machine. There is now an optional argument for the SharedFolder.ShareExistingFolder method where you can specify the remote machine name - do not use "\\ServerName", just "ServerName" will do it.
If you want to create the share locally then just either miss this argument off or specify Nothing for it.

Here is an example of sharing the folder D:\Shared as "Shared Folder" on a remote server called TestServer:
  SharedFolder.ShareExistingFolder("Shared Folder", "This is an example comment", "D:\Shared", PermissionsList, "TestServer")
