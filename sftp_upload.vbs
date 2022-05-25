' Add-Type -Path "C:\Users\mtadmin\Desktop\Test_sFTP\WinSCPnet.dll"
Option Explicit
' Setup session options
Dim sessionOptions, Protocol_Sftp
Set sessionOptions = WScript.CreateObject("WinSCP.SessionOptions")
With sessionOptions
    .Protocol = Protocol_Sftp
    .HostName = "xxxxxxxxxxxxxxxxx.com"
    .UserName = "xxxxxxxxxxx"
    .Password = "xxxxxxxx"
    .SshHostKeyFingerprint = "ssh-rsa 1024 XXXXXXXXXXXXXXXXXXXxxxx"
End With
Dim session
Set session = WScript.CreateObject("WinSCP.Session")
' Connect
session.Open sessionOptions
' Upload files
Dim transferOptions, TransferMode_Binary
Set transferOptions = WScript.CreateObject("WinSCP.TransferOptions")
transferOptions.TransferMode = TransferMode_Binary
Dim transferResult
Set transferResult = session.PutFiles("D:\testing-file.txt", "/", False, transferOptions)
' Throw on any error
transferResult.Check
' Print results
Dim transfer
For Each transfer In transferResult.Transfers
    WScript.Echo "Upload of " & transfer.FileName & " succeeded"
Next
' Disconnect, clean up
session.Dispose