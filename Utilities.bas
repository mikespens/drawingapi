Attribute VB_Name = "Utilities"
Option Explicit
'add a reference to Microsoft Scripting Runtime

Public Function ChangeExtension(filePath As String, extension As String) As String
    Dim fso As New FileSystemObject
    
    Dim currentExt As String
    currentExt = fso.GetExtensionName(filePath)
    
    ChangeExtension = Strings.Replace(filePath, currentExt, extension)
End Function
