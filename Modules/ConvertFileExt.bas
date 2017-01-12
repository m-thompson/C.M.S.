Attribute VB_Name = "ConvertFileExt"
Option Explicit


Public Sub f1()
Dim fso As New FileSystemObject
Dim fil As File
    
    For Each fil In fso.GetFolder("C:\Program Files\Congregation Management System\Songs").Files
        If fso.GetExtensionName(fil.Path) = "wma" Then
            fil.Name = fso.GetBaseName(fil.Path) & ".cms"
        End If
    Next


End Sub
