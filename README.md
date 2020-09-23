<div align="center">

## Prevent your application from unauthorised person


</div>

### Description

Prevent your files from copying by unauthorised person.

By default any file's Attribute is Archive.

You have to just remove this Attribute from it's property.

When any one copy this file,at this time the file is again

became Archive we can use this to identifying Unauthorised user.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Milind Patil](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/milind-patil.md)
**Level**          |Advanced
**User Rating**    |4.2 (50 globes from 12 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/milind-patil-prevent-your-application-from-unauthorised-person__1-33998/archive/master.zip)





### Source Code

```
' * ************************************************************
' * Programmer Name : Milind M. Patil
' * E-Mail      :       milind_7001@rediffmail.com
' * Date       :       08/29/2001
' **********************************************************************
' * Comments     : Prevent your files from copying by unauthorised person.
' *
' *               By default any file's Attribute is Archive.
' *               You have to just remove this Attribute from it's property.
' *               When any one copy this file,at this time the file is again
' *               became Archive we can use this to identifying Unauthorised user.
' *
' **********************************************************************
' Frist Remove The Archive Attribute Of The File And Follow The Simple Code.
Private Sub command1_Click()
Dim result
result = GetAttr("c:\myprojects\project1.exe") And vbArchive
If result = 0 Then
MsgBox "ok"
Else
MsgBox "sorry unautohrised user"
unload me
End If
End Sub
OR
'''''''You can also use API to check files attributes''''''
''''place the code in module1
Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
''''place the code on form's command1
Private Sub command1_Click()
dim val as long
val = GetFileAttributes("c:\myprojects\project1.vbp") ' read file attributes
If (attribs And FILE_ATTRIBUTES_ARCHIVE) <> 0 Then msgbox "Sorry Unauthorised Person" ........
.
.
.
End Sub
```

