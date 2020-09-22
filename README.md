<div align="center">

## FilesSearch


</div>

### Description

This sub/function searches your hard drive(s) or directories for file(s) like the Windows 'Find Files or Folders...'. It uses mainly the Dir() command and can be used with any programs and visual basic I have encountered. This helps uses to quickly find a file or program for their applications.
 
### More Info
 
It needs two parameters, the start directory or drive and the extension. Example: FilesSearch "C:\", "*.txt".

None that I am aware of.

Finds the file(s) with a particular extension from a start directory.

None. Slightly slower than the windows 'Find Files or Folder...' function.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andrew Tang](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andrew-tang.md)
**Level**          |Unknown
**User Rating**    |5.0 (40 globes from 8 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andrew-tang-filessearch__1-900/archive/master.zip)





### Source Code

```
Sub FilesSearch(DrivePath As String, Ext As String)
Dim XDir() As String
Dim TmpDir As String
Dim FFound As String
Dim DirCount As Integer
Dim X As Integer
'Initialises Variables
DirCount = 0
ReDim XDir(0) As String
XDir(DirCount) = ""
If Right(DrivePath, 1) <> "\" Then
  DrivePath = DrivePath & "\"
End If
'Enter here the code for showing the path being
'search. Example: Form1.label2 = DrivePath
'Search for all directories and store in the
'XDir() variable
DoEvents
TmpDir = Dir(DrivePath, vbDirectory)
Do While TmpDir <> ""
  If TmpDir <> "." And TmpDir <> ".." Then
    If (GetAttr(DrivePath & TmpDir) And vbDirectory) = vbDirectory Then
      XDir(DirCount) = DrivePath & TmpDir & "\"
      DirCount = DirCount + 1
      ReDim Preserve XDir(DirCount) As String
    End If
  End If
  TmpDir = Dir
Loop
'Searches for the files given by extension Ext
FFound = Dir(DrivePath & Ext)
Do Until FFound = ""
  'Code in here for the actions of the files found.
  'Files found stored in the variable FFound.
  'Example: Form1.list1.AddItem DrivePath & FFound
  FFound = Dir
Loop
'Recursive searches through all sub directories
For X = 0 To (UBound(XDir) - 1)
  FilesSearch XDir(X), Ext
Next X
End Sub
```

