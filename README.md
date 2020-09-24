<div align="center">

## Deleting Directory with all the Subdirectories


</div>

### Description

Delete a folder, with all it's subdirectories.
 
### More Info
 
The folder name, for example:

"c:\My documents\stuff" will delete the "stuff" folder.

Simple recursion.

System and hidden files will not be deleted, so the directory won't be deleted either. However You cen rewrite the function a little to work with hidden file.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Vlad Azarkhin](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vlad-azarkhin.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vlad-azarkhin-deleting-directory-with-all-the-subdirectories__1-6741/archive/master.zip)





### Source Code

```
Public Sub CleanAllPath(sPath As String)
Dim sName As String
Dim sFullName As String
' Array used for holding the directories,
' however collection may be used as well
Dim Dirs() As String
Dim DirsNo As Integer
Dim i As Integer
 If Not Right(sPath, 1) = "\" Then
 sPath = sPath & "\"
 End If
 ' clean all files in the directory
 sName = Dir(sPath & "*.*")
 While Len(sName) > 0
 sFullName = sPath & sName
 SetAttr sFullName, vbNormal
 Kill sFullName
 sName = Dir
 Wend
 sName = Dir(sPath & "*.*", vbHidden)
 While Len(sName) > 0
 sFullName = sPath & sName
 SetAttr sFullName, vbNormal
 Kill sFullName
 sName = Dir
 Wend
 ' read all the directories into array
 DirsNo = 0
 sName = Dir(sPath, vbDirectory)
 While Len(sName) > 0
 If sName <> "." And sName <> ".." Then
  DirsNo = DirsNo + 1
  ReDim Preserve Dirs(DirsNo) As String
  Dirs(DirsNo - 1) = sName
 End If
 sName = Dir
 Wend
 For i = 0 To DirsNo - 1
 CleanAllPath (sPath & Dirs(i) & "\")
 RmDir sPath & Dirs(i)
 Next
End Sub
```

