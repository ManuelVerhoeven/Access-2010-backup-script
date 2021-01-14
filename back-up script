
Private Sub Form_Load()
Dim str As String
Dim buf As String
Dim MD_Date As Variant
Dim fs As Object
Dim source As String
Const conPATH_FILE_ACCESS_ERROR = 75
On Error GoTo Make_Backup_Folder
'buf = Back Up Folder
'buf is created if it does not exist
'CurrentProject.Path = the path that the FE is located
buf = "W:\12-exp (exploration)\9. Mn Exploration\1. Group Admin\Database\Exploration Access Database\2021\Backups\"
'map = CurrentProject.Path & "\MapInfo\" changed how map info export works. maybe figure out this way later
MkDir buf
    Resume Make_Backup_Folder
Make_Backup_Folder:
'On Error Resume Next - old attempt
'Use yyyy-mm-dd hh-mm-ss as folder name. Change as needed.
MD_Date = Format(Date, "dd_mm_yyyy")
'& Format(Time, "Ham/pm")
str = "W:\12-exp (exploration)\9. Mn Exploration\1. Group Admin\Database\Exploration Access Database\2021\Backups\" & MD_Date
If Len(Dir(str, vbDirectory)) = 0 Then
   MkDir str
End If
'map = CurrentProject.Path & "\MapInfo\" changed how map info export works. maybe figure out this way later
'If Len(Dir(map, vbDirectory)) = 0 Then changed how map info export works. maybe figure out this way later
 '  MkDir map
'End If
'MkDir str - original
'On Error GoTo 0 - old attemptS
'Copy_Database: - old attempt
'Source = where the data is stored
'map = CurrentProject.Path & "\MapInfo\" changed how map info export works. maybe figure out this way later
source = "W:\12-exp (exploration)\9. Mn Exploration\1. Group Admin\Database\Exploration Access Database\2021" & "\"
Set fs = CreateObject("Scripting.FileSystemObject")
'Change the file extension as needed
fs.CopyFile source & "*.accdb", str, True
'copy file to mapinfo folder to map .tab files
'fs.CopyFile source & "*.accdb", map, True  -changed how map info export works. maybe figure out this way later
Set fs = Nothing
MsgBox "Data backup at " & vbCrLf & MD_Date & vbCrLf & "successfully!", _
        vbInformation, "Backup Successful"
Exit_Button_Backup:
  Exit Sub
Err_Button_Backup:
  If Err.Number = conPATH_FILE_ACCESS_ERROR Then
    MsgBox "The following Path, " & str & ", already exists or there was an Error " & _
           "accessing it!", vbExclamation, "Path/File Access Error"
  Else
    MsgBox Err.Description, vbExclamation, "Error Creating " & str
  
End If
    Resume Exit_Button_Backup


End Sub

Private Sub NavigationButton102_Click()
Dim str As String
Dim buf As String
Dim MD_Date As Variant
Dim fs As Object
Dim source As String
Const conPATH_FILE_ACCESS_ERROR = 75
On Error GoTo Make_Backup_Folder
'buf = Back Up Folder
'buf is created if it does not exist
'CurrentProject.Path = the path that the FE is located
buf = "W:\12-exp (exploration)\9. Mn Exploration\1. Group Admin\Database\Exploration Access Database\2021\Backups\"
MkDir buf
    Resume Make_Backup_Folder
Make_Backup_Folder:
'On Error Resume Next - old attempt
'Use yyyy-mm-dd hh-mm-ss as folder name. Change as needed.
MD_Date = Format(Date, "dd_mm_yyyy ") & Format(Time, "hh-mmam/pm")
str = "W:\12-exp (exploration)\9. Mn Exploration\1. Group Admin\Database\Exploration Access Database\2021\Backups\" & MD_Date
If Len(Dir(str, vbDirectory)) = 0 Then
   MkDir str
End If
'MkDir str - original
'On Error GoTo 0 - old attempt
'Copy_Database: - old attempt
'Source = where the data is stored
source = "W:\12-exp (exploration)\9. Mn Exploration\1. Group Admin\Database\Exploration Access Database\2021" & "\"
Set fs = CreateObject("Scripting.FileSystemObject")
'Change the file extension as needed
fs.CopyFile source & "*.accdb", str, True
Set fs = Nothing
MsgBox "Data backup at " & vbCrLf & MD_Date & vbCrLf & "successfully!", _
        vbInformation, "Backup Successful"
Exit_Button_Backup:
  Exit Sub
Err_Button_Backup:
  If Err.Number = conPATH_FILE_ACCESS_ERROR Then
    MsgBox "The following Path, " & str & ", already exists or there was an Error " & _
           "accessing it!", vbExclamation, "Path/File Access Error"
  Else
    MsgBox Err.Description, vbExclamation, "Error Creating " & str
  
End If
    Resume Exit_Button_Backup
End Sub
