Attribute VB_Name = "Backup_Creator"
'=====================================================================================================================
'                                                 File Backup Creator
'                                          ---------------------------------
' Creates a backup of any files on the list in Column A where the date modified of the original file is >
' the date modified of the most recent backup.
'   - Stores "most recent backup" in a "Most Recent" folder created at the root direcotry specified
'   - When it makes a new backup of a file (org. date modified > backup date modified) it moves the
'     previous backup to a yyyy\yy-mm-dd directory
'   - It deletes or moves any backups older than a specified number of days and have a backup newer than that.
'
'
'
'
'   - Set "RootBackupPath" in CreateBackups() ("C:\Backups") (will create if it doesn't already exist)
'   - Set OutdatedAction in CreateBackups() (default to -1)
'   - Set DaysOld in CheckForOutdatedBackups() (default to 180)
'=====================================================================================================================

Public FSO                          As New FileSystemObject     'FSO object used to interact with the filesystem
Public MostRecentDateModified       As Date                     'Most recent date modified for a specific file, used across multiple subs so must be public
Public RootBackupPath               As String                   'The root backup used, actually defined in "CreateBackups"
Public OutdatedPath                 As String
Public OutdatedAction               As Integer
       

Sub CreateBackups()
Dim LastRow As Long, MostRecentPath As String, PotentialMostRecentFilePath As String, ThisYearFolder As String
Dim OrgFile As File, MRBackup As File

'------------------------------------------------------
' What to do with outdated backups (older than date specified)
    OutdatedAction = -1
'  -1 = do nothing
'   0 = move to "Outdated" folder
'   1 = delete
'------------------------------------------------------

'------------------------------------------------------
RootBackupPath = "C:\Backups"
If Not FSO.FolderExists(RootBackupPath) Then FSO.CreateFolder (RootBackupPath)
'------------------------------------------------------

'Guarantee outdated action is set properly
    If OutdatedAction < -1 Or OutdatedAction > 1 Then Err.Raise Number:=513, Description:="The ""OutdatedAction"" is incorrectly set."
    
    
'Make sure data is valid and every file exists, and exit the sub if there is an error
    If Not DataIsValid Then Exit Sub
    If WorksheetFunction.CountIf(Range("B:B"), False) > 0 Then GoTo MissingFilesError
        
    
'Ensure that "Most Recent" folder, "Outdated", and this years backup folder exists
    MostRecentPath = RootBackupPath & "\Most Recent"
    If Not FSO.FolderExists(MostRecentPath) Then FSO.CreateFolder (MostRecentPath)
    
    'This one is public because it is used in CheckForOutdatedBackups
    OutdatedPath = RootBackupPath & "\Outdated"
    If Not FSO.FolderExists(OutdatedPath) Then FSO.CreateFolder (OutdatedPath)
    
    ThisYearFolder = RootBackupPath & "\" & Year(Now)
    If Not FSO.FolderExists(ThisYearFolder) Then FSO.CreateFolder (ThisYearFolder)


'Loop through every file
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
    For i = 2 To LastRow
        Set OrgFile = FSO.GetFile(Cells(i, 1))
        PotentialMostRecentFilePath = MostRecentPath & "\" & OrgFile.Name
        If FSO.FileExists(PotentialMostRecentFilePath) Then                         ' It HAS a most recent backup: check if the MRDM (most recent date modified) is > OrgFile DM
            Set MRBackup = FSO.GetFile(PotentialMostRecentFilePath)
            If Not OrgFile.DateLastModified > MRBackup.DateLastModified Then GoTo CheckForOlderBackupsToDelete  'It has a most recent backup, but the file has not been updated since then, jump to check for older backups
            
         
            '-----------------------------------------------------------
            'The backup is outdated, move it and create a new backup
            BackupFolderPath = RootBackupPath & "\" & Year(Now) & "\" & Format(Now, "yy-mm-dd")
            If Not FSO.FolderExists(BackupFolderPath) Then FSO.CreateFolder (BackupFolderPath)      'Create the backup directory for today if it doesn't already exist
            FSO.MoveFile MRBackup.Path, BackupFolderPath & "\" & MRBackup.Name                      'Move MRBackup to todays backup directory
            
        End If
        
        
    'Create backup - skipped if file in "Most Recent" is still Most Recent
        OrgFile.Copy (MostRecentPath & "\" & OrgFile.Name)
        
        
CheckForOlderBackupsToDelete:
    'Guaranteed that the file in MostRecent is actually the MostRecent, safe to delete outdated backups
        If OutdatedAction > -1 Then Call CheckForOutdatedBackups(OrgFile)
    
    Next i

Exit Sub

MissingFilesError:
    MsgBox "One or more files do not exist. Please correct the missing files before proceeding.", vbCritical, "Validation Error"
    Exit Sub
    
End Sub
Sub CheckForOutdatedBackups(OrgFile As File)
Dim YearFolder As Folder, DateFolder As Folder, CutOffDate As Date, PotentialFilePath As String, OldBackup As File

'----------------------------------------------------------
'"If a file is older than {$DaysOld} days, move/delete it"
 DaysOld = 180
'----------------------------------------------------------

'Set cut off date
CutOffDate = Now - DaysOld

'Sacrifice AND for readability
For Each YearFolder In FSO.GetFolder(RootBackupPath).SubFolders                                                                                         ' Every year folder...
    If YearFolder.Name <= Format(CutOffDate, "yyyy") Then   'will evaluate to false on any folders not named after the year                             ' Older than cutoff date
        For Each DateFolder In YearFolder.SubFolders                                                                                                    ' Every date folder...
            FolderDate = DateFolder.DateCreated
            If FolderDate <= CutOffDate Then                                                                                                            ' Older than cutoff date...
                PotentialFilePath = DateFolder.Path & "\" & OrgFile.Name
                If FSO.FileExists(PotentialFilePath) Then                                                                                               ' If the file exists...
                    Set OldBackup = FSO.GetFile(PotentialFilePath)
                    '...and the date modified is older than cut off date...
                    If OldBackup.DateLastModified <= CutOffDate Then 'delete or move                                                                    ' ...and is outdated...
                        If OutdatedAction = 0 Then                                                                                                      ' Move or delete it.
                            ' Increment a version number so nothing gets overwritten and move the file to "Outdated" folder created earlier
                            VNumber = 0
                            Do
                                VNumber = VNumber + 1
                                NewFilePath = OutdatedPath & "\" & OldBackup.ShortName & "-" & VNumber & "." & FSO.GetExtensionName(OldBackup.Path)
                            Loop While FSO.FileExists(NewFilePath)
                            FSO.MoveFile OldBackup.Path, NewFilePath
                        ElseIf OutdatedAction = 1 Then
                            FSO.DeleteFile (OldBackup.Path)
                        End If
                    End If
                End If
            End If
        Next DateFolder
    End If
Next YearFolder

Exit Sub

End Sub

Function DataIsValid()
' Check that all files from A2 to bottom exist
Dim LastRow As Long
LastRow = Range("A" & Rows.Count).End(xlUp).Row
If Not Cells(1, "A") = "Filepath" Then GoTo NoHeaderError   '(1,1) must be "Filepath" to make this list easily convertable to a Seach Everything file list (*.efu)
    
For i = 2 To LastRow
    FilePath = Cells(i, 1)
    If FilePath = "" Then GoTo EmptyCellError
    If Not FSO.FileExists(FilePath) Then
        Cells(i, 2) = False
    ElseIf Cells(i, 2) <> "" Then
        Cells(i, 2) = ""
    End If
Next i

DataIsValid = True
Exit Function


'------------------
' Errors
'------------------
EmptyCellError:
    MsgBox "Row " & i & " is blank. Either delete this row or enter a valid File Path.", vbCritical, "Empty Cell Error"
    Cells(i, 1).Select
    DataIsValid = False
    Exit Function

NoHeaderError:
    MsgBox "A1 must = ""Filepath"". Ensure that there is no actual File Path in this cell and then type ""Filepath"" here.", vbCritical, "Validation Error"
    DataIsValid = False
    Exit Function

End Function
Sub GetMostRecentBackup(FileName, RootFilePath)
'=================================================== OBSOLETE ===================================================
'Recursively find the most recent version of a file starting at a root folder

Dim ThisFolder As Folder
Dim Folders As Folders
Dim Folder As Folder
Dim File As File


Set ThisFolder = FSO.GetFolder(RootFilePath)

For Each File In ThisFolder.Files
    If File.Name = FileName Then
        If File.DateLastModified > MostRecentDateModified Then MostRecentDateModified = File.DateLastModified
    End If
Next File

For Each Folder In ThisFolder.SubFolders
    If Folder.DateLastModified > MostRecentDateModified Then Call GetMostRecentBackup(FileName, Folder.Path)
Next Folder

End Sub


