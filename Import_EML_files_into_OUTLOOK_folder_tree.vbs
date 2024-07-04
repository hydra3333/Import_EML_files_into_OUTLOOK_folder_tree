Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const local_disk_folder_tree_path As String = "D:\EXPORTED_FROM_THUNDERBIRD"
Const local_disk_folder_tree_DONE_path As String = "D:\EXPORTED_FROM_THUNDERBIRD_DONE"
Const specific_top_level_OUTLOOK_folder_name As String = "NEW_IMPORTED_FROM_THUNDERBIRD"

' Module-level variables
Public originalEntryID As String
Public objOutlookApp As Outlook.Application
Public objOutlookNamespace As Outlook.NameSpace
Public fso As Object
Public email_count As Integer

Const olMailItem As Integer = 0 ' This assumes you have olMailItem defined elsewhere in your code
Const olMIME As Integer = 41 ' Defined in the Outlook library
Const olDiscard As Integer = 1 ' olDiscard is a constant for discarding changes

Sub Main()
    Dim objFolder As Outlook.Folder
    Dim specificFolderObj As Outlook.Folder

    Debug.Print "Starting."

    ' Initialize Outlook and FileSystemObject
    Set objOutlookApp = New Outlook.Application
    Set objOutlookNamespace = objOutlookApp.GetNamespace("MAPI")
    Set fso = CreateObject("Scripting.FileSystemObject")
  
    ' Call function to find and return a specific top-level folder object
    Set specificFolderObj = FindSpecificTopLevelFolder(specific_top_level_OUTLOOK_folder_name, objOutlookNamespace.folders)
    If specificFolderObj Is Nothing Then
        Debug.Print "Specific Top Level Folder not found: " & specific_top_level_OUTLOOK_folder_name
        End ' Quits the VBA execution immediately
    Else
        Debug.Print "Specific Top Level Folder found: " & specificFolderObj.folderPath
    End If
    
    ' Call function to recreate folder tree under the specific top-level folder in Outlook
    RecreateFolderTreeInOutlook local_disk_folder_tree_path, specificFolderObj
    
    ' Set the global originalEntryID to enable checking for new email being opened successfully across folders as well
    originalEntryID = ""
    email_count = 0
    
    ' Call function to import .eml files into corresponding Outlook folders
    ImportEmailsIntoOutlookFolders local_disk_folder_tree_path, specificFolderObj
    
    MsgBox "Imported emails: " & email_count, vbInformation Or vbOKOnly
    
    ' Cleanup
    Set specificFolderObj = Nothing
    Set objOutlookNamespace = Nothing
    Set objOutlookApp = Nothing
    Set fso = Nothing

    Debug.Print "Finished."
End Sub

' Function to find and return a specific top-level folder object
Function FindSpecificTopLevelFolder(folderName As String, folders As Outlook.folders) As Outlook.Folder
    Dim objFolder As Outlook.Folder
    For Each objFolder In folders
        If objFolder.Name = folderName Then
            Set FindSpecificTopLevelFolder = objFolder
            Exit Function
        End If
    Next
    ' Return Nothing if folder not found
    Set FindSpecificTopLevelFolder = Nothing
End Function

' Function to recreate folder tree from local disk under a specific Outlook folder
Sub RecreateFolderTreeInOutlook(localDiskFolderPath As String, parentFolderInOutlook As Outlook.Folder)
    Dim localDiskFolder As Object
    Dim objSubFolder As Object
    Dim newOutlookFolder As Outlook.Folder
  
    ' Check if the local folder exists
    If Not fso.FolderExists(localDiskFolderPath) Then
        Debug.Print "Local disk folder to copy does not exist: " & localDiskFolderPath
        End ' Quits the VBA execution immediately
    End If
    
    ' Get the local folder object
    Set localDiskFolder = fso.GetFolder(localDiskFolderPath)
    
    ' Iterate through subfolders in the local disk folder
    For Each objSubFolder In localDiskFolder.SubFolders
        ' Check if the folder already exists in Outlook
        If Not FolderExistsInOutlook(objSubFolder.Name, parentFolderInOutlook.folders) Then
            ' Create a corresponding folder in Outlook under the parent folder
            Set newOutlookFolder = parentFolderInOutlook.folders.Add(objSubFolder.Name)
            Debug.Print "Created Outlook tree subfolder: " & newOutlookFolder.folderPath
        Else
            ' Folder already exists in Outlook, retrieve the existing folder
            Set newOutlookFolder = GetFolderFromOutlook(objSubFolder.Name, parentFolderInOutlook.folders)
            Debug.Print "Outlook folder already exists in tree, using it: '" & newOutlookFolder.folderPath & "'"
        End If
        
        ' Recursively call to handle subfolders
        RecreateFolderTreeInOutlook objSubFolder.Path, newOutlookFolder
    Next
    
End Sub

' Function to check if a folder exists in Outlook
Function FolderExistsInOutlook(folderName As String, folders As Outlook.folders) As Boolean
    Dim objFolder As Outlook.Folder
    For Each objFolder In folders
        If objFolder.Name = folderName Then
            FolderExistsInOutlook = True
            Exit Function
        End If
    Next
    ' Return False if folder not found
    FolderExistsInOutlook = False
End Function

' Function to retrieve a folder from Outlook
Function GetFolderFromOutlook(folderName As String, folders As Outlook.folders) As Outlook.Folder
    Dim objFolder As Outlook.Folder
    For Each objFolder In folders
        If objFolder.Name = folderName Then
            Set GetFolderFromOutlook = objFolder
            Exit Function
        End If
    Next
    ' Return Nothing if folder not found
    Set GetFolderFromOutlook = Nothing
End Function

' Function to import .eml files into corresponding Outlook folders
Sub ImportEmailsIntoOutlookFolders(localDiskFolderPath As String, parentFolderInOutlook As Outlook.Folder)
    Dim localDiskFolder As Object
    Dim emlFile As Object
    Dim outlookFolder As Outlook.Folder
    Dim objSubFolder As Object
    
    Debug.Print "Started: ImportEmailsIntoOutlookFolders " & localDiskFolderPath
    
    ' Check if the local folder exists
    If Not fso.FolderExists(localDiskFolderPath) Then
        Debug.Print "Local disk folder to import .eml files does not exist: " & localDiskFolderPath
        End ' Quits the VBA execution immediately
    End If
    
    ' Get the local folder object
    Set localDiskFolder = fso.GetFolder(localDiskFolderPath)
    
    ' Iterate through subfolders in the local disk folder
    For Each objSubFolder In localDiskFolder.SubFolders
        ' Check if the folder exists in Outlook
        Set outlookFolder = GetFolderFromOutlook(objSubFolder.Name, parentFolderInOutlook.folders)
        
        If Not outlookFolder Is Nothing Then
            ' Iterate through .eml files in the local folder
            Debug.Print "Looking for and Importing .eml files into Outlook folder: '" & outlookFolder.folderPath & "'"
            
            For Each emlFile In objSubFolder.Files
                ' Check if the file is an .eml file
                If LCase(fso.GetExtensionName(emlFile.Path)) = "eml" Then
                    ' Import each .eml file into Outlook folder
                    'Debug.Print "Importing file: '" & emlFile.Path & "'"
                    ImportEmlFileIntoOutlook emlFile.Path, outlookFolder
                Else
                    'Debug.Print "Ignoring file: '" & emlFile.Path & "'"
                End If
            Next
            
            ' Recursively process subfolders as well
            ImportEmailsIntoOutlookFolders objSubFolder.Path, outlookFolder
        Else
            Debug.Print "Outlook folder does not exist for '" & objSubFolder.Path & "'. Quitting..."
            End ' Quits the VBA execution immediately
        End If
    Next
    
    Debug.Print "Finished: ImportEmailsIntoOutlookFolders " & localDiskFolderPath
End Sub

Sub ImportEmlFileIntoOutlook(emlFilePath As String, outlookFolder As Outlook.Folder)
    Dim insp As Outlook.Inspector
    Dim m As Object, m2 As Object, m3 As Object
    Dim p As Integer, retries As Integer, i As Integer
    Const max_retries As Integer = 100

    ' Ensure the .eml file is accessible
    If Not fso.FileExists(emlFilePath) Then
        Debug.Print ".eml File does not exist: " & emlFilePath
        End ' Quits the VBA execution immediately
    End If

    Debug.Print "Attempting to import .eml file '" & emlFilePath & "' into Outlook folder '" & outlookFolder.folderPath & "'"
    
    ' Open the .eml file using Shell to ensure it opens in Outlook
    Shell "explorer.exe """ & emlFilePath & """"
    
    ' Allow some time for the .eml file to open
    For i = 1 To 5
        DoEvents
    Next
    Sleep 250 ' You might need to adjust this value depending on your system performance
    
    ' Get the active inspector
    retries = 0
    Set insp = objOutlookApp.ActiveInspector
    ' Wait for the inspector to become available
    While TypeName(insp) = "Nothing"
        Sleep 50
        DoEvents
        Sleep 50
        Set insp = objOutlookApp.ActiveInspector
        retries = retries + 1
        If retries > max_retries Then
            Debug.Print "Error: Could not find an open inspector for importing email."
            End ' Quits the VBA execution immediately
        End If
    Wend
    ' Ensure there is an active inspector
    If TypeName(insp) = "Nothing" Then
        Debug.Print "Error: After loop, could not find an open inspector for importing email."
        End ' Quits the VBA execution immediately
    End If
    
    ' Get the current item in the inspector (the .eml file)
    'Set m = insp.CurrentItem
    ' Copy the item
    'Set m2 = m.Copy
    ' Move the copied item to the target Outlook folder
    'Set m3 = m2.Move(outlookFolder)
    ' Save the moved item
    'm3.Save

    ' Get the current item in the inspector (the .eml file)
    Set m = insp.CurrentItem
    ' Move the copied item to the target Outlook folder
    Set m3 = m.Move(outlookFolder)
    ' Save the moved item
    m3.Save
    
    ' Close the inspector without saving
    insp.Close olDiscard
    
    ' Move the .eml file to the corresponding location in DONE folder tree
    MoveEmlFileToDoneFolder emlFilePath, local_disk_folder_tree_path, local_disk_folder_tree_DONE_path
    ' Optionally, success message
    Debug.Print "Imported .eml file '" & emlFilePath & "' into Outlook folder '" & outlookFolder.folderPath & "'"

    email_count = email_count + 1
    'If email_count Mod 5 = 0 Then
    '    MsgBox "5 done", vbInformation Or vbOKOnly
    '    End
    'End If

    ' Clean up
    Set m = Nothing
    Set m2 = Nothing
    Set m3 = Nothing
    Set insp = Nothing
End Sub

Sub MoveEmlFileToDoneFolder(emlFilePath As String, originalTreePath As String, doneTreePath As String)
    Dim relativePath As String
    Dim doneFilePath As String
    
    ' Construct the new file path in the DONE folder tree
    doneFilePath = Replace(emlFilePath, originalTreePath, doneTreePath)

    ' Ensure the destination folder exists, create it if it does not
    If Not fso.FolderExists(fso.GetParentFolderName(doneFilePath)) Then
        fso.CreateFolder fso.GetParentFolderName(doneFilePath)
    End If
    
    ' Move the .eml file to the DONE folder tree
    fso.MoveFile emlFilePath, doneFilePath
    Debug.Print "Moved .eml file to DONE folder: " & doneFilePath
End Sub

Function EscapeSpecialCharacters(inputString As String) As String
    Dim resultString As String
    
    ' List of special characters to escape (add more as needed)
    Dim specialChars As String
    specialChars = "\/:*?""<>|"
    
    ' Initialize result string
    resultString = inputString
    
    ' Loop through each character in the input string
    For i = 1 To Len(specialChars)
        resultString = Replace(resultString, Mid(specialChars, i, 1), "\" & Mid(specialChars, i, 1))
    Next i
    
    ' Return the escaped string
    EscapeSpecialCharacters = resultString
End Function


