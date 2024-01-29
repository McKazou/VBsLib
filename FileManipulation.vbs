Option Explicit
On Error Resume Next
'CLASS OBJECT : https://www.tutorialspoint.com/vbscript/vbscript_class_objects.htm
call test()
'To USE ME call the add a "include" sub inside the main script like the one at the bottom of this page
'Include "".\FileManipulation.vbs"

Class FileManipulation

    'Class imports for shell stuff (invoke function)
    Function getUserFolder()
        Dim oShell
        Set oShell = CreateObject("WScript.Shell")
        getUserFolder = oShell.ExpandEnvironmentStrings("%USERPROFILE%")
        Call ErrorHandler
    End Function

    Function moveFileTo(fileToMove, destinationAsPath)
        On Error Resume Next

        if TypeName(fileToMove) = "String" Then
            set fileToMove = getFileFromPath(fileToMove)
        end if

        with CreateObject("Scripting.FileSystemObject")
                Dim fileName
                fileName = .GetFileName(fileToMove)

                WScript.Echo "[INFORMATION]: {fileManipulation} MoveFileto - Moving file : " & fileToMove & " to " & destinationAsPath+"\"+fileName
                'msgbox "Moving: " & fileToMove & " To " & destinationAsPath+"\"+fileName
                .MoveFile fileToMove, destinationAsPath+"\"+fileName
        end With
        
        Call ErrorHandler
    End Function

    'Rename file
    function renameFileTo(ByRef selectedfile, ByVal newName)
        'On Error Resume Next

        if TypeName(selectedfile) = "String" Then
            set selectedfile = getFileFromPath(selectedfile)
        end if

        with CreateObject("Scripting.FileSystemObject")
            'On Error Resume Next
                Dim oldName
                oldName = .GetFileName(selectedfile)

            if( Err.number <> 0 ) then
                'If we have an error here it's probably because we have a duplicated files
                Call Err.Raise(vbObjectError + 10, "{FileManipulation} Rename File", "we may have file that have been renamed to the same name")
                Err.clear()
            else
        
            end if
            'MsgBox("Renaming "& .GetFileName(selectedfile) & " to " & selectedfile.ParentFolder+"\"+newName)
        end With

        WScript.Echo "[INFORMATION]: {fileManipulation} RenameFileTo : " & selectedfile.Name & " to " & newName
        selectedfile.Move(selectedfile.ParentFolder+"\"+newName)

        Call ErrorHandler
    end function

     '---------- RECURSIVE FILE FINDER ----------
    'Recursive : https://stackoverflow.com/questions/14950475/recursively-access-subfolder-files-inside-a-folder
    'List : https://gist.github.com/simply-coded/feb68898fad5b900116b479035ab7f05

    function FileFinder (ByRef selectedFolder, ByVal isRecursif, recursifLimit)
        On Error Resume Next

        if TypeName(selectedFolder) = "String" Then
            set selectedFolder = getFolderFromPath(selectedFolder)
        end if

        'MsgBox "TypeName(selectedFolder)" & TypeName(selectedFolder)

        Wscript.Echo "[INFORMATION]: {fileManipulation} - fecthing in folder: " & selectedFolder
        'Msgbox "[INFORMATION]: {fileManipulation} - fecthing in folder: " & selectedFolder

        Dim filesFound
        set filesFound = CreateObject("System.Collections.ArrayList")

        'MsgBox "TypeName(selectedFolder)" & TypeName(selectedFolder)

        if recursifLimit = 0 Then
            FileFinder = Nothing
            Exit Function
        Else
            Dim file
            For Each file In selectedFolder.Files
                filesFound.add file
                'Wscript.Echo "[INFORMATION]: {fileManipulation} - File found: " & file.Path
            next
            if isRecursif then
                Dim subFolder
                Dim filesFoundinSubFolder
                set filesFoundinSubFolder = CreateObject("System.Collections.ArrayList")
            
                'Call the same function for subfolder
                for Each subFolder in selectedFolder.SubFolders
                    set filesFoundinSubFolder = FileFinder(subFolder, true, recursifLimit-1)
                next

                'if their is actually files found add it to the current list
                if filesFoundinSubFolder.Count > 0 then
                    For Each file in filesFoundinSubFolder 
                    filesFound.add file
                    next
                end if
            end if
        set Filefinder = filesFound
        end if

        Call ErrorHandler
    end function

    function getPathWhereScriptIsRun()
        '------GET CURRENT FOLDER OR SPECIFIC FOLDER IN TEST MODE-------
        Dim WshShell
        Set WshShell = CreateObject("WScript.Shell")
        getPathWhereScriptIsRun = WshShell.CurrentDirectory
        Call ErrorHandler
    end function

    function getFolderWhereScriptIsRun()
        '------GET CURRENT FOLDER OR SPECIFIC FOLDER IN TEST MODE-------
        Dim objFso
        Set objFso = CreateObject("Scripting.FileSystemObject")
        Dim curPath
        curPath = getPathWhereScriptIsRun()
        Dim selectedFolder
            '1. Get the current path where this script is located
            Set selectedFolder = objFSO.GetFolder(curPath)
            getFolderWhereScriptIsRun = selectedFolder
        set WshShell = Nothing
        set objFso = nothing
        '----------------------------------------------------------
        Call ErrorHandler
    end function

    Function getFolderFromPath(pathAsString)
        Dim objFso
        Set objFso = CreateObject("Scripting.FileSystemObject")
        Set getFolderFromPath = objFSO.GetFolder(pathAsString)
        set objFso = nothing

        Call ErrorHandler
    End Function

    function getFileFromPath(pathAsString)
        Dim objFso
        Set objFso = CreateObject("Scripting.FileSystemObject")
        Set getFileFromPath = objFSO.GetFile(pathAsString)
        set objFso = nothing

        Call ErrorHandler
    end function

    function FileExists(fileToTest)
        'source : https://www.tachytelic.net/2017/10/function-check-file-exists-vbscript/
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(fileToTest) Then
            FileExists=true
        Else
            FileExists=false
        End If
    End Function

    Private Function ErrorHandler()
        If Err.Number <> 0 Then
            ' MsgBox or whatever. You may want to display or log your error there
            Call Err.Raise(vbObjectError + 10, "File Manipulation", "Unknow error with description: "&Err.Description)
            WScript.Echo("[ERROR] {File Manipulation} Unknow error with description: "&Err.Description)
            Err.Clear
        End If
    end function

    '----------------EVENTS----------------
    Private Sub Class_Initialize(  )
        'Getting all processRunning at object creation
        Call ErrorHandler
    End Sub
    
    'When Object is Set to Nothing
    Private Sub Class_Terminate(  )
        'Termination code goes here
        Call ErrorHandler
    End Sub

    Private Sub Include( scriptName )
        WScript.Echo "[LOADING]: "&scriptName
        Dim sScript
        Dim oStream
        With CreateObject( "Scripting.FileSystemobject" )
            Set oStream = .OpenTextFile(scriptName)
        End With
        sScript = oStream.ReadAll()
        oStream.Close
        ExecuteGlobal sScript
        WScript.Echo "[LOADED]: "&scriptName

        Call ErrorHandler
    End Sub

end class

'---------TESTING-----------
function test()

    Dim fileManip
    set fileManip = new FileManipulation

    Dim activefolder
    activefolder = fileManip.getPathWhereScriptIsRun()
    WScript.Echo "[TESTING] The script is run from: " & activefolder

    Dim userFolder
    userFolder = fileManip.getUserFolder()
    WScript.Echo "[TESTING] User's home folder is: " & userFolder
    WScript.Echo "[TESTING RESULT] getFolderFromPath: " & fileManip.getFolderFromPath(userFolder & "\Downloads\").path

    'filePresence
    WScript.Echo "[TESTING] if file exist <.\VBsLib\renameMe.txt> "
    Dim filePresence
    filePresence = fileManip.FileExists(".\VBsLib\renameMe.txt")
    WScript.Echo "[TESTING RESULT] file exist " & filePresence

    Dim file, files, path
    path = ".\VBsLib\"
    set files = fileManip.FileFinder(path,false,1)
    For each file in files
        WScript.Echo "[TESTING] File found: " & file
    Next

    Dim filePathToRename
    filepathToRename = ".\VBsLib\renameMe.txt"
    fileManip.renameFileTo filepathToRename, "RenameMe.txt"
    fileManip.renameFileTo filepathToRename, "renameMe.txt"

    Dim moveFileTo
    fileManip.moveFileTo filepathToRename, movefileTo
    fileManip.moveFileTo ".\VBsLib\renameMe.txt", "..\"
end function
