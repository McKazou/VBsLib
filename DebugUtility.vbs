Option Explicit
'On Error Resume Next
'To USE ME call the add a "include" sub inside the main script like the one at the bottom of this page
'Include .\FileManipulation.vbs
'------------ReUSE OTHER SCRIPT----------------
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")
dim currentDir
currentDir = WshShell.CurrentDirectory
'WScript.Echo "[modificationWatcher] Started from : "& currentDir&"\FileManipulation.vbs"
Include currentDir&"\FileManipulation.vbs"

'CLASS OBJECT : https://www.tutorialspoint.com/vbscript/vbscript_class_objects.htm
'testDebugUtility()

Class DebugUtility

    'Class imports for shell stuff (invoke function)

    'Key: Printers name, Value: PORT
    private isWrittingToLogFile

    Private logPath
    Public Function setLogPath(pathWhereLogShouldBe)
        logPath = pathWhereLogShouldBe
    End Function
    'Could add also the timming stuff to track the time something ?
    
    Public Function printSimpleString(ByVal stringToPrint)
        'https://stackoverflow.com/questions/17194375/i-need-to-write-vbs-wscript-echo-output-to-text-or-cvs
        'The idea here is to replace Wscript.Echo with Debug.print and be able to differentiate between the app running in debug like VS code ?
        if isWrittingToLogFile Then
            WScript.Echo stringToPrint
        else
            with CreateObject("Scripting.FileSystemObject")
                Dim fileManip, f
                set fileManip = new FileManipulation
                
                if fileManip.FileExists(logPath&"\log.txt") Then
                    Set f = .OpenTextFile( logPath&"\log.txt", 8)
                    f.WriteLine stringToPrint
                    f.Close
                else
                    Set f = .CreateTextFile( logPath&"\log.txt", 2)
                    f.WriteLine stringToPrint
                    f.Close
                end if  
            end with
        end if
    End Function

    Public Function printTable(ByVal tableToPrint)
        'https://stackoverflow.com/questions/17194375/i-need-to-write-vbs-wscript-echo-output-to-text-or-cvs
        'The idea here is to replace Wscript.Echo with Debug.print and be able to differentiate between the app running in debug like VS code ?
        Dim itemNumber 
        itemNumber = 0
        Dim item
        for each item in tableToPrint
            printSimpleString(itemNumber &":"& item)
            itemNumber = itemNumber+1
        Next

    end function

    function startWrittingInConsole()
        isWrittingToLogFile = true
    end function

    function stopWrittingInConsole()
        isWrittingToLogFile = false
    end function


    '----------------EVENTS----------------
    Private Sub Class_Initialize(  )
        Dim WshShell
        Set WshShell = CreateObject("WScript.Shell")
         logPath = WshShell.CurrentDirectory&"\logs"
        stopWrittingInConsole()
    End Sub
    
    'When Object is Set to Nothing
    Private Sub Class_Terminate(  )
        
    End Sub

end class

'---------TESTING-----------
function testDebugUtility()

    Dim oDebug
    set oDebug = new DebugUtility

    WScript.Echo "[TESTING] Logging into the log.txt"
    oDebug.printSimpleString("This should be inside a log.txt file")

    Dim TableToPrint(3) 
    TableToPrint(1)="coucou"
    TableToPrint(2)="raté"
    TableToPrint(3)="réussite"
    oDebug.printTable(TableToPrint)

    WScript.Echo "[TESTING] Logging into the console"
    oDebug.startWrittingInConsole()
    oDebug.printSimpleString("This shoulb be in a console log window")
    oDebug.printTable(TableToPrint)
end function

Private Sub Include( scriptName )
    WScript.Echo "[LOADING]: "&scriptName
    Dim sScript
    Dim oStream
    With CreateObject( "Scripting.FileSystemobject" )
        set oStream = .OpenTextFile(scriptName)
    End With
    sScript = oStream.ReadAll()
    oStream.Close
    ExecuteGlobal sScript
    WScript.Echo "[LOADED]: "&scriptName

    Call ErrorHandler
End Sub

Private Function ErrorHandler()
    If Err.Number <> 0 Then
        ' MsgBox or whatever. You may want to display or log your error there
        Call Err.Raise(vbObjectError + 10, "Debug Utility", "Unknow error with description: "&Err.Description)
        WScript.Echo("[ERROR] {Debug Utility} Unknow error with description: "&Err.Description)
        Err.Clear
    End If
end function