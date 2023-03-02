Option Explicit
'CLASS OBJECT : https://www.tutorialspoint.com/vbscript/vbscript_class_objects.htm
'call test()

Class ProcessManipulation

    'Class imports for shell stuff (invoke function)

    'Key: Printers name, Value: PORT
    Public Property Get ProcessAvailable
        set ProcessAvailable = updateListProcessRunning()
    End Property

    
    'https://stackoverflow.com/questions/15261317/vbscript-to-check-if-a-process-is-running-if-its-not-then-copy-a-file-from-netw
    Public function isProcessRunning(ByVal processName)
        Dim processList, process
        Set processList = ProcessAvailable

        For Each process In processList
            if process.Name = processName then
                'WScript.Echo "[INFORMATION]: {ProcessManipulation} - isProcessRunning : Process to find: - " & processName & " - " & process.Name
                isProcessRunning = true
                Exit function
            end if
        Next
        WScript.Echo "[INFORMATION]: {ProcessManipulation} - isProcessRunning : Process to find: - " & processName & " - " & process.Name
        
    end function

    private function updateListProcessRunning()
        'Fecthing all process running
        Dim sComputerName, objWMIService, sQuery
        sComputerName = "."
        Set objWMIService = GetObject("winmgmts:\\" & sComputerName & "\root\cimv2")
        sQuery = "SELECT * FROM Win32_Process"
        set updateListProcessRunning = objWMIService.ExecQuery(sQuery)
    end function

    function findWindowTitle(srchstr)
        Dim filterSrc, strCommand, cmdout
        filterSrc = replace(srchstr, "*", "")
        strCommand = "tasklist /v | find /i """ & filterSrc & """"
        cmdout = CreateObject("Wscript.Shell").Exec("cmd /c """ & strCommand & " 2>&1 """).stdout.readall
        wscript.sleep 500
        findwindowtitle = matchtitle(srchstr, cmdout)
    End Function

    Function matchTitle(srchstr, input)
        Dim filterSrc, filterstrpatt, regex, matches, m
        matchtitle = false
        if instr(1, srchstr, "*", 1) <> 0 Then 
            filterSrc = replace(srchstr, "*", "")
            filterstrpatt = replace(srchstr, "*", "[a-zA-Z0-9\. ]*")
        end if
        Set regex = CreateObject("VBScript.RegExp")
        regex.MultiLine = True
        regex.Global = True
        regex.IgnoreCase = True
        regex.Pattern = "(?:.*)(?:\d\d?\d?:\d\d:\d\d\s)(\b" & filterstrpatt & "\b)"
        Set matches = regex.Execute(input)
        for m = 0 to matches.count - 1
            Set SubMatches = matches.item(m).SubMatches
            for i = 0 to (Submatches.count - 1)
                if instr(1, Submatches.item(i), filterSrc, 1) <> 0 then matchtitle = Submatches.item(i)
            Next
        Next
        if (matchtitle = false) then
            wscript.echo "Could not find process with title matching, '" & srchstr & "'"
            'wscript.quit
        end if
    End Function

    Function killProcess( myProcess )    

        Dim blnRunning, colProcesses, objProcess    
        blnRunning = False
        killProcess = False    
         
        Set colProcesses = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery( "Select * From Win32_Process", , 48 ) 
        For Each objProcess in colProcesses     
            Set colProcesses = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery( "Select * From Win32_Process", , 48 )      
            If LCase( myProcess ) = LCase( objProcess.Name ) Then   
            ' Confirm that the process was actually running             
            'blnRunning = True             
            ' Get exact case for the actual process name            
                myProcess  = objProcess.Name             
                if isProcessRunning(myProcess) then
                    ' Kill all instances of the process            
                    objProcess.Terminate() 
                    killProcess = True
                end if
            End If    
        Next    
        'If blnRunning Then        
            'Do Until Not blnRunning            
                'Set colProcesses = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery( "Select * From Win32_Process Where Name = '"& myProcess & "'" )            
                'WScript.Sleep 100 'Wait for 100 MilliSeconds            
                'If colProcesses.Count = 0 Then 'If no more processes are running, exit loop                
                    'blnRunning = False            
                'End If        
            'Loop      
        'End If
            
    End Function


    '----------------EVENTS----------------
    Private Sub Class_Initialize(  )
        'Getting all processRunning at object creation
        updateListProcessRunning()
    End Sub
    
    'When Object is Set to Nothing
    Private Sub Class_Terminate(  )
        'Termination code goes here
    End Sub

end class

'---------TESTING-----------
function test()
'Dim printObject
'set printObject = new PrintFileObject

'Dim printersAvailable
'set printersAvailable = printObject.Printers


Dim processManip
set processManip = new ProcessManipulation

Dim allProcessRunning, process
set allProcessRunning = processManip.ProcessAvailable

For Each process In allProcessRunning
    WScript.Echo "Process [Name:" & process.Name & "]"
Next

if processManip.isProcessRunning("notepad.exe") Then
    WScript.Echo "Process [Name:" & "notepad.exe" & "] = FOUND "
else
    WScript.Echo "Process [Name:" & "notepad.exe" & "] = NOT FOUND "
end if

'WScript.Echo "Process (PDF Creator)* has been found with the name : " & processManip.findwindowtitle("(PDF Creator)*")

WScript.Echo "Process have been killed successfully : " & processManip.killProcess("AcroRd32.exe")
end function
