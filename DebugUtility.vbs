Option Explicit
'CLASS OBJECT : https://www.tutorialspoint.com/vbscript/vbscript_class_objects.htm
'call test()

Class DebugUtility

    'Class imports for shell stuff (invoke function)

    'Key: Printers name, Value: PORT
    private isDebugging

    Private currentPath

    'Could add also the timming stuff to track the time something ?
    
    Public function print(ByVal stringToPrint)
        'https://stackoverflow.com/questions/17194375/i-need-to-write-vbs-wscript-echo-output-to-text-or-cvs
        'The idea here is to replace Wscript.Echo with Debug.print and be able to differentiate between the app running in debug like VS code ?
        if isDebugging Then
            WScript.Echo stringToPrint
        else
            Dim fso, f
            Set fso = WScript.CreateObject("Scripting.Filesystemobject")
            Set f = fso.CreateTextFile(currentPath&"\log.txt", 2)
            f.WriteLine stringToPrint
            f.Close
        end if
    end function

    function startDebugging()
        isDebugging = true
    end function

    function stopDebugging()
        isDebugging = false
    end function


    '----------------EVENTS----------------
    Private Sub Class_Initialize(  )
        Dim WshShell
        Set WshShell = CreateObject("WScript.Shell")
        currentPath = WshShell.CurrentDirectory
    End Sub
    
    'When Object is Set to Nothing
    Private Sub Class_Terminate(  )
        
    End Sub

end class

'---------TESTING-----------
function test()

    Dim oDebug
    set oDebug = new DebugUtility

    oDebug.print("help it's not working")

    oDebug.startDebugging()

    oDebug.print("This is working ! ")
end function
