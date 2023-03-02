Option Explicit
'On Error Goto 0
'CLASS OBJECT : https://www.tutorialspoint.com/vbscript/vbscript_class_objects.htm
call test()


Class WorkingWithDatesAndTime

    Public function convertDateToValue(DateAsString)
        convertDateToValue = CDate(DateAsString)
    end function

    '----------------EVENTS----------------
    Private Sub Class_Initialize(  )
        'Getting all processRunning at object creation
    End Sub
    
    'When Object is Set to Nothing
    Private Sub Class_Terminate(  )
        'Termination code goes here
    End Sub

end class

'---------TESTING-----------
function test()
    Dim a,b,diff

    a = Now()
    b = Timer()
    wscript.echo "Is ""a"" a date : " & IsDate(a) _ 
                & " with the value : " & a _ 
                & "convertedTo : " & Cdate(a)
    wscript.echo "Is ""b"" a date : " & IsDate(b) _ 
                & " with the value : " & b _ 
                & "convertedTo : " & Cdate(b)
     diff = Cdate(a) - Cdate(b)
     wscript.echo "Is ""b"" a date : " & IsDate(diff) _ 
                & " with the value : " & diff _ 
                & "convertedTo : " & Cdate(diff)


end function
