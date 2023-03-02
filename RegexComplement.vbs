'TESTING : 
'call classTest()

'--------Regex Complement--------.
'Definition : https://www.tutorialspoint.com/vbscript/vbscript_class_objects.htm
Class RegexComplement


'This regex function will replace all string in another using a pattern. If the pattern isn't found the string will stay the same
Public Function RegexReplace(strInput, strPattern, replacedBy)
    Dim Regex 
    set Regex = new RegExp
    Dim strReplace 
    Dim strOutput 

    If strPattern <> "" Then
        
        With Regex
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        
        If Regex.Test(strInput) Then
            RegexReplace = Regex.Replace(strInput, replacedBy)
        Else
            RegexReplace = strInput
        End If
    End If
    Set Regex = Nothing

    ErrorHandler
End Function

Public Function RegexReplaceFirst(strInput, strPattern, replacedBy)
    Dim Regex 
    set Regex = new RegExp
    Dim strReplace 
    Dim strOutput 

    If strPattern <> "" Then
        
        With Regex
            .Global = False
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        
        If Regex.Test(strInput) Then
            RegexReplaceFirst = Regex.Replace(strInput, replacedBy)
        Else
            RegexReplaceFirst = strInput
        End If
    End If
    Set Regex = Nothing

    ErrorHandler
End Function

Public Function StringReverse(s)
    StringReverse = StrReverse(s)
End Function

Public Function RegexFindFirst(strInput, strPattern)
    Dim Regex 
    set Regex = new RegExp
    Dim strReplace 
    Dim strOutput 
    
    If strPattern <> "" Then
        
        With Regex
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        
        If Regex.Test(strInput) Then
            RegexFindFirst = Regex.Execute(strInput)(0)
            
        Else
            'if we dont find the pattern return ""
            RegexFindFirst = null
        End If
    End If
    ErrorHandler
End Function

    '--------------FUNCTION TO HANDLE ERRORS------------
    Sub ErrorHandler()
    
        If Err.Number <> 0 Then
        ' MsgBox or whatever. You may want to display or log your error there
            Call Err.Raise(vbObjectError + 10, "Start Converting", "Unknow error with description: "&Err.Description)
            WScript.Echo("[ERROR] Error in Errohandler.vbs")
            Err.Clear
        End If

    end sub

End Class

Function classTest()
    Dim regComp
    set regComp = new RegexComplement

    Dim result
     result = regComp.RegexReplace("blabla","(bla)","bui")
    WScript.Echo "Test remplace function :" & result

    Dim Indice
    'https://stackoverflow.com/questions/10768924/match-sequence-using-regex-after-a-specified-character
    Indice = regComp.RegexFindFirst("000100026A","[a-zA-Z]+")
    WScript.Echo "Test Find first function :" & Indice

    dim partNumberWithSpaces 
    partNumberWithSpaces = regComp.RegexReplace("00100026","((\d){3})","$1")
    WScript.Echo "Test Replace with occurence function :" & partNumberWithSpaces

    indice = regComp.RegexFindFirst("00100026 A","[a-zA-Z]+")
    WScript.Echo "Test Find first function :" & indice
end function