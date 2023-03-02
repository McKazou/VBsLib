Option Explicit
'CLASS OBJECT : https://www.tutorialspoint.com/vbscript/vbscript_class_objects.htm
'call test()

Class PrintFileObject

    'Class imports for shell stuff (invoke function)

    'Key: Printers name, Value: PORT
    Public Property Get Printers
        Dim WshNetwork, oPrinters
        Set WshNetwork = WScript.CreateObject("WScript.Network")
        Set oPrinters = WshNetwork.EnumPrinterConnections

        'Imprimante disponible :
        Dim i
        Dim printersList 
        Set printersList = CreateObject("Scripting.Dictionary")
        For i = 0 to oPrinters.Count - 1 Step 2
            'WScript.Echo "Port " & oPrinters.Item(i) & " = " & oPrinters.Item(i+1)
            printersList.Add oPrinters.Item(i+1), oPrinters.Item(i)
        Next
        set Printers = printersList 

        'Mettre une boite de dialog pour selectionner l'imprimante

        'Letting object go
        set WshNetwork = Nothing
        set oPrinters = Nothing
    End Property
    

    Public function PrintFile(ByVal file)
        'USING ACRORD FOR PDF:
        'Dim oWsh
        'set oWsh = CreateObject ("Wscript.Shell")
        'oWsh.run """AcroRd32.exe"" /t /n /h " &file,,true

        'GETTING TE PATH TO THE FILE :
        'Dim splitPath 
        'splitPath = Split(file.path,"\")
        'Dim folderPath
        'folderPath = splitPath(UBound(splitPath))
        'folderPath = Replace(file.path,folderPath,"")
        'msgbox(folderPath)

        'INVOKE THE RIGHT CLIC PRINT BUTTON
        'Dim objShell, objFolder, objFile
        'Set objShell = CreateObject("Shell.Application")
        'Set objFolder = objShell.NameSpace(folderPath)
        'Set objFile = objFolder.ParseName(file)

        'LOL
        WScript.Echo "[INFORMATION]: {PrintFileObject} - PrintFile : Printing File - " & file.Path
        
        
'----------PRINT USING DEFAULT PRINTER-------
'Printing one file : https://community.spiceworks.com/topic/823781-print-file-from-network-drive-with-vbs-or-bat-script
'Printing multiple files : https://stackoverflow.com/questions/9013941/how-to-run-batch-file-from-network-share-without-unc-path-are-not-supported-me
'Print pdf using vbs : https://stackoverflow.com/questions/50920097/printing-pdf-files-using-vbs



        Dim objShell, objFolder, objFile
        Set objShell = CreateObject("Shell.Application")
        Dim folderPath
        folderPath = Split(file.Path, "\")
        '   Get the folder path (parent's path)
        Redim Preserve folderPath(UBound(folderPath) - 1)
        folderPath = Join(folderPath,"\")&"\"
        'WScript.Echo "[INFORMATION]: {PrintFileObject} - PrintFile : File's parents folder - " & folderPath
        Set objFolder = objShell.NameSpace(folderPath)
        Set objFile = objFolder.ParseName(file.Name)
        objFile.InvokeVerb("Print")


        

        'It may be possible to change the folder using send key and "tab"
        WScript.Echo "[INFORMATION]: {PrintFileObject} - PrintFile : File Printed - " & file
    end function

    Private previousDefaultPrinter

    'SET DEFAULT PRINTER TO
    'Work with printers : https://ss64.com/vb/setdefaultprinter.html
    Public function setDefaultPrinterTo(printerNames)
        Dim WshNetwork
        Set WshNetwork = CreateObject("WScript.Network")    
        WshNetwork.SetDefaultPrinter printerNames
    end function

    Public function getPrinterAvailable()

    end function

    '----------------EVENTS----------------
    Private Sub Class_Initialize(  )
        'Initalization code goes here

        Dim printersAvailable 
        Set printersAvailable = CreateObject("Scripting.Dictionary")

        'STORE THE DEFAULT PRINTER
        'https://stackoverflow.com/questions/2273458/vbs-get-default-printer
        Dim oShell, strValue, strPrinter
        Set oShell = CreateObject("WScript.Shell")
        strValue = "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Device"
        strPrinter = oShell.RegRead(strValue)
        strPrinter = Split(strPrinter, ",")(0)
        WScript.Echo "[INFORMATION]: {PrintFileObject} - Initialize : Previous default printer - " & strPrinter

        previousDefaultPrinter = strPrinter

        if not strPrinter = "PDFCreator" Then
            WScript.Echo "[INFORMATION]: {PrintFileObject} - Initialize : Setting default printer to - PDFCreator"
            setDefaultPrinterTo("PDFCreator")
        end if
    End Sub
    
    'When Object is Set to Nothing
    Private Sub Class_Terminate(  )
        'Termination code goes here
        setDefaultPrinterTo(previousDefaultPrinter)
    End Sub

    Function printFromFolder()
    Dim shApp
    Dim shFIcol
    Dim shFIx
    Dim shFLDx
    Dim lngX

    Set shApp = New Shell32.Shell
    Set shFLDx = shApp.BrowseForFolder(0, "Select Folder to print from", 0, 0)
    Set shFIcol = shFLDx.Items()

    For Each shFIx In shFIcol
        If Not shFIx.IsFolder Then    ' Print only if is file
            shFIx.InvokeVerb ("&Print")
            DoEvents
        End If
    Next
    End Function

end class

'---------TESTING-----------
function test()
Dim printObject
set printObject = new PrintFileObject

Dim printersAvailable
set printersAvailable = printObject.Printers

dim key
For Each key in printersAvailable
    WScript.Echo "Key: " & key & " = " & printersAvailable(key)
Next

Dim shApp, currentPath, shFolder,files
set shApp = CreateObject("shell.application")
currentPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".") 
set shFolder = shApp.NameSpace( currentPath )
set files = shFolder.Items()

'2. Get all the files in one array in the current folder and any sub folder to
Dim file
for each file in files
    'msgbox(file)
    printObject.PrintFile(file)
next
end function
