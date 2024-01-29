# VBsLib
This is a central storage for commun Lib files or function or method i can use


1. Debug Utility.vbs
  The main goal is to create a class able to handle debugging trough a file or in a console.
  This will basically create a way to swap easily between the two using "startWrittingInConsole" and "stopWrittingInConsole". By default it is using the current directory&\logs.
  
2. File Manipulation.vbs
  This will help to manipulate files trough a class object.
   - getUserFolder : get the "user" folder
   - moveFileTo : will move a file specified in _fileToMove_[STRING] into _DestinationAsPath_[STRING]
   - renameFileTo : will rename a _selectedFile_[STRING] with _newName_[STRING]
   - FileFinder : will return a list of file found in a folder, can be recursive or not using _isRecursif_[BOOL] with a limit _recursifLimit_[INTEGER] into a specified path _SelectedFolder_[STRING]
   - getPathWhereScriptIsRun: will return the current directory where the script is run
   - getFolderWhereScriptIsRun : will return the current folder where the script is run
   - getFolderFromPath : will return the folders path part of a file path string using _pathAsString_[STRING]
   - getFileFromPath : will extract the file name from a full path using _pathAsString_[STRING]
   - FileExists : return if a specified file exist in the _fileToTest_[STRING] exist as [BOOL]
     
3. MultiTreading.vbs
   - PlaceHolder is i need to use multitreading; should contain a class to create, handle and return multithread script
  
4. Notes how to do lib file
   - PlaceHolder for how to create proper lib file using a "COM File"
   - also contain a teplace for "include" function to be added in other vbs file if needed
     
5. PrintFile.vs
   - This was used to convert file to pdf; but in it has been disabled
     
6. ProcessManipulation.vbs
   This will handle process to look for a specific program running and kill it if needed
   - isProcessRunning: will look for a specific process name
   - updateListProcessRunning: will get the process running using a query
   - findWindowTitle : will return a [BOOL] of a window already opening in windows
   - matchTitle: will return a [BOOL] after trying to match the name of all windows open using a _input_[REGEX] code as limiting
   - killProcess: will kill the process based on the _myprocess_[STRING] name
     
7.  RegexComplement.vbs
   This is a tool to use Regex in VBS
   - RegexReplace: will replace all
   - RegexReplaceFirst: will replace the first match
   - StringReverse: will reverse a string
   - RegexFindFirst: will return the first match found
     
8. WorkingWithDatesAndTime.vbs
   - This is to capitalize the conversion of number to date
  
