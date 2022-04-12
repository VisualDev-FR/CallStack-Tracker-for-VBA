Attribute VB_Name = "CallStackTracker"
Option Explicit

Public Const histEND As String = "End"      'Constant specifying the normal exit of a sub / function to the CallStack Tracker
Public Const histEXIT As String = "Exit"    'Constant specifying the anormal exit of a sub / function to the CallStack Tracker
Public Const histERROR As String = "Error"  'Constant specifying than an error occured in a sub / function to the CallStack Tracker

Public histo_IgnorePrint As Boolean         'Allows to fully disable the CallStack Tracker

Private strFile_Headers As String           'Full path of the file Headers.html
Private strFile_Style As String             'Full path of the file Darkstyle.css
Private strFile_Script As String            'Full path of the file arbo.js

Private strDir_CrashReport As String        'Full path of the directory wich will contain the reports with at least one crash
Private strDir_StackReport As String        'Full path of the directory wich will contain all the reports (be careful to the siez of the generated reports)

Private histoIndex As Long                  'Allows to assign one single ID to all html lists, in order to apply the CSS tree style
Private mHistFile As String                 'Full path of the report wich is being written
Private isRunning As Boolean                'Allows to know if a root sub / function is launched
Private histoLevel As Long                  'Allows to know the current level indentation of the callStack

Private dicoIndex As Dictionary             'Dictionnary containing the current parent of a specified indentation level. Each parent is dynamicly erased when a same indentation level is reached
Private dicoError As Dictionary             'Dictionnary which contains all ocurred errors during the runtime. Allow to make an error report at the end of the html report printing

Private mCurMod As String                   'Current module kept in memory
Private mCurSub As String                   'Current sub / function kept in memory
Private mCurParent As String                'Current parent kept in memory (module + sub where was launched the last sub / function)

Private mLastMod As String                  'Last module kept in memory
Private mLastSub As String                  'Last sub / function kept in memory

Private histoChrono As Double               'Allows to follow the elapsed time since the root sub launched
Private mStream As TextStream               'Allows to write into the html file, a printing is running


Private Function InitFileSystem() As Boolean
    '---------------------------------------------------
    'LAST MODIF : 22/03/2022
    'AUTHOR : VisualDev-FR
    'PURPOSE : Initialize all directories / files path variables (must be adapted to the user fileSystem)
    'RETURN (Boolean) : True if eveything is OK / False on error
    '---------------------------------------------------

    On Error GoTo errHandle

    Dim thisPath As String: thisPath = ThisWorkbook.Path & "\HTML\"
    
    strDir_CrashReport = ThisWorkbook.Path & "\CRASH\"
    strDir_StackReport = ThisWorkbook.Path & "\REPORT\"

    strFile_Headers = thisPath & "HistoHeaders.html"
    strFile_Style = thisPath & "DarkStyle.css"
    strFile_Script = thisPath & "arbo.js"

    InitFileSystem = True
    Exit Function

errHandle:
    InitFileSystem = False
    histo_IgnorePrint = True
End Function

'-----------------------------------------------------------------------------------------------------------------
'PUBLIC FUNCTIONS
'-----------------------------------------------------------------------------------------------------------------

Public Sub XX_PRINT_HISTO_XX(module_ As String, Optional procedure_ As String):
    '---------------------------------------------------
    'LAST MODIF : 22/03/2022
    'AUTHOR : VisualDev-FR
    'PURPOSE : Main Sub, to call in all sub or functions where you want to track the CallStack. Must be closed by one of the 3 strings below
    'PARAMETERS
    '    * String (procedure_): Name of the current sub / function you want to trace. (Nothing in case of exit)
    '    * String (module_): Name of the current module / of the exit instruction (3 constants allowed) :
    '                       - histEND   : normal exit of a function / sub
    '                       - histExit  : anormal exit from a sub / function, to place before "Exit Sub" or "Exit function"
    '                       - histError : exit from a sub / function after a fail, to insert in exeptions management
    '---------------------------------------------------
    If histo_IgnorePrint Then Exit Sub
    
    If Not isRunning Then
    
        If InitFileSystem = False Then Exit Sub

        mCurMod = module_
        mCurSub = procedure_
        mCurParent = mCurMod & mCurSub
        
        Set dicoIndex = New Dictionary

        Call Histo_OpenFile(module_, procedure_)
        Exit Sub
        
    End If
    
    mLastMod = mCurMod
    mLastSub = mCurSub

    mCurMod = module_
    mCurSub = procedure_
    mCurParent = mCurMod & mCurSub

    Select Case mLastMod
    
        Case histEND, histEXIT, histERROR
        
            Select Case mCurParent
            
                Case histEND
                
                    Call Histo_CloseList

                    If histoLevel <= 0 Then Call Histo_CloseFile
                     
                Case histEXIT
                
                    Call Histo_CloseList
                    Call Histo_WriteExit(dicoIndex(histoLevel)(0), dicoIndex(histoLevel)(1))

                    If histoLevel <= 0 Then Call Histo_CloseFile
                     
                Case histERROR
                
                    Call Histo_CloseList
                    Call Histo_WriteError(dicoIndex(histoLevel)(0), dicoIndex(histoLevel)(1))

                    If histoLevel <= 0 Then Call Histo_CloseFile
                    
                Case Else 'Do nothing

                End Select

        Case Else
        
            Select Case mCurParent
            
                Case histEND
            
                    Call Histo_WriteLine(mLastMod, mLastSub)
                    
                    If histoLevel <= 0 Then Call Histo_CloseFile
                    
                Case histEXIT
            
                    Call Histo_WriteLine(mLastMod, mLastSub)
                    Call Histo_WriteExit(mLastMod, mLastSub)

                    If histoLevel <= 0 Then Call Histo_CloseFile
                    
                Case histERROR
            
                    Call Histo_WriteLine(mLastMod, mLastSub)
                    Call Histo_WriteError(mLastMod, mLastSub)

                    If histoLevel <= 0 Then Call Histo_CloseFile

                Case Else

                    Call Histo_OpenList(mLastMod, mLastSub)
                    
            End Select

    End Select
    
End Sub

Public Sub FatalError(Optional displayMsgbox As Boolean = True, Optional sendReport As Boolean = True)
    '---------------------------------------------------
    'LAST MODIF : 22/03/2022
    'AUTHOR : VisualDev-FR
    'PURPOSE : Ending all running programs and send the error report to the directory specified by the constant 'strDir_CrashReport'
    'PARAMETERS
    '    * Boolean (displayMsgbox): Allows to show or not a msgbox to signal the crash to the user
    '    * Boolean (sendReport): Allows to enable or not the sending of the crash report to the specified directory
    '---------------------------------------------------
    If histo_IgnorePrint Then Exit Sub

    Call XX_PRINT_HISTO_XX(histERROR)
    Call Histo_CloseFile

    Dim fso As New FileSystemObject, crashFile As File
    
    Set crashFile = fso.GetFile(mHistFile)
    
    If sendReport Then crashFile.Copy (strDir_CrashReport)
    If displayMsgbox Then MsgBox "Fatal Error", vbOKOnly Or vbCritical, "Crash"

    End

End Sub

'-----------------------------------------------------------------------------------------------------------------
'INTERNAL FUNCTIONS
'-----------------------------------------------------------------------------------------------------------------
Private Sub Histo_WriteError(mModule_, mProcedure_)
    '---------------------------------------------------
    'LAST MODIF : 12/04/2022
    'AUTHOR : VisualDev-FR
    'PURPOSE : Write the location and the description of the error in a special html tag 'error' in order to apply a specific style
    'PARAMETERS
    '    *  (mModule_): module where the error occured
    '    *  (mProcedure_): sub or function where the error noccured
    '---------------------------------------------------
    Dim strKey As String: strKey = mModule_ & " / " & mProcedure_

    mStream.WriteLine "<error>" & strKey & " : " & err.Description & "</error>"

    dicoError.Add key:=dicoError.Count + 1 & ". " & strKey, Item:=err.Description

    err.Clear

End Sub

Private Sub Histo_WriteExit(mModule_, mProcedure_)
    '---------------------------------------------------
    'LAST MODIF : 12/04/2022
    'AUTHOR : VisualDev-FR
    'PURPOSE : Write the location of an anormal exit from a sub or function, in a special html tag 'exit' in order to apply a specific style
    'PARAMETERS
    '    *  (mModule_): module where the exit was forced
    '    *  (mProcedure_): sub or function where the exit was forced
    '---------------------------------------------------

    Dim strKey As String: strKey = mModule_ & " / " & mProcedure_

    mStream.WriteLine "<exit>" & strKey & " : Exit </exit>"
    
    dicoError.Add key:=dicoError.Count + 1 & ". " & strKey, Item:="Exit"

End Sub

Private Sub Histo_OpenFile(mainModule As String, mainSub As String)
    '---------------------------------------------------
    'LAST MODIF : 12/04/2022
    'AUTHOR : VisualDev-FR
    'PURPOSE : Create the html file, open the private textStream and insert the headers in the report
    'PARAMETERS
    '    * String (mainModule): Name of the root module where is launched the root sub
    '    * String (mainSub): Name of the root sub
    '---------------------------------------------------

    histoChrono = Timer

    mHistFile = strDir_StackReport & Format(Now, "yyyy-mm-dd_hhmmss_") & mainModule & "_" & mainSub & ".html"
    
    Set dicoError = New Dictionary
    
    Dim fso As New FileSystemObject
    Set mStream = fso.OpenTextFile(mHistFile, ForWriting, True)
    
    Dim strHeader As String
    strHeader = Histo_ParseHeaders( _
        mainModule_:=mainModule, _
        mainSub_:=mainSub, _
        Title_:=Dir(mHistFile))
        
    With mStream
    
        .WriteLine strHeader
        .WriteBlankLines (1)
        .WriteLine "<div id = ""historique"">"
        .WriteLine "    <ul>"

    End With
    
    isRunning = True

End Sub

Private Sub Histo_CloseFile()
    '---------------------------------------------------
    'LAST MODIF : 12/04/2022
    'AUTHOR : VisualDev-FR
    'PURPOSE : Insert the finals data into the headers (execution time + error report), close all html lists and close the public textStream
    ' * Might be improved by finding a cleviest way to modify the html file headers, once the entire callStack was written
    '---------------------------------------------------
    
    Dim start As Double, i
    Dim fso As New FileSystemObject, strFile As String
    
    start = Timer

    mStream.Close

    Set mStream = fso.OpenTextFile(mHistFile, ForReading, False)

    strFile = mStream.ReadAll
    strFile = Replace(strFile, "<cfg_Chrono>", Histo_GetChrono)
    strFile = Replace(strFile, "<cfg_ErrReport>", Histo_ParseErrors)

    mStream.Close

    Set mStream = fso.OpenTextFile(mHistFile, ForWriting, False)
    
    With mStream
    
        .WriteLine strFile

        For i = histoLevel To 0 Step -1
        
            .WriteLine "</ul>"
            
        Next i

        .WriteLine "</div>"
        
    End With

    mStream.Close

    histoChrono = 0
    isRunning = False
    histoLevel = 0
    histoIndex = 0
    
    Set dicoIndex = Nothing

    If dicoError.Count > 0 Then fso.GetFile(mHistFile).Copy Destination:=strDir_CrashReport

End Sub

Private Function Histo_ParseErrors() As String
    '---------------------------------------------------
    'LAST MODIF : 12/04/2022
    'AUTHOR : VisualDev-FR
    'PURPOSE : Read the public error Dictionnary, and concatenate all items in one html list
    'RETURN (String) : A string containing the html concatened list of errors
    '---------------------------------------------------
    
    Dim strTemp As String, k
    
    For Each k In dicoError.Keys
    
        strTemp = strTemp & "<li>" & k & " : " & dicoError(k) & "</li>" & vbCrLf
    
    Next k
    
    Histo_ParseErrors = IIf(strTemp <> "", strTemp, "N/A")

End Function

Private Function Histo_ParseHeaders( _
    Optional mainModule_ As String, _
    Optional mainSub_ As String, _
    Optional Title_ As String) As String
    '---------------------------------------------------
    'LAST MODIF : 12/04/2022
    'AUTHOR : VisualDev-FR
    'PURPOSE : Parse the original html header file, by adding all spcifics data + the CSS style ans the JS script
    'RETURN (String) : A string containing all the initial html header
    'PARAMETERS :
    '    * String (mainModule_): Root module of the callStack
    '    * String (mainSub_): Root sub or function of the callStack
    '    * String (Title_): Title of the html page
    '---------------------------------------------------
                      
    Dim strTemp As String
    
    Dim fso As New FileSystemObject
    Dim headerStream As TextStream, styleStream As TextStream, scriptStream As TextStream
    
    Set headerStream = fso.OpenTextFile(strFile_Headers, ForReading, False)
    Set styleStream = fso.OpenTextFile(strFile_Style, ForReading, False)
    Set scriptStream = fso.OpenTextFile(strFile_Script, ForReading, False)

    strTemp = headerStream.ReadAll

    strTemp = Replace(strTemp, "<cfg_MainModule>", mainModule_)
    strTemp = Replace(strTemp, "<cfg_MainSub>", mainSub_)
    strTemp = Replace(strTemp, "<cfg_Title>", Title_)

    strTemp = Replace(strTemp, "<cfg_User>", Environ("username"))
    strTemp = Replace(strTemp, "<cfg_Date>", Format(Now, "dd/mm/yyyy"))
    strTemp = Replace(strTemp, "<cfg_Hour>", Format(Now, "hh:mm:ss"))
    
    strTemp = Replace(strTemp, "<script></script>", "<script>" & scriptStream.ReadAll & "</script>")
    strTemp = Replace(strTemp, "<style></style>", "<style>" & styleStream.ReadAll & "</style>")

    Histo_ParseHeaders = strTemp

    headerStream.Close
    scriptStream.Close
    styleStream.Close
    
    Set headerStream = Nothing
    Set fso = Nothing
    
End Function

Private Function Histo_GetLineStr(module_ As String, procedure_ As String) As String
    '---------------------------------------------------
    'LAST MODIF : 12/04/2022
    'AUTHOR : VisualDev-FR
    'PURPOSE : Write a line to send into the report, according to the current module / sub + write the elapsed time since the running of the root procedure
    'RETURN (String) : A string containging all values between html tags, to insert between two <li> tags
    'PARAMETERS :
    '    * String (module_): current module
    '    * String (procedure_): current sub / function
    '---------------------------------------------------

    Dim strModule As String, strSub As String, strchrono As String

    strModule = "<Module>" & module_ & "</Module>"
    strSub = "<Procedure>" & procedure_ & "</Procedure>"
    strchrono = "<Chrono>" & Histo_GetChrono() & "</Chrono>"

    Histo_GetLineStr = strModule & " : " & strSub & "  " & strchrono


End Function

Private Sub Histo_WriteLine(module_ As String, procedure_ As String)
    '---------------------------------------------------
    'LAST MODIF : 12/04/2022
    'AUTHOR : VisualDev-FR
    'PURPOSE : Write a simple line into the public stream
    'PARAMETERS
    '    * String (module_): current module
    '    * String (procedure_): current sub / function
    '---------------------------------------------------
    
    With mStream
    
        .WriteLine "<li><span>" & Histo_GetLineStr(module_, procedure_) & "</span></li>"

    End With

End Sub

Private Sub Histo_OpenList(mModule_ As String, mProcedure_ As String)
    '---------------------------------------------------
    'LAST MODIF : 12/04/2022
    'AUTHOR : VisualDev-FR
    'PURPOSE : Open a list in the public textStream with html format allowing to fold / unfold the lines below with CSS code + increase the current indentation level of the callStack
    'PARAMETERS
    '    * String (mModule_): current module
    '    * String (mProcedure_): current sub / function
    '---------------------------------------------------

    Dim strInput As String
    
    strInput = "<li><input type = ""checkbox"" id=""" & histoIndex & """><label for=""" & histoIndex & """>" & Histo_GetLineStr(mModule_, mProcedure_) & "</label>"

    With mStream
        
        .WriteLine strInput
        .WriteLine "<ul>"

    End With
    
    dicoIndex(histoLevel) = Array(mModule_, mProcedure_)
    
    histoIndex = histoIndex + 1
    histoLevel = histoLevel + 2

End Sub

Private Sub Histo_CloseList()
    '---------------------------------------------------
    'LAST MODIF : 12/04/2022
    'AUTHOR : VisualDev-FR
    'PURPOSE : Close a html list in the public textStream + reduce the current indentation level of the callStack
    '---------------------------------------------------

    histoLevel = histoLevel - 2

    With mStream

        .WriteLine "</ul>"
        .WriteLine "</li>"
            
    End With

End Sub

Private Function Histo_GetChrono() As String
    '---------------------------------------------------
    'LAST MODIF : 12/04/2022
    'AUTHOR : VisualDev-FR
    'PURPOSE : Gives the elapsed time since the launching of the root sub/function + change the unity according to the elapsed time
    'RETURN (String) : A formatted string, containg the elapsed time + the time unity
    '---------------------------------------------------

    Dim exeChrono As Double
    exeChrono = (Timer - histoChrono)

    Select Case exeChrono
    
        Case Is < 1
        
            Histo_GetChrono = Format(exeChrono * 1000, "0.000 ms")
            
        Case Else
        
            Histo_GetChrono = Format(exeChrono, "0.000 s")
        
    End Select

End Function


















