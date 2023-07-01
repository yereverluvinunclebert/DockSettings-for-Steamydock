Attribute VB_Name = "common3"
' .01 DAEB 31/01/2021 common3.bas Added new checkbox to determine if a post initiation dialog should appear
' .02 DAEB 20/05/2021 common.bas Added new check box to allow a quick launch of the chosen app

Option Explicit

'------------------------------------------------------------
' common3.bas
'
' Public procedures that appear in just two of the programs as an included module common3.bas,
' specifically, enhanced icon settings and steamy dock itself.
'
' Note: If you make a change here it affects the two programs dynamically
'------------------------------------------------------------



 Private Type STARTUPINFO
      cb As Long
      lpReserved As String
      lpDesktop As String
      lpTitle As String
      dwX As Long
      dwY As Long
      dwXSize As Long
      dwYSize As Long
      dwXCountChars As Long
      dwYCountChars As Long
      dwFillAttribute As Long
      dwFlags As Long
      wShowWindow As Integer
      cbReserved2 As Integer
      lpReserved2 As Long
      hStdInput As Long
      hStdOutput As Long
      hStdError As Long
   End Type

   Private Type PROCESS_INFORMATION
      hProcess As Long
      hThread As Long
      dwProcessId As Long
      dwThreadId As Long
   End Type

   Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
   Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
   Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
   Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

   Private Const NORMAL_PRIORITY_CLASS = &H20&
   Private Const INFINITE = -1&
   Private Const SW_HIDE = 0
   Private Const SW_SHOWMINNOACTIVE = 7

'---------------------------------------------------------------------------------------
' Procedure : readIconSettingsIni
' Author    : beededea
' Date      : 21/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub readIconSettingsIni(location As String, ByVal iconNumberToRead As Integer, settingsFile As String)
    'Reads an .INI File (SETTINGS.INI)
        
   On Error GoTo readIconSettingsIni_Error
   'If debugflg = 1 Then debugLog "%readIconSettingsIni"
                                '"Software\RocketDock\Icons"
        sFilename = GetINISetting(location, iconNumberToRead & "-FileName", settingsFile)
        sFileName2 = GetINISetting(location, iconNumberToRead & "-FileName2", settingsFile)
        sTitle = GetINISetting(location, iconNumberToRead & "-Title", settingsFile)
        sCommand = GetINISetting(location, iconNumberToRead & "-Command", settingsFile)
        sArguments = GetINISetting(location, iconNumberToRead & "-Arguments", settingsFile)
        sWorkingDirectory = GetINISetting(location, iconNumberToRead & "-WorkingDirectory", settingsFile)
        sShowCmd = GetINISetting(location, iconNumberToRead & "-ShowCmd", settingsFile)
        sOpenRunning = GetINISetting(location, iconNumberToRead & "-OpenRunning", settingsFile)
        sIsSeparator = GetINISetting(location, iconNumberToRead & "-IsSeparator", settingsFile)
        sUseContext = GetINISetting(location, iconNumberToRead & "-UseContext", settingsFile)
        sDockletFile = GetINISetting(location, iconNumberToRead & "-DockletFile", settingsFile)
        
        'If defaultDock = 1 Then
        sUseDialog = GetINISetting(location, iconNumberToRead & "-UseDialog", settingsFile)
        sUseDialogAfter = GetINISetting(location, iconNumberToRead & "-UseDialogAfter", settingsFile)  ' .01 DAEB 31/01/2021 rdIconConfig.frm Added new checkbox to determine if a post initiation dialog should appear
        sQuickLaunch = GetINISetting(location, iconNumberToRead & "-QuickLaunch", settingsFile) ' .02 DAEB 20/05/2021 common.bas Added new check box to allow a quick launch of the chosen app
        sAutoHideDock = GetINISetting(location, iconNumberToRead & "-AutoHideDock", settingsFile)       ' .12 DAEB 20/05/2021 common3.bas Added new check box to allow autohide of the dock after launch of the chosen app
        sSecondApp = GetINISetting(location, iconNumberToRead & "-SecondApp", settingsFile)      ' .11 DAEB 21/05/2021 common.bas Added new field for second program to be run
        
   On Error GoTo 0
   Exit Sub

readIconSettingsIni_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readIconSettingsIni of Module Module2"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readIconRegistryWriteSettings
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Read the registry one line at a time and create a temporary settings file
'---------------------------------------------------------------------------------------
'
Public Sub readIconRegistryWriteSettings(settingsFile As String)
    Dim useloop As Integer
    
    On Error GoTo readIconRegistryWriteSettings_Error
    'If debugflg = 1 Then debugLog "%" & "readIconRegistryWriteSettings"
 
    For useloop = 0 To rdIconMaximum
         ' get the relevant entries from the registry
         readRegistryIconValues (useloop)
         ' write the rocketdock alternative settings.ini
         Call writeIconSettingsIni("Software\RocketDock\Icons", useloop, settingsFile) ' the alternative settings.ini exists when RD is set to use it
     Next useloop

   On Error GoTo 0
   Exit Sub

readIconRegistryWriteSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readIconRegistryWriteSettings of Module common3"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : writeRegistryOnce
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub writeRegistryOnce(ByVal iconNumberToWrite As Integer)
        
   On Error GoTo writeRegistryOnce_Error
    'If debugflg = 1 Then debugLog "%" & "writeRegistryOnce"
    
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-FileName", sFilename)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-FileName2", sFileName2)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-Title", sTitle)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-Command", sCommand)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-Arguments", sArguments)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-WorkingDirectory", sWorkingDirectory)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-ShowCmd", sShowCmd)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-OpenRunning", sOpenRunning)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-IsSeparator", sIsSeparator)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-UseContext", sUseContext)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-DockletFile", sDockletFile)


    'If defaultDock = 1 Then
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-UseDialog", sUseDialog)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-UseDialog", sUseDialogAfter) ' .01 DAEB 31/01/2021 common3.bas Added new checkbox to determine if a post initiation dialog should appear
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-QuickLaunch", sQuickLaunch) ' .02 DAEB 20/05/2021 common.bas Added new check box to allow a quick launch of the chosen app
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-AutoHideDock", sAutoHideDock) ' .12 DAEB 20/05/2021 common3.bas Added new check box to allow autohide of the dock after launch of the chosen app
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-SecondApp", sSecondApp)  ' .11 DAEB 21/05/2021 common.bas Added new field for second program to be run
      
   On Error GoTo 0
   Exit Sub

writeRegistryOnce_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeRegistryOnce of Module common3"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : removeSettingsIni
' Author    : beededea
' Date      : 21/09/2019
' Purpose   : 'effectively removes data from the ini file at the given location by writing nulls to each value
'---------------------------------------------------------------------------------------
'
Public Sub removeSettingsIni(ByVal iconNumberToWrite As Integer)
       
   On Error GoTo removeSettingsIni_Error
   'If debugflg = 1 Then debugLog "%removeSettingsIni"

        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-FileName", vbNullString, rdSettingsFile
        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-FileName2", vbNullString, rdSettingsFile
        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Title", vbNullString, rdSettingsFile
        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Command", vbNullString, rdSettingsFile
        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Arguments", vbNullString, rdSettingsFile
        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-WorkingDirectory", vbNullString, rdSettingsFile
        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-ShowCmd", vbNullString, rdSettingsFile
        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-OpenRunning", vbNullString, rdSettingsFile
        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-IsSeparator", vbNullString, rdSettingsFile
        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-UseContext", vbNullString, rdSettingsFile
        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-DockletFile", vbNullString, rdSettingsFile

        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-UseDialog", vbNullString, rdSettingsFile
        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-UseDialogAfter", vbNullString, rdSettingsFile
        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-QuickLaunch", vbNullString, rdSettingsFile
        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-AutoHideDock", vbNullString, rdSettingsFile
        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-SecondApp", vbNullString, rdSettingsFile
                
   On Error GoTo 0
   Exit Sub

removeSettingsIni_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure removeSettingsIni of Module common3"
    
End Sub



'   Public Function ExecCmd(cmdline As String, workdir As String) As Integer
'      Dim proc As PROCESS_INFORMATION
'      Dim start As STARTUPINFO
'      Dim ret As Long
'
'        ChDrive Left(workdir, 1) & ":"
'        ChDir workdir
'
'        start.cb = Len(start)
'        start.wShowWindow = SW_SHOWMINNOACTIVE
'
'        Call CreateProcessA(0&, cmdline, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
'        Call WaitForSingleObject(proc.hProcess, INFINITE)
'        Call GetExitCodeProcess(proc.hProcess, ret)
'        Call CloseHandle(proc.hThread)
'        Call CloseHandle(proc.hProcess)
'        ExecCmd = ret
'   End Function



' .54 DAEB 19/04/2021 frmMain.frm Added new function to identify an icon to assign to the entry
'---------------------------------------------------------------------------------------
' Procedure : identifyAppIcons
' Author    : beededea
' Date      : 19/04/2021
' Purpose   : identify an icon to assign to the entry
'---------------------------------------------------------------------------------------
'
Public Function identifyAppIcons(iconCommand As String) As String
    Dim iconFileName As String
    
    Dim identFileName As String
    Dim sDataLine As String
    Dim strDelimiter As String
    
    Dim appName As String
     Dim appIdent1 As String
    Dim appIdent2  As String
    Dim appIcon  As String
    Dim appIdent1Bool As Boolean
    Dim appIdent2Bool As Boolean
    
    Dim fileH As Long
    
   On Error GoTo identifyAppIcons_Error

    'initialise the vars
    
    iconFileName = ""
    identFileName = ""
    sDataLine = ""
    strDelimiter = ""
    appName = ""
    appIdent1 = ""
    appIdent2 = ""
    appIcon = ""
    appIdent1Bool = False
    appIdent2Bool = False
    fileH = 0
    
    identFileName = sdAppPath & "\appIdent.csv"

    
    strDelimiter = ","
    If FExists(identFileName) Then
      fileH = FreeFile() 'get the next free file handle
      ' open the identFileName file
      Open identFileName For Input As #fileH
      ' loop through the identFileName file
      Do While Not EOF(fileH)
          appIdent1Bool = False
          appIdent2Bool = False
         
          ' extract the line from the appIdent file
          Input #fileH, appName, appIdent1, appIdent2, appIcon ' read the four values
          ' set the first two factors to a unlikely value to avoid matching on a ""
          If appIdent1 = "" Then appIdent1 = "XXXXXXXXXXXXXXXXXXXXXX"
          If appIdent2 = "" Then appIdent2 = "XXXXXXXXXXXXXXXXXXXXXX"
          
          ' search for these two factors in the icon command
          If InStr(LCase(iconCommand), LCase(appIdent1)) > 0 Then
              appIdent1Bool = True
          End If
          If InStr(LCase(iconCommand), LCase(appIdent2)) > 0 Then
              appIdent2Bool = True
          End If
          ' if there is a match then read the icon location
          If appIdent1Bool = True And appIdent2Bool = True Then
              ' set that icon as the iconFileName to use
              
              
                  iconFileName = sdAppPath & "\" & appIcon
 
              
                'iconFileName = App.Path & appIcon
                Exit Do ' now found exit the loop
          End If
      Loop
      Close #fileH
    End If
    
    If Not iconFileName = "" Then
       ' check the icon exists


        If FExists(iconFileName) Then
            identifyAppIcons = iconFileName
        End If
    End If

   On Error GoTo 0
   Exit Function

identifyAppIcons_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure identifyAppIcons of Form dock"

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetShortcutInfo
' Author    : Jacques Lebrun
' Date      : 19/04/2021
' Purpose   : Resolve a shortcut
'---------------------------------------------------------------------------------------
' Not that you must add a reference to the COM library "Microsoft Shell Controls and Automation."
' CREDIT http://www.vb-helper.com/howto_get_shortcut_info.html
Public Function GetShortcutInfo(Path As String, Shortcut As Link) As Boolean

    Dim FileNo As Integer
    Dim LongValue As Long
    Dim IntValue As Integer
    Dim LinkFlags As Long
    Dim NextPtr As Long
    Dim Ptr(6) As Long
    Dim Idx As Integer
    Dim PtrBasePath As Long
    Dim PtrNetworkVolumeInfo As Long
    Dim PtrFilename As Long
    Dim Str As String
    
    ' Initialise link results
   On Error GoTo GetShortcutInfo_Error

    With Shortcut
        .Filename = ""
        .Description = ""
        .RelPath = ""
        .WorkingDir = ""
        .Arguments = ""
        .CustomIcon = ""
    End With
    For Idx = 0 To 6
        Ptr(Idx) = 0
    Next
    
    ' Open file with .lnk extension
    FileNo = FreeFile
    Str = Path
    If Right(Str, 4) <> ".lnk" Then Str = Str & ".lnk"
    Open Str For Binary Access Read As FileNo
    
    ' First double-word of link file must be 'L'
    Get FileNo, 1, LongValue
    If LongValue = 76 Then
        ' Skip 16 bytes file GUID and get File Flags
        Get FileNo, 21, LinkFlags
        
        ' Read File Attributes
        Get FileNo, , Shortcut.Attributes
        
        ' Check if ID List section is defined (Ignored)
        NextPtr = 77
        If LinkFlags And 1 Then
            ' Position pointer to next block
            Get FileNo, NextPtr, IntValue
            NextPtr = NextPtr + IntValue + 2
        End If
        
        ' Check if Filename section is defined (no longer mandatory as user created links may not have this defined - see bottom)
        If LinkFlags And 2 Then
            Get FileNo, NextPtr + 16, PtrBasePath
            Get FileNo, , PtrNetworkVolumeInfo
            Get FileNo, , PtrFilename
            
            ' Read base path
            If PtrBasePath Then
                Shortcut.Filename = ReadSingleString(FileNo, NextPtr + PtrBasePath)
            ' Or network path
            ElseIf PtrNetworkVolumeInfo Then
                Shortcut.Filename = ReadSingleString(FileNo, NextPtr + PtrNetworkVolumeInfo + &H14)
            End If
            
            ' Read remaining filename
            If PtrFilename Then
                Str = ReadSingleString(FileNo, NextPtr + PtrFilename)
                If Str <> "" Then
                    If Right(Shortcut.Filename, 1) <> "\" Then
                        Shortcut.Filename = Shortcut.Filename & "\"
                    End If
                    Shortcut.Filename = Shortcut.Filename & Str
                End If
            End If
            
            ' Position pointer to next block
            Get FileNo, NextPtr, IntValue
            NextPtr = NextPtr + IntValue
        End If
    End If
        
    ' Check if Description section is defined (Optional)
    If LinkFlags And 4 Then
        ' Read string length followed by double-byte string
        Get FileNo, NextPtr, IntValue
        NextPtr = NextPtr + IntValue * 2 + 2
        Shortcut.Description = ReadDoubleString(FileNo, IntValue)
    End If
    
    ' Check if Relative Path section is defined (Optional)
    If LinkFlags And 8 Then
        ' Read string length followed by double-byte string
        Get FileNo, NextPtr, IntValue
        NextPtr = NextPtr + IntValue * 2 + 2
        Shortcut.RelPath = ReadDoubleString(FileNo, IntValue)
    End If
    
    ' Check if Working Directory section is defined (Optional)
    If LinkFlags And 16 Then
        ' Read string length followed by double-byte string
        Get FileNo, NextPtr, IntValue
        NextPtr = NextPtr + IntValue * 2 + 2
        Shortcut.WorkingDir = ReadDoubleString(FileNo, IntValue)
    End If
    
     ' Check if Arguments section is defined (Optional)
    If LinkFlags And 32 Then
        ' Read string length followed by double-byte string
        Get FileNo, NextPtr, IntValue
        NextPtr = NextPtr + IntValue * 2 + 2
        Shortcut.Arguments = ReadDoubleString(FileNo, IntValue)
    End If
    
    ' Check if CustomIcon section is defined (Optional)
    If LinkFlags And 64 Then
        ' Read string length followed by double-byte string
        Get FileNo, NextPtr, IntValue
        NextPtr = NextPtr + IntValue * 2 + 2
        Shortcut.CustomIcon = ReadDoubleString(FileNo, IntValue)
    End If
        
    Close FileNo
    'GetShortcutInfo = (Shortcut.FileName <> "") 'this line has been disabled to ensure that even shortcuts where we cannot extract a filename can be utilised, ie. no longer mandatory.

   On Error GoTo 0
   Exit Function

GetShortcutInfo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetShortcutInfo of Module mdlMain"
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadSingleString
' Author    : Jacques Lebrun
' Date      : 28/05/2021
' Purpose   :  Read a single-byte string from the file
'---------------------------------------------------------------------------------------
'
Private Function ReadSingleString(FileNo As Integer, Offset As Long) As String

    Dim Str As String
    Dim ByteValue As Byte
    
    On Error GoTo ReadSingleString_Error

    Seek FileNo, Offset
    Get FileNo, , ByteValue
    Str = ""
    
    Do While ByteValue <> 0
        Str = Str & ChrW$(ByteValue)
        Get FileNo, , ByteValue
    Loop
    
    ReadSingleString = Str

    On Error GoTo 0
    Exit Function

ReadSingleString_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReadSingleString of Module mdlMain"

End Function

'---------------------------------------------------------------------------------------
' Procedure : ReadDoubleString
' Author    : Jacques Lebrun
' Date      : 28/05/2021
' Purpose   : Read a double-byte string value preceded by its length
'---------------------------------------------------------------------------------------
'
Private Function ReadDoubleString(FileNo As Integer, StrLen As Integer) As String

    Dim IntValue As Integer
    Dim Str As String
    
    On Error GoTo ReadDoubleString_Error

    Str = ""
    Do While StrLen > 0
        Get FileNo, , IntValue
        Str = Str & ChrW$(IntValue)
        StrLen = StrLen - 1
    Loop
    
    ReadDoubleString = Str

    On Error GoTo 0
    Exit Function

ReadDoubleString_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReadDoubleString of Module mdlMain"
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetShellShortcutInfo
' Author    : beededea
' Date      : 16/04/2021
' Purpose   : This is our old standard method of getting the path from a LNK file using shell scripting that can cause a/v tools to FP.
'---------------------------------------------------------------------------------------
' CREDIT http://www.vb-helper.com/howto_get_shortcut_info.html
Public Function GetShellShortcutInfo(ByVal full_name As String, _
    ByRef name As String, ByRef Path As String, ByVal descr _
    As String, ByRef working_dir As String, ByRef args As _
    String) As String

    Dim shl As Shell32.Shell
    Dim shortcut_path
    Dim shortcut_name As String

    Dim shortcut_folder As Shell32.folder
    Dim folder_item As Shell32.FolderItem
    Dim lnk As Shell32.ShellLinkObject

    'On Error GoTo GetShellShortcutInfo_Error

    ' Make a Shell object.
    
    Set shl = New Shell32.Shell

    ' Get the shortcut's folder and name.
    shortcut_path = Left$(full_name, InStrRev(full_name, _
        "\"))
    shortcut_name = Mid$(full_name, InStrRev(full_name, _
        "\") + 1)
    If Not Right$(shortcut_name, 4) = ".lnk" Then _
        shortcut_name = shortcut_name & ".lnk"

    ' Get the shortcut's folder.
    Set shortcut_folder = shl.NameSpace(shortcut_path)

    ' Get the shortcut's file.
    Set folder_item = _
        shortcut_folder.Items.Item(shortcut_name)
    If folder_item Is Nothing Then
        GetShellShortcutInfo = "Cannot find shortcut file '" & _
            full_name & "'"
    ElseIf Not folder_item.IsLink Then
        ' It's not a link.
        GetShellShortcutInfo = "File '" & full_name & "' isn't a " & _
            "shortcut."
    Else
        ' Display the shortcut's information.
        Set lnk = folder_item.GetLink
        name = folder_item.name
        descr = lnk.Description
        Path = lnk.Path
        working_dir = lnk.WorkingDirectory
        args = lnk.Arguments
        GetShellShortcutInfo = vbNullString
    End If
    

   On Error GoTo 0
   Exit Function

GetShellShortcutInfo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetShellShortcutInfo of Form dock"
End Function


