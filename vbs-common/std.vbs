'------------------------------------------------
' \file std.vbs
' \brief vbs 基类库
' 
' 提供vbs常用函数和类
' 
' \author wangyang 
' \date 2015/03/11 
' \version 2.1
'------------------------------------------------
Const COM_FSO           = "Scripting.FileSystemObject"
Const COM_SHELL         = "WScript.Shell"
Const COM_SHELLAPP      = "Shell.Application"
Const COM_HTML          = "htmlfile"
Const COM_HTTP          = "Msxml2.XMLHTTP"
Const COM_XMLHTTP       = "Msxml2.ServerXMLHTTP"
Const COM_ADOSTREAM     = "Adodb.Stream"
Const COM_XMLDOM        = "Microsoft.XMLDOM"
Const COM_WMI           = "winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2"
Const COM_WMP           = "WMPlayer.ocx"
Const COM_WORD          = "Word.Application"
Const COM_EXCEL         = "Excel.Application"
Const COM_ACCESS        = "Access.Application"
Const COM_PHOTOSHOP     = "PHOTOSHOP.APPLICATION"
Const COM_DICT          = "Scripting.Dictionary"
Const COM_ADO_CONN      = "ADODB.Connection"
Const COM_ADO_RECORDSET = "ADODB.Recordset"
Const COM_ADO_COMMAND   = "ADODB.Command"
Const COM_ADO_CATALOG   = "ADOX.Catalog"
Const COM_COMMONDIALOG  = "UserAccounts.CommonDialog"
Const COM_IE            = "InternetExplorer.Application"
Const COM_TYPELIB       = "Scriptlet.TypeLib"
Const COM_POCKET_HTTP   = "pocket.HTTP"
Const COM_CAPICOM_UTIL  = "CAPICOM.Utilities"
Const COM_CAPICOM_HASH  = "CAPICOM.HashedData"
Const COM_REGEXP        = "VBSCRIPT.REGEXP"

'------------------------------------------------
' VB常数
' vbCrLf        Chr(13) + Chr(10)   回车/换行组合符
' vbCr          Chr(13)             回车符
' vbLf          Chr(10)             换行符
' vbNewLine     Chr(13) + Chr(10)   换行符
' vbNullChar    Chr(0)              值为 0 的字符
' vbNullString  值为 0 的字符串
' vbObjectError -2147221504         错误号。用户定义的错误号应当大于该值。例如：Err.Raise(Number) = vbObjectError + 1000
' vbTab         Chr(9)              Tab 字符
' vbBack        Chr(8)              退格字符

'------------------------------------------------
' FSO
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const CreateIfNeeded = true



Const DESKTOP = &H10&
Const LOCAL_APPLICATION_DATA = &H1c&
Const TEMPORARY_INTERNET_FILES = &H20&
Const FOF_CREATEPROGRESSDLG = &H0&

'------------------------------------------------
' Registry
Const HKEY_CLASSES_ROOT     = &H80000000
Const HKEY_CURRENT_USER     = &H80000001
Const HKEY_LOCAL_MACHINE    = &H80000002
Const HKEY_USERS            = &H80000003
Const HKEY_CURRENT_CONFIG   = &H80000005
Const HKCR = &H80000000 'HKEY_CLASSES_ROOT
Const HKCU = &H80000001 'HKEY_CURRENT_USER
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
Const HKU  = &H80000003 'HKEY_USERS
Const HKCC = &H80000005 'HKEY_CURRENT_CONFIG
Const REG_SZ        = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY    = 3
Const REG_DWORD     = 4
Const REG_MULTI_SZ  = 7

'------------------------------------------------
' Valid Charset values for ADODB.Stream
Const CdoBIG5        = "big5"
Const CdoEUC_JP      = "euc-jp"
Const CdoEUC_KR      = "euc-kr"
Const CdoGB2312      = "gb2312"
Const CdoISO_2022_JP = "iso-2022-jp"
Const CdoISO_2022_KR = "iso-2022-kr"
Const CdoISO_8859_1  = "iso-8859-1"
Const CdoISO_8859_2  = "iso-8859-2"
Const CdoISO_8859_3  = "iso-8859-3"
Const CdoISO_8859_4  = "iso-8859-4"
Const CdoISO_8859_5  = "iso-8859-5"
Const CdoISO_8859_6  = "iso-8859-6"
Const CdoISO_8859_7  = "iso-8859-7"
Const CdoISO_8859_8  = "iso-8859-8"
Const CdoISO_8859_9  = "iso-8859-9"
Const cdoKOI8_R      = "koi8-r"
Const cdoShift_JIS   = "shift-jis"
Const CdoUS_ASCII    = "us-ascii"
Const CdoUTF_7       = "utf-7"
Const CdoUTF_8       = "utf-8"

'------------------------------------------------
' Constants used by MS ADO.DB 

'---- CursorTypeEnum Values ----
Const adOpenForwardOnly     = 0
Const adOpenKeyset          = 1
Const adOpenDynamic         = 2
Const adOpenStatic          = 3

'---- LockTypeEnum Values ----
Const adLockReadOnly        = 1
Const adLockPessimistic     = 2
Const adLockOptimistic      = 3
Const adLockBatchOptimistic = 4

'---- CursorLocationEnum Values ----
Const adUseServer           = 2
Const adUseClient           = 3

'---- SearchDirection Values ----
Const adSearchForward       = 1
Const adSearchBackward      = -1

'---- CommandTypeEnum Values ----
Const adCmdUnknown          = &H0008
Const adCmdText             = &H0001
Const adCmdTable            = &H0002
Const adCmdStoredProc       = &H0004



'------------------------------------------------
' ADODB.Stream file I/O constants
Const adTypeBinary          = 1
Const adTypeText            = 2
Const adSaveCreateNotExist  = 1
Const adSaveCreateOverWrite = 2
Const adModeUnknown         = 0
Const adModeRead            = 1
Const adModeWrite           = 2
Const adModeReadWrite       = 3


'------------------------------------------------
' CAPICOM
Const CAPICOM_HASH_ALGORITHM_SHA1   = 0
Const CAPICOM_HASH_ALGORITHM_MD2    = 1
Const CAPICOM_HASH_ALGORITHM_MD4    = 2
Const CAPICOM_HASH_ALGORITHM_MD5    = 3
Const CAPICOM_HASH_ALGORITHM_SHA256 = 4
Const CAPICOM_HASH_ALGORITHM_SHA384 = 5
Const CAPICOM_HASH_ALGORITHM_SHA512 = 6

'------------------------------------------------
' Base64
Const sBASE_64_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"  


'------------------------------------------------
' 进制数
' Oct(1234)     8进制
' Hex(1234)     16进制



'------------------------------------------------
' IIf条件表达式
Function IIf(condition, resTrue, resFalse)
    If condition Then
        IIf = resTrue
    Else
        IIf = resFalse
    End if
End Function

'------------------------------------------------
' 打印字符串
Sub Echo(message)
    WScript.Echo message
End Sub

'------------------------------------------------
' 打印字符串，带换行符
Sub Println(message)
    Dim stdout
    Set stdout = WScript.StdOut
    stdout.WriteLine message
End Sub

'------------------------------------------------
' 打印字符串，带换行符
Sub Pause(message)
    WScript.Echo message
    z = WScript.StdIn.Read(1)
End Sub

'------------------------------------------------
' HTA Sleep
Function HTA_Sleep(n)
    Dim SHELL
    Set SHELL = CreateObject(COM_SHELL)
    Call SHELL.Run("%comspec% /c ping -n " + n + " 127.0.0.1 > nul", 0, 1)
    Set SHELL = Nothing
End Function

'------------------------------------------------
' 字符串添加引号
Function Quotes(strQuotes)
    Quotes = chr(34) & strQuotes & chr(34)
End Function


'------------------------------------------------
' 脚本目录
Function ScriptPath
    ScriptPath = left(Wscript.ScriptFullName,len(Wscript.ScriptFullName)-len(Wscript.ScriptName))
    'ScriptPath = Replace(WScript.ScriptFullName, "\" & WScript.ScriptName, "")
End Function

'------------------------------------------------
' 文件夹是否存在
Function FolderExists(dir)
    Dim FSO 
    Set FSO = CreateObject(COM_FSO) 
    FolderExists = FSO.FolderExists(dir)
    Set FSO = Nothing 
End Function

'------------------------------------------------
' 不需要实时输出，执行，返回errorCode
Function Run(Cmd)
    Dim objShell, errorCode
    Set objShell = CreateObject(COM_SHELL)
    errorCode = objShell.Run(Cmd, 0, True)
    Run = errorCode
    Set objShell = Nothing
End Function

'------------------------------------------------
' 执行，实时输出
Sub Exec(Cmd)
    Dim objShell, objExec, comspec
    Set objShell = CreateObject(COM_SHELL)	
    comspec = objShell.ExpandEnvironmentStrings("%comspec%")
    Set objExec = objShell.Exec(comspec & " /c ipconfig")
    Do
        WScript.StdOut.WriteLine(objExec.StdOut.ReadLine())
    Loop While Not objExec.Stdout.atEndOfStream
    WScript.StdOut.WriteLine(objExec.StdOut.ReadAll)
    Set objShell = Nothing
End Sub

'------------------------------------------------
' 执行PHP脚本
Function ExecPHP(phpFile)
    Dim objShell, objExec, php, arrStr
    Set objShell = CreateObject(COM_SHELL)
    php = config("PHP")
    Set objExec = objShell.Exec(php & " " & phpFile)
    ExecPHP = objExec.StdOut.ReadAll
End Function

'------------------------------------------------
' 执行Jar文件
Function ExecJar(jarFile)
    Dim objShell, objExec, java, arrStr
    Set objShell = CreateObject(COM_SHELL)
    java = config("JAVA")
    Set objExec = objShell.Exec(java & " -jar " & jarFile)
    ExecJar = objExec.StdOut.ReadAll
End Function


'------------------------------------------------
' 枚举当前脚本目录的子目录
Sub EnumCurrentDirectory(ByRef arrFolders)
    Dim objShell, objFSO, objFolder, currentDirectory, folder
    Set objShell = CreateObject(COM_SHELL)
    currentDirectory = objShell.CurrentDirectory
    Set objFSO = CreateObject(COM_FSO)
    Set objFolder = objFSO.GetFolder(currentDirectory)
    Set arrFolders = objFolder.SubFolders	
    Set objShell = Nothing
    Set objFSO = Nothing
End Sub

'------------------------------------------------
' 重命名文件夹
Sub RenameFolders(folder1, folder2)
    On Error Resume Next
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    FSO.MoveFolder (folder1),(folder2)
    If Err.Number <> 0 Then
        WScript.Echo "sorry you have a file open in that directory"
        WScript.Echo Err.Description
        WScript.Echo Err.Number
        Err.Clear 
    End If
End Sub

'------------------------------------------------
' 重命名文件
Sub RenameFile(sourcefile, destfile)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    FSO.MoveFile sourcefile, destfile
End Sub

'------------------------------------------------
' 创建级联目录
' Example:
'   ForceCreateFolder("C:\d\e\f\g\h")
Sub ForceCreateFolder(dir)
    Dim FSO, dirpath
    Set FSO = CreateObject(COM_FSO)
    dirpath = FSO.GetAbsolutePathName(dir)
    If (Not FSO.folderExists(FSO.GetParentFolderName(dirpath))) then    
        Call ForceCreateFolder(fso.GetParentFolderName(dirpath))
    End If
    
    FSO.CreateFolder(dirpath)
End Sub

'------------------------------------------------
' 删除目录
' Example:
'   ForceDeleteFolder("C:\d")
Sub ForceDeleteFolder(dir)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    dir = FSO.GetAbsolutePathName(dir)
    If (FSO.FolderExists(dir)) Then
        FSO.DeleteFolder(dir)
    End If
End Sub

'------------------------------------------------
' 拷贝文件
Sub CopyFile(SourceFile, DestinationFile)
    
    Set FSO = CreateObject(COM_FSO)
    
    'Check to see if the file already exists in the destination folder
    Dim wasReadOnly
    wasReadOnly = False
    If FSO.FileExists(DestinationFile) Then
        'Check to see if the file is read-only
        If fso.GetFile(DestinationFile).Attributes And 1 Then 
            'The file exists and is read-only.
            WScript.Echo "Removing the read-only attribute"
            'Remove the read-only attribute
            FSO.GetFile(DestinationFile).Attributes = FSO.GetFile(DestinationFile).Attributes - 1
            wasReadOnly = True
        End If
        
        WScript.Echo "Deleting the file"
        FSO.DeleteFile DestinationFile, True
    End If
    
    'Copy the file
    WScript.Echo "Copying " & SourceFile & " to " & DestinationFile
    FSO.CopyFile SourceFile, DestinationFile, True
    
    If wasReadOnly Then
        'Reapply the read-only attribute
        FSO.GetFile(DestinationFile).Attributes = FSO.GetFile(DestinationFile).Attributes + 1
    End If
    
    Set FSO = Nothing
    
End Sub

'------------------------------------------------
' 枚举System环境变量
Sub EnumSystemEnvironment(ByRef arrEnvironment)
    Dim objShell, objEnv
    Set objShell = CreateObject(COM_SHELL)
    Set arrEnvironment = objShell.Environment("SYSTEM")
End Sub

'------------------------------------------------
' 桌面文件夹
Function DesktopDir
    Dim objShell
    Set objShell = CreateObject(COM_SHELL)
    DesktopDir = objShell.SpecialFolders("desktop")
    Set objShell = Nothing
End Function


'------------------------------------------------
' 获取屏幕分辨率
Sub GetScreenWidthHeight(ByRef width, ByRef height)
    Dim objHTML, objScreen
    Set objHTML = CreateObject(COM_HTML)
    Set objScreen = objHTML.parentwindow.screen
    width = objScreen.width
    height = objScreen.height
    Set objHTML = Nothing
End Sub

'------------------------------------------------
' 显示桌面
Sub ShowDesktop
    Dim objShell
    Set objShell = CreateObject(COM_SHELLAPP)
    objShell.ToggleDesktop
    Set objShell = Nothing
End Sub

'------------------------------------------------
' 暂停
Sub Pause(message)
    Dim char
    WScript.Echo(message)
    char = WScript.StdIn.Read(1)
End Sub

'------------------------------------------------
' 重启计算机
Sub ShutDown
    Dim Result, SHELL
    Set SHELL = CreateObject(COM_SHELL)
    Result = MsgBox("你确定要重起计算机吗?",vbokcancel+vbexclamation,"注意！") 
    If Result = vbOk Then
        SHELL.Run("Shutdown.exe -r -t 0")
    End If
End Sub

Sub ImportRegistry
End Sub

Function IsX64()
    Dim objWMI, colItems, objItem, computer
    IsX64 = False
    computer = "."
    Set objWMI = CreateObject("winmgmts:{impersonationLevel=impersonate}!\\"&computer&"\root\cimv2")
    Set colItems = objWMI.ExecQuery("Select * from Win32_ComputerSystem",,48)
    For Each objItem in colItems		
        If InStr(objItem.SystemType, "64") <> 0 Then
            IsX64 = True
            Exit For
        End If
    Next
    Set objWMI = Nothing
End Function

'------------------------------------------------
' 播放MP3
Function PlayMp3(FileName)
    Dim objWWP, objShell 
    Set objShell = CreateObject(COM_SHELL)
    Set objWMP = CreateObject(COM_WMP)
    objWMP.url = FileName
    Do Until objWMP.playState = 1
        objShell.Sleep 100
    Loop
    Set objShell = Nothing
    Set objWMP = Nothing
End Function

'------------------------------------------------
' 选择文件
Function SelectFile
    Dim objDialog
    Set objDialog = CreateObject(COM_COMMONDIALOG)
    objDialog.Filter = "Windows Media 音频(*.wma;*.wav)|*.wma;*.wav|MP3(*.mp3)|*.mp3|All Files(*.*)|*.*"
    objDialog.InitialDir = ScriptPath
    intResult = objDialog.ShowOpen
    If intResult = 0 Then
        SelectFile = ""
    Else
        SelectFile = objDialog.FileName
    End If
End Function


'------------------------------------------------
' 下载文件
Sub DwonloadFile(url,target)    
    Dim http, adodbStream
    Set http = CreateObject(COM_HTTP)
    http.open "GET",url,False
    http.send
    Set adodbStream = createobject(COM_ADOSTREAM)
    adodbStream.Type = adTypeBinary
    adodbStream.Open
    adodbStream.Write http.responseBody
    adodbStream.SaveToFile target
    adodbStream.Close
    Set adodbStream = Nothing
End Sub

'------------------------------------------------
' 读文本文件
Function ReadFile(ByVal filename)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    
    If InStr(filename, ":\") = 0 And Left(filename, 2) <> "\\" Then 
        filename = FSO.GetSpecialFolder(0) & "\" & filename
    End If
    
    On Error Resume Next
    ReadFile = FSO.OpenTextFile(filename).ReadAll
End Function

'------------------------------------------------
' 写文本文件
Function WriteFile(ByVal filename, ByVal Contents)
    
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    
    If InStr(filename, ":\") = 0 And Left(filename, 2) <> "\\" Then 
        filename = FSO.GetSpecialFolder(0) & "\" & filename
    End If
    
    Dim OutStream
    Set OutStream = FSO.OpenTextFile(filename, 2, True)
    OutStream.Write Contents
End Function

'------------------------------------------------
' 读文本文件到数组
Function ReadFile2Array(ByVal filename)
    Dim arrFileLines(), FSO, file, I
    I = 0    
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.OpenTextFile(filename, ForReading)    
    Do Until file.AtEndOfStream
        Redim Preserve arrFileLines(i)
        arrFileLines(i) = file.ReadLine
        I = I + 1
    Loop    
    file.Close
    Set FSO = Nothing        
    ReadFile2Array = arrFileLines
End Function


'------------------------------------------------
' 读二进制文件
Function ReadBinary(FileName)    
    Dim adodbStream, xmldom, node
    Set xmldom = CreateObject(COM_XMLDOM)
    Set node = xmldom.CreateElement("binary")
    node.DataType = "bin.hex"
    Set adodbStream = CreateObject(COM_ADOSTREAM)
    adodbStream.Type = adTypeBinary
    adodbStream.Open
    adodbStream.LoadFromFile FileName
    node.NodeTypedValue = adodbStream.Read
    adodbStream.Close
    Set adodbStream = Nothing
    ReadBinary = node.Text
    Set node = Nothing
    Set xmldom = Nothing
End Function

'------------------------------------------------
' 写二进制文件
Function WriteBinary(FileName, Buf)    
    Dim adodbStream, xmldom, node
    Set xmldom = CreateObject(COM_XMLDOM)
    Set node = xmldom.CreateElement("binary")
    node.DataType = "bin.hex"
    node.Text = Buf
    Set adodbStream = CreateObject(COM_ADOSTREAM)
    adodbStream.Type = adTypeBinary
    adodbStream.Open
    adodbStream.write node.NodeTypedValue
    adodbStream.saveToFile FileName, adSaveCreateOverWrite
    adodbStream.Close
    Set adodbStream = Nothing
    Set node = Nothing
    Set xmldom = Nothing
End Function 

'------------------------------------------------
' 读文本文件
Function ReadTextFile(filename, charset)
    Dim adodbStream, retval    
    Set adodbStream = CreateObject(COM_ADOSTREAM)
    adodbStream.Type = adTypeText '以本模式读取
    adodbStream.mode = adModeReadWrite 
    adodbStream.charset = charset
    adodbStream.Open
    adodbStream.loadfromfile filename
    retval = adodbStream.readtext
    adodbStream.Close
    Set adodbStream = Nothing
    ReadTextFile = retval
End Function 


'------------------------------------------------
' 写文本文件
Function WriteTextFile(filename, byval Str, charset) 
    Dim adodbStream
    Set adodbStream = CreateObject(COM_ADOSTREAM)
    adodbStream.Type = adTypeText '以本模式读取
    adodbStream.mode = adModeReadWrite
    adodbStream.charset = charset
    adodbStream.Open
    adodbStream.WriteText str
    adodbStream.SaveToFile filename, 2 
    adodbStream.flush
    adodbStream.Close
    Set adodbStream = nothing
End Function 

'------------------------------------------------
' 字符串转字节数组
Function Str2Bytes(str, charset)
    Dim adodbStream, strRet 
    Set adodbStream = CreateObject(COM_ADOSTREAM)     
    adodbStream.Type = adTypeText              
    adodbStream.Charset = charset    
    adodbStream.Open                     
    adodbStream.WriteText str                  
    adodbStream.Position = 0         
    adodbStream.Type = adTypeBinary        
    vout = adodbStream.Read(adodbStream.Size)   
    adodbStream.Close                
    Set adodbStream = nothing 
    Str2Bytes = vout 
End Function

'------------------------------------------------
' 字节数组转字符串
Function BytesToBstr(str, charset)
    If LenB(str) = 0 Then  
        BytesToBstr = "" 
        Exit Function 
    End If 
    
    Dim adodbStream 
    Set adodbStream = CreateObject(COM_ADOSTREAM) 
    adodbStream.Type = adTypeBinary 
    adodbStream.Mode = adModeReadWrite 
    adodbStream.Open 
    adodbStream.Write str 
    adodbStream.Position = 0 
    adodbStream.Type = adTypeText 
    adodbStream.Charset = charset 
    BytesToBstr = adodbStream.ReadText 
    adodbStream.Close 
    Set adodbStream = nothing 
End Function

'------------------------------------------------
' WriteINIString
' example:
'   WriteINIString "Mail", "MAPI", "1", "win.ini"
'   wscript.echo GetINIString("Mail", "MAPI", "-", "win.ini")
Sub WriteINIString(Section, KeyName, Value, FileName)
    Dim INIContents, PosSection, PosEndSection
    
    
    INIContents = ReadFile(FileName)
    
    'Find section
    PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
    If PosSection > 0 Then
        'Section exists. Find end of section
        PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
        '?Is this last section?
        If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
        
        'Separate section contents
        Dim OldsContents, NewsContents, Line
        Dim sKeyName, Found
        OldsContents = Mid(INIContents, PosSection, PosEndSection - PosSection)
        OldsContents = split(OldsContents, vbCrLf)
        
        'Temp variable To find a Key
        sKeyName = LCase(KeyName & "=")
        
        'Enumerate section lines
        For Each Line In OldsContents
            If LCase(Left(Line, Len(sKeyName))) = sKeyName Then
                Line = KeyName & "=" & Value
                Found = True
            End If
            NewsContents = NewsContents & Line & vbCrLf
        Next
        
        If isempty(Found) Then
            'key Not found - add it at the end of section
            NewsContents = NewsContents & KeyName & "=" & Value
        Else
            'remove last vbCrLf - the vbCrLf is at PosEndSection
            NewsContents = Left(NewsContents, Len(NewsContents) - 2)
        End If
        
        'Combine pre-section, new section And post-section data.
        INIContents = Left(INIContents, PosSection-1) & _
        NewsContents & Mid(INIContents, PosEndSection)
    Else        'if PosSection>0 Then
        'Section Not found. Add section data at the end of file contents.
        If Right(INIContents, 2) <> vbCrLf And Len(INIContents)>0 Then 
            INIContents = INIContents & vbCrLf 
        End If
        INIContents = INIContents & "[" & Section & "]" & vbCrLf & _
        KeyName & "=" & Value
    End If      'if PosSection>0 Then
    WriteFile FileName, INIContents
End Sub

'------------------------------------------------
' GetINIString
Function GetINIString(Section, KeyName, Default, FileName)
    Dim INIContents, PosSection, PosEndSection, sContents, Value, Found
    
    
    INIContents = ReadFile(FileName)
    
    'Find section
    PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
    If PosSection > 0 Then
        'Section exists. Find end of section
        PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
        '?Is this last section?
        If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
        
        'Separate section contents
        sContents = Mid(INIContents, PosSection, PosEndSection - PosSection)
        
        If InStr(1, sContents, vbCrLf & KeyName & "=", vbTextCompare)>0 Then
            Found = True
            'Separate value of a key.
            Value = SeparateField(sContents, vbCrLf & KeyName & "=", vbCrLf)
        End If
    End If
    
    If isempty(Found) Then Value = Default
    
    GetINIString = Value
End Function

'------------------------------------------------
' Separates one field between sStart And sEnd
Function SeparateField(ByVal sFrom, ByVal sStart, ByVal sEnd)
    Dim PosB: PosB = InStr(1, sFrom, sStart, 1)
    If PosB > 0 Then
        PosB = PosB + Len(sStart)
        Dim PosE: PosE = InStr(PosB, sFrom, sEnd, 1)
        If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf, 1)
        If PosE = 0 Then PosE = Len(sFrom) + 1
        SeparateField = Mid(sFrom, PosB, PosE - PosB)
    End If
End Function

'------------------------------------------------
' 文件是否存在
Function FileExists(filename)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    FileExists = FSO.FileExists(filename)
End Function

'------------------------------------------------
' 目录是否存在
Function DirExists(dirname)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    DirExists = FSO.FolderExists(dirname)
End Function

'------------------------------------------------
' 移动文件夹
' sourcedir = "C:\Scripts"
' destdir = "D:\Archive"
Sub MoveFolder(sourcedir, destdir)    
    Dim objShell, objFolder
    Set objShell = CreateObject(COM_SHELLAPP)
    Set objFolder = objShell.NameSpace(destdir) 
    objFolder.MoveHere sourcedir, FOF_CREATEPROGRESSDLG
End Sub

'------------------------------------------------
' 删除文件
' 删除.txt文件，"C:\FSO\*.txt"
Function DeleteFiles(filename)
    
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    
    If FSO.FileExists(filename) Then
        FSO.DeleteFile filename, True
        DeleteFiles = True
    Else
        DeleteFiles = False
    End If
    
    Set FSO = Nothing
    
End Function 

'------------------------------------------------
' 删除特定文件
' @delfilesname         文件列表"test1.txt|test2.txt"
' @dirname            文件目录
Sub DelFiles(delfilesname, dirname) 
    Dim FSO, files, fullpath, I
    If Right(dirname, 1) <> "\" Then dirname = dirname & "\"
    If delfilesname <> "" And Not IsNull(delfilesname) Then
        Set FSO = CreateObject(COM_FSO)
        files = Split(delfilesname & "|", "|")
        For I = 0 to Ubound(files) - 1
            fullpath = dirname + files(I)
            If FSO.FileExists(fullpath) Then FSO.DeleteFile(fullpath)
        Next
    End If
End Sub

'------------------------------------------------
' 删除特定文件
' @dir          文件目录
' @days         当前日期减去多少天
Sub DeleteFilesByDate(dir, days)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    Call DeleteSubFolders(FSO.GetFolder(dir), days, FSO)
End Sub

Sub DeleteSubFolders(folder, days, fso)
    Dim subfolder, files
    For Each subfolder in folder.SubFolders
        Set files = subfolder.Files
        If files.Count <> 0 Then
            For Each file in Files
                If file.DateLastModified < (Now - days) Then
                    fso.DeleteFile(subfolder.Path & "\" & file.Name)    
                End If
            Next
        End If
        Call DeleteSubFolders(subfolder, days, fso)
    Next
End Sub

'------------------------------------------------
' 文件重命名
Function ReFilename(filename, name)
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.GetFile(filename)
    file.Name = name
    Set FSO = Nothing
End Function 

'------------------------------------------------
' 文件夹重命名
Function ReDir(source, dest)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    FSO.MoveFolder source, dest
    Set FSO = Nothing
End Function

'------------------------------------------------
' 获取文件路径
Function GetFilePath(filename)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    GetFilePath = DisposePath(FSO.GetParentFolderName(filename))
End Function 


'------------------------------------------------
' 获取文件绝对路径
Function GetAbsolutePathName(filename)
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.GetFile(filename)
    GetAbsolutePathName = FSO.GetAbsolutePathName(file)
End Function

'------------------------------------------------
' 获取文件名
Function GetFileName(filename)
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.GetFile(filename)
    GetFileName = FSO.GetFileName(file)
End Function

'------------------------------------------------
' 获取基本文件名
Function GetBaseName(filename)
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.GetFile(filename)
    GetBaseName = FSO.GetBaseName(file)
End Function

'------------------------------------------------
' 获取文件扩展名
Function GetExtensionName(filename)
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.GetFile(filename)
    GetExtensionName = FSO.GetExtensionName(file)
End Function

'------------------------------------------------
' 获取文件扩展名
Function GetAnExtension(filename)
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    GetAnExtension = FSO.GetExtensionName(filename)
End Function

'------------------------------------------------
' 获取工作目录
Function GetCurrentDirectory() 
    Dim objShell
    Set objShell = CreateObject(COM_SHELL)
    GetCurrentDirectory = objShell.CurrentDirectory 
End Function 

'------------------------------------------------
' 获取脚本工作目录
Function GetScriptPath()
    GetScriptPath = left(Wscript.ScriptFullName,len(Wscript.ScriptFullName)-len(Wscript.ScriptName))
End Function



'------------------------------------------------
' 获取GUID值
Function NewGUID
    Set TypeLib = CreateObject(COM_TYPELIB) 
    NewGUID = Left(TypeLib.Guid, 38)
    Set TypeLib = Nothing
End Function 

'------------------------------------------------
' 获取GUID值, 不带{}
Function NewGUID2  
    Set TypeLib = CreateObject(COM_TYPELIB)
    NewGUID2 = Mid(TypeLib.Guid, 2, 36)
    Set TypeLib = Nothing
End Function 


'------------------------------------------------
' 区间随机数
' @lowerbound       下限
' @upperbound       上限
Function RandomNum(lowerbound, upperbound)
    Randomize
    RandomNum = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function

'------------------------------------------------
' 创建随机密码
Function CreatePassword(numchar)
    Dim avail, parola, f, i
    
    avail = "abcdefghijklmnopqrstuvwxyz1234567890"
    Randomize
    parola = ""
    for f = 1 to numchar
        i = (CInt(len(avail) * Rnd + 1) mod len(avail)) + 1
        parola = parola & mid(avail, i, 1)
    next
    CreatePassword = parola
End Function


'------------------------------------------------
' 字符串转数字
' @strS         字符串
' @return       Integer (>=0)
Function CID(strS)
    Dim intI
    intI = 0
    If IsNull(strS) Or strS = "" Then
        intI = 0
    Else
        If Not IsNumeric(strS) Then
            intI = 0
        Else
            Dim intk
            On Error Resume Next
            intk = Abs(Clng(strS))
            If Err.Number = 6 Then intk = 0  '数据溢出
            Err.Clear
            intI = intk
        End If
    End If
    CID = intI
End Function

'------------------------------------------------
' 判断用户名是否合法
' @username        用户名
Function IsTrueName(username)
    Dim Hname, I
    IsTrueName = False
    Hname = Array("=", "%", chr(32), "?", "&", ";", ",", "'", ",", chr(34), chr(9), "", "$", "|")
    For I = 0 To Ubound(Hname)
        If InStr(username, Hname(I)) > 0 Then
            Exit Function
        End If
    Next
    IsTrueName = True 
End Function

'------------------------------------------------
' 路径末尾添加\
Function DisposePath(sPath)
    On Error Resume Next
    
    If Right(sPath, 1) = "\" Then
        DisposePath = sPath
    Else
        DisposePath = sPath & "\"
    End If
    
    DisposePath = Trim(DisposePath)
End Function 

'------------------------------------------------
' 替换文件内容
Function ReplaceFileContent(filepath, pattern, text, is_utf8)
    Set objFSO = CreateObject(COM_FSO)
    Set objFile = objFSO.GetFile(filepath)
    Dim objStream
    
    If objFile.Size > 0 Then
        
        If is_utf8 = 1 Then			
            Set objStream = CreateObject(COM_ADOSTREAM)
            objStream.Open
            objStream.Type = adTypeText
            objStream.Position = 0
            objStream.Charset = CdoUTF_8
            objStream.LoadFromFile filepath
            strContents = objstream.ReadText
            objStream.Close
            Set objStream = Nothing
        Else
            Set objReadFile = objFSO.OpenTextFile(filepath, 1)
            strContents = objReadFile.ReadAll
            objReadFile.Close
        End If
    End If
    
    Dim re
    Set re = new RegExp
    re.IgnoreCase = False
    re.Global = True
    re.MultiLine = True
    re.Pattern = pattern
    strContents = re.replace(strContents, text)
    
    're.Pattern="^Public\s+Const\s+APP_VERSION.*""$"
    'strContents = re.replace(strContents,"Public Const APP_VERSION = ""Version: " & appversion & """")
    
    Set re = Nothing
    
    If is_utf8 = 1 Then
        Set objStream = CreateObject(COM_ADOSTREAM)
        objStream.Open
        objStream.Type = adTypeText
        objStream.Position = 0
        objStream.Charset = CdoUTF_8
        objStream.WriteText = strContents
        objStream.SaveToFile filepath, adSaveCreateOverWrite
        objStream.Close
        Set objStream = Nothing
    Else
        Set objWriteFile = objFSO.OpenTextFile(filepath, 2, False)
        objWriteFile.Write(strContents)
        objWriteFile.Close
    End If
End Function 

'------------------------------------------------
' 获取桌面路径
Function GetDesktopPath()
    Set objShell = CreateObject(COM_SHELLAPP)
    Set objFolder = objShell.Namespace(DESKTOP)
    Set objFolderItem = objFolder.Self
    GetDesktopPath = objFolderItem.Path
End Function

'------------------------------------------------
' 获取应用程序数据路径
Function GetApplicationDataPath()
    Dim SHELL, folder, folder_item
    Set SHELL = CreateObject(COM_SHELLAPP)
    Set folder = SHELL.Namespace(LOCAL_APPLICATION_DATA)
    Set folder_item = folder.Self
    GetApplicationDataPath = folder_item.Path
End Function 


'------------------------------------------------
' 获取临时文件夹路径
Function GetTempPath()
    Set objShell = CreateObject(COM_SHELLAPP)
    Set objFolder = objShell.Namespace(TEMPORARY_INTERNET_FILES)
    Set objFolderItem = objFolder.Self
    GetTempPath = objFolderItem.Path
End Function 

'------------------------------------------------
' 创建临时文件
Function CreateTempFile(dir)
    Dim FSO, tempname, fullname, file
    Set FSO = CreateObject(COM_FSO)
    tempname = FSO.GetTempName
    fullname = FSO.BuildPath(dir, tempname)
    Set file = FSO.CreateTextFile(fullname)
    file.Close
End Function



'------------------------------------------------
' 获取正则匹配内容
Function GetMatchText(filename, pattern)
    Dim text, re, matches, tmpstr
    text = ReadTextFile(filename, "gb2312")
    
    Set re = new RegExp
    re.IgnoreCase = False
    re.Global = True
    re.MultiLine = True
    re.Pattern = pattern
    
    Set matches = re.Execute(text)
    If matches.Count > 0 Then
        For Each m In matches
            If m.SubMatches.Count > 0 Then
                GetMatchText = m.SubMatches(0)
            End If
        Next
    End If
End Function 

'------------------------------------------------
' 获取文件行数
Function GetFileLines(filename)
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)
    Set file = FSO.OpenTextFile(filename, ForReading)
    ' Skip lines one by one
    Do While file.AtEndOfStream <> True
        file.SkipLine
    Loop
    
    GetFileLines = file.Line
    
    Set FSO = Nothing
End Function



'------------------------------------------------
' 遍历文件夹
'	Function testfile(filename)
'		WScript.Echo filename
'	End Function
'
'	Call EachFiles("D:\tools\7-Zip", "\.txt", "testfile")
Sub EachFiles(dir, pattern, method)
    Dim FSO, re
    Set FSO = CreateObject(COM_FSO)
    Set root = FSO.GetFolder(dir)
    Set re = new RegExp
    re.Pattern    = pattern
    re.IgnoreCase = True
    
    Call EachSubFolder(root, re, method)
    
    Set FSO = Nothing
    Set re = Nothing
End Sub

Sub EachSubFolder(root, re, method)
    Dim subfolder, file, script
    
    For Each file In root.Files
        If re.Test(file.Name) Then
            script = "Call " & method & "(""" & file.Path & """)"
            ExecuteGlobal script
        End If
    Next
    
    For Each subfolder In root.SubFolders
        Call EachSubFolder(subfolder, re, method)    
    Next
End Sub

'------------------------------------------------
' 根据原文件名，自动以日期YYYY-MM-DD-RANDOM格式生成新文件名
Function GetfileExt(byval filename)
    Dim fileExt_a
    fileExt_a = Split(filename,".")
    GetfileExt = Lcase(fileExt_a(Ubound(fileExt_a)))
End Function

'------------------------------------------------
' 根据原文件名，自动以日期YYYY-MM-DD-RANDOM格式生成新文件名
Function GenerateRandomFileName(ByVal filename)
    Randomize
    ranNum = Int(90000 * Rnd) + 10000
    If Month(Now) < 10 Then c_month = "0" & Month(Now) Else c_month = Month(Now)
    If Day(Now) < 10 Then c_day = "0" & Day(Now) Else c_day = Day(Now)
    If Hour(Now) < 10 Then c_hour = "0" & Hour(Now) Else c_hour = Hour(Now)
    If Minute(Now) < 10 Then c_minute = "0" & Minute(Now) Else c_minute = Minute(Now)
    If Second(Now) < 10 Then c_second = "0" & Second(Now) Else c_second = Minute(Now)
    fileExt_a = Split(filename, ".")
    FileExt = LCase(fileExt_a(UBound(fileExt_a)))
    GenerateRandomFileName = Year(Now) & c_month & c_day & c_hour & c_minute & c_second & "_" & ranNum & "." & FileExt
End Function


'------------------------------------------------
' 建立目录的程序，如果有多级目录，则一级一级的创建
Function CreateDir(ByVal LocalPath) 
    On Error Resume Next
    Dim FSO
    LocalPath = Replace(LocalPath, "\", "/")
    Set FSO = CreateObject(COM_FSO)
    patharr = Split(LocalPath, "/")
    path_level = UBound(patharr)
    For I = 0 To path_level
        If I = 0 Then pathtmp = patharr(0) & "/" Else pathtmp = pathtmp & patharr(I) & "/"
        cpath = Left(pathtmp, Len(pathtmp) - 1)
        If Not FSO.FolderExists(cpath) Then FSO.CreateFolder cpath
    Next
    Set FSO = Nothing
    If Err.Number <> 0 Then
        CreateDir = False
        Err.Clear
    Else
        CreateDir = True
    End If
End Function

'------------------------------------------------
' NewZip
Sub NewZip(filename) 
    'WScript.Echo "Newing up a zip file (" & pathToZipFile & ") "
    
    Dim FSO, file
    Set FSO = CreateObject(COM_FSO)   
    Set file = FSO.CreateTextFile(filename)
    
    file.Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0) 
    file.Close
    Set FSO = Nothing
    Set FSO = Nothing 
    WScript.Sleep 500 
End Sub

'------------------------------------------------
' CreateZip         空目录无法压缩
' Example:
'   CreateZip "results.zip", "results"
Sub CreateZip(filename, dir) 
    'WScript.Echo "Creating zip  (" & pathToZipFile & ") from (" & dirToZip & ")"
    
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    
    filename = FSO.GetAbsolutePathName(filename)
    dir = FSO.GetAbsolutePathName(dir)
    
    If FSO.FileExists(filename) Then
        'WScript.Echo "That zip file already exists - deleting it."
        FSO.DeleteFile filename
    End If
    
    If Not FSO.FolderExists(dir) Then
        'WScript.Echo "The directory to zip does not exist."
        Exit Sub
    End If
    
    NewZip filename
    
    Dim SHELLAPP, zip, d
    Set SHELLAPP = CreateObject(COM_SHELLAPP)   
    Set zip = SHELLAPP.NameSpace(filename) 
    
    'WScript.Echo "opening dir  (" & dir & ")" 
    
    Set d = SHELLAPP.NameSpace(dir)
    
    ' Look at http://msdn.microsoft.com/en-us/library/bb787866(VS.85).aspx
    ' for more information about the CopyHere function.
    zip.CopyHere d.items, 4
    
    Do Until d.Items.Count <= zip.Items.Count
        Wscript.Sleep(200)
    Loop
    
End Sub

'------------------------------------------------
' ExtractFilesFromZip
' Example:
'   ExtractFilesFromZip "results.zip", "."
Sub ExtractFilesFromZip(filename, dir)
    
    Dim FSO
    Set FSO = CreateObject(COM_FSO)
    
    filename = fso.GetAbsolutePathName(filename)
    dir = fso.GetAbsolutePathName(dir)
    
    If (Not fso.FileExists(filename)) Then
        WScript.Echo "Zip file does not exist: " & filename
        Exit Sub
    End If
    
    If Not fso.FolderExists(dir) Then
        WScript.Echo "Directory does not exist: " & dir
        Exit Sub
    End If
    
    Dim SHELLAPP, zip, d
    set SHELLAPP = CreateObject("Shell.Application")   
    Set zip = SHELLAPP.NameSpace(filename)  
    Set d = SHELLAPP.NameSpace(dir)
    
    ' Look at http://msdn.microsoft.com/en-us/library/bb787866(VS.85).aspx
    ' for more information about the CopyHere function.
    d.CopyHere zip.items, 4
    
    Do Until zip.Items.Count <= d.Items.Count
        Wscript.Sleep(200)
    Loop
    
End Sub

'------------------------------------------------
' ZipBy7Zip
' @archive_file_name        压缩文件名
' @filelist                 文件列表
' Example:
'   Call ZipBy7Zip("results_01.zip", "111.txt 222.txt") 文件列表
'   Call ZipBy7Zip("files.zip", """c:\program files\text files\*.txt""") 文件列表
'   Call ZipBy7Zip("resutls_02.zip", "dadfasd")     文件夹
Function ZipBy7Zip(archive_file_name, filelist)
    Dim FSO, SHELL, sWorkingDirectory
    Set FSO = CreateObject(COM_FSO)
    Set SHELL = CreateObject(COM_SHELL)   
    
    sWorkingDirectory = FSO.GetParentFolderName(Wscript.ScriptFullName) 
    
    '-------Ensure we can find 7za.exe------
    If FSO.FileExists(sWorkingDirectory & "\" & "7z.exe") Then
        s7zLocation = ""
    ElseIf FSO.FileExists("D:\tools\7-Zip\7z.exe") Then
        s7zLocation = "D:\tools\7-Zip\"
    Else
        ZipBy7Zip = "Error: Couldn't find 7za.exe"
        Exit Function
    End If
    '--------------------------------------
    
    SHELL.Run """" & s7zLocation & "7z.exe"" a -tzip -y """ & archive_file_name & """ " _
    & filelist, 0, True   
    
    If FSO.FileExists(archive_file_name) Then
        ZipBy7Zip = 1
    Else
        ZipBy7Zip = "Error: Archive Creation Failed."
    End If
End Function

'------------------------------------------------
' UnZipBy7Zip
' @archive_file_name        压缩文件名
' @dir                      解压目录
' Example:
'   Call UnZipBy7Zip("results_01.zip", "C:\ddddd\dddd\ddd")
Function UnZipBy7Zip(archive_file_name, dir)  
    Dim FSO, SHELL, sWorkingDirectory
    Set FSO = CreateObject(COM_FSO)
    Set SHELL = CreateObject(COM_SHELL)   
    
    sWorkingDirectory = FSO.GetParentFolderName(Wscript.ScriptFullName) 
    '--------------------------------------
    
    '-------Ensure we can find 7za.exe------
    If FSO.FileExists(sWorkingDirectory & "\" & "7z.exe") Then
        s7zLocation = ""
    ElseIf FSO.FileExists("D:\tools\7-Zip\7z.exe") Then
        s7zLocation = "D:\tools\7-Zip\"
    Else
        UnZipBy7Zip = "Error: Couldn't find 7za.exe"
        Exit Function
    End If
    '--------------------------------------
    
    '-Ensure we can find archive to uncompress-
    If Not FSO.FileExists(archive_file_name) Then
        UnZipBy7Zip = "Error: File Not Found."
        Exit Function
    End If
    '--------------------------------------
    
    SHELL.Run """" & s7zLocation & "7z.exe"" e -y -o""" & dir & """ """ & _
    archive_file_name & """", 0, True
    UnZipBy7Zip = 1
End Function


'------------------------------------------------
' BStr2UStr
Function BStr2UStr(BStr)
    'Byte string to Unicode string conversion
    Dim lngLoop
    BStr2UStr = ""
    For lngLoop = 1 to LenB(BStr)
        BStr2UStr = BStr2UStr & Chr(AscB(MidB(BStr,lngLoop,1))) 
    Next
End Function

'------------------------------------------------
' UStr2Bstr
Function UStr2Bstr(UStr)
    'Unicode string to Byte string conversion
    Dim lngLoop
    Dim strChar
    UStr2Bstr = ""
    For lngLoop = 1 to Len(UStr)
        strChar = Mid(UStr, lngLoop, 1)
        UStr2Bstr = UStr2Bstr & ChrB(AscB(strChar))
    Next
End Function

'------------------------------------------------
' Base64encode
Function Base64Encode(str)  
    Dim CAPIUtil
    Set CAPIUtil = CreateObject(COM_CAPICOM_UTIL)
    Base64encode = CAPIUtil.Base64Encode(str)
    Set CAPIUtil = Nothing
End Function

'------------------------------------------------
' Base64decode
Function Base64Decode(str) 
    Dim CAPIUtil
    Set CAPIUtil = CreateObject(COM_CAPICOM_UTIL)
    Base64Decode = CAPIUtil.Base64Decode(str)
    Set CAPIUtil = Nothing
End Function 

'------------------------------------------------
' MD5
Function MD5(str) 
    Dim CAPIHASH
    Set CAPIHASH = CreateObject(COM_CAPICOM_HASH)
    CAPIHASH.Algorithm = CAPICOM_HASH_ALGORITHM_MD5
    CAPIHASH.Hash UStr2Bstr(str)
    MD5 = CAPIHASH.Value
    Set CAPIHASH = Nothing
End Function 

'------------------------------------------------
' MD5_File
Function MD5_File(filename, raw_output)
    Dim HashedData, Utility, Stream
    Set HashedData = CreateObject(COM_CAPICOM_HASH)
    Set Utility = CreateObject(COM_CAPICOM_UTIL)
    Set Stream = CreateObject(COM_ADOSTREAM)
    HashedData.Algorithm = CAPICOM_HASH_ALGORITHM_MD5
    Stream.Type = 1
    Stream.Open
    Stream.LoadFromFile filename
    Do Until Stream.EOS
        HashedData.Hash Stream.Read(1024)
    Loop
    If raw_output Then
        MD5_File = Utility.HexToBinary(HashedData.Value)
    Else
        MD5_File = HashedData.Value
    End If
End Function

'------------------------------------------------
' SHA1
Function SHA1(str) 
    Dim CAPIHASH
    Set CAPIHASH = CreateObject(COM_CAPICOM_HASH)
    CAPIHASH.Algorithm = CAPICOM_HASH_ALGORITHM_SHA1
    CAPIHASH.Hash UStr2Bstr(str)
    SHA1 = CAPIHASH.Value
    Set CAPIHASH = Nothing
End Function 

'------------------------------------------------
' URLEncoding
Function URLEncoding(vstrIn) 
    Dim strReturn, ThisChr, innerCode, Hight8, Low8
    strReturn = "" 
    For i = 1 To Len(vstrIn) 
        ThisChr = Mid(vStrIn,i,1) 
        If Abs(Asc(ThisChr)) < &HFF Then 
            strReturn = strReturn & ThisChr 
        Else 
            innerCode = Asc(ThisChr) 
            If innerCode < 0 Then 
                innerCode = innerCode + &H10000 
            End If 
            Hight8 = (innerCode And &HFF00) OR &HFF 
            Low8 = innerCode And &HFF 
            strReturn = strReturn & "%" & Hex(Hight8) &  "%" & Hex(Low8) 
        End If 
    Next 
    URLEncoding = strReturn 
End Function 


'------------------------------------------------
' GetHttp
Function GetHttp(url) 
    Dim xmlhttp
    Set xmlhttp = CreateObject(COM_XMLHTTP)  
    postdata = "" 
    xmlhttp.Open "GET", url, False 
    xmlhttp.setRequestHeader "Authorization", "Basic " & Base64encode("test:pass") 
    'xmlhttp.setRequestHeader("Referer","来路的绝对地址") 
    'xmlhttp.setRequestHeader "Cookie",Cookies   'Cookie 
    xmlhttp.Send postdata 
    Wscript.echo xmlhttp.status & ":" & xmlhttp.statusText 
    respStr = BytesToBstr(xmlhttp.responseBody, "UTF-8") 
    Wscript.echo respStr 
    Set xmlhttp = nothing 
End Function 

'------------------------------------------------
' HttpGet
' @url          URL地址
' @charset      网页编码(gb2312, utf-8)
Function HttpGet(url, charset)
    Dim xmlhttp
    Set xmlhttp = CreateObject(COM_XMLHTTP)    
    xmlhttp.Open "GET", url, False     
    xmlhttp.Send() 
    If xmlhttp.readystate <> 4 Then
        Exit Function
    End If
    HttpGet = BytesToBstr(xmlhttp.responseBody, charset)     
    Set xmlhttp = nothing 
End Function


'------------------------------------------------
' PostHttp
Function PostHttp(url) 
    Set xmlhttp = CreateObject(COM_XMLHTTP)  
    postdata = "" 
    xmlhttp.Open "POST", url1, False 
    xmlhttp.setRequestHeader "CONTENT-TYPE","application/x-www-form-urlencoded" 
    xmlhttp.setRequestHeader "Authorization", "Basic " & Base64encode("test:pass") 
    'xmlhttp.setRequestHeader("Referer","来路的绝对地址") 
    'xmlhttp.setRequestHeader "Cookie",Cookies   'Cookie 
    xmlhttp.Send postdata 
    Wscript.echo xmlhttp.status & ":" & xmlhttp.statusText 
    respStr = BytesToBstr(xmlhttp.responseBody, "GB2312") 
    Wscript.echo respStr 
    Set xmlhttp = nothing 
End Function 


'------------------------------------------------
' 过滤html标签
Function FilterHtml(str)
    Dim re    
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.MultiLine = True
    re.Pattern = "<.+?>"
    FilterHtml = re.Replace(str, "")
    Set re = Nothing
End Function

'------------------------------------------------
' 过滤html标签
Function StripHTML(ByRef sHTML)
    Dim re 
    Set re = New RegExp
    re.Pattern = "<[^>]*>" 
    re.IgnoreCase = True  
    re.Global = True    
    StripHTML = re.Replace(sHTML, " ")   
    Set re = Nothing
End Function

'------------------------------------------------
' 过滤指定html标签
Function DecodeFilter(html, filter)
    html = LCase(html)
    filter = split(filter, ",")
    For Each i In filter
        Select Case i
            Case "SCRIPT"   ' 去除所有客户端脚本javascipt,vbscript,jscript,js,vbs,event,...
            html = exeRE("(javascript|jscript|vbscript|vbs):", "#", html)
            html = exeRE("</?script[^>]*>", "", html)
            html = exeRE("on(mouse|exit|error|click|key)", "", html)
            Case "TABLE":   ' 去除表格<table><tr><td><th>
            html = exeRE("</?table[^>]*>", "", html)
            html = exeRE("</?tr[^>]*>", "", html)
            html = exeRE("</?th[^>]*>", "", html)
            html = exeRE("</?td[^>]*>", "", html)
            html = exeRE("</?tbody[^>]*>", "", html)
            Case "CLASS"    ' 去除样式类class=""
            html = exeRE("(<[^>]+) class=[^ |^>]*([^>]*>)", "$1 $2", html) 
            Case "STYLE"    ' 去除样式style=""
            html = exeRE("(<[^>]+) style=""[^""]*""([^>]*>)", "$1 $2", html)
            html = exeRE("(<[^>]+) style='[^']*'([^>]*>)", "$1 $2", html)
            Case "IMG"      ' 去除样式style=""
            html = exeRE("</?img[^>]*>", "", html)
            Case "XML"      ' 去除XML<?xml>
            html = exeRE("<\\?xml[^>]*>", "", html)
            Case "NAMESPACE"    ' 去除命名空间<o:p></o:p>
            html = exeRE("<\/?[a-z]+:[^>]*>", "", html)
            Case "FONT"     ' 去除字体<font></font>
            html = exeRE("</?font[^>]*>", "", html)
            Case "MARQUEE"  ' 去除字幕<marquee></marquee>
            html = exeRE("</?marquee[^>]*>", "", html)
            Case "OBJECT"   ' 去除对象<object><param><embed></object>
            html = exeRE("</?object[^>]*>", "", html)
            html = exeRE("</?param[^>]*>", "", html)
            html = exeRE("</?embed[^>]*>", "", html)
            Case "DIV"      ' 去除对象<object><param><embed></object>
            html = exeRE("</?div([^>])*>", "$1", html)
        End Select
    Next
    'html = Replace(html,"<table","<")
    'html = Replace(html,"<tr","<")
    'html = Replace(html,"<td","<")
    DecodeFilter = html
End Function

'------------------------------------------------
' 字符串转Unicode
Function Chinese2Unicode(str) 
    Dim i 
    Dim Str_one 
    Dim Str_unicode 
    For i = 1 To Len(str) 
        Str_one = Mid(str, i, 1) 
        Str_unicode = Str_unicode & chr(38) 
        Str_unicode = Str_unicode & chr(35) 
        Str_unicode = Str_unicode & chr(120) 
        Str_unicode = Str_unicode & Hex(ascw(Str_one)) 
        Str_unicode = Str_unicode & chr(59) 
    Next 
    
    str = Str_unicode
End Function

'------------------------------------------------
' 正则表达式替换
' @content  文本
' @pattern  正则表达式模式
' @str      替换字符串
Function ReplaceText(content, pattern, str)
    Dim re
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = pattern
    ReplaceText = re.Replace(content, str)
    Set re = Nothing    
End Function


'------------------------------------------------
' HTMLEncode
Function HTMLEncode(text)
    If text = "" or IsNull(text) Then 
        Exit Function
    Else
        If Instr(text, "'") > 0 Then 
            text = replace(text, "'", "&#39;")
        End If
        text = replace(text, ">", "&gt;")
        text = replace(text, "<", "&lt;")
        text = Replace(text, CHR(32), "&nbsp;")
        text = Replace(text, CHR(9), "&nbsp;")
        text = Replace(text, CHR(34), "&quot;")
        text = Replace(text, CHR(13),"")
        text = Replace(text, CHR(10) & CHR(10), "</P><P>")
        text = Replace(text, CHR(10), "<BR>")
        text = Replace(text, CHR(39), "&#39;")
        text = Replace(text, CHR(0), "")
        text = ChkBadWords(text)
        HTMLEncode = text
    End If
End Function


'------------------------------------------------
' HTMLDecode
Public Function HTMLDecode(text)
    If text = "" or IsNull(text) Then 
        Exit Function
    Else
        If Instr(text, "'")>0 Then 
            text = replace(text, "'", "&#39;")
        End If
        text = replace(text, "&gt;", ">")
        text = replace(text, "&lt;", "<")
        text = Replace(text, "&nbsp;", CHR(32))
        text = Replace(text, "&nbsp;", CHR(9))
        text = Replace(text, "&quot;", CHR(34))
        text = Replace(text, "", CHR(13))
        text = Replace(text, "</P><P>", CHR(10) & CHR(10))
        text = Replace(text, "<BR>", CHR(10))
        text = Replace(text, "", CHR(0))
        text = Replace(text, "&#39;", CHR(39))
        text = ChkBadWords(text)
        HTMLDecode = text
    End If
End Function


'------------------------------------------------
' 日期格式化
Function DateToStr(DateTime, ShowType)
    Dim DateMonth, DateDay, DateHour, DateMinute
    DateMonth = Month(DateTime)
    DateDay = Day(DateTime)
    DateHour = Hour(DateTime)
    DateMinute = Minute(DateTime)
    If Len(DateMonth) < 2 Then DateMonth = "0" & DateMonth
    If Len(DateDay) < 2 Then DateDay = "0" & DateDay
    Select Case ShowType
        Case "Y-m"
        DateToStr = Year(DateTime) & "-" & Month(DateTime)
        Case "Y-m-d"
        DateToStr = Year(DateTime) & "-" & DateMonth & "-" & DateDay
        Case "Y-m-d H:I A"
        Dim DateAMPM
        If DateHour > 12 Then
            DateHour = DateHour - 12
            DateAMPM = "PM"
        Else
            DateHour = DateHour
            DateAMPM = "AM"
        End If
        If Len(DateHour) < 2 Then DateHour = "0" & DateHour
        If Len(DateMinute) < 2 Then DateMinute = "0" & DateMinute
        DateToStr = Year(DateTime) & "-" & DateMonth & "-" & DateDay & " " & DateHour & ":" & DateMinute & " " & DateAMPM
        Case "Y-m-d H:I:S"
        Dim DateSecond
        DateSecond = Second(DateTime)
        If Len(DateHour) < 2 Then DateHour = "0" & DateHour
        If Len(DateMinute) < 2 Then DateMinute = "0" & DateMinute
        If Len(DateSecond) < 2 Then DateSecond = "0" & DateSecond
        DateToStr = Year(DateTime) & "-" & DateMonth & "-" & DateDay & " " & DateHour & ":" & DateMinute & ":" & DateSecond
        Case "YmdHIS"
        DateSecond = Second(DateTime)
        If Len(DateHour) < 2 Then DateHour = "0" & DateHour
        If Len(DateMinute) < 2 Then DateMinute = "0" & DateMinute
        If Len(DateSecond) < 2 Then DateSecond = "0" & DateSecond
        DateToStr = Year(DateTime) & DateMonth & DateDay & DateHour & DateMinute & DateSecond
        Case "Ymd"			
        DateToStr = Year(DateTime) & DateMonth & DateDay 
        Case "ym"
        DateToStr = Right(Year(DateTime), 2) & DateMonth
        Case "d"
        DateToStr = DateDay
        Case Else
        If Len(DateHour) < 2 Then DateHour = "0" & DateHour
        If Len(DateMinute) < 2 Then DateMinute = "0" & DateMinute
        DateToStr = Year(DateTime) & "-" & DateMonth & "-" & DateDay & " " & DateHour & ":" & DateMinute
    End Select
End Function



'------------------------------------------------
' 根据年份及月份得到每月的总天数
Function GetDaysInMonth(iMonth, iYear) 
    Select Case iMonth 
        Case 1, 3, 5, 7, 8, 10, 12 
        GetDaysInMonth = 31 
        Case 4, 6, 9, 11 
        GetDaysInMonth = 30 
        Case 2 
        If IsDate("February 29, " & iYear) Then 
            GetDaysInMonth = 29 
        Else 
            GetDaysInMonth = 28 
        End If 
    End Select 
End Function 

'------------------------------------------------
' 得到一个月开始的日期
Function GetWeekdayMonthStartsOn(dAnyDayInTheMonth) 
    Dim dTemp 
    dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth) 
    GetWeekdayMonthStartsOn = WeekDay(dTemp) 
End Function 

'------------------------------------------------
' 得到当前一个月的上一个月
Function SubtractOneMonth(dDate) 
    SubtractOneMonth = DateAdd("m", -1, dDate) 
End Function 

'------------------------------------------------
' 得到当前一个月的下一个月
Function AddOneMonth(dDate) 
    AddOneMonth = DateAdd("m", 1, dDate) 
End Function 

'------------------------------------------------
' 中文日期格式化
Function Date2Chinese(iDate)
    Dim num(10)
    Dim iYear
    Dim iMonth
    Dim iDay
    
    num(0) = "〇"
    num(1) = "一"
    num(2) = "二"
    num(3) = "三"
    num(4) = "四"
    num(5) = "五"
    num(6) = "六"
    num(7) = "七"
    num(8) = "八"
    num(9) = "九"
    
    iYear = Year(iDate)
    iMonth = Month(iDate)
    iDay = Day(iDate)
    Date2Chinese = (num(iYear \ 1000) + num((iYear \ 100) Mod 10) + num((iYear\ 10) Mod 10) + num(iYear Mod 10)) & "年"
    
    If iMonth >= 10 Then
        If iMonth = 10 Then
            Date2Chinese = Date2Chinese & "十" & "月"
        Else
            Date2Chinese = Date2Chinese & "十" & num(iMonth Mod 10) & "月"
        End If
    Else
        Date2Chinese = Date2Chinese & num(iMonth Mod 10) & "月"
    End If
    
    If iDay >= 10 Then
        If iDay = 10 Then
            Date2Chinese = Date2Chinese & "十" & "日"
        ElseIf iDay = 20 or iDay = 30 Then
            Date2Chinese = Date2Chinese & num(iDay \ 10) & "十" & "日"
        ElseIf iDay > 20 Then
            Date2Chinese = Date2Chinese & num(iDay \ 10) & "十" & num(iDay Mod 10) & "日"
        Else
            Date2Chinese = Date2Chinese & "十" & num(iDay Mod 10) & "日"
        End If
    Else
        Date2Chinese = Date2Chinese & num(iDay Mod 10) & "日"
    End If
    
End Function

'------------------------------------------------
' Date2ChineseRSS
Function Date2ChineseRSS(iDate)
    Dim num(10)
    Dim iYear
    Dim iMonth
    Dim iDay
    
    num(0) = "〇"
    num(1) = "一"
    num(2) = "二"
    num(3) = "三"
    num(4) = "四"
    num(5) = "五"
    num(6) = "六"
    num(7) = "七"
    num(8) = "八"
    num(9) = "九"
    
    iYear = Year(iDate)
    iMonth = Month(iDate)
    iDay = Day(iDate)
    Date2ChineseRSS = iYear & "年"
    
    If iMonth >= 10 Then
        If iMonth = 10 Then
            Date2ChineseRSS = Date2ChineseRSS & "十" & "月"
        Else
            Date2ChineseRSS = Date2ChineseRSS & "十" & num(iMonth Mod 10) & "月"
        End If
    Else
        Date2ChineseRSS = Date2ChineseRSS & num(iMonth Mod 10) & "月"
    End If
    
End Function


'------------------------------------------------
' Convert a string to a date or datetime
' IN  : sDate (string) : source (format YYYYMMDD HH:MM:SS or YYYYMMDD)
' OUT : (datetime) : destination
Function StringToDate(strDate)
    Dim dDate, sDate
    
    sDate = Trim(strDate)
    Select Case Len(sDate)
        Case 17
        dDate = DateSerial(Left(sDate, 4), Mid(sDate, 5, 2), Mid(sDate, 7, 2)) + TimeSerial(Mid(sDate, 10, 2), Mid(sDate, 13, 2), Mid(sDate, 16, 2))
        Case 8
        dDate = DateSerial(Left(sDate, 4), Mid(sDate, 5, 2), Mid(sDate, 7, 2))
        Case Else
        If isDate(sDate) Then
            dDate = CDate(sDate)
        End If
    End Select
    StringToDate = dDate
End Function


'------------------------------------------------
' 取字段数据每个汉字的拼音首字母
Function getpychar(char)
    tmp = 65536 + Asc(char)
    If (tmp>= 45217 And tmp<= 45252) Then
        getpychar = "A"
    ElseIf (tmp>= 45253 And tmp<= 45760) Then
        getpychar = "B"
    ElseIf (tmp>= 47761 And tmp<= 46317) Then
        getpychar = "C"
    ElseIf (tmp>= 46318 And tmp<= 46825) Then
        getpychar = "D"
    ElseIf (tmp>= 46826 And tmp<= 47009) Then
        getpychar = "E"
    ElseIf (tmp>= 47010 And tmp<= 47296) Then
        getpychar = "F"
    ElseIf (tmp>= 47297 And tmp<= 47613) Then
        getpychar = "G"
    ElseIf (tmp>= 47614 And tmp<= 48118) Then
        getpychar = "H"
    ElseIf (tmp>= 48119 And tmp<= 49061) Then
        getpychar = "J"
    ElseIf (tmp>= 49062 And tmp<= 49323) Then
        getpychar = "K"
    ElseIf (tmp>= 49324 And tmp<= 49895) Then
        getpychar = "L"
    ElseIf (tmp>= 49896 And tmp<= 50370) Then
        getpychar = "M"
    ElseIf (tmp>= 50371 And tmp<= 50613) Then
        getpychar = "N"
    ElseIf (tmp>= 50614 And tmp<= 50621) Then
        getpychar = "O"
    ElseIf (tmp>= 50622 And tmp<= 50905) Then
        getpychar = "P"
    ElseIf (tmp>= 50906 And tmp<= 51386) Then
        getpychar = "Q"
    ElseIf (tmp>= 51387 And tmp<= 51445) Then
        getpychar = "R"
    ElseIf (tmp>= 51446 And tmp<= 52217) Then
        getpychar = "S"
    ElseIf (tmp>= 52218 And tmp<= 52697) Then
        getpychar = "T"
    ElseIf (tmp>= 52698 And tmp<= 52979) Then
        getpychar = "W"
    ElseIf (tmp>= 52980 And tmp<= 53640) Then
        getpychar = "X"
    ElseIf (tmp>= 53689 And tmp<= 54480) Then
        getpychar = "Y"
    ElseIf (tmp>= 54481 And tmp<= 62289) Then
        getpychar = "Z"
    Else '如果不是中文，则不处理
        getpychar = char
    End If
End Function

'------------------------------------------------
' 获取拼音 
Function GetPinYin(Str)
    Dim I
    For I = 1 To Len(Str)
        GetPinYin = GetPinYin & getpychar(Mid(Str, i, 1))
    Next
End Function

'------------------------------------------------
' 验证Email 
Function CheckEmail(Str)
    CheckEmail = False
    Dim re, match
    Set re = New RegExp
    re.Pattern = "^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$"
    re.IgnoreCase = True
    Set match = re.Execute(Str)
    If match.Count Then CheckEmail = True
    Set re = Nothing
End Function

'------------------------------------------------
' 验证用户名
Function CheckUserName(str)
    Dim re
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.MultiLine = True
    re.Pattern = "^[a-z0-9_]{2,20}$"
    CheckUserName = re.Test(str)
    Set re = Nothing
End Function

'------------------------------------------------
' 获取计算机名
Function GetComputerName()
    Dim shell, regpath
    Set shell = CreateObject(COM_SHELL)
    regpath = "HKLM\System\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName"
    GetComputerName = shell.RegRead(regpath)
End Function

'------------------------------------------------
' 杀死进程
' @name         进程名
' example:
'   Call KillProcess("rar.exe")
Sub KillProcess(name)
    Dim computer, WMI, processlist, process
    computer = "."
    Set WMI = GetObject("winmgmts:\" & computer & "\root\cimv2")
    Set processlist = WMI.ExecQuery("Select * from Win32_Process Where Name = '" & name & "'")
    For Each process in processlist
        process.Terminate()
    Next
End Sub

'------------------------------------------------
' 删除注册表键
Sub RegDelete(fullkey)
    Set objShell = CreateObject(COM_SHELL)
    objShell.RegDelete fullkey
End Sub

'------------------------------------------------
' 删除注册表键
Sub RegDeleteKey(rootkey, key)
    Dim computer   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv")
         
    oReg.DeleteKey rootkey, key
End Sub 

'------------------------------------------------
' 删除注册表键值
Sub RegDeleteValue(rootkey, key, name)
    Dim computer   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv") 
    oReg.DeleteValue rootkey, key, name
End Sub


'------------------------------------------------
' 写注册表MultiString值
Sub RegWriteMultiStringValue(rootkey, key, name, ByRef values)
    Dim computer   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv")
         
    oReg.SetMultiStringValue rootkey, key, name, values
End Sub

'------------------------------------------------
' 读注册表MultiString值
Function RegReadMultiString(rootkey, key, name)
    Dim computer, arrValues   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv") 

    oReg.GetMultiStringValue rootkey, key, name, arrValues
    RegReadMultiString = arrValues
End Function

'------------------------------------------------
' 写注册表String值
Sub RegWriteStringValue(rootkey, key, name, value)
    Dim computer   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv") 

    oReg.SetStringValue rootkey, key, name, value
End Sub

'------------------------------------------------
' 读注册表String值
Function RegReadStringValue(rootkey, key, name)
    Dim computer, value   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv") 

    oReg.GetStringValue rootkey, key, name, value
    RegReadStringValue = value
End Function

'------------------------------------------------
' 写注册表DWORD值
Sub RegWriteDWORDValue(rootkey, key, name, value)
    Dim computer   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv") 

    oReg.SetDWORDValue rootkey, key, name, value
End Sub

'------------------------------------------------
' 读注册表DWORD值
Function RegReadDWORDValue(rootkey, key, name)
    Dim computer, value   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv") 

    oReg.GetDWORDValue rootkey, key, name, value
    RegReadDWORDValue = value
End Function


'------------------------------------------------
' 枚举注册表键
Function RegEnumKeys(rootkey, key)
    Dim computer, arrSubKeys   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv") 
    oReg.EnumKey rootkey, key, arrSubKeys 
    RegEnumKeys = arrSubKeys
End Function

'------------------------------------------------
' 枚举注册表值
Function RegEnumValues(rootkey, key, ByRef arrValueTypes)
    Dim computer, arrValueNames   
    computer = "."     
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
        computer & "\root\default:StdRegProv")
    oReg.EnumValues rootkey, key, arrValueNames, arrValueTypes
    RegEnumValues = arrValueNames
End Function


'------------------------------------------------
' Doc2PDF
' Example:
'   Doc2PDF "C:\Documents and Settings\MyUserID\My Documents\resume.doc"
Sub Doc2PDF( myFile )
    ' This subroutine opens a Word document, then saves it as PDF, and closes Word.
    ' If the PDF file exists, it is overwritten.
    ' If Word was already active, the subroutine will leave the other document(s)
    ' alone and close only its "own" document.
    '
    ' Requirements:
    ' This script requires the "Microsoft Save as PDF or XPS Add-in for 2007
    ' Microsoft Office programs", available at:
    ' http://www.microsoft.com/downloads/details.aspx?
    '        familyid=4D951911-3E7E-4AE6-B059-A2E79ED87041&displaylang=en
    '
    ' Written by Rob van der Woude
    ' http://www.robvanderwoude.com
    
    ' Standard housekeeping
    Dim objDoc, objFile, objFSO, objWord, strFile, strPDF
    
    Const wdFormatDocument                    =  0
    Const wdFormatDocument97                  =  0
    Const wdFormatDocumentDefault             = 16
    Const wdFormatDOSText                     =  4
    Const wdFormatDOSTextLineBreaks           =  5
    Const wdFormatEncodedText                 =  7
    Const wdFormatFilteredHTML                = 10
    Const wdFormatFlatXML                     = 19
    Const wdFormatFlatXMLMacroEnabled         = 20
    Const wdFormatFlatXMLTemplate             = 21
    Const wdFormatFlatXMLTemplateMacroEnabled = 22
    Const wdFormatHTML                        =  8
    Const wdFormatPDF                         = 17
    Const wdFormatRTF                         =  6
    Const wdFormatTemplate                    =  1
    Const wdFormatTemplate97                  =  1
    Const wdFormatText                        =  2
    Const wdFormatTextLineBreaks              =  3
    Const wdFormatUnicodeText                 =  7
    Const wdFormatWebArchive                  =  9
    Const wdFormatXML                         = 11
    Const wdFormatXMLDocument                 = 12
    Const wdFormatXMLDocumentMacroEnabled     = 13
    Const wdFormatXMLTemplate                 = 14
    Const wdFormatXMLTemplateMacroEnabled     = 15
    Const wdFormatXPS                         = 18
    
    ' Create a File System object
    Set objFSO = CreateObject( "Scripting.FileSystemObject" )
    
    ' Create a Word object
    Set objWord = CreateObject( "Word.Application" )
    
    With objWord
        ' True: make Word visible; False: invisible
        .Visible = True
        
        ' Check if the Word document exists
        If objFSO.FileExists( myFile ) Then
            Set objFile = objFSO.GetFile( myFile )
            strFile = objFile.Path
        Else
            WScript.Echo "FILE OPEN ERROR: The file does not exist" & vbCrLf
            ' Close Word
            .Quit
            Exit Sub
        End If
        
        ' Build the fully qualified HTML file name
        strPDF = objFSO.BuildPath( objFile.ParentFolder, _
        objFSO.GetBaseName( objFile ) & ".pdf" )
        
        ' Open the Word document
        .Documents.Open strFile
        
        ' Make the opened file the active document
        Set objDoc = .ActiveDocument
        
        ' Save as HTML
        objDoc.SaveAs strPDF, wdFormatPDF
        
        ' Close the active document
        objDoc.Close
        
        ' Close Word
        .Quit
    End With
End Sub

'------------------------------------------------
' ObjectList

Class ObjectList
    Public List
    
    Sub Class_Initialize()
        Set List = CreateObject(COM_DICT)
    End Sub
    
    Sub Class_Terminate()
        Set List = Nothing
    End Sub
    
    Function Append(Anything) 
        List.Add CStr(List.Count + 1), Anything 
        Set Append = Anything
    End Function
    
    Function Item(id) 
        If List.Exists(CStr(id)) Then
            Set Item = List(CStr(id))
        Else
            Set Item = Nothing
        End If
    End Function
End Class


'------------------------------------------------
' XML Upload Class
' Example:
'   Dim UploadData
'   Set UploadData = New XMLUpload
'   UploadData.Charset = "utf-8"
'   UploadData.AddForm "content", "Hello world" '文本域的名称和内容
'   UploadData.AddFile "file", "test.jpg", "image/jpg", "test.jpg"
'   WScript.Echo UploadData.Upload("http://example.com/takeupload.php")
'   Set UploadData = Nothing
Class XMLUpload
    Private xmlHttp
    Private objTemp
    Private adTypeBinary, adTypeText
    Private strCharset, strBoundary
    
    Private Sub Class_Initialize()
        adTypeBinary = 1
        adTypeText = 2
        Set xmlHttp = CreateObject(COM_HTTP)
        Set objTemp = CreateObject(COM_ADOSTREAM)
        objTemp.Type = adTypeBinary
        objTemp.Open
        strCharset = "utf-8"
        strBoundary = GetBoundary()
    End Sub
    
    Private Sub Class_Terminate()
        objTemp.Close
        Set objTemp = Nothing
        Set xmlHttp = Nothing
    End Sub
    
    '指定字符集的字符串转字节数组
    Public Function StringToBytes(ByVal strData, ByVal strCharset)
        Dim objFile
        Set objFile = CreateObject(COM_ADOSTREAM)
        objFile.Type = adTypeText
        objFile.Charset = strCharset
        objFile.Open
        objFile.WriteText strData
        objFile.Position = 0
        objFile.Type = adTypeBinary
        If UCase(strCharset) = "UNICODE" Then
            objFile.Position = 2 'delete UNICODE BOM
        ElseIf UCase(strCharset) = "UTF-8" Then
            objFile.Position = 3 'delete UTF-8 BOM
        End If
        StringToBytes = objFile.Read(-1)
        objFile.Close
        Set objFile = Nothing
    End Function
    
    '获取文件内容的字节数组
    Private Function GetFileBinary(ByVal strPath)
        Dim objFile
        Set objFile = CreateObject(COM_ADOSTREAM)
        objFile.Type = adTypeBinary
        objFile.Open
        objFile.LoadFromFile strPath
        GetFileBinary = objFile.Read(-1)
        objFile.Close
        Set objFile = Nothing
    End Function
    
    '获取自定义的表单数据分界线
    Private Function GetBoundary()
        Dim ret(12)
        Dim table
        Dim i
        table = "abcdefghijklmnopqrstuvwxzy0123456789"
        Randomize
        For i = 0 To UBound(ret)
            ret(i) = Mid(table, Int(Rnd() * Len(table) + 1), 1)
        Next
        GetBoundary = "---------------------------" & Join(ret, Empty)
    End Function 
    
    '设置上传使用的字符集
    Public Property Let Charset(ByVal strValue)
    strCharset = strValue
    End Property
    
    '添加文本域的名称和值
    Public Sub AddForm(ByVal strName, ByVal strValue)
        Dim tmp
        tmp = "\r\n--$1\r\nContent-Disposition: form-data; name=""$2""\r\n\r\n$3"
        tmp = Replace(tmp, "\r\n", vbCrLf)
        tmp = Replace(tmp, "$1", strBoundary)
        tmp = Replace(tmp, "$2", strName)
        tmp = Replace(tmp, "$3", strValue)
        objTemp.Write StringToBytes(tmp, strCharset)
    End Sub
    
    '设置文件域的名称/文件名称/文件MIME类型/文件路径或文件字节数组
    Public Sub AddFile(ByVal strName, ByVal strFileName, ByVal strFileType, ByVal strFilePath)
        Dim tmp
        tmp = "\r\n--$1\r\nContent-Disposition: form-data; name=""$2""; filename=""$3""\r\nContent-Type: $4\r\n\r\n"
        tmp = Replace(tmp, "\r\n", vbCrLf)
        tmp = Replace(tmp, "$1", strBoundary)
        tmp = Replace(tmp, "$2", strName)
        tmp = Replace(tmp, "$3", strFileName)
        tmp = Replace(tmp, "$4", strFileType)
        objTemp.Write StringToBytes(tmp, strCharset)
        objTemp.Write GetFileBinary(strFilePath)
    End Sub
    
    '设置multipart/form-data结束标记
    Private Sub AddEnd()
        Dim tmp
        tmp = "\r\n--$1--\r\n" 
        tmp = Replace(tmp, "\r\n", vbCrLf) 
        tmp = Replace(tmp, "$1", strBoundary)
        objTemp.Write StringToBytes(tmp, strCharset)
        objTemp.Position = 2
    End Sub
    
    '上传到指定的URL，并返回服务器应答
    Public Function Upload(ByVal strURL)
        Call AddEnd
        xmlHttp.Open "POST", strURL, False
        xmlHttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & strBoundary
        'xmlHttp.setRequestHeader "Content-Length", objTemp.size
        xmlHttp.Send objTemp
        Upload = xmlHttp.responseText
    End Function
End Class

'------------------------------------------------
' StringBuilder Class
Class StringBuilder
    Private strArray()
    Private intGrowRate
    Private intItemCount
    
    Private Sub Class_Initialize()
        intGrowRate = 50
        intItemCount = 0
    End Sub
    
    Public Property Get GrowRate
    GrowRate = intGrowRate
    End Property
    
    Public Property Let GrowRate(value)
    intGrowRate = value
    End Property
    
    Private Sub InitArray()
        Redim Preserve strArray(intGrowRate)
    End Sub
    
    Public Sub Clear()
        intItemCount = 0
        Erase strArray
    End Sub
    
    Public Sub Append(str)
        
        If intItemCount = 0 Then
            Call InitArray
        ElseIf intItemCount > UBound(strArray) Then
            Redim Preserve strArray(Ubound(strArray) + intGrowRate)
        End If
        
        strArray(intItemCount) = str
        
        intItemCount = intItemCount + 1
        
    End Sub
    
    Public Function FindString(str)
        Dim x,mx
        mx = intItemCount - 1
        For x = 0 To mx
            If strArray(x) = str Then
                FindString = x
                Exit Function
            End If
        Next
        FindString = -1
    End Function
    
    Public Function ToString2(sep)
        If intItemCount = 0 Then
            ToString2 = ""
        Else
            Redim Preserve strArray(intItemCount)
            ToString2 = Join(strArray,sep)
        End If
    End Function
    
    Public Default Function ToString()
    If intItemCount = 0 Then
        ToString = ""
    Else
        ToString = Join(strArray,"")
    End If
    End Function

End Class


'------------------------------------------------
' DBControl Class
Class DBControl

    Private m_connectionString
    Private m_conn
    Private m_dbType
    
    Private Sub Class_Initialize
        m_dbType = "ACCESS"
        m_connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & config("database_connectionstring")
    End Sub

    Private Sub Class_Terminate        
    End Sub

    Public Property Get ConnectionString()
        ConnectionString = m_connectionString 
    End Property

    Public Function Open()
        On Error Resume Next
        Set m_conn = CreateObject(COM_ADO_CONN)
        m_conn.Open m_connectionString	
        If Err Then
            Err.Clear
            Set m_conn = Nothing
            Response.Write "数据库连接出错，请检查连接字串。"
            Response.End
        End If
    End Function


    Public Function Close()
        m_conn.Close
        Set m_conn = Nothing
    End Function


    Public Function CreateRS()
        Set CreateRS = CreateObject(COM_ADO_RECORDSET)
    End Function

    Public Function BeginTrans()
        m_conn.BeginTrans 
        on error resume next
    End Function

    Public Function EndTrans()
        If Err.number = 0 Then
            m_conn.CommitTrans  
        Else
            m_conn.RollbackTrans 
            strerr = Err.Description
            Response.Write "数据库错误！错误日志：<font color=red>"&strerr &"</font>"
            Response.End
        End If
    End Function


    '函数:根据当前数据库类型转换Sql脚本
    '参数:Sql串
    '返回:转换结果Sql串
    Public Function SqlTran(Sql)
        If m_dbType = "ACCESS" Then
            SqlTran = SqlServer_To_Access(Sql)
        Else
            SqlTran = Sql
        End If
    End Function

    '函数:数据库脚本执行(代Sql转换)
    '参数:Sql脚本
    '返回:执行结果
    '说明:本执行可自动根据数据库类型对部分Sql基础语法进行转换执行
    Public Function ExeCute(Sql)        
        If config("isdebug") = 0 Then 
            On Error Resume Next
            Sql = SqlTran(Sql)
            Set ExeCute = m_conn.ExeCute(Sql)
            If Err Then
                    Err.Clear
                    Set m_conn = Nothing
                    Response.Write "查询数据的时候发现错误,请检查您的查询代码是否正确.<br /><li>"
                    Response.Write Sql
                    Response.End
            End If
        Else
            Set ExeCute = m_conn.ExeCute(Sql)
        End If
        SQL_QUERY_NUM = SQL_QUERY_NUM + 1
    End Function

    '函数:数据库脚本执行
    '参数:Sql脚本
    '返回:执行结果
    Public Function ExeCute2(Sql)
        Set ExeCute2 = m_conn.ExeCute(Sql)
    End Function

    Public Function ExeCute3(sql_proc, ByRef parameters)
        Set cmd = CreateObject(COM_ADO_COMMAND)
        With cmd
            .ActiveConnection = m_conn
            .CommandType = &H0004 '存储过程
            .CommandText = sql_proc
        End With
        Set ExeCute3 = cmd.Execute(, parameters)
    End Function

    '函数:SqlServer(97-2000) to Access(97-2000)
    '参数:Sql,数据库类型(ACCESS,SQLSERVER)
    '说明:
    Public Function SqlServer_To_Access(Sql)
        Dim regEx, Matches, Match
        '创建正则对象
        Set regEx = New RegExp
        regEx.IgnoreCase = True
        regEx.Global = True
        regEx.MultiLine = True

        '转:GetDate()
        regEx.Pattern = "(?=[^']?)GETDATE\(\)(?=[^']?)"
        Sql = regEx.Replace(Sql,"NOW()")

        '转:UPPER()
        regEx.Pattern = "(?=[^']?)UPPER\([\s]?(.+?)[\s]?\)(?=[^']?)"
        Sql = regEx.Replace(Sql,"UCASE($1)")

        '转:日期表示方式
        '说明:时间格式必须是2004-23-23 11:11:10 标准格式
        regEx.Pattern = "'([\d]{4,4}\-[\d]{1,2}\-[\d]{1,2}(?:[\s][\d]{1,2}:[\d]{1,2}:[\d]{1,2})?)'"
        Sql = regEx.Replace(Sql,"#$1#")
        
        regEx.Pattern = "DATEDIFF\([\s]?(second|minute|hour|day|month|year)[\s]?\,[\s]?(.+?)[\s]?\,[\s]?(.+?)([\s]?\)[\s]?)"
        Set Matches = regEx.ExeCute(Sql)
        Dim temStr
        For Each Match In Matches
            temStr = "DATEDIFF("
            Select Case lcase(Match.SubMatches(0))
                Case "second" :
                    temStr = temStr & "'s'"
                Case "minute" :
                    temStr = temStr & "'n'"
                Case "hour" :
                    temStr = temStr & "'h'"
                Case "day" :
                    temStr = temStr & "'d'"
                Case "month" :
                    temStr = temStr & "'m'"
                Case "year" :
                    temStr = temStr & "'y'"
            End Select
            temStr = temStr & "," & Match.SubMatches(1) & "," &  Match.SubMatches(2) & Match.SubMatches(3)
            Sql = Replace(Sql,Match.Value,temStr,1,1)
        Next

        '转:Insert函数
        regEx.Pattern = "CHARINDEX\([\s]?'(.+?)'[\s]?,[\s]?'(.+?)'[\s]?\)[\s]?"
        Sql = regEx.Replace(Sql,"INSTR('$2','$1')")

        Set regEx = Nothing
        SqlServer_To_Access = Sql
    End Function    
End Class

'------------------------------------------------
' Pager Class
Class Pager
    Private m_id 
    Private m_currentpage 
    Private m_recordcount  
    Private m_pagecount  
    Private m_pagesize 
    Private m_endfix

    Public  Function Init(id, currentpage, recordcount, pagesize, endfix)
        m_id = id
        m_currentpage = currentpage
        m_recordcount = recordcount
        m_pagesize = pagesize
        m_endfix = endfix

        If recordcount mod pagesize <> 0 Then
            m_pagecount = Int((recordcount / pagesize) + 1)
        Else 
            m_pagecount = Int(recordcount / pagesize)
        End If
    End Function

    Public Function PageSize()
        PageSize = Int(m_pagesize)
    End Function

    Public Function getHTML() 
        If m_currentpage < 1 Then
            m_currentpage = 1
        End If
        If m_pagecount < 1 Then
            m_pagecount = 1
        End If
        If m_currentpage > m_pagecount Then
            m_currentpage = m_pagecount
        End If


        Dim prevpage 
        prevpage =  m_currentpage - 1 

        Dim nextpage  
        nextpage =  m_currentpage + 1 



        Dim retval 
        Dim sbPager 
        Set sbPager =  New StringBuilder
        sbPager.Append("<span class=""count"">Pages: ")
        sbPager.Append(m_currentpage)
        sbPager.Append("/")
        sbPager.Append(m_pagecount)
        sbPager.Append("</span>")

        sbPager.Append("<b>")

        If prevpage < 1 Then
            sbPager.Append(" &laquo; First")
            sbPager.Append(" &laquo;")
        Else 
            sbPager.Append(" <a href=""" & m_id & "1" & m_endfix & """>&laquo; First</a>")
            sbPager.Append(" <a href=""" & m_id & prevpage & m_endfix & """>&laquo;</a>")
        End If


            Dim startpage 
            If (m_currentpage mod 10) = 0 Then
                startpage = m_currentpage - 9
            Else 
                startpage = m_currentpage - CInt((m_currentpage mod 10)) + 1
            End If

            If startpage > 10 Then
                sbPager.Append(" <a href=""" & m_id & (startpage - 1) & m_endfix & """>...</a>")
            End If

            Dim i 
            For  i = startpage To  startpage + 10- 1  Step  i + 1
                If i > m_pagecount Then
                    Exit For
                End If
                If i = m_currentpage Then
                    sbPager.Append(" <span title=""Page " & i & """>" & i & "</span>")
                Else 
                    sbPager.Append(" <a href=""" & m_id & i & m_endfix & """>" & i & "</a>")
                End If
            Next

            If m_pagecount >= m_currentpage + 10 Then
                sbPager.Append(" <a href=""" & m_id & (startpage + 10) & m_endfix & """>...</a>")
            End If


        If nextpage > m_pagecount Then
            sbPager.Append(" &raquo;")
            sbPager.Append(" Last &raquo;")
        Else 
            sbPager.Append(" <a href=""" & m_id & nextpage & m_endfix & """>&raquo;</a>")
            sbPager.Append(" <a href=""" & m_id & m_pagecount & m_endfix & """>Last &raquo;</a>")
        End If

        sbPager.Append("</b>")

        retval = sbPager.ToString()
        getHTML = retval
    End Function
End Class

'------------------------------------------------
' TagParser Class
Class TagParser

    Private TempContent     ' 临时模版
    Private ResColl         ' 字典对象, 存放标记和标记要替换的内容

    Private Sub Class_Initialize
        Set ResColl = CreateObject(COM_DICT)        
    End Sub

    Private Sub Class_Terminate
        Set ResColl = Nothing
    End Sub

    Public Function Parser(Str)
        TempContent = Str

        ' 开始解析模版
        Tag_Parser()
                
        Parser = TempContent
    End Function


    Private Function Tag_Parser()
        Dim regex, matches, match
        set regex = New RegExp
        regex.IgnoreCase = False
        regex.Global = True
        regex.MultiLine = True

        regex.Pattern = "<cms:file>([^\b]+?)</cms:file>"
        Set matches2 = regex.Execute(TempContent)
        For Each match2 In matches2
            retVal = GetCacheValue(match2.Value)
            If retVal = "" Then 
                If match2.SubMatches(0) <> "" Then
                    retVal = Tag_File_Parser(match2.SubMatches(0))
                End If
            End If
            TempContent = Replace(TempContent, match2.Value,  retVal)
            SetCacheValue match2.Value, retVal, 5
        Next

        regex.Pattern = "<cms:list>([^\b]+?)</cms:list>"
        Set matches = regex.Execute(TempContent)

        Dim strContent, tmpItem
        For Each match In matches
            If match.SubMatches(0) <> "" Then
                'TempContent = match.SubMatches(0)
                'ResColl.Add match.Value Tag_Parser2(match.SubMatches(0))

                TempContent = Replace(TempContent, match.Value,  Tag_Parser2(match.SubMatches(0)))
            End If
        Next

        regex.Pattern = "<cms:function>([^\b]+?)</cms:function>"
        Set matches2 = regex.Execute(TempContent)
        For Each match2 In matches2
            retVal = GetCacheValue(match2.Value)
            If retVal = "" Then 
                If match2.SubMatches(0) <> "" Then
                    Execute("retVal = " & match2.SubMatches(0))
                End If
            End If
            TempContent = Replace(TempContent, match2.Value,  retVal)
            SetCacheValue match2.Value, retVal, 120	
        Next

        regex.Pattern = "<cms:pager>(.*?)</cms:pager>"
        Set matches2 = regex.Execute(TempContent)
        For Each match2 In matches2
            If match2.SubMatches(0) <> "" Then
                TempContent = Replace(TempContent, match2.Value,  "-------------pager-------------------")
            End If
        Next


    End Function

    Private Function Tag_File_Parser(strCommand)
        Dim regex, matches, match, retVal
        set regex = New RegExp
        regex.IgnoreCase = False
        regex.Global = True
        regex.MultiLine = True
        regex.Pattern = "\$([^\b]+?)\$"
        Set matches = regex.Execute(strCommand)
        For Each match In matches
            If match.SubMatches(0) <> "" Then
                filepath = Server.MapPath(".") & "\system\" & Application_PATH & "\views\" & match.SubMatches(0)
                Set filestream = Server.CreateObject(COM_ADOSTREAM)
                    With filestream
                        .Type = 2 '以本模式读取
                        .Mode = 3 
                        .Charset = "utf-8"
                        .Open
                        .Loadfromfile filepath
                        retVal = .readtext
                        .Close
                    End With
                Set filestream = Nothing
            End If
        Next
        Tag_File_Parser = retVal
    End Function

    Private Function Tag_Parser2(strCommand)
        Dim regex, matches, match, retVal, temp
        set regex = New RegExp
        regex.IgnoreCase = False
        regex.Global = True
        regex.MultiLine = True
        regex.Pattern = "<sql>([^\b]+?)</sql>[^\b]*?<template>([^\b]+?)</template>[^\b]*?<cache>([^\b]+?)</cache>"
        Set matches = regex.Execute(strCommand)
        For Each match In matches
            retVal = GetCacheValue(match.Value)
            If retVal = "" Then 
                If match.SubMatches(0) <> "" And match.SubMatches(1) <> "" And match.SubMatches(2) <> "" Then
                    Dim sql, strTemplate, rs, strHTML, strTemplate2
                    sql = match.SubMatches(0)
                    strTemplate = match.SubMatches(1)

                    Dim matches2, match2
                    
                    regex.Pattern = "\$(\w+?)\$"
                    set matches2 = regex.Execute(strTemplate)

                    Dim matches3, match3
                    
                    regex.Pattern = "\$(\w+?)\[(\d+?)\]\$"
                    set matches3 = regex.Execute(strTemplate)

                    Dim matches4, match4

                    regex.Pattern = "\$(\w+?)\((.+?)\)\$"
                    set matches4 = regex.Execute(strTemplate)


                    Set rs = Db.ExeCute(sql)
                    While Not rs.Eof
                        
                        strTemplate2 = strTemplate

                        For Each match4 In matches4
                            'Response.Write match4.SubMatches(1)
                            Dim tempArray, strA
                            tempArray = Split(match4.SubMatches(1), ",")
                            strA = "temp = " & match4.SubMatches(0) & "("
                            For i = 0 To UBound(tempArray)
                                tempArray(i) = rs(Trim(tempArray(i)))
                                If i <> UBound(tempArray) Then
                                    strA = strA & "tempArray(" & i & ")," 
                                Else
                                    strA = strA & "tempArray(" & i & ")"
                                End If
                            Next
                            strA = strA & ")"
                            
                            
                            Execute(strA)
                            strTemplate2 = Replace(strTemplate2, match4.Value, temp)
                        Next

                        For Each match3 In matches3
                            strTemplate2 = Replace(strTemplate2, match3.Value, Left(rs(match3.SubMatches(0)), match3.SubMatches(1)))
                        Next
                                        
                        For Each match2 In matches2
                                strTemplate2 = Replace(strTemplate2, match2.Value, rs(match2.SubMatches(0)))
                        Next

                        strHTML = strHTML & strTemplate2 & vbCrLf
                        rs.MoveNext()
                    Wend
                    rs.Close
                    Set rs = Nothing
                    retVal = strHTML
                    SetCacheValue match.Value, retVal, Int(match.SubMatches(2))
                End If
            End If 
        Next

        Tag_Parser2 = retVal

    End Function
End Class



'------------------------------------------------
' clsThief Class
Class clsThief
    Private value_      ' 窃取到的内容
    Private src_        ' 要偷的目标URL地址
    Private isGet_      ' 判断是否已经偷过
    Private cookie_ 

    ' 赋值—要偷的目标URL地址/属性

    Public Property Let src(Str)
        src_ = Str
    End Property

    '返回值—最终窃取并应用类方法加工过的内容/属性

    Public Property Get Value
        Value = value_
    End Property

    Public Property Get Cookie
        Cookie = cookie_
    End Property

    Public Property Get Version
        Version = "先锋海盗类 Version 2004"
    End Property

    Private Sub class_initialize()
        value_ = ""
        src_ = ""
        isGet_ = False
    End Sub

    Private Sub class_terminate()
    End Sub

    ' 中文处理

    Private Function BytesToBstr(body, Cset)
        Dim objstream
        Set objstream = CreateObject(COM_ADOSTREAM)
        objstream.Type = 1
        objstream.Mode = 3
        objstream.Open
        objstream.Write body
        objstream.Position = 0
        objstream.Type = 2
        objstream.Charset = Cset
        BytesToBstr = objstream.ReadText
        objstream.Close
        Set objstream = Nothing
    End Function

    ' 窃取目标URL地址的HTML代码/方法

    Public Sub steal(encode)
        If src_<>"" Then
            Dim Http
            Set Http = CreateObject(COM_HTTP)
            Http.Open "GET", src_ , false
            Http.send()
            'cookie = Http.getResponseHeader("Set-Cookie")
            If Http.readystate<>4 Then
                Exit Sub
            End If
            value_ = BytesToBSTR(Http.responseBody, encode)
            isGet_ = True
            Set http = Nothing
            If Err.Number<>0 Then Err.Clear
        Else
            response.Write("<script>alert(""请先设置src属性！"")</script>")
        End If
    End Sub

    ' 删除偷到的内容中里面的换行、回车符以便进一步加工/方法

    Public Sub noReturn()
        If isGet_ = false Then Call steal()
        value_ = Replace(Replace(value_ , vbCr, ""), vbLf, "")
    End Sub

    ' 对偷到的内容中的个别字符串用新值更换/方法
    ' 参数分别是旧字符串,新字符串
    Public Sub change(oldStr, Str) 
        If isGet_ = false Then Call steal()
        value_ = Replace(value_ , oldStr, Str)
    End Sub

    ' 按指定首尾字符串对偷取的内容进行裁减（不包括首尾字符串）/方法
    ' 参数分别是首字符串,尾字符串

    Public Sub cut(head, bot)
        If isGet_ = false Then Call steal()
        value_ = Mid(value_ , InStr(value_ , head) + Len(head), InStr(value_ , bot) - InStr(value_ , head) - Len(head))
    End Sub

    ' 按指定首尾字符串对偷取的内容进行裁减（包括首尾字符串）/方法
    ' 参数分别是首字符串,尾字符串

    Public Sub cutX(head, bot)
        If isGet_ = false Then Call steal()
        value_ = Mid(value_ , InStr(value_ , head), InStr(value_ , bot) - InStr(value_ , head) + Len(bot))
    End Sub

    '按指定首尾字符串位置偏移指针对偷取的内容进行裁减/方法
    '参数分别是首字符串,首偏移值,尾字符串,尾偏移值,左偏移用负值,偏移指针单位为字符数

    Public Sub cutBy(head, headCusor, bot, botCusor)
        If isGet_ = false Then Call steal()
        value_ = Mid(value_ , InStr(value_ , head) + Len(head) + headCusor, InStr(value_ , bot) -1 + botCusor - InStr(value_ , head) - Len(head) - headcusor)
    End Sub

    '按指定首尾字符串对偷取的内容用新值进行替换（不包括首尾字符串）/方法
    '参数分别是首字符串,尾字符串,新值,新值位空则为过滤

    Public Sub filt(head, bot, Str)
        If isGet_ = false Then Call steal()
        value_ = Replace(value_, Mid(value_ , InStr(value_ , head) + Len(head), InStr(value_ , bot) -1), Str)
    End Sub

    '按指定首尾字符串对偷取的内容用新值进行替换（包括首尾字符串）/方法
    '参数分别是首字符串,尾字符串,新值,新值为空则为过滤

    Public Sub filtX(head, bot, Str)
        If isGet_ = false Then Call steal()
        value_ = Replace(value_, Mid(value_ , InStr(value_ , head), InStr(value_ , bot) + Len(bot) -1), Str)
    End Sub

    '按指定首尾字符串位置偏移指针对偷取的内容新值进行替换/方法
    '参数分别是首字符串,首偏移值,尾字符串,尾偏移值,新值,左偏移用负值,偏移指针单位为字符数,新值为空则为过滤

    Public Sub filtBy(head, headCusor, bot, botCusor, Str)

        If isGet_ = false Then Call steal()
        value_ = Replace(value_ , Mid(value_ , InStr(value_ , head) + Len(head) + headCusor, InStr(value_ , bot) -1 + botCusor - InStr(value_ , head) - Len(head) - headcusor), Str)
    End Sub

    '将偷取的内容中的绝对URL地址改为本地相对地址

    Public Sub local()
        Dim tempReg
        Set tempReg = New RegExp
        tempReg.IgnoreCase = true
        tempReg.Global = true
        tempReg.Pattern = "^(http|https|ftp):(\/\/|////)(\w+.)+(com|net|org|cc|tv|cn|biz|com.cn|net.cn|sh.cn)\/"
        value_ = tempReg.Replace(value_ , "")
        Set tempReg = Nothing
    End Sub

    '对偷到的内容中的符合正则表达式的字符串用新值进行替换/方法
    '参数是你自定义的正则表达式,新值

    Public Sub replaceByReg(patrn, Str)
        If isGet_ = false Then Call steal()
        Dim tempReg
        Set tempReg = New RegExp
        tempReg.IgnoreCase = true
        tempReg.Global = true
        tempReg = patrn
        value_ = tempReg.Replace(value_ , Str)
        Set tempReg = Nothing
    End Sub

    '应用正则表达式对符合条件的内容进行分块采集并组合,最终内容为以<!--lkstar-->隔断的大文本/方法
    '通过属性value得到此内容后你可以用split(value,"<!--lkstar-->")得到你需要的数组
    '参数是你自定义的正则表达式

    Public Sub pickByReg(patrn)
        If isGet_ = false Then Call steal()
        Dim tempReg, match, matches, content
        Set tempReg = New RegExp
        tempReg.IgnoreCase = true
        tempReg.Global = true
        tempReg = patrn
        Set matches = tempReg.Execute(value_)
        For Each match in matches
            content = content&match.Value&"<!--lkstar-->"
        Next
        value_ = content
        Set matches = Nothing
        Set tempReg = Nothing
    End Sub

    '类排错模式——在类释放之前应用此方法可以随时查看你截获的内容HTML代码和页面显示效果/方法

    Public Sub debug()
        Dim tempstr
        tempstr = "<SCRIPT>function runEx(){var winEx2 = window.open("""", ""winEx2"", ""width=500,height=300,status=yes,menubar=no,scrollbars=yes,resizable=yes""); winEx2.document.open(""text/html"", ""replace""); winEx2.document.write(unescape(event.srcElement.parentElement.children[0].value)); winEx2.document.close(); }function saveFile(){var win=window.open('','','top=10000,left=10000');win.document.write(document.all.asdf.innerText);win.document.execCommand('SaveAs','','javascript.htm');win.close();}</SCRIPT><center><TEXTAREA id=asdf name=textfield rows=32  wrap=VIRTUAL cols=""120"">"&value_&"</TEXTAREA><BR><BR><INPUT name=Button onclick=runEx() type=button value=""查看效果"">&nbsp;&nbsp;<INPUT name=Button onclick=asdf.select() type=button value=""全选"">&nbsp;&nbsp;<INPUT name=Button onclick=""asdf.value=''"" type=button value=""清空"">&nbsp;&nbsp;<INPUT onclick=saveFile(); type=button value=""保存代码""></center>"
        'response.Write(tempstr)
        document.Write tempstr
    End Sub

End Class

'------------------------------------------------
' Vector Class
Class Vector
    Private myStack
    Private myCount

    Private Sub Class_Initialize()
        Redim myStack(8)
        myCount = -1
    End Sub

    Private Sub Class_Terminate()
    End Sub

    Public Property Let Dimension(pDim)
        Redim myStack(pDim)
    End Property

    Public Property Get Count()
        Count = myCount + 1
    End Property

    Public Sub Push(pElem)
        myCount = myCount + 1
        If (UBound(myStack) < myCount) Then
            Redim Preserve myStack(UBound(myStack) * 2)
        End If
        Call SetElementAt(myCount, pElem)
    End Sub

    Public Function Pop()
        If IsObject(myStack(myCount)) Then
            Set Pop = myStack(myCount)
        Else
            Pop = myStack(myCount)
        End If
        myCount = myCount - 1
    End Function

    Public Function Top()
        If IsObject(myStack(myCount)) Then
            Set Top = myStack(myCount)
        Else
            Top = myStack(myCount)
        End If
    End Function

    Public Function ElementAt(pIndex)
        If IsObject(myStack(pIndex)) Then
            Set ElementAt = myStack(pIndex)
        Else
            ElementAt = myStack(pIndex)
        End If
    End Function

    Public Sub SetElementAt(pIndex, pValue)
        If IsObject(pValue) Then
            Set myStack(pIndex) = pValue
        Else
            myStack(pIndex) = pValue
        End If
    End Sub

    Public Sub RemoveElementAt(pIndex)
        Do While pIndex < myCount
            Call SetElementAt(pIndex, ElementAt(pIndex + 1))
            pIndex = pIndex + 1
        Loop
        myCount = myCount - 1
    End Sub

    Public Function IsEmpty()
        IsEmpty = (myCount < 0)
    End Function
End Class

'------------------------------------------------
' CLogger Class
' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

'! Create an error message with hexadecimal error number from the given Err
'! object's properties. Formatted messages will look like "Foo bar (0xDEAD)".
'!
'! Implemented as a global function due to general lack of class methods in
'! VBScript.
'!
'! @param  e   Err object
'! @return Formatted error message consisting of error description and
'!         hexadecimal error number. Empty string if neither error description
'!         nor error number are available.
Public Function FormatErrorMessage(e)
  Dim re : Set re = New RegExp
  re.Global = True
  re.Pattern = "\s+"
  FormatErrorMessage = Trim(Trim(re.Replace(e.Description, " ")) & " (0x" & Hex(e.Number) & ")")
End Function

'! Create an error message with decimal error number from the given Err
'! object's properties. Formatted messages will look like "Foo bar (42)".
'!
'! Implemented as a global function due to general lack of class methods in
'! VBScript.
'!
'! @param  e   Err object
'! @return Formatted error message consisting of error description and
'!         decimal error number. Empty string if neither error description
'!         nor error number are available.
Public Function FormatErrorMessageDec(e)
  Dim re : Set re = New RegExp
  re.Global = True
  re.Pattern = "\s+"
  FormatErrorMessage = Trim(Trim(re.Replace(e.Description, " ")) & " (" & e.Number & ")")
End Function

'! Class for abstract logging to one or more logging facilities. Valid
'! facilities are:
'!
'! - interactive desktop/console
'! - log file
'! - eventlog
'!
'! Note that this class does not do any error handling at all. Taking care of
'! errors is entirely up to the calling script.
'!
'! @author  Ansgar Wiechers <ansgar.wiechers@planetcobalt.net>
'! @date    2011-03-13
'! @version 2.0
Class CLogger
	Private validLogLevels
	Private logToConsoleEnabled
	Private logToFileEnabled
	Private logFileName
	Private logFileHandle
	Private overwriteFile
	Private sep
	Private logToEventlogEnabled
	Private sh
	Private addTimestamp
	Private debugEnabled
	Private vbsDebug

	'! Enable or disable logging to desktop/console. Depending on whether the
	'! script is run via wscript.exe or cscript.exe, the message is either
	'! displayed as a MsgBox() popup or printed to the console. This facility
	'! is enabled by default when the script is run interactively.
	'!
	'! Console output is printed to StdOut for Info and Debug messages, and to
	'! StdErr for Warning and Error messages.
	Public Property Get LogToConsole
		LogToConsole = logToConsoleEnabled
	End Property

	Public Property Let LogToConsole(ByVal enable)
		logToConsoleEnabled = CBool(enable)
	End Property

	'! Indicates whether logging to a file is enabled or disabled. The log file
	'! facility is disabled by default. To enable it, set the LogFile property
	'! to a non-empty string.
	'!
	'! @see #LogFile
	Public Property Get LogToFile
		LogToFile = logToFileEnabled
	End Property

	'! Enable or disable logging to a file by setting or unsetting the log file
	'! name. Logging to a file ie enabled by setting this property to a non-empty
	'! string, and disabled by setting it to an empty string. If the file doesn't
	'! exist, it will be created automatically. By default this facility is
	'! disabled.
	'!
	'! Note that you MUST set the property Overwrite to False BEFORE setting
	'! this property to prevent an existing file from being overwritten!
	'!
	'! @see #Overwrite
	Public Property Get LogFile
		LogFile = logFileName
	End Property

	Public Property Let LogFile(ByVal filename)
		Dim fso, ioMode

		filename = Trim(Replace(filename, vbTab, " "))
		If filename = "" Then
			' Close a previously opened log file.
			If Not logFileHandle Is Nothing Then
				logFileHandle.Close
				Set logFileHandle = Nothing
			End If
			logToFileEnabled = False
		Else
			Set fso = CreateObject("Scripting.FileSystemObject")
			filename = fso.GetAbsolutePathName(filename)
			If logFileName <> filename Then
				' Close a previously opened log file.
				If Not logFileHandle Is Nothing Then logFileHandle.Close

				If overwriteFile Then
					ioMode = 2  ' open for (over)writing
				Else
					ioMode = 8  ' open for appending
				End If

				' Open log file either as ASCII or Unicode, depending on system settings.
				Set logFileHandle = fso.OpenTextFile(filename, ioMode, -2)

				logToFileEnabled = True
			End If
			Set fso = Nothing
		End If

		logFileName = filename
	End Property

	'! Enable or disable overwriting of log files. If disabled, log messages
	'! will be appended to an already existing log file (this is the default).
	'! The property affects only logging to a file and is ignored by all other
	'! facilities.
	'!
	'! Note that changes to this property will not affect already opened log
	'! files until they are re-opened.
	'!
	'! @see #LogFile
	Public Property Get Overwrite
		Overwrite = overwriteFile
	End Property

	Public Property Let Overwrite(ByVal enable)
		overwriteFile = CBool(enable)
	End Property

	'! Separate the fields of log file entries with the given character. The
	'! default is to use tabulators. This property affects only logging to a
	'! file and is ignored by all other facilities.
	'!
	'! @raise  Separator must be a single character (5)
	'! @see http://msdn.microsoft.com/en-us/library/xe43cc8d (VBScript Run-time Errors)
	Public Property Get Separator
		Separator = sep
	End Property

	Public Property Let Separator(ByVal char)
		If Len(char) <> 1 Then
			Err.Raise 5, WScript.ScriptName, "Separator must be a single character."
		Else
			sep = char
		End If
	End Property

	'! Enable or disable logging to the Eventlog. If enabled, messages are
	'! logged to the Application Eventlog. By default this facility is enabled
	'! when the script is run non-interactively, and disabled when the script
	'! is run interactively.
	'!
	'! Logging messages to this facility produces eventlog entries with source
	'! WSH and one of the following IDs:
	'! - Debug:       ID 0 (SUCCESS)
	'! - Error:       ID 1 (ERROR)
	'! - Warning:     ID 2 (WARNING)
	'! - Information: ID 4 (INFORMATION)
	Public Property Get LogToEventlog
		LogToEventlog = logToEventlogEnabled
	End Property

	Public Property Let LogToEventlog(ByVal enable)
		logToEventlogEnabled = CBool(enable)
		If sh Is Nothing And logToEventlogEnabled Then
			Set sh = CreateObject("WScript.Shell")
		ElseIf Not (sh Is Nothing Or logToEventlogEnabled) Then
			Set sh = Nothing
		End If
	End Property

	'! Enable or disable timestamping of log messages. If enabled, the current
	'! date and time is logged with each log message. The default is to not
	'! include timestamps. This property has no effect on Eventlog logging,
	'! because eventlog entries are always timestamped anyway.
	Public Property Get IncludeTimestamp
		IncludeTimestamp = addTimestamp
	End Property

	Public Property Let IncludeTimestamp(ByVal enable)
		addTimestamp = CBool(enable)
	End Property

	'! Enable or disable debug logging. If enabled, debug messages (i.e.
	'! messages passed to the LogDebug() method) are logged to the enabled
	'! facilities. Otherwise debug messages are silently discarded. This
	'! property is disabled by default.
	Public Property Get Debug
		Debug = debugEnabled
	End Property

	Public Property Let Debug(ByVal enable)
		debugEnabled = CBool(enable)
	End Property

	' - Constructor/Destructor ---------------------------------------------------

	'! @brief Constructor.
	'!
	'! Initialize logger objects with default values, i.e. enable console
	'! logging when a script is run interactively or eventlog logging when
	'! it's run non-interactively, etc.
	Private Sub Class_Initialize()
		logToConsoleEnabled = WScript.Interactive

		logToFileEnabled = False
		logFileName = ""
		Set logFileHandle = Nothing
		overwriteFile = False
		sep = vbTab

		logToEventlogEnabled = Not WScript.Interactive

		Set sh = Nothing

		addTimestamp = False
		debugEnabled = False
		vbsDebug = &h0050

		Set validLogLevels = CreateObject("Scripting.Dictionary")
		validLogLevels.Add vbInformation, True
		validLogLevels.Add vbExclamation, True
		validLogLevels.Add vbCritical, True
		validLogLevels.Add vbsDebug, True
	End Sub

	'! @brief Destructor.
	'!
	'! Clean up when a logger object is destroyed, i.e. close file handles, etc.
	Private Sub Class_Terminate()
		If Not logFileHandle Is Nothing Then
			logFileHandle.Close
			Set logFileHandle = Nothing
			logFileName = ""
		End If

		Set sh = Nothing
	End Sub

	' ----------------------------------------------------------------------------

	'! An alias for LogInfo(). This method exists for convenience reasons.
	'!
	'! @param  msg   The message to log.
	'!
	'! @see #LogInfo(msg)
	Public Sub Log(msg)
		LogInfo msg
	End Sub

	'! Log message with log level "Information".
	'!
	'! @param  msg   The message to log.
	Public Sub LogInfo(msg)
		LogMessage msg, vbInformation
	End Sub

	'! Log message with log level "Warning".
	'!
	'! @param  msg   The message to log.
	Public Sub LogWarning(msg)
		LogMessage msg, vbExclamation
	End Sub

	'! Log message with log level "Error".
	'!
	'! @param  msg   The message to log.
	Public Sub LogError(msg)
		LogMessage msg, vbCritical
	End Sub

	'! Log message with log level "Debug". These messages are logged only if
	'! debugging is enabled, otherwise the messages are silently discarded.
	'!
	'! @param  msg   The message to log.
	'!
	'! @see #Debug
	Public Sub LogDebug(msg)
		If debugEnabled Then LogMessage msg, vbsDebug
	End Sub

	'! Log the given message with the given log level to all enabled facilities.
	'!
	'! @param  msg       The message to log.
	'! @param  logLevel  Logging level (Information, Warning, Error, Debug) of the message.
	'!
	'! @raise  Undefined log level (51)
	'! @see http://msdn.microsoft.com/en-us/library/xe43cc8d (VBScript Run-time Errors)
	Private Sub LogMessage(msg, logLevel)
		Dim tstamp, prefix

		If Not validLogLevels.Exists(logLevel) Then Err.Raise 51, _
			WScript.ScriptName, "Undefined log level '" & logLevel & "'."

		tstamp = Now
		prefix = ""

		' Log to facilite "Console". If the script is run with cscript.exe, messages
		' are printed to StdOut or StdErr, depending on log level. If the script is
		' run with wscript.exe, messages are displayed as MsgBox() pop-ups.
		If logToConsoleEnabled Then
			If InStr(LCase(WScript.FullName), "cscript") <> 0 Then
				If addTimestamp Then prefix = tstamp & vbTab
				Select Case logLevel
					Case vbInformation: WScript.StdOut.WriteLine prefix & msg
					Case vbExclamation: WScript.StdErr.WriteLine prefix & "Warning: " & msg
					Case vbCritical:    WScript.StdErr.WriteLine prefix & "Error: " & msg
					Case vbsDebug:      WScript.StdOut.WriteLine prefix & "DEBUG: " & msg
				End Select
			Else
				If addTimestamp Then prefix = tstamp & vbNewLine & vbNewLine
				If logLevel = vbsDebug Then
					MsgBox prefix & msg, vbOKOnly Or vbInformation, WScript.ScriptName & " (Debug)"
				Else
					MsgBox prefix & msg, vbOKOnly Or logLevel, WScript.ScriptName
				End If
			End If
		End If

		' Log to facility "Logfile".
		If logToFileEnabled Then
			If addTimestamp Then prefix = tstamp & sep
			Select Case logLevel
				Case vbInformation: logFileHandle.WriteLine prefix & "INFO" & sep & msg
				Case vbExclamation: logFileHandle.WriteLine prefix & "WARN" & sep & msg
				Case vbCritical:    logFileHandle.WriteLine prefix & "ERROR" & sep & msg
				Case vbsDebug:      logFileHandle.WriteLine prefix & "DEBUG" & sep & msg
			End Select
		End If

		' Log to facility "Eventlog".
		' Timestamps are automatically logged with this facility, so addTimestamp
		' can be ignored.
		If logToEventlogEnabled Then
			Select Case logLevel
				Case vbInformation: sh.LogEvent 4, msg
				Case vbExclamation: sh.LogEvent 2, msg
				Case vbCritical:    sh.LogEvent 1, msg
				Case vbsDebug:      sh.LogEvent 0, "DEBUG: " & msg
			End Select
		End If
	End Sub
End Class

'MsgBox "111"
'Call ForceCreateFolder("C:\d\e\f\g\h")
'Call ForceDeleteFolder("C:\d")
'Call CreateZip("results.zip", "dadfasd")
'Call ZipBy7Zip("results_01.zip", "222.txt")
'Call UnZipBy7Zip("results_01.zip", "C:\ddddd\dddd\ddd")
'Call ZipBy7Zip("resutls_02.zip", "dadfasd")
'Call ZipBy7Zip("files.zip", """*.txt""") 