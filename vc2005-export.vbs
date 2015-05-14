'------------------------------------------------
' \file vc2005-export.vbs
' \brief 用于vc2005工程导出
' 
' 分析vc2005工程文件，获取文件依赖，并创建include依赖目录，拷贝必要文件
' 
' \author wangyang 
' \date 2015/03/11 
' \version 0.7
' 
' Example:
'   cscript.exe vc2005-export.vbs "×××××.vcproj" "E:\111\src\test\×××××\"
'------------------------------------------------

Dim lib : Set lib = New LibLoader
lib.path = "../lib"
lib.Import "cactus.vbs"

Dim objShell : Set objShell = wscript.createObject("wscript.shell")
objShell.CurrentDirectory = get_project_dir()

Dim Dict : Set Dict = CreateObject("Scripting.Dictionary")

Call Main()

Set Dict = Nothing
Set myLog = Nothing
Set objShell = Nothing


Function get_project_dir()
    get_project_dir = GetFilePath(WScript.Arguments(0))    
End Function

Function get_project_file()
    get_project_file = WScript.Arguments(0)
End Function

Function repace_file_path(filename)
    repace_file_path = replace(filename, "/", "\")
End Function

Function get_dest_dir()
    get_dest_dir = DisposePath(WScript.Arguments(1))
End Function

Function Main()
    Dim ProjectFileDir	    
    ProjectFileDir = get_project_dir()   


    Dim doc
    Dim root 
    Dim hfiles
    Dim cfiles
    Dim I

    Set doc = CreateObject("MSXML2.DOMDocument")
    doc.async = False
    doc.load(get_project_file())
    If doc.parseError.errorCode = 0 Then
        Set root = doc.documentElement     

        Set cfiles = root.selectNodes("//File")
        If Not (cfiles is Nothing) Then		
            Dim ext, filename2, dirs, dir         
            For I = 0 To cfiles.length-1
                On Error Resume Next                                  
                filename2 = repace_file_path(ProjectFileDir & cfiles(I).getAttribute("RelativePath"))
                set fnMyFunction = GetRef("testfile")                
                Call each_files_matches(filename2, "\#include\s+[\""\<](.*?)[\""\>]", fnMyFunction, root)                        
            Next
        End If

    End If

    Call CopyFile(get_project_file(), get_dest_dir())

    Set doc = Nothing     

End Function

Function get_dirs(filename, ByRef root)
    Dim dirs(), I, dirs2
    Redim Preserve dirs(0)
    dirs(0) = GetFilePath(filename)
    dirs2 = GetIncludeDirs(root)
    For I = 0 to Ubound(dirs2)         
        Redim Preserve dirs(I + 1)
        If InStr(dirs2(I), ":\") > 0 Then
            dirs(I + 1) =  DisposePath(dirs2(I))              
        Else
            dirs(I + 1) = DisposePath(get_project_dir() & dirs2(I))
        End If
    Next
    
    get_dirs = dirs
End function



Function GetIncludeDirs(ByRef root)
    Dim dir
    Set include_node = root.selectSingleNode("//Configuration[@Name=""Release|Win32""]//Tool[@Name=""VCCLCompilerTool""]")
    If Not (include_node is Nothing) Then        
        dir = include_node.getAttribute("AdditionalIncludeDirectories") 	
    End If
    GetIncludeDirs = Split(dir, ";")    
End Function

Sub each_files_matches(filename, pattern, method, ByRef root)  
    On Error Resume Next
    'filename = GetAbsolutePathName(filename)
    If InStr(filename, ":\") > 0 Then        
        Dim dest_filename : dest_filename = Replace(filename, get_project_dir(), get_dest_dir())        
        dest_filename = GetAbsolutePathName(dest_filename)        
        Dim dest_dir : dest_dir = GetFilePath(dest_filename)
        Call ForceCreateFolder(dest_dir) 
        filename = GetAbsolutePathName(filename)
        Echo "原始文件:" & filename
        Echo "目的文件:" & dest_filename
        Call CopyFile(filename, dest_dir)
        Dict.Add filename, filename    
    End If
    
    
    Dim content   
    content = ReadTextFile(filename)    

    Dim regex, matches, match
    set regex = New RegExp
    regex.IgnoreCase = False
    regex.Global = True
    regex.MultiLine = True

    regex.Pattern = pattern
    Set matches = regex.Execute(content)    
    Call method(matches, filename, root)  
    
End Sub

Function testfile(ByRef matches, filename, ByRef root)      
    Dim match, fullpath     
    For Each match In matches
        If match.SubMatches(0) <> "" Then                        
            fullpath = match.SubMatches(0)    
            Echo "-------------" & fullpath
            If has_file(filename, root, fullpath) Then                
                set fnMyFunction = GetRef("testfile")                
                Call each_files_matches(fullpath, "\#include\s+[\""\<](.*?)[\""\>]", fnMyFunction, root)
            Else
                Echo fullpath                
            End If            
        End If
    Next
End Function

Function has_file(filename, ByRef root, ByRef filename2)
    Dim dir, dirs, fullpath, retval
    retval = False
    dirs = get_dirs(filename, root)
    For Each dir in dirs        
        Echo dir
        fullpath = repace_file_path(dir & filename2)        
        If (Not Dict.Exists(fullpath)) and FileExists(fullpath) Then
            'Dict.Add fullpath, fullpath            
            filename2 = fullpath
            retval = True
            Exit For
        End If
    Next    
    has_file = retval
End Function




Class LibLoader    
    Private lib_dir_
    
    Private Sub Class_Initialize()
        Dim objShell
        lib_dir_ = left(Wscript.ScriptFullName,len(Wscript.ScriptFullName)-len(Wscript.ScriptName))        
        Set objShell = wscript.createObject("wscript.shell")
        objShell.CurrentDirectory = lib_dir_
    End Sub
    
    Private Sub Class_Terminate()        
    End Sub

    Public Property Get Path
        Path = lib_dir_
    End Property
    
    Public Property Let Path(value)
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        lib_dir_ = fso.GetAbsolutePathName(value)
        Set fso = Nothing
    End Property
    
    Public Function Import(ByVal filename) 
        Dim fso, sh, file, code, dir, basename

        ' Create my own objects, so the function is self-contained and can be called
        ' before anything else in the script.
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set sh = CreateObject("WScript.Shell")

        filename = Trim(sh.ExpandEnvironmentStrings(filename))
        If Not (Left(filename, 2) = "\\" Or Mid(filename, 2, 2) = ":\") Then
            ' filename is not absolute
            If Not fso.FileExists(fso.GetAbsolutePathName(filename)) Then
                If fso.FileExists(fso.BuildPath(lib_dir_, filename)) Then
                    filename = fso.BuildPath(lib_dir_, filename)                    
                End If                
            End If
            filename = fso.GetAbsolutePathName(filename)
        End If

        'WScript.Echo filename

        On Error Resume Next
        Set file = fso.GetFile(filename)
        basename = fso.GetBaseName(file)
        ExecuteGlobal "Const " & basename & "_vbs_loading = 1"
        If Err = 0 Then
            On Error Resume Next
            Set file = fso.OpenTextFile(filename, 1, False)
            If Err Then
                WScript.Echo "Cannot import '" & filename & "': " & Err.Description & " (0x" & Hex(Err.Number) & ")"
                WScript.Quit 1
            End If
            On Error Goto 0
            code = file.ReadAll
            file.Close
            ExecuteGlobal(code)        
        End If
        Set sh = Nothing
        Set fso = Nothing
    End Function    
End Class

