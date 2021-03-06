Option Explicit
Dim myArg
Dim i

Dim myDate
Dim myTime

If  WScript.Arguments.Count > 0 Then
    myDate  =  Replace(Date,"/","")
    myTime  =  Replace(Time,":","")
    For i = 0 To WScript.Arguments.Count - 1
        Call RemoveBom( WScript.Arguments(i) )
    Next
    WScript.Echo("処理が終了しました。")
Else
    WScript.Echo("ドラッグ・ドロップされたファイルを" & _
    "処理します（複数可）")
End If
WScript.Quit

Sub RemoveBOM(INfile)
    Dim OUTfile
    Dim fileName
    
    Dim dirPos
    Dim pathOUT
    Dim OUTName
    
    Dim oADOST_R     'As Object
    Dim readPos      'As Long ' or Currency or Double
    
    Dim oADOST_W     'As Object
    
    Dim dathead01
    Dim dathead02
    Dim dathead03
    
    dirPos   = InStrRev(INfile,"\")
    pathOUT  = Left(INfile,dirPos)
    fileName = Mid(INfile , dirPos + 1 , 999)
    
    OUTName  = myDate & myTime & "_NoBom_" & fileName
    OUTfile  = pathOUT & OUTName
    
    Set oADOST_R = CreateObject("ADODB.Stream")
    
    oADOST_R.Type = 1   '1=adTypeBinary 2=adTypeText
    oADOST_R.Open
    oADOST_R.LoadFromFile INfile
    readPos = 0
    oADOST_R.Position = readPos '読込開始位置
    
    dathead01  = UCase(Right("0" & Hex(AscB(oADOST_R.Read(1) )) , 2)) 
    readPos = readPos + 1
    dathead02  = UCase(Right("0" & Hex(AscB(oADOST_R.Read(1) )) , 2))  
    readPos = readPos + 1
    dathead03  = UCase(Right("0" & Hex(AscB(oADOST_R.Read(1) )) , 2)) 
    
    If  ( dathead01 = "FF" And dathead02 = "FE" ) Or _
        ( dathead01 = "FE" And dathead02 = "FF" ) Then
        readPos = 2  'utf-16
    ElseIf ( dathead01 = "EF"   And _
        dathead02 = "BB"   And _
        dathead03 = "BF" ) Then
        readPos = 3  'utf-8
    Else
        readPos = 0  'BOM無し
    End If
    oADOST_R.Position = readPos '読込開始位置
    
    '書込Object設定
    Set oADOST_W = CreateObject("ADODB.Stream")
    
    oADOST_W.Type = 1 '1=adTypeBinary 2=adTypeText
    'oADOST_W.Charset = "iso-8859-1" 'キャラクタセット＝Latin-1
    oADOST_W.Open
    
    'Do Until oADOST_R.EOS = True
    '   oADOST_W.Write oADOST_R.Read(1)
    '   readPos = readPos + 1
    '   oADOST_R.Position = readPos
    'Loop
    
    oADOST_W.Write oADOST_R.Read() '省略時、全部読込
    
    '既にファイルが存在する場合　1=実行時エラー、2=上書保存
    oADOST_W.SaveToFile OUTfile, 2
    
    oADOST_R.Close
    oADOST_W.Close
    Set oADOST_R = Nothing
    Set oADOST_W = Nothing
End Sub

