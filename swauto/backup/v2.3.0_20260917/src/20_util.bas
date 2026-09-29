
'==========================================================================================
'  第 2 节  通用工具（路径 / 文件 / 对话框 / 结果记录）
'==========================================================================================

'---- 2.1 路径与文件名 ----
Private Function U_NormPath(ByVal sPath As String) As String
    Dim s As String
    s = Trim$(sPath)
    Do While Len(s) > 3 And Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop
    U_NormPath = s
End Function

Private Function U_JoinPath(ByVal sDir As String, ByVal sName As String) As String
    Dim s As String
    s = U_NormPath(sDir)
    If Len(s) = 0 Then
        U_JoinPath = sName
        Exit Function
    End If
    If Right$(s, 1) <> "\" Then s = s & "\"
    U_JoinPath = s & sName
End Function

Private Function U_FileName(ByVal sPath As String) As String
    Dim i As Long
    i = InStrRev(sPath, "\")
    If i = 0 Then U_FileName = sPath Else U_FileName = Mid$(sPath, i + 1)
End Function

Private Function U_Parent(ByVal sPath As String) As String
    Dim i As Long
    i = InStrRev(sPath, "\")
    If i <= 0 Then Exit Function
    If i = 3 Then
        U_Parent = Left$(sPath, 3)          ' 盘根，如 D:\
    Else
        U_Parent = Left$(sPath, i - 1)
    End If
End Function

Private Function U_BaseName(ByVal sPath As String) As String
    Dim s As String, i As Long
    s = U_FileName(sPath)
    i = InStrRev(s, ".")
    If i > 1 Then s = Left$(s, i - 1)
    U_BaseName = s
End Function

' 把字符串里 Windows 文件名不允许的字符换成下划线。
' 用途：拿"来源名"（文档名 / 文件夹名）去拼清单文件名时兜一下底。
' 正常来源名本来就取自磁盘上的真实名字，理论上合法；这里是防御性处理，
' 万一来源名里有路径分隔符之类的字符，不至于拼出一个建不出来的文件名。
Private Function U_SafeName(ByVal s As String) As String
    Dim s1 As String
    s1 = s
    s1 = Replace(s1, "\", "_")
    s1 = Replace(s1, "/", "_")
    s1 = Replace(s1, ":", "_")
    s1 = Replace(s1, "*", "_")
    s1 = Replace(s1, "?", "_")
    s1 = Replace(s1, """", "_")
    s1 = Replace(s1, "<", "_")
    s1 = Replace(s1, ">", "_")
    s1 = Replace(s1, "|", "_")
    U_SafeName = Trim$(s1)
End Function

Private Function U_Ext(ByVal sPath As String) As String
    Dim s As String, i As Long
    s = U_FileName(sPath)
    i = InStrRev(s, ".")
    If i > 0 Then U_Ext = UCase$(Mid$(s, i)) Else U_Ext = ""
End Function

' 是否绝对路径（D:\... 或 \\server\share\...）。
' 关键用途：GetDocumentDependencies2 返回的数组是「名称 + 全路径」成对出现，
' 而"名称"那一项在某些保存方式下会带 "\"（例如 ".\零件.SLDPRT"、"..\零件.SLDPRT"）。
' 如果只判断"含不含 \"，这种名称就会被当成一条独立路径收进来，
' 同一个零件于是被收录两次 → 导出成 零件.STEP 和 零件_2.STEP。必须挡掉。
Private Function U_IsAbsPath(ByVal sPath As String) As Boolean
    If Len(sPath) < 3 Then Exit Function
    If Mid$(sPath, 2, 1) = ":" Then U_IsAbsPath = True      ' D:\...
    If Left$(sPath, 2) = "\\" Then U_IsAbsPath = True       ' UNC \\server\share
End Function

' 把一个路径压成"同一文件只有一个写法"的 key，用于去重。
' 统一分隔符 + 去掉末尾反斜杠 + 转大写，这样 D:\a\X.SLDPRT 与 D:\A\x.sldprt 会被认成同一个。
Private Function U_FileKey(ByVal sPath As String) As String
    U_FileKey = UCase$(U_NormPath(Trim$(sPath)))
End Function

' 改用 FileSystemObject 判断存在性：
'   原来的 Dir$ 会把路径里的 "[" 当成通配符，零件名带方括号时会被误判成"文件不存在"
Private Function U_FileExists(ByVal sPath As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    If Len(sPath) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    U_FileExists = fso.FileExists(sPath)
End Function

Private Function U_DirExists(ByVal sPath As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    If Len(sPath) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    U_DirExists = fso.FolderExists(sPath)
End Function

' 返回磁盘上的真实大小写路径，便于清单里显示
Private Function U_ResolveCase(ByVal sPath As String) As String
    Dim sDir As String, sName As String, sFound As String
    If Not U_FileExists(sPath) Then Exit Function
    If InStr(sPath, "\") = 0 Then
        U_ResolveCase = sPath
        Exit Function
    End If
    sDir = Left$(sPath, InStrRev(sPath, "\"))
    sName = Mid$(sPath, Len(sDir) + 1)
    sFound = Dir$(sDir & sName, vbNormal)
    If Len(sFound) > 0 Then U_ResolveCase = sDir & sFound Else U_ResolveCase = sPath
End Function

Private Function U_IsModel(ByVal sPath As String) As Boolean
    Dim e As String
    e = U_Ext(sPath)
    U_IsModel = (e = ".SLDPRT" Or e = ".SLDASM")
End Function

Private Function U_IsDrawing(ByVal sPath As String) As Boolean
    U_IsDrawing = (U_Ext(sPath) = ".SLDDRW")
End Function

Private Function U_DocType(ByVal sPath As String) As Long
    Select Case U_Ext(sPath)
        Case ".SLDPRT": U_DocType = swDocumentTypes_e.swDocPART
        Case ".SLDASM": U_DocType = swDocumentTypes_e.swDocASSEMBLY
        Case ".SLDDRW": U_DocType = swDocumentTypes_e.swDocDRAWING
        Case Else:      U_DocType = -1
    End Select
End Function

' 逐级创建目录，支持 UNC；已存在返回 True
Private Function U_MakeDir(ByVal sPath As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    sPath = U_NormPath(sPath)
    If Len(sPath) = 0 Then Exit Function
    If fso.FolderExists(sPath) Then
        U_MakeDir = True
        Exit Function
    End If
    If Len(sPath) <= 3 Then Exit Function          ' 连盘根都不在，放弃
    U_MakeDir = U_MakeDir(fso.GetParentFolderName(sPath))
    If U_MakeDir Then fso.CreateFolder sPath
    U_MakeDir = fso.FolderExists(sPath)
End Function

' 给一个源文件算出输出路径；若配置为不覆盖且已存在，返回 "" 表示跳过
Private Function U_TargetPath(ByVal sSrcPath As String, ByVal sExt As String, ByVal sOutDir As String) As String
    Dim sBase As String, sTry As String, i As Long
    sBase = U_BaseName(sSrcPath)
    sTry = U_JoinPath(sOutDir, sBase & sExt)
    If (Not CFG_OVERWRITE) And U_FileExists(sTry) Then
        U_TargetPath = ""
        Exit Function
    End If
    ' 本次运行内的重名（不同目录的同名文件）自动加序号，避免互相覆盖
    If m_UsedNames.Exists(UCase$(sTry)) Then
        For i = 2 To 999
            sTry = U_JoinPath(sOutDir, sBase & "_" & i & sExt)
            If Not m_UsedNames.Exists(UCase$(sTry)) Then Exit For
        Next i
    End If
    m_UsedNames.Add UCase$(sTry), True
    U_TargetPath = sTry
End Function

'---- 2.2 文本 / CSV ----
Private Function U_Csv(ByVal s As String) As String
    U_Csv = """" & Replace(s, """", """""") & """"
End Function

Private Function U_WriteText(ByVal sFile As String, ByVal sText As String) As Boolean
    Dim st As Object, h As Integer
    On Error Resume Next
    Set st = CreateObject("ADODB.Stream")
    If Not st Is Nothing Then
        st.Type = 2                 ' adTypeText
        st.Charset = "utf-8"
        st.Open
        st.WriteText sText
        st.SaveToFile sFile, 2      ' adSaveCreateOverWrite
        st.Close
    End If
    If Err.Number = 0 And U_FileExists(sFile) Then
        U_WriteText = True
        Exit Function
    End If
    ' 回退：ANSI 直写
    Err.Clear
    h = FreeFile
    Open sFile For Output As #h
    Print #h, sText
    Close #h
    U_WriteText = U_FileExists(sFile)
End Function

Private Function U_ReadLines(ByVal sFile As String) As Variant
    Dim fso As Object, ts As Object, col As Collection, s As String
    Dim a() As String, i As Long
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(sFile, 1, False, 0)
    If ts Is Nothing Then Exit Function
    Set col = New Collection
    Do While Not ts.AtEndOfStream
        s = ts.ReadLine
        If Left$(s, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then s = Mid$(s, 4)   ' 去 BOM
        s = Trim$(s)
        If Len(s) > 0 Then col.Add s
    Loop
    ts.Close
    If col.Count = 0 Then Exit Function
    ReDim a(col.Count - 1)
    For i = 1 To col.Count
        a(i - 1) = col(i)
    Next i
    U_ReadLines = a
End Function

'---- 2.3 外部进程 ----
' 把字符串包成 PowerShell 单引号字面量（内部单引号翻倍）
Private Function U_PsQuote(ByVal s As String) As String
    U_PsQuote = "'" & Replace(s, "'", "''") & "'"
End Function

Private Function U_ShellRunWait(ByVal sCmd As String, ByVal nStyle As Long) As Long
    Dim wsh As Object
    On Error Resume Next
    Set wsh = CreateObject("WScript.Shell")
    U_ShellRunWait = wsh.Run(sCmd, nStyle, True)
End Function

'  注：文件夹选择框（U_PickFolder）已挪到第 2.5 节，改用 Windows 通用对话框。

'---- 2.4 文件多选框（PowerShell 原生对话框，支持多选 + 记住初始目录）----
Private Function U_PickFiles(ByVal sDefaultDir As String) As Variant
    Dim sTmp As String, sPs As String, sCmd As String, sInit As String
    Dim v As Variant
    U_PickFiles = Empty
    If Not CFG_USE_PS_DIALOG Then Exit Function

    sTmp = Environ$("TEMP")
    If Len(sTmp) = 0 Then sTmp = Environ$("TMP")
    If Len(sTmp) = 0 Then Exit Function
    sTmp = U_JoinPath(sTmp, "swauto_selection.txt")
    On Error Resume Next
    If Len(Dir$(sTmp, vbNormal)) > 0 Then Kill sTmp
    Err.Clear

    If U_DirExists(sDefaultDir) Then
        sInit = "$d.InitialDirectory = " & U_PsQuote(sDefaultDir) & ";"
    End If
    sPs = "Add-Type -AssemblyName System.Windows.Forms;" & _
          "$o = New-Object System.Windows.Forms.Form;" & _
          "$o.TopMost = $true; $o.ShowInTaskbar = $false; $o.FormBorderStyle = 'None';" & _
          "$o.Width = 1; $o.Height = 1; $o.StartPosition = 'Manual';" & _
          "$o.Left = -3000; $o.Top = -3000; $o.Show();" & _
          "$d = New-Object System.Windows.Forms.OpenFileDialog;" & _
          "$d.Multiselect = $true;" & _
          "$d.Title = '【2/3 · 自选文件】请选择要导出的 SolidWorks 文件（可多选）';" & _
          "$d.Filter = 'SolidWorks 文件|*.sldprt;*.sldasm;*.slddrw|所有文件 (*.*)|*.*';" & _
          sInit & _
          "if ($d.ShowDialog($o) -eq [System.Windows.Forms.DialogResult]::OK) {" & _
          "$d.FileNames | Set-Content -LiteralPath " & U_PsQuote(sTmp) & " -Encoding Default };" & _
          "$o.Close()"
    sCmd = "powershell.exe -NoProfile -STA -WindowStyle Hidden -Command """ & sPs & """"
    U_ShellRunWait sCmd, 0

    If Len(Dir$(sTmp, vbNormal)) = 0 Then Exit Function
    v = U_ReadLines(sTmp)
    On Error Resume Next
    Kill sTmp
    U_PickFiles = v
End Function

'---- 2.6 结果记录 ----
Private Sub U_AddResult(ByVal sKind As String, ByVal sSrc As String, _
                       ByVal sOut As String, ByVal sState As String, ByVal sNote As String)
    m_Results.Add Array(sKind, sSrc, sOut, sState, sNote)
End Sub

Private Function U_CountBy(ByVal sKind As String, ByVal sState As String) As Long
    Dim i As Long, it As Variant, n As Long
    For i = 1 To m_Results.Count
        it = m_Results(i)
        If it(0) = sKind And it(3) = sState Then n = n + 1
    Next i
    U_CountBy = n
End Function

Private Function U_BriefList(ByVal col As Collection, ByVal nMax As Long) As String
    Dim i As Long, s As String
    For i = 1 To col.Count
        If i > nMax Then
            s = s & "… 等共 " & col.Count & " 个"
            Exit For
        End If
        If Len(s) > 0 Then s = s & "，"
        s = s & U_FileName(col(i))
    Next i
    U_BriefList = s
End Function

'---- 2.7 关键字匹配（标准件过滤用）----
' 分三种情况，目的都是"该命中的命中、不该命中的别误伤"：
'
'   ① 关键字含中文        → 直接子串匹配
'        中文不存在"词中词"问题；而「内六角螺钉M8x30」这类把尺寸代号紧贴在后面的命名极其常见，
'        加词边界反而会漏判（"螺钉"后面紧跟字母 M，会被当成词内）。
'
'   ② 关键字含 / - _ 空格 → 直接子串匹配
'        这类关键字自带边界，不可能藏在别的词中间：GB/T、GB-T、Spring Pin、O-Ring、Snap Ring。
'
'   ③ 其余（纯 ASCII 字母数字）→ 词边界匹配
'        前一个字符不能是字母数字；后一个字符不能是字母（数字放行）。
'        · 挡住：Spindle 不命中 Pin、Walnut 不命中 Nut、ISOLATION 不命中 ISO、
'                Bolted 不命中 Bolt、Screwdriver 不命中 Screw、Nutshell 不命中 Nut
'        · 放行：GBT5782、ISO4762、Nut8 —— 关键字紧贴数字的写法照样能命中
'
' 改动这条规则前，先跑 python check.py，里面有一节"标准件关键字规则抽检"会立刻告诉你有没有踩坑。
Private Function U_HitKeyword(ByVal sHay As String, ByVal sNeedle As String) As Boolean
    If Len(sHay) = 0 Or Len(sNeedle) = 0 Then Exit Function
    If U_HasNonAscii(sNeedle) Or U_HasSeparator(sNeedle) Then
        U_HitKeyword = (InStr(1, sHay, sNeedle, vbTextCompare) > 0)
    Else
        U_HitKeyword = U_HitBounded(sHay, sNeedle)
    End If
End Function

' 情况③：纯字母数字关键字，按"前边界严格、后边界放行数字"匹配
Private Function U_HitBounded(ByVal sHay As String, ByVal sNeedle As String) As Boolean
    Dim p As Long, n As Long, cL As String, cR As String
    If Len(sHay) = 0 Or Len(sNeedle) = 0 Then Exit Function
    n = Len(sNeedle)
    p = InStr(1, sHay, sNeedle, vbTextCompare)
    Do While p > 0
        cL = ""
        cR = ""
        If p > 1 Then cL = Mid$(sHay, p - 1, 1)
        If p + n <= Len(sHay) Then cR = Mid$(sHay, p + n, 1)
        If Not U_IsAsciiAlnum(cL) And Not U_IsAsciiAlpha(cR) Then
            U_HitBounded = True
            Exit Function
        End If
        p = InStr(p + 1, sHay, sNeedle, vbTextCompare)
    Loop
End Function

Private Function U_IsAsciiAlnum(ByVal c As String) As Boolean
    Dim u As Long
    If Len(c) = 0 Then Exit Function
    u = AscW(c)
    If u < 0 Then u = u + 65536            ' VBA 的 AscW 对 U+8000 以上返回负数
    U_IsAsciiAlnum = ((u >= 48 And u <= 57) Or (u >= 65 And u <= 90) Or (u >= 97 And u <= 122))
End Function

Private Function U_IsAsciiAlpha(ByVal c As String) As Boolean
    Dim u As Long
    If Len(c) = 0 Then Exit Function
    u = AscW(c)
    If u < 0 Then u = u + 65536
    U_IsAsciiAlpha = ((u >= 65 And u <= 90) Or (u >= 97 And u <= 122))
End Function

Private Function U_HasNonAscii(ByVal s As String) As Boolean
    Dim i As Long, u As Long
    For i = 1 To Len(s)
        u = AscW(Mid$(s, i, 1))
        If u < 0 Then u = u + 65536
        If u > 126 Then
            U_HasNonAscii = True
            Exit Function
        End If
    Next i
End Function

' 含非字母数字的 ASCII 字符（/ - _ 空格 . 等）—— 这种关键字自带边界
Private Function U_HasSeparator(ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        If Not U_IsAsciiAlnum(Mid$(s, i, 1)) Then
            U_HasSeparator = True
            Exit Function
        End If
    Next i
End Function

' 把 CFG_* 里逗号分隔的关键字表逐个去空
Private Function U_KeywordList(ByVal sCsv As String) As Variant
    Dim v As Variant, i As Long
    Dim col As Collection, a() As String
    Set col = New Collection
    If Len(Trim$(sCsv)) > 0 Then
        v = Split(sCsv, ",")
        For i = LBound(v) To UBound(v)
            If Len(Trim$(v(i))) > 0 Then col.Add Trim$(v(i))
        Next i
    End If
    If col.Count = 0 Then Exit Function
    ReDim a(col.Count - 1)
    For i = 1 To col.Count
        a(i - 1) = col(i)
    Next i
    U_KeywordList = a
End Function
