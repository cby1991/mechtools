Option Explicit

'==========================================================================================
'  SWautoExport  v1.0.0
'  SolidWorks 一键导出：模型 -> STEP，工程图 -> PDF
'------------------------------------------------------------------------------------------
'  运行入口（工具 > 宏 > 运行，建议绑定快捷键或工具条按钮）：
'      SwAuto_ExportActiveDoc   导出当前打开的文档（连同它的工程图 / 被引用模型）
'      SwAuto_PickFiles         弹文件对话框，多选文件后排入队列
'      SwAuto_PickFolder        弹文件夹对话框，批量导出整个文件夹
'------------------------------------------------------------------------------------------
'  所有可调参数集中在【第 1 节 配置区】，改完直接保存即可，无需重新导入。
'  每次运行都会在输出目录生成 _导出清单.csv，并弹窗汇总（含"未找到工程图"的模型）。
'==========================================================================================

'---- 全局状态（单次运行内有效）----
Private m_swApp          As SldWorks.SldWorks
Private m_OutRoot        As String           ' 用户选择的输出根目录
Private m_StepDir        As String           ' STEP 实际输出目录
Private m_PdfDir         As String           ' PDF  实际输出目录
Private m_UsedNames      As Object           ' Dictionary：本次运行已占用的输出路径(大写)
Private m_Results        As Collection       ' 结果行：Array(类别, 源文件, 输出文件, 结果, 备注)
Private m_NoDrawing      As Collection       ' 未找到工程图的模型路径
Private m_OriginalActive As String           ' 运行前的活动文档标题，跑完还原

'---- 工程图索引（惰性构建，批量时只建一次）----
Private m_IdxByPath As Object                ' Dictionary：模型全路径(大写) -> 工程图全路径
Private m_IdxByName As Object                ' Dictionary：模型文件名(大写) -> 工程图全路径
Private m_IdxBuilt  As Boolean


'==========================================================================================
'  第 1 节  配置区  ——  按需修改
'==========================================================================================

'---- 1.1 输出格式 ----
Private Const CFG_STEP_AP       As Long = 203       ' STEP 协议：203 或 214（swStepAP 直接接受这两个值）
Private Const CFG_STEP_EXT      As String = ".STEP"
Private Const CFG_PDF_EXT       As String = ".PDF"

'---- 1.2 输出目录结构 ----
Private Const CFG_USE_SUBFOLDER As Boolean = True   ' True：输出目录下自动建 STEP\ 与 PDF\ 两个子目录
Private Const CFG_OVERWRITE     As Boolean = True   ' True：覆盖同名文件；False：跳过并在清单里标注

'---- 1.3 工程图查找 ----
Private Const CFG_FIND_DRAWING   As Boolean = True  ' 模型导出后，是否自动找它的工程图一起导出 PDF
Private Const CFG_DRW_ALSO_MODEL As Boolean = True  ' 处理工程图时，是否把它引用的模型也导出 STEP
Private Const CFG_DEEP_SEARCH    As Boolean = True  ' 同名找不到时，扫描目录反查引用关系（慢但准）
Private Const CFG_DEEP_LIMIT     As Long = 1500     ' 深搜最多扫描多少个工程图，防超大目录卡死
Private Const CFG_EXTRA_DIRS     As String = ""     ' 额外搜索目录，多个用 ; 分隔，如 "D:\图纸库"

'---- 1.4 文件夹模式 ----
Private Const CFG_RECURSE_SUBDIR As Boolean = False ' 是否递归子目录

'---- 1.5 其它 ----
Private Const CFG_WRITE_REPORT  As Boolean = True   ' 在输出目录生成 _导出清单.csv
Private Const CFG_USE_PS_DIALOG As Boolean = True   ' 多选文件用 PowerShell 原生对话框
Private Const CFG_QUIET_OPEN    As Boolean = True   ' 静默打开文档（不弹配置 / 参考丢失对话框）

'---- 1.6 上次目录记忆（存于 HKCU\Software\VB and VBA Program Settings\SWauto\Export）----
Private Const CFG_APP     As String = "SWauto"
Private Const CFG_SEC     As String = "Export"
Private Const CFG_KEY_OUT As String = "LastOutputDir"
Private Const CFG_KEY_IN  As String = "LastInputDir"

Private Function Cfg_GetSetting(ByVal sKey As String) As String
    On Error Resume Next
    Cfg_GetSetting = GetSetting(CFG_APP, CFG_SEC, sKey, "")
End Function

Private Sub Cfg_PutSetting(ByVal sKey As String, ByVal sVal As String)
    On Error Resume Next
    SaveSetting CFG_APP, CFG_SEC, sKey, sVal
End Sub


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

Private Function U_Ext(ByVal sPath As String) As String
    Dim s As String, i As Long
    s = U_FileName(sPath)
    i = InStrRev(s, ".")
    If i > 0 Then U_Ext = UCase$(Mid$(s, i)) Else U_Ext = ""
End Function

Private Function U_FileExists(ByVal sPath As String) As Boolean
    On Error Resume Next
    If Len(sPath) = 0 Then Exit Function
    U_FileExists = (Len(Dir$(sPath, vbNormal)) > 0)
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

'---- 2.4 文件夹选择框（默认定位到 sDefault，即"上次使用的目录"）----
Private Function U_PickFolder(ByVal sTitle As String, ByVal sDefault As String) As String
    Dim sh As Object, fld As Object
    Dim sOld As String, sRes As String
    On Error Resume Next
    sOld = CurDir
    If U_DirExists(sDefault) Then
        If Mid$(sDefault, 2, 1) = ":" Then        ' UNC 不做 ChDir，会报错
            ChDrive Left$(sDefault, 1)
            ChDir sDefault
        End If
    End If
    Set sh = CreateObject("Shell.Application")
    Set fld = sh.BrowseForFolder(0, sTitle, &H41, 0)   ' 1 = 只返回目录, 64 = 新式对话框（可新建文件夹）
    If Not fld Is Nothing Then
        sRes = fld.Self.Path
        If Len(sRes) = 0 Then sRes = fld.Items().Item().Path    ' 新式对话框下 Self.Path 可能为空
    End If
    If Len(sOld) > 0 Then ChDir sOld
    On Error GoTo 0
    U_PickFolder = U_NormPath(sRes)
End Function

'---- 2.5 文件多选框（PowerShell 原生对话框，支持多选 + 记住初始目录）----
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
          "$d.Title = '选择要导出的 SolidWorks 文件（可多选）';" & _
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


'==========================================================================================
'  第 3 节  工程图查找（复刻 SolidWorks「打开工程图」的查找顺序）
'          1) 模型同目录下的同名工程图
'          2) CFG_EXTRA_DIRS 里的同名工程图
'          3) 扫描目录，读文件内的引用关系反查（GetDocumentDependencies2 不需要打开文件，很快）
'==========================================================================================

Private Sub F_InitIndex()
    If m_IdxByPath Is Nothing Then
        Set m_IdxByPath = CreateObject("Scripting.Dictionary")
        m_IdxByPath.CompareMode = 1
        Set m_IdxByName = CreateObject("Scripting.Dictionary")
        m_IdxByName.CompareMode = 1
    End If
End Sub

Private Function F_ExtraDirs() As Variant
    Dim v As Variant, col As Collection, i As Long, a() As String
    Set col = New Collection
    If Len(Trim$(CFG_EXTRA_DIRS)) > 0 Then
        v = Split(CFG_EXTRA_DIRS, ";")
        For i = LBound(v) To UBound(v)
            If U_DirExists(Trim$(v(i))) Then col.Add U_NormPath(Trim$(v(i)))
        Next i
    End If
    If col.Count = 0 Then Exit Function
    ReDim a(col.Count - 1)
    For i = 1 To col.Count
        a(i - 1) = col(i)
    Next i
    F_ExtraDirs = a
End Function

' 入口：返回工程图全路径；找不到返回 ""，sNote 给出原因
Private Function F_Find(ByVal sModelPath As String, ByRef sNote As String) As String
    Dim sDir As String, sHit As String, vDirs As Variant, i As Long
    sNote = ""
    sDir = U_Parent(sModelPath)

    ' (1) 同目录同名
    sHit = U_ResolveCase(U_JoinPath(sDir, U_BaseName(sModelPath) & ".SLDDRW"))
    If Len(sHit) > 0 Then
        F_Find = sHit
        Exit Function
    End If

    ' (2) 额外搜索目录里的同名
    vDirs = F_ExtraDirs()
    If IsArray(vDirs) Then
        For i = LBound(vDirs) To UBound(vDirs)
            sHit = U_ResolveCase(U_JoinPath(vDirs(i), U_BaseName(sModelPath) & ".SLDDRW"))
            If Len(sHit) > 0 Then
                F_Find = sHit
                Exit Function
            End If
        Next i
    End If

    ' (3) 深搜引用关系
    If CFG_DEEP_SEARCH Then
        F_BuildIndex sDir
        sHit = F_Lookup(sModelPath)
        If Len(sHit) > 0 Then
            F_Find = sHit
            Exit Function
        End If
    End If

    sNote = "未找到对应工程图"
End Function

' 扫目录建索引：遍历工程图，读出它们引用的模型，反向映射 模型 -> 工程图
' 批量时只建一次，之后每次查找都是字典命中
Private Sub F_BuildIndex(ByVal sRootDir As String)
    Dim colDirs As Collection, vRoots As Variant, i As Long, n As Long, k As Long
    Dim fso As Object, fld As Object, f As Object, sf As Object
    Dim sDir As String
    If m_IdxBuilt Then Exit Sub
    F_InitIndex

    Set colDirs = New Collection
    If U_DirExists(sRootDir) Then colDirs.Add U_NormPath(sRootDir)
    vRoots = F_ExtraDirs()
    If IsArray(vRoots) Then
        For i = LBound(vRoots) To UBound(vRoots)
            colDirs.Add U_NormPath(vRoots(i))
        Next i
    End If

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    k = 1
    Do While k <= colDirs.Count
        If n >= CFG_DEEP_LIMIT Then Exit Do
        sDir = colDirs(k)
        k = k + 1
        If fso.FolderExists(sDir) Then
            Set fld = fso.GetFolder(sDir)
            For Each f In fld.Files
                If UCase$(f.Name) Like "*.SLDDRW" Then
                    n = n + 1
                    If n > CFG_DEEP_LIMIT Then Exit For
                    F_IndexOne f.Path
                End If
            Next f
            For Each sf In fld.SubFolders
                colDirs.Add U_NormPath(sf.Path)
            Next sf
        End If
    Loop
    m_IdxBuilt = True
End Sub

Private Sub F_IndexOne(ByVal sDrwPath As String)
    Dim vDeps As Variant, i As Long, sItem As String, sKey As String
    On Error Resume Next
    vDeps = m_swApp.GetDocumentDependencies2(sDrwPath, False, False, False)
    If Not IsArray(vDeps) Then Exit Sub
    For i = LBound(vDeps) To UBound(vDeps)
        sItem = CStr(vDeps(i))
        ' 该接口返回 (名称, 全路径) 成对出现，这里一并兼容纯路径数组
        If Len(sItem) > 0 Then
            If U_IsModel(sItem) And InStr(sItem, "\") > 0 Then
                If Not m_IdxByPath.Exists(UCase$(sItem)) Then
                    m_IdxByPath.Add UCase$(sItem), sDrwPath
                End If
                sKey = UCase$(U_FileName(sItem))
                If Not m_IdxByName.Exists(sKey) Then
                    m_IdxByName.Add sKey, sDrwPath
                End If
            End If
        End If
    Next i
End Sub

Private Function F_Lookup(ByVal sModelPath As String) As String
    Dim sKey As String
    If m_IdxByPath Is Nothing Then Exit Function
    sKey = UCase$(sModelPath)
    If m_IdxByPath.Exists(sKey) Then
        F_Lookup = m_IdxByPath(sKey)
        Exit Function
    End If
    sKey = UCase$(U_FileName(sModelPath))
    If m_IdxByName.Exists(sKey) Then F_Lookup = m_IdxByName(sKey)
End Function

' 反过来：一个工程图引用了哪些模型（用于"工程图 -> 模型"模式）
Private Function F_ModelsOfDrawing(ByVal sDrwPath As String) As Collection
    Dim vDeps As Variant, i As Long, sItem As String
    Dim col As New Collection
    On Error Resume Next
    Set F_ModelsOfDrawing = col
    vDeps = m_swApp.GetDocumentDependencies2(sDrwPath, False, False, False)
    If Not IsArray(vDeps) Then Exit Function
    For i = LBound(vDeps) To UBound(vDeps)
        sItem = CStr(vDeps(i))
        If U_IsModel(sItem) And InStr(sItem, "\") > 0 Then
            If U_FileExists(sItem) Then
                If Not F_Contains(col, sItem) Then col.Add sItem
            End If
        End If
    Next i
End Function

Private Function F_Contains(ByVal col As Collection, ByVal sVal As String) As Boolean
    Dim i As Long
    For i = 1 To col.Count
        If StrComp(col(i), sVal, vbTextCompare) = 0 Then
            F_Contains = True
            Exit Function
        End If
    Next i
End Function


'==========================================================================================
'  第 4 节  导出核心（STEP / PDF）
'==========================================================================================

'---- 4.1 打开 / 关闭文档 ----
Private Function X_FindOpenDoc(ByVal sPath As String) As Object
    Dim vDocs As Variant, i As Long, d As Object
    On Error Resume Next
    vDocs = m_swApp.GetDocuments
    If IsArray(vDocs) Then
        For i = LBound(vDocs) To UBound(vDocs)
            Set d = vDocs(i)
            If Not d Is Nothing Then
                If StrComp(d.GetPathName, sPath, vbTextCompare) = 0 Then
                    Set X_FindOpenDoc = d
                    Exit Function
                End If
            End If
        Next i
    End If
    On Error GoTo 0
End Function

' 已在内存里的文档直接复用；否则静默打开，bOurs=True 表示是我们打开的（跑完要关掉）
Private Function X_Open(ByVal sPath As String, ByVal nType As Long, ByRef bOurs As Boolean) As Object
    Dim d As Object, lErr As Long, lWarn As Long
    bOurs = False
    Set d = X_FindOpenDoc(sPath)
    If Not d Is Nothing Then
        Set X_Open = d
        Exit Function
    End If
    On Error Resume Next
    If CFG_QUIET_OPEN Then
        Set d = m_swApp.OpenDoc6(sPath, nType, swOpenDocOptions_e.swOpenDocOptions_Silent, "", lErr, lWarn)
    Else
        Set d = m_swApp.OpenDoc6(sPath, nType, 0, "", lErr, lWarn)
    End If
    If Not d Is Nothing Then bOurs = True
    Set X_Open = d
End Function

Private Sub X_Close(ByVal sPath As String, ByVal bOurs As Boolean)
    Dim d As Object
    If Not bOurs Then Exit Sub
    On Error Resume Next
    Set d = X_FindOpenDoc(sPath)
    If Not d Is Nothing Then m_swApp.CloseDoc d.GetTitle
End Sub

'---- 4.2 模型 -> STEP ----
Private Function X_ToStep(ByVal d As Object, ByVal sOut As String, ByRef sMsg As String) As Boolean
    Dim ext As Object, lErr As Long, lWarn As Long, bRet As Boolean, sNote As String
    On Error GoTo EH

    ' STEP 协议版本：swStepAP 直接接受 203 / 214
    m_swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swStepAP, CFG_STEP_AP
    If m_swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swStepAP) <> CFG_STEP_AP Then
        sNote = "AP" & CFG_STEP_AP & " 未能写入用户选项，已按当前选项导出"
    End If

    ' 装配体先把轻化零件还原，否则可能导出空几何
    If d.GetType = swDocumentTypes_e.swDocASSEMBLY Then
        On Error Resume Next
        d.ResolveAllLightWeightComponents True
        Err.Clear
        On Error GoTo EH
    End If

    Set ext = d.Extension
    bRet = ext.SaveAs(sOut, swSaveAsVersion_e.swSaveAsCurrentVersion, _
                      swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, lErr, lWarn)
    If bRet And lErr = 0 Then
        X_ToStep = True
        sMsg = sNote
    Else
        sMsg = "SaveAs 失败（err=" & lErr & " / warn=" & lWarn & "）"
    End If
    Exit Function
EH:
    sMsg = "异常 " & Err.Number & ": " & Err.Description
End Function

'---- 4.3 工程图 -> PDF ----
Private Function X_ToPdf(ByVal d As Object, ByVal sOut As String, ByRef sMsg As String) As Boolean
    Dim pdfData As Object, lErr As Long, lWarn As Long, bRet As Boolean
    On Error GoTo EH
    Set pdfData = m_swApp.GetExportFileData(swExportDataFileType_e.swExportPDFData)
    If pdfData Is Nothing Then
        sMsg = "无法获取 PDF 导出设置对象"
        Exit Function
    End If
    On Error Resume Next
    pdfData.ExportAsSinglePDF = True        ' 多张图纸合并成一个 PDF
    pdfData.ViewPdfAfterSaving = False      ' 导出后不自动打开阅读器
    Err.Clear
    On Error GoTo EH

    bRet = d.Extension.SaveAs(sOut, swSaveAsVersion_e.swSaveAsCurrentVersion, _
                              swSaveAsOptions_e.swSaveAsOptions_Silent, pdfData, lErr, lWarn)
    If bRet And lErr = 0 Then
        X_ToPdf = True
    Else
        sMsg = "SaveAs 失败（err=" & lErr & " / warn=" & lWarn & "）"
    End If
    Exit Function
EH:
    sMsg = "异常 " & Err.Number & ": " & Err.Description
End Function

'---- 4.4 处理一个模型文件：导出 STEP ----
Private Sub X_DoModel(ByVal sPath As String)
    Dim d As Object, bOurs As Boolean, sOut As String, sMsg As String
    If Not U_FileExists(sPath) Then
        U_AddResult "模型", sPath, "", "失败", "源文件不存在"
        Exit Sub
    End If
    sOut = U_TargetPath(sPath, CFG_STEP_EXT, m_StepDir)
    If Len(sOut) = 0 Then
        U_AddResult "模型", sPath, U_JoinPath(m_StepDir, U_BaseName(sPath) & CFG_STEP_EXT), "跳过", "输出文件已存在"
        Exit Sub
    End If
    Set d = X_Open(sPath, U_DocType(sPath), bOurs)
    If d Is Nothing Then
        U_AddResult "模型", sPath, sOut, "失败", "无法打开文档（请检查参考引用是否齐全）"
        Exit Sub
    End If
    If X_ToStep(d, sOut, sMsg) Then
        U_AddResult "模型", sPath, sOut, "成功", sMsg
    Else
        U_AddResult "模型", sPath, sOut, "失败", sMsg
    End If
    X_Close sPath, bOurs
End Sub

'---- 4.5 处理一个工程图文件：导出 PDF ----
Private Sub X_DoDrawing(ByVal sPath As String)
    Dim d As Object, bOurs As Boolean, sOut As String, sMsg As String
    If Not U_FileExists(sPath) Then
        U_AddResult "图纸", sPath, "", "失败", "源文件不存在"
        Exit Sub
    End If
    sOut = U_TargetPath(sPath, CFG_PDF_EXT, m_PdfDir)
    If Len(sOut) = 0 Then
        U_AddResult "图纸", sPath, U_JoinPath(m_PdfDir, U_BaseName(sPath) & CFG_PDF_EXT), "跳过", "输出文件已存在"
        Exit Sub
    End If
    Set d = X_Open(sPath, swDocumentTypes_e.swDocDRAWING, bOurs)
    If d Is Nothing Then
        U_AddResult "图纸", sPath, sOut, "失败", "无法打开工程图（请检查其引用的模型是否可访问）"
        Exit Sub
    End If
    If X_ToPdf(d, sOut, sMsg) Then
        U_AddResult "图纸", sPath, sOut, "成功", sMsg
    Else
        U_AddResult "图纸", sPath, sOut, "失败", sMsg
    End If
    X_Close sPath, bOurs
End Sub


'==========================================================================================
'  第 5 节  入口程序
'    SwAuto_ExportActiveDoc   导出当前打开的文档
'    SwAuto_PickFiles         弹文件对话框，多选文件
'    SwAuto_PickFolder        弹文件夹对话框，批量导出
'==========================================================================================

Public Sub SwAuto_ExportActiveDoc()
    Dim d As Object, sPath As String, col As Collection
    If Not M_Init() Then Exit Sub
    Set d = m_swApp.ActiveDoc
    If d Is Nothing Then
        MsgBox "当前没有打开的文档。", vbExclamation, "SWauto"
        M_Cleanup
        Exit Sub
    End If
    sPath = d.GetPathName
    If Len(sPath) = 0 Then
        MsgBox "当前文档还没保存到磁盘，请先保存。", vbExclamation, "SWauto"
        M_Cleanup
        Exit Sub
    End If
    Set col = New Collection
    col.Add sPath
    M_Run col, "当前文档：" & U_FileName(sPath)
    M_Cleanup
End Sub

Public Sub SwAuto_PickFiles()
    Dim v As Variant, col As Collection, i As Long
    If Not M_Init() Then Exit Sub
    v = U_PickFiles(Cfg_GetSetting(CFG_KEY_IN))
    If IsEmpty(v) Then
        M_Cleanup
        Exit Sub
    End If
    Set col = New Collection
    For i = LBound(v) To UBound(v)
        If U_IsModel(v(i)) Or U_IsDrawing(v(i)) Then
            If Not F_Contains(col, v(i)) Then col.Add v(i)
        End If
    Next i
    If col.Count = 0 Then
        MsgBox "没有选中可处理的 SolidWorks 文件（.sldprt / .sldasm / .slddrw）。", vbExclamation, "SWauto"
        M_Cleanup
        Exit Sub
    End If
    Cfg_PutSetting CFG_KEY_IN, U_Parent(col(1))
    M_Run col, "已选 " & col.Count & " 个文件"
    M_Cleanup
End Sub

Public Sub SwAuto_PickFolder()
    Dim sDir As String, col As Collection
    If Not M_Init() Then Exit Sub
    sDir = U_PickFolder("请选择要导出的源文件夹", Cfg_GetSetting(CFG_KEY_IN))
    If Len(sDir) = 0 Then
        M_Cleanup
        Exit Sub
    End If
    Cfg_PutSetting CFG_KEY_IN, sDir
    Set col = M_ScanFolder(sDir)
    If col.Count = 0 Then
        MsgBox "该文件夹里没有找到 .sldprt / .sldasm / .slddrw 文件。" & vbCrLf & _
               "（如需包含子目录，请把 CFG_RECURSE_SUBDIR 改为 True）", vbExclamation, "SWauto"
        M_Cleanup
        Exit Sub
    End If
    M_Run col, "文件夹：" & sDir
    M_Cleanup
End Sub

'---- 5.1 主流程 ----
Private Sub M_Run(ByVal colPaths As Collection, ByVal sDesc As String)
    Dim colModels As Collection, colDrawings As Collection, colDep As Collection
    Dim i As Long, j As Long, sPath As String, sNote As String, sDrw As String

    ' (1) 输出目录：默认上次使用的目录
    m_OutRoot = M_AskOutDir(U_Parent(colPaths(1)))
    If Len(m_OutRoot) = 0 Then Exit Sub              ' 用户取消
    If Not U_MakeDir(m_OutRoot) Then
        MsgBox "无法创建输出目录：" & m_OutRoot, vbCritical, "SWauto"
        Exit Sub
    End If
    If CFG_USE_SUBFOLDER Then
        m_StepDir = U_JoinPath(m_OutRoot, "STEP")
        m_PdfDir = U_JoinPath(m_OutRoot, "PDF")
        U_MakeDir m_StepDir
        U_MakeDir m_PdfDir
    Else
        m_StepDir = m_OutRoot
        m_PdfDir = m_OutRoot
    End If

    ' (2) 分类去重
    Set colModels = New Collection
    Set colDrawings = New Collection
    For i = 1 To colPaths.Count
        sPath = colPaths(i)
        If U_IsModel(sPath) Then
            If Not F_Contains(colModels, sPath) Then colModels.Add sPath
        ElseIf U_IsDrawing(sPath) Then
            If Not F_Contains(colDrawings, sPath) Then colDrawings.Add sPath
        End If
    Next i

    ' (3) 确认
    If MsgBox("来源：" & sDesc & vbCrLf & vbCrLf & _
              "模型：" & colModels.Count & " 个  →  导出 STEP (AP" & CFG_STEP_AP & ")" & vbCrLf & _
              "工程图：" & colDrawings.Count & " 个  →  导出 PDF" & IIf(CFG_FIND_DRAWING, vbCrLf & _
              "（模型还会自动查找对应工程图一起导出）", "") & vbCrLf & vbCrLf & _
              "输出目录：" & m_OutRoot & vbCrLf & vbCrLf & "开始导出？", _
              vbYesNo + vbQuestion, "SWauto 导出确认") <> vbYes Then
        Exit Sub
    End If

    ' (4) 为每个模型查找对应工程图
    If CFG_FIND_DRAWING Then
        For i = 1 To colModels.Count
            sDrw = F_Find(colModels(i), sNote)
            If Len(sDrw) > 0 Then
                If Not F_Contains(colDrawings, sDrw) Then colDrawings.Add sDrw
            Else
                m_NoDrawing.Add colModels(i)
            End If
        Next i
    End If

    ' (5) 工程图反查它引用的模型
    If CFG_DRW_ALSO_MODEL Then
        For i = 1 To colDrawings.Count
            Set colDep = F_ModelsOfDrawing(colDrawings(i))
            For j = 1 To colDep.Count
                If Not F_Contains(colModels, colDep(j)) Then colModels.Add colDep(j)
            Next j
        Next i
    End If

    ' (6) 执行
    For i = 1 To colModels.Count
        X_DoModel colModels(i)
    Next i
    For i = 1 To colDrawings.Count
        X_DoDrawing colDrawings(i)
    Next i

    ' (7) 清单 + 汇总
    If CFG_WRITE_REPORT Then M_WriteReport
    M_RestoreActive
    M_Summary
End Sub

'---- 5.2 输出目录选择 ----
Private Function M_AskOutDir(ByVal sSeed As String) As String
    Dim sDef As String
    sDef = Cfg_GetSetting(CFG_KEY_OUT)
    If Not U_DirExists(sDef) Then sDef = sSeed
    If Not U_DirExists(sDef) Then sDef = ""
    M_AskOutDir = U_PickFolder("请选择导出文件的输出目录", sDef)
    If Len(M_AskOutDir) > 0 Then Cfg_PutSetting CFG_KEY_OUT, M_AskOutDir
End Function

'---- 5.3 扫描文件夹 ----
Private Function M_ScanFolder(ByVal sDir As String) As Collection
    Dim col As New Collection, colDirs As New Collection
    Dim fso As Object, fld As Object, f As Object, sf As Object
    Dim k As Long, s As String
    Set M_ScanFolder = col
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    colDirs.Add U_NormPath(sDir)
    k = 1
    Do While k <= colDirs.Count
        s = colDirs(k)
        k = k + 1
        If fso.FolderExists(s) Then
            Set fld = fso.GetFolder(s)
            For Each f In fld.Files
                If InStr(f.Name, "~$") = 0 Then
                    If U_IsModel(f.Path) Or U_IsDrawing(f.Path) Then col.Add f.Path
                End If
            Next f
            If CFG_RECURSE_SUBDIR Then
                For Each sf In fld.SubFolders
                    colDirs.Add U_NormPath(sf.Path)
                Next sf
            End If
        End If
    Loop
End Function

'---- 5.4 导出清单 ----
Private Sub M_WriteReport()
    Dim sb As String, i As Long, it As Variant
    sb = "SWauto 导出清单" & vbCrLf
    sb = sb & U_Csv("导出时间") & "," & U_Csv(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & vbCrLf
    sb = sb & U_Csv("输出目录") & "," & U_Csv(m_OutRoot) & vbCrLf
    sb = sb & U_Csv("STEP 协议") & "," & U_Csv("AP" & CStr(CFG_STEP_AP)) & vbCrLf
    sb = sb & U_Csv("模型 STEP 成功") & "," & U_CountBy("模型", "成功") & vbCrLf
    sb = sb & U_Csv("模型 STEP 失败") & "," & U_CountBy("模型", "失败") & vbCrLf
    sb = sb & U_Csv("模型 STEP 跳过") & "," & U_CountBy("模型", "跳过") & vbCrLf
    sb = sb & U_Csv("工程图 PDF 成功") & "," & U_CountBy("图纸", "成功") & vbCrLf
    sb = sb & U_Csv("工程图 PDF 失败") & "," & U_CountBy("图纸", "失败") & vbCrLf
    sb = sb & U_Csv("工程图 PDF 跳过") & "," & U_CountBy("图纸", "跳过") & vbCrLf
    sb = sb & U_Csv("未找到工程图的模型") & "," & m_NoDrawing.Count & vbCrLf
    sb = sb & vbCrLf
    sb = sb & "序号,类别,源文件,输出文件,结果,备注" & vbCrLf
    For i = 1 To m_Results.Count
        it = m_Results(i)
        sb = sb & i & "," & U_Csv(it(0)) & "," & U_Csv(it(1)) & "," & _
             U_Csv(it(2)) & "," & U_Csv(it(3)) & "," & U_Csv(it(4)) & vbCrLf
    Next i
    sb = sb & vbCrLf & U_Csv("以下模型未找到对应工程图") & vbCrLf
    If m_NoDrawing.Count = 0 Then
        sb = sb & U_Csv("（无）") & vbCrLf
    Else
        For i = 1 To m_NoDrawing.Count
            sb = sb & U_Csv(m_NoDrawing(i)) & vbCrLf
        Next i
    End If
    U_WriteText U_JoinPath(m_OutRoot, "_导出清单.csv"), sb
End Sub

'---- 5.5 汇总弹窗 ----
Private Sub M_Summary()
    Dim s As String, i As Long, nFail As Long, it As Variant
    s = "导出完成" & vbCrLf & vbCrLf
    s = s & "输出目录：" & m_OutRoot & vbCrLf & vbCrLf
    s = s & "模型 → STEP：" & U_CountBy("模型", "成功") & " 成功"
    If U_CountBy("模型", "失败") > 0 Then s = s & "，" & U_CountBy("模型", "失败") & " 失败"
    If U_CountBy("模型", "跳过") > 0 Then s = s & "，" & U_CountBy("模型", "跳过") & " 跳过"
    s = s & vbCrLf
    s = s & "工程图 → PDF：" & U_CountBy("图纸", "成功") & " 成功"
    If U_CountBy("图纸", "失败") > 0 Then s = s & "，" & U_CountBy("图纸", "失败") & " 失败"
    If U_CountBy("图纸", "跳过") > 0 Then s = s & "，" & U_CountBy("图纸", "跳过") & " 跳过"

    s = s & vbCrLf & vbCrLf
    If m_NoDrawing.Count = 0 Then
        s = s & "所有模型都找到了对应工程图。"
    Else
        s = s & "未找到对应工程图（" & m_NoDrawing.Count & " 个）：" & vbCrLf & _
            U_BriefList(m_NoDrawing, 8)
    End If

    nFail = U_CountBy("模型", "失败") + U_CountBy("图纸", "失败")
    If nFail > 0 Then
        s = s & vbCrLf & vbCrLf & "失败明细：" & vbCrLf
        i = 0
        For Each it In m_Results
            If it(3) = "失败" Then
                i = i + 1
                If i > 5 Then
                    s = s & "…"
                    Exit For
                End If
                s = s & "· " & U_FileName(it(1)) & "：" & it(4) & vbCrLf
            End If
        Next it
    End If

    If CFG_WRITE_REPORT Then
        s = s & vbCrLf & "明细清单：" & U_JoinPath(m_OutRoot, "_导出清单.csv")
    End If
    MsgBox s, IIf(nFail > 0, vbExclamation, vbInformation), "SWauto"
End Sub

'---- 5.6 初始化 / 收尾 ----
Private Function M_Init() As Boolean
    On Error Resume Next
    Set m_swApp = Application.SldWorks
    If m_swApp Is Nothing Then Set m_swApp = CreateObject("SldWorks.Application")
    On Error GoTo 0
    If m_swApp Is Nothing Then
        MsgBox "无法连接 SolidWorks。", vbCritical, "SWauto"
        Exit Function
    End If
    On Error Resume Next
    Set m_UsedNames = CreateObject("Scripting.Dictionary")
    m_UsedNames.CompareMode = 1
    On Error GoTo 0
    Set m_Results = New Collection
    Set m_NoDrawing = New Collection
    Set m_IdxByPath = Nothing
    Set m_IdxByName = Nothing
    m_IdxBuilt = False
    m_OutRoot = ""
    m_StepDir = ""
    m_PdfDir = ""
    m_OriginalActive = ""
    If Not m_swApp.ActiveDoc Is Nothing Then m_OriginalActive = m_swApp.ActiveDoc.GetTitle
    M_Init = True
End Function

Private Sub M_Cleanup()
    Set m_Results = Nothing
    Set m_NoDrawing = Nothing
    Set m_IdxByPath = Nothing
    Set m_IdxByName = Nothing
    Set m_UsedNames = Nothing
End Sub

' 还原用户原来的活动文档，避免批量导出后界面停在最后一个文件上
Private Sub M_RestoreActive()
    Dim lErr As Long
    If Len(m_OriginalActive) = 0 Then Exit Sub
    On Error Resume Next
    m_swApp.ActivateDoc3 m_OriginalActive, False, 0, lErr    ' 0 = 激活时不重建
    On Error GoTo 0
End Sub
