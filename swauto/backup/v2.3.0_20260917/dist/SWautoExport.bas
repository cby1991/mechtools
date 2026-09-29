Option Explicit

'==========================================================================================
'  SWautoExport  v2.2.0
'  SolidWorks 一键导出：模型 -> STEP，工程图 -> PDF
'------------------------------------------------------------------------------------------
'  ★ 平时只用这一个：SwAuto_ExportActiveDoc  （模式 1/3）
'      （打开要导出的零件/装配体 -> 运行它 -> 选输出目录 -> 完成）
'
'  运行入口（工具 > 宏 > 运行，建议绑定快捷键或工具条按钮）：
'      SwAuto_ExportActiveDoc   【模式 1/3 · 当前文档】【主入口】
'                                 零件   -> 本零件 STEP + 对应工程图 PDF
'                                 装配体 -> 只导它下面的零件 STEP + 各自工程图 PDF
'                                           （装配体/子装配体本身不出，配置 1.5 可改）
'                                           （螺钉螺母等标准件自动过滤，配置 1.6）
'                                           （隐藏/压缩/封套：弹勾选框逐类询问，配置 1.7）
'                                 工程图 -> 本图 PDF + 它引用的模型 STEP
'      SwAuto_PickFiles         【模式 2/3 · 自选文件】多选文件，想只导几个时用
'      SwAuto_PickFolder        【模式 3/3 · 文件夹批量】整个文件夹批量导出
'      SwAuto_Help              【使用说明】宏一览 + 快捷键绑定方法 + 当前打开的文档
'------------------------------------------------------------------------------------------
'  ！绑快捷键时的大坑（已经踩过一次）：
'     工具 > 自定义 > 键盘 > 类别选「宏」>「新建宏按钮」里有一个【方法】下拉框，
'     因为本模块含多个入口，**必须**把「方法」选成 SwAuto_ExportActiveDoc。
'     漏掉这一步，SolidWorks 会挑一个方法运行 → 按快捷键却弹出「文件夹批量模式」。
'     每个模式的所有弹窗标题都会写明自己是几号模式，运行起来一眼就能看出有没有绑错。
'------------------------------------------------------------------------------------------
'  所有可调参数集中在【第 1 节 配置区】，改完直接保存即可，无需重新导入。
'  每次运行都会在输出目录生成 _导出清单.csv，并弹窗汇总（含"未找到工程图"和被过滤的标准件）。
'==========================================================================================

'---- 全局状态（单次运行内有效）----
Private m_swApp          As SldWorks.SldWorks
Private m_OutRoot        As String           ' 用户选择的输出根目录
Private m_StepDir        As String           ' STEP 实际输出目录
Private m_PdfDir         As String           ' PDF  实际输出目录
Private m_UsedNames      As Object           ' Dictionary：本次运行已占用的输出路径(大写)
Private m_DoneSrc        As Object           ' Dictionary：源文件规范化路径 -> 已生成的输出文件（同一文件只导一次）
Private m_Results        As Collection       ' 结果行：Array(类别, 源文件, 输出文件, 结果, 备注)
Private m_NoDrawing      As Collection       ' 未找到工程图的模型路径
Private m_SkippedStd     As Collection       ' 被当作标准件过滤掉的零件路径
Private m_SkippedState   As Collection       ' 因"隐藏/压缩/封套"被跳过的组件标识（同一次运行内去重）
Private m_SkipWhy        As Object           ' Dictionary：被跳过的组件标识 -> 跳过原因（与上面同步维护）

'---- 装配体展开时，"不参与当前状态"的组件按类别暂存（等用户在勾选框里拍板）----
' 每个元素是 Array(显示标识, 磁盘路径, 原因文字)。磁盘路径为空 = 拿不到路径，勾了也导不出来。
' 之所以要"先存后定"：勾选框里要显示每一类的实际个数，而这个数只有扫完才知道；
' 用户勾选后又要把对应的那类放回待导出列表 —— 所以不能一边扫一边就把它写进"跳过"。
Private m_ListHidden     As Collection       ' 已隐藏
Private m_ListSuppressed As Collection       ' 已压缩
Private m_ListEnvelope   As Collection       ' 封套
Private m_SrcName        As String           ' 本次导出的"来源名"，用于给导出清单命名（文档名 / 文件夹名 / "多选文件"）
Private m_OriginalActive As String           ' 运行前的活动文档标题，跑完还原
Private m_ModeTag        As String           ' 当前模式的标识，如 "1/3 · 当前文档"；写进所有弹窗标题，防止绑错宏后看不出来

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
Private Const CFG_USE_SUBFOLDER As Boolean = False  ' False：STEP 与 PDF 都放进你选的同一个目录（推荐）
                                                    ' True ：在输出目录下自动建 STEP\ 与 PDF\ 两个子目录
Private Const CFG_OVERWRITE     As Boolean = True   ' True：覆盖同名文件；False：跳过并在清单里标注

'---- 1.3 工程图查找 ----
Private Const CFG_FIND_DRAWING   As Boolean = True  ' 模型导出后，是否自动找它的工程图一起导出 PDF
Private Const CFG_DRW_ALSO_MODEL As Boolean = True  ' 处理工程图时，是否把它引用的模型也导出 STEP
Private Const CFG_DEEP_SEARCH    As Boolean = True  ' 同名找不到时，扫描目录反查引用关系（慢但准）
Private Const CFG_DEEP_LIMIT     As Long = 1500     ' 深搜最多扫描多少个工程图，防超大目录卡死
Private Const CFG_EXTRA_DIRS     As String = ""     ' 额外搜索目录，多个用 ; 分隔，如 "D:\图纸库"

'---- 1.4 文件夹模式 ----
Private Const CFG_RECURSE_SUBDIR As Boolean = False ' 是否递归子目录

'---- 1.5 装配体展开（只影响 SwAuto_ExportActiveDoc 在"当前是装配体"时的行为）----
Private Const CFG_ASM_EXPAND    As Boolean = True   ' True ：装配体 → 展开导出它下面所有零件
                                                    ' False：装配体 → 只导装配体自己，不碰里面的零件
Private Const CFG_ASM_SELF      As Boolean = False  ' 展开时，装配体本身（以及各级子装配体）是否也各出一份
                                                    ' 默认 False：只交付零件 —— 装配体/子装配体自身
                                                    ' 以及它们的工程图都不导出（想要就改 True）

'---- 1.6 标准件过滤（螺钉、螺母、垫圈……自动跳过，只作用于"装配体展开"）----
'  判定规则（命中任意一条即算标准件）：
'    ① SolidWorks 组件自带的 Toolbox 标记（最准）
'    ② 文件路径里含 CFG_STD_PATH_HINTS
'    ③ 文件名里含 CFG_STD_KEYWORDS（按"词边界"匹配，不会把 Spindle 当成 Pin、Walnut 当成 Nut）
'  想放行某一类，把对应关键字从列表里删掉即可；整个功能不想用就把 CFG_EXCLUDE_STD_PARTS 改 False。
Private Const CFG_EXCLUDE_STD_PARTS As Boolean = True
Private Const CFG_STD_USE_API       As Boolean = True   ' 是否用 Toolbox 标记判定

' 文件名关键字（逗号分隔，中英文都可以加）
Private Const CFG_STD_KEYWORDS As String = _
    "螺钉,螺母,螺栓,螺柱,垫圈,垫片,挡圈,卡簧,弹性挡圈,开口销,销钉,圆柱销,圆锥销,铆钉,自攻," & _
    "Bolt,Nut,Washer,Screw,Stud,Rivet,Circlip,Snap Ring,Retaining Ring,O-Ring,Spring Pin"

' 国标 / 国际标准号（逗号分隔）——标准件文件名基本都带这些前缀
Private Const CFG_STD_STDCODES As String = "GB/T,GB-T,GBT,ISO,DIN,JIS,ANSI,BS-"

' 路径关键字（逗号分隔）——Toolbox 数据库的典型路径
Private Const CFG_STD_PATH_HINTS As String = "SOLIDWORKS Data,Toolbox,标准件,紧固件"

'---- 1.7 装配体展开时的"参与状态"过滤（隐藏 / 压缩 / 封套，只作用于"装配体展开"）----
'  这三种组件在装配体里都不参与当前的设计状态，导出它们只会污染交付物：
'    隐藏 —— 视觉上根本看不到（Component2::IsHidden）
'    压缩 —— 不参与重建、也不进 BOM（Component2::IsSuppressed）
'    封套 —— 只是空间占位的参考体，不是真实零件（Component2::IsEnvelope）
'
'  判定会沿"父级链"向上走：某个子装配体被隐藏/压缩时，它下面的零件同样跳过
'  （子零件自身往往是"显示"的，但实际跟着父级一起看不见）。
'
'  ！可见性随【当前激活的配置】变化 —— 这里过滤的是运行那一刻的实际状态。
'     同一个装配体换个配置再跑，结果可能不同，这是符合预期的。
'  ！轻化（Lightweight）不算：那只是加载方式，零件本身有效，照常导出。
Private Const CFG_SKIP_HIDDEN     As Boolean = True ' 跳过 已隐藏 的零件
Private Const CFG_SKIP_SUPPRESSED As Boolean = True ' 跳过 已压缩 的零件
Private Const CFG_SKIP_ENVELOPE   As Boolean = True ' 跳过 封套 零件
'  被跳过的组件不会静默消失：清单里记"跳过"并写明是哪种状态（含"上级装配体"字样）

Private Const CFG_ASK_SKIP_OPTIONS As Boolean = True
'  True ：装配体展开后弹一个勾选框，把上面三类**逐行列出**（带各自的实际个数），
'          由用户当场决定要哪几类。勾选框的**初始状态取自上面三个 CFG_SKIP_***
'          —— 比如 CFG_SKIP_HIDDEN = False 时，"也导出已隐藏"一进框就是勾上的。
'          用户在框里改的只影响这一次运行，不写回配置。
'  False：不询问，直接按上面三个 CFG_SKIP_* 的设定执行
'  只有装配体才会弹（零件 / 工程图没有"隐藏的零件"这回事）。
'  万一勾选框弹不出来（PowerShell 被安全策略拦了之类），会按上面的配置继续，不会卡住。

'---- 1.8 其它 ----
Private Const CFG_WRITE_REPORT  As Boolean = True   ' 是否在输出目录生成导出清单（CSV，Excel 可直接双击打开）

Private Const CFG_REPORT_WITH_SRC As Boolean = True ' 清单文件名带上"来源名"，便于一眼看出是哪个文档导出的
Private Const CFG_REPORT_SUFFIX   As String = "_导出清单"
'  清单文件名 = <来源名><CFG_REPORT_SUFFIX>.csv
'    来源名：当前文档模式 = 文档名（如 家庭服务机器人-重组）
'            文件夹模式   = 源文件夹名
'            多选文件模式 = "多选文件"
'  例：家庭服务机器人-重组_导出清单.csv
'  CFG_REPORT_WITH_SRC 改成 False 就回到原来的固定名 _导出清单.csv
Private Const CFG_USE_PS_DIALOG As Boolean = True   ' 多选文件用 PowerShell 原生对话框
Private Const CFG_QUIET_OPEN    As Boolean = True   ' 静默打开文档（不弹配置 / 参考丢失对话框）

Private Const CFG_MODERN_FOLDER_PICK As Boolean = True
'  选文件夹时用哪种对话框：
'    True ：Windows 通用对话框（和 SolidWorks「另存为」同款）——
'           左侧有"快速访问"，顶部地址栏能直接**粘贴路径**，可新建文件夹
'    False：旧式树形对话框（BrowseForFolder）
'  True 时如果脚本被安全策略拦住，会自动退回旧式对话框，不会卡住

Private Const CFG_VER           As String = "v2.3.0" ' 版本号（只用于弹窗标题显示）

'---- 1.9 上次目录记忆（存于 HKCU\Software\VB and VBA Program Settings\SWauto\Export）----
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


'==========================================================================================
'  第 2.5 节  现代对话框（借 PowerShell 实现）
'------------------------------------------------------------------------------------------
'  为什么这两个界面非要用 PowerShell，而不是纯 VBA：
'
'  ① 选文件夹 —— VBA 里能用的只有 Shell.Application.BrowseForFolder，那是**旧式树形**对话框：
'     没有"快速访问"侧栏，地址栏也不能粘贴路径。用户要的是和 SolidWorks「另存为」同款的
'     Windows 通用对话框，它对应 COM 接口 IFileOpenDialog + FOS_PICKFOLDERS。
'     这个接口在 VBA 里没法直接声明（要按 vtable 顺序写一堆方法），只能借 PowerShell
'     编译一小段 C# 去调它。
'
'  ② 逐类勾选框 —— VBA 的 MsgBox 最多给三个按钮，做不出复选框。同样借 WinForms。
'
'  两个脚本都写成临时 .ps1 再用 -File 运行，而不用 -Command，原因有两条：
'     · -Command 有命令行长度上限，脚本一长就被截断；
'     · .ps1 存成 **UTF-8 带 BOM** 时 PowerShell 才认它是 UTF-8，
'       否则按系统 ANSI 代码页解释，界面上的中文会变乱码
'       （和 VBE 导入 .bas 是同一类坑，只是宿主换成了 PowerShell）。
'
'  ★★ 参数一律内嵌在脚本正文里，命令行上不传任何参数 ★★
'
'     这个坑在 v2.2.0 真踩过一次，症状极具迷惑性：
'     「新式对话框明明弹出来了、也选好了，紧接着又弹一次旧式的」。
'
'     根因：WScript.Shell.Run 是经 **cmd** 启动 powershell 的，而 **cmd 根本不认单引号**。
'     原来我用 U_PsQuote 把参数包成 'xxx'，cmd 原样丢给 PowerShell；
'     PowerShell 的 -File 参数解析又不剥掉这层单引号，于是脚本收到的 $Out 是
'     "'C:\...\swauto_pickdir.txt'"（**带着引号**）。Set-Content -LiteralPath 拿到它，
'     当成一个相对路径，结果文件就写到了别处。VBA 这边读不到结果文件，
'     判定为"这套方案用不了"，默默回退到旧式对话框 —— 于是用户看到两个框连着弹。
'
'     实测确认（同一台机器上跑对照）：-File + 单引号参数 = 文件写不出来；
'     换成内嵌之后，命令行只剩 -File "<脚本路径>"，不存在任何解析歧义。
'     标题里带引号、带空格、带 $ 也都安全 —— 内嵌用的是 PowerShell 单引号字面量。
'
'  结果通过临时文本文件回传（UTF-8），VBA 侧用 U_ReadTextUtf8 读回。
'==========================================================================================

'---- 2.5.1 临时文件与 PowerShell 运行封装 ----

Private Function U_TempDir() As String
    Dim s As String
    s = Environ$("TEMP")
    If Len(s) = 0 Then s = Environ$("TMP")
    If Len(s) = 0 Then s = "C:\Windows\Temp"
    If Not U_DirExists(s) Then s = ""
    U_TempDir = s
End Function

Private Sub U_DelFile(ByVal sFile As String)
    On Error Resume Next
    If Len(sFile) > 0 Then
        If Len(Dir$(sFile, vbNormal)) > 0 Then Kill sFile
    End If
    Err.Clear
End Sub

' 写 UTF-8 带 BOM 的文本文件。
' ADODB.Stream 的 utf-8 到底带不带 BOM，各版本/补丁不一致，所以写完统一检查一次，
' 缺了就自己补上 —— 这个 BOM 是 PowerShell 正确识别中文的唯一依据，不能省。
Private Function U_WriteUtf8Bom(ByVal sFile As String, ByVal sText As String) As Boolean
    Dim st As Object, b() As Byte, n As Long
    On Error Resume Next

    Set st = CreateObject("ADODB.Stream")
    If st Is Nothing Then Exit Function
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText sText
    st.SaveToFile sFile, 2
    st.Close
    Err.Clear
    If Not U_FileExists(sFile) Then Exit Function

    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.LoadFromFile sFile
    b = st.Read
    st.Close

    n = -1
    On Error Resume Next
    n = UBound(b)
    Err.Clear
    If n >= 2 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then
            U_WriteUtf8Bom = True                  ' 已经有 BOM，不用动
            Exit Function
        End If
    End If

    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write ChrB(&HEF) & ChrB(&HBB) & ChrB(&HBF)
    If n >= 0 Then st.Write b
    st.SaveToFile sFile, 2
    st.Close
    U_WriteUtf8Bom = U_FileExists(sFile)
End Function

' 读 UTF-8 文本（ADODB.Stream 会自动跳过 BOM）
Private Function U_ReadTextUtf8(ByVal sFile As String) As String
    Dim st As Object
    On Error Resume Next
    If Not U_FileExists(sFile) Then Exit Function
    Set st = CreateObject("ADODB.Stream")
    If st Is Nothing Then Exit Function
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile sFile
    U_ReadTextUtf8 = st.ReadText(-1)               ' -1 = adReadAll
    st.Close
End Function

' 把脚本回传的结果清洗成一行干净文本。**所有**读结果的地方都必须过这一道。
'
' ！不要用 VBA 的 Trim$ —— 它**只去空格(Chr 32)，不去换行**。
'   而 PowerShell 的 Set-Content 默认会在文件末尾补一个 CRLF。
'   v2.2.1 就栽在这上面，症状相当绕：
'     · 选好文件夹 → 读回的路径末尾带 CRLF → 拿去建目录必然失败 →
'       弹「无法创建输出目录：D:\...」，而那个换行在弹窗里根本看不见；
'     · 用户点取消 → 读回的是 "C" & vbCrLf，和 "C" 比较不相等 →
'       「取消」判定失效，于是把 "C" 当成一个路径继续往下走。
'   所以在源头（脚本侧）加了 -NoNewline 之后，这里再兜一道，
'   把 CR / LF / Tab 和可能残留的 BOM 一并剔掉。
Private Function U_CleanResult(ByVal s As String) As String
    Dim s1 As String, u As Long
    If Len(s) = 0 Then Exit Function
    s1 = Replace(s, vbCr, "")
    s1 = Replace(s1, vbLf, "")
    s1 = Replace(s1, vbTab, "")
    If Len(s1) > 0 Then
        u = AscW(Left$(s1, 1))
        If u < 0 Then u = u + 65536                 ' VBA 的 AscW 对 U+8000 以上返回负数
        If u = &HFEFF Then s1 = Mid$(s1, 2)         ' 万一 ADODB 没跳过 BOM
    End If
    U_CleanResult = Trim$(s1)
End Function

' 把一段 PowerShell 脚本写成临时 .ps1 并运行。
'   -ExecutionPolicy Bypass 是必需的：默认执行策略会直接拒绝运行 .ps1
'   （实测本机默认就是 Restricted，不加这个开关脚本根本不会执行，而且是**静默失败**）。
'   -WindowStyle Hidden 只藏控制台窗口，WinForms 窗体照常显示。
'
'  ！命令行上只给脚本路径，**不传任何业务参数** —— 理由见文件头部那段 ★★ 说明。
'     所以调用方要先把参数值用 U_PsQuote 内嵌进 sBody 里再调这里。
Private Function U_RunPs(ByVal sBody As String) As Boolean
    Dim sDir As String, sPs As String, sCmd As String
    sDir = U_TempDir()
    If Len(sDir) = 0 Then Exit Function
    sPs = U_JoinPath(sDir, "swauto_ps.ps1")
    If Not U_WriteUtf8Bom(sPs, sBody) Then Exit Function
    sCmd = "powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & sPs & """"
    U_ShellRunWait sCmd, 0
    U_RunPs = True
End Function

' 追加一条诊断信息到 %TEMP%\swauto_diag.txt。
' 只在这类"静默失败"时调用：外部方案没给出结果、但界面上看不出来发生了什么。
' （v2.2.0 那个 cmd 单引号的问题就是典型 —— 用户只看到多弹了一个框，
'   完全没法知道是参数没传进去。有这条日志就能直接看到症结。）
Private Sub U_Diag(ByVal sMsg As String)
    Dim sF As String, h As Integer
    sF = U_JoinPath(U_TempDir(), "swauto_diag.txt")
    If Len(sF) < 4 Then Exit Sub
    On Error Resume Next
    h = FreeFile
    Open sF For Append As #h
    Print #h, Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & sMsg
    Close #h
    Err.Clear
End Sub

'---- 2.5.2 选文件夹 ----
' 对外只有 U_PickFolder 一个入口：默认走"现代"对话框，用不了才退回旧式。

Private Function U_PickFolder(ByVal sTitle As String, ByVal sDefault As String) As String
    Dim sOut As String, bOk As Boolean, bUsable As Boolean
    If CFG_MODERN_FOLDER_PICK Then
        bUsable = False
        bOk = U_PickFolderModern(sTitle, sDefault, sOut, bUsable)
        If bOk Then
            U_PickFolder = U_NormPath(sOut)
            Exit Function
        End If
        ' bUsable = True 说明对话框确实弹出来了，只是用户点了取消 —— 这时**绝不能**
        ' 再退回旧式重新弹一次，那是骚扰。只有脚本根本没跑起来（bUsable = False）才回退。
        If bUsable Then Exit Function
    End If
    U_PickFolder = U_PickFolderClassic(sTitle, sDefault)
End Function

' Windows 通用文件夹对话框（和「另存为」同款：左侧快速访问、地址栏可粘贴路径）。
'   sOut    选中路径
'   bUsable True  = 对话框正常弹出来了（不管用户选还是取消）
'            False = 这套方案不可用，调用方应回退到旧式对话框
Private Function U_PickFolderModern(ByVal sTitle As String, ByVal sDefault As String, _
                                    ByRef sOut As String, ByRef bUsable As Boolean) As Boolean
    Dim sDir As String, sRes As String, sResp As String
    bUsable = False
    sDir = U_TempDir()
    If Len(sDir) = 0 Then Exit Function

    sRes = U_JoinPath(sDir, "swauto_pickdir.txt")
    U_DelFile sRes

    ' 三个参数全部内嵌进脚本正文，命令行不传任何参数（理由见文件头部 ★★ 那段）
    If Not U_RunPs(U_FolderPickerScript(sTitle, sDefault, sRes)) Then Exit Function

    sResp = U_CleanResult(U_ReadTextUtf8(sRes))
    If Len(sResp) = 0 Then
        ' 结果文件没生成：脚本没跑起来，或者它把文件写到别处去了。
        ' 这条日志是排查这类"静默失败"的唯一线索，别删。
        U_Diag "文件夹选择器没返回结果。脚本=" & U_JoinPath(sDir, "swauto_ps.ps1") & _
               "  期望结果文件=" & sRes
        Exit Function                                ' 允许回退
    End If
    U_DelFile sRes

    If sResp = "E" Then
        U_Diag "文件夹选择器脚本自报不可用（Add-Type 编译失败等），回退旧式对话框"
        Exit Function                                ' 同样允许回退
    End If
    bUsable = True
    If sResp = "C" Then Exit Function                ' 用户点了取消
    sOut = sResp
    U_PickFolderModern = True
End Function

' 旧式树形文件夹对话框（仅作回退）。
Private Function U_PickFolderClassic(ByVal sTitle As String, ByVal sDefault As String) As String
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
    ' &H51 = BIF_RETURNONLYFSDIRS(1) + BIF_EDITBOX(16) + BIF_NEWDIALOGSTYLE(64)
    ' 加上 EDITBOX 是为了**至少能粘贴路径** —— 现代对话框不可用时，这条通道的体验就靠它。
    Set fld = sh.BrowseForFolder(0, sTitle, &H51, 0)
    If Not fld Is Nothing Then
        sRes = fld.Self.Path
        If Len(sRes) = 0 Then sRes = fld.Items().Item().Path    ' 新式对话框下 Self.Path 可能为空
    End If
    If Len(sOld) > 0 Then ChDir sOld
    On Error GoTo 0
    U_PickFolderClassic = U_NormPath(sRes)
End Function

' 文件夹选择器的 PowerShell 脚本正文。
' 关键就一句：给通用对话框加上 FOS_PICKFOLDERS 标志，把它从"选文件"变成"选文件夹"。
Private Function U_FolderPickerScript(ByVal sTitle As String, ByVal sInitial As String, _
                                      ByVal sOut As String) As String
    Dim s As String

    ' 三个参数直接内嵌成 PowerShell 单引号字面量，命令行不传参（见文件头部 ★★ 那段）
    s = "$Title = " & U_PsQuote(sTitle) & vbCrLf
    s = s & "$Initial = " & U_PsQuote(sInitial) & vbCrLf
    s = s & "$Out = " & U_PsQuote(sOut) & vbCrLf
    s = s & "if ([string]::IsNullOrEmpty($Out)) { exit 1 }" & vbCrLf
    s = s & "$cs = @'" & vbCrLf
    s = s & "using System;" & vbCrLf
    s = s & "using System.Runtime.InteropServices;" & vbCrLf
    s = s & "public static class SwFolderPicker" & vbCrLf
    s = s & "{" & vbCrLf
    s = s & "    [ComImport, Guid(""DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7"")]" & vbCrLf
    s = s & "    private class FileOpenDialog { }" & vbCrLf
    s = s & "    [ComImport, Guid(""42f85136-db7e-439c-85f1-e4075d135fc8""), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]" & vbCrLf
    s = s & "    private interface IFileDialog" & vbCrLf
    s = s & "    {" & vbCrLf
    s = s & "        [PreserveSig] int Show(IntPtr parent);" & vbCrLf
    s = s & "        void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);" & vbCrLf
    s = s & "        void SetFileTypeIndex(uint iFileType);" & vbCrLf
    s = s & "        void GetFileTypeIndex(out uint piFileType);" & vbCrLf
    s = s & "        void Advise(IntPtr pfde, out uint pdwCookie);" & vbCrLf
    s = s & "        void Unadvise(uint dwCookie);" & vbCrLf
    s = s & "        void SetOptions(uint fos);" & vbCrLf
    s = s & "        void GetOptions(out uint pfos);" & vbCrLf
    s = s & "        void SetDefaultFolder(IShellItem psi);" & vbCrLf
    s = s & "        void SetFolder(IShellItem psi);" & vbCrLf
    s = s & "        void GetFolder(out IShellItem ppsi);" & vbCrLf
    s = s & "        void GetCurrentSelection(out IShellItem ppsi);" & vbCrLf
    s = s & "        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);" & vbCrLf
    s = s & "        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);" & vbCrLf
    s = s & "        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);" & vbCrLf
    s = s & "        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);" & vbCrLf
    s = s & "        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);" & vbCrLf
    s = s & "        void GetResult(out IShellItem ppsi);" & vbCrLf
    s = s & "        void AddPlace(IShellItem psi, int fdap);" & vbCrLf
    s = s & "        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);" & vbCrLf
    s = s & "        void Close(int hr);" & vbCrLf
    s = s & "        void SetClientGuid(ref Guid guid);" & vbCrLf
    s = s & "        void ClearClientData();" & vbCrLf
    s = s & "        void SetFilter(IntPtr pFilter);" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "    [ComImport, Guid(""43826d1e-e718-42ee-bc55-a1e261c37bfe""), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]" & vbCrLf
    s = s & "    private interface IShellItem" & vbCrLf
    s = s & "    {" & vbCrLf
    s = s & "        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);" & vbCrLf
    s = s & "        void GetParent(out IShellItem ppsi);" & vbCrLf
    s = s & "        void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);" & vbCrLf
    s = s & "        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);" & vbCrLf
    s = s & "        void Compare(IShellItem psi, uint hint, out int piOrder);" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "    [DllImport(""shell32.dll"", CharSet = CharSet.Unicode, PreserveSig = false)]" & vbCrLf
    s = s & "    private static extern void SHCreateItemFromParsingName(" & vbCrLf
    s = s & "        [MarshalAs(UnmanagedType.LPWStr)] string pszPath," & vbCrLf
    s = s & "        IntPtr pbc, ref Guid riid," & vbCrLf
    s = s & "        [MarshalAs(UnmanagedType.Interface)] out IShellItem ppv);" & vbCrLf
    s = s & "    [DllImport(""user32.dll"")]" & vbCrLf
    s = s & "    private static extern IntPtr GetForegroundWindow();" & vbCrLf
    s = s & "    public static string Pick(string title, string initial)" & vbCrLf
    s = s & "    {" & vbCrLf
    s = s & "        const uint FOS_PICKFOLDERS = 0x20;" & vbCrLf
    s = s & "        const uint FOS_FORCEFILESYSTEM = 0x40;" & vbCrLf
    s = s & "        const uint SIGDN_FILESYSPATH = 0x80058000;" & vbCrLf
    s = s & "        IFileDialog dlg = (IFileDialog)(new FileOpenDialog());" & vbCrLf
    s = s & "        uint opts;" & vbCrLf
    s = s & "        dlg.GetOptions(out opts);" & vbCrLf
    s = s & "        dlg.SetOptions(opts | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);" & vbCrLf
    s = s & "        if (!string.IsNullOrEmpty(title)) dlg.SetTitle(title);" & vbCrLf
    s = s & "        if (!string.IsNullOrEmpty(initial))" & vbCrLf
    s = s & "        {" & vbCrLf
    s = s & "            try" & vbCrLf
    s = s & "            {" & vbCrLf
    s = s & "                IShellItem start;" & vbCrLf
    s = s & "                Guid iid = typeof(IShellItem).GUID;" & vbCrLf
    s = s & "                SHCreateItemFromParsingName(initial, IntPtr.Zero, ref iid, out start);" & vbCrLf
    s = s & "                if (start != null) dlg.SetFolder(start);" & vbCrLf
    s = s & "            }" & vbCrLf
    s = s & "            catch { }" & vbCrLf
    s = s & "        }" & vbCrLf
    s = s & "        if (dlg.Show(GetForegroundWindow()) != 0) return """";" & vbCrLf
    s = s & "        IShellItem result;" & vbCrLf
    s = s & "        dlg.GetResult(out result);" & vbCrLf
    s = s & "        if (result == null) return """";" & vbCrLf
    s = s & "        string path;" & vbCrLf
    s = s & "        result.GetDisplayName(SIGDN_FILESYSPATH, out path);" & vbCrLf
    s = s & "        return path == null ? """" : path;" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "'@" & vbCrLf
    s = s & "try {" & vbCrLf
    s = s & "    Add-Type -TypeDefinition $cs -Language CSharp -ErrorAction Stop" & vbCrLf
    s = s & "} catch {" & vbCrLf
    s = s & "    Set-Content -LiteralPath $Out -Value 'E' -NoNewline -Encoding ASCII" & vbCrLf
    s = s & "    exit 0" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "$picked = ''" & vbCrLf
    s = s & "$err = $false" & vbCrLf
    s = s & "try { $picked = [SwFolderPicker]::Pick($Title, $Initial) } catch { $err = $true }" & vbCrLf
    s = s & "if ($err) {" & vbCrLf
    s = s & "    Set-Content -LiteralPath $Out -Value 'E' -NoNewline -Encoding ASCII" & vbCrLf
    s = s & "} elseif ([string]::IsNullOrEmpty($picked)) {" & vbCrLf
    s = s & "    Set-Content -LiteralPath $Out -Value 'C' -NoNewline -Encoding ASCII" & vbCrLf
    s = s & "} else {" & vbCrLf
    s = s & "    Set-Content -LiteralPath $Out -Value $picked -NoNewline -Encoding UTF8" & vbCrLf
    s = s & "}" & vbCrLf

    U_FolderPickerScript = s
End Function

'---- 2.5.3 装配体展开：逐类勾选 ----
' 返回：-1 = 用户点了「取消导出」；-2 = 对话框弹不出来；否则 0..7 位掩码
'        （bit0 = 也导出已隐藏，bit1 = 也导出已压缩，bit2 = 也导出封套）
Private Function U_AskSkipOptions(ByVal nHid As Long, ByVal nSup As Long, ByVal nEnv As Long) As Long
    Dim sDir As String, sRes As String, sResp As String
    U_AskSkipOptions = -2
    sDir = U_TempDir()
    If Len(sDir) = 0 Then Exit Function

    sRes = U_JoinPath(sDir, "swauto_ask.txt")
    U_DelFile sRes

    ' 个数、初始勾选状态、结果路径**全部内嵌**进脚本正文，命令行不传参（见文件头部 ★★ 那段）。
    ' 初始勾选状态取自配置：CFG_SKIP_XXX = False 的类别（"总是要导"）默认就打上勾，
    ' 否则用户一进这个框就会被"都不勾 = 都不导"的默认值覆盖掉配置意图。
    If Not U_RunPs(U_AskSkipScript(nHid, nSup, nEnv, sRes)) Then Exit Function

    sResp = U_CleanResult(U_ReadTextUtf8(sRes))
    U_DelFile sRes

    ' ！这段用 If/ElseIf/Else 分三级，级别顺序就是判定的优先级，**不能调换**。
    '   「取消」回传的就是**一个字符 C** —— 如果先做"长度必须是 3"的校验，
    '   C 会被当成"没拿到结果"，函数返回 -2（对话框不可用），
    '   调用方于是按默认继续往下导出，用户按的「取消导出」就彻底失效了。
    '   v2.2.0 到 v2.2.2 一直是这个毛病，而且表面上完全看不出问题在哪 ——
    '   用户只会看到「点了取消却没停下来」。check.py 里现在有一条专门盯这个顺序。
    If sResp = "C" Then
        U_AskSkipOptions = -1                      ' 级别①：用户明确取消，调用方要中止整个导出
    ElseIf Len(sResp) = 3 Then
        U_AskSkipOptions = 0                       ' 级别②：正常的三位掩码
        If Mid$(sResp, 1, 1) = "1" Then U_AskSkipOptions = U_AskSkipOptions + 1
        If Mid$(sResp, 2, 1) = "1" Then U_AskSkipOptions = U_AskSkipOptions + 2
        If Mid$(sResp, 3, 1) = "1" Then U_AskSkipOptions = U_AskSkipOptions + 4
    Else
        ' 级别③：空串、'E'、或任何意外内容 —— 保持 -2，
        ' 让调用方按默认（三类都不导）继续，别因为一个界面问题把整个导出卡死。
        U_Diag "展开选项勾选框没返回有效结果。读到=[" & sResp & "] 长度=" & Len(sResp)
    End If
End Function

' 勾选框窗体的 PowerShell 脚本正文。
' 结果写成三位 0/1（顺序固定：隐藏 / 压缩 / 封套），取消则写 'C'。
Private Function U_AskSkipScript(ByVal nHid As Long, ByVal nSup As Long, ByVal nEnv As Long, _
                                 ByVal sOut As String) As String
    Dim s As String

    ' 全部内嵌成字面量，命令行不传参（见文件头部 ★★ 那段）。
    ' 数字用 [int] 显式转一次：内嵌走的是单引号字符串，转完才是真正的整数。
    s = "$nHid = [int]" & U_PsQuote(CStr(nHid)) & vbCrLf
    s = s & "$nSup = [int]" & U_PsQuote(CStr(nSup)) & vbCrLf
    s = s & "$nEnv = [int]" & U_PsQuote(CStr(nEnv)) & vbCrLf
    s = s & "$dHid = [int]" & U_PsQuote(CStr(IIf(CFG_SKIP_HIDDEN, 0, 1))) & vbCrLf
    s = s & "$dSup = [int]" & U_PsQuote(CStr(IIf(CFG_SKIP_SUPPRESSED, 0, 1))) & vbCrLf
    s = s & "$dEnv = [int]" & U_PsQuote(CStr(IIf(CFG_SKIP_ENVELOPE, 0, 1))) & vbCrLf
    s = s & "$Out = " & U_PsQuote(sOut) & vbCrLf
    s = s & "if ([string]::IsNullOrEmpty($Out)) { exit 1 }" & vbCrLf
    s = s & "Add-Type -AssemblyName System.Windows.Forms" & vbCrLf
    s = s & "Add-Type -AssemblyName System.Drawing" & vbCrLf
    s = s & "$f = New-Object System.Windows.Forms.Form" & vbCrLf
    s = s & "$f.Text = 'SWauto - 装配体展开选项'" & vbCrLf
    s = s & "$f.StartPosition = 'CenterScreen'" & vbCrLf
    s = s & "$f.FormBorderStyle = 'FixedDialog'" & vbCrLf
    s = s & "$f.MaximizeBox = $false" & vbCrLf
    s = s & "$f.MinimizeBox = $false" & vbCrLf
    s = s & "$f.ShowInTaskbar = $false" & vbCrLf
    s = s & "$f.TopMost = $true" & vbCrLf
    s = s & "$f.ClientSize = New-Object System.Drawing.Size(486, 258)" & vbCrLf
    s = s & "$lb = New-Object System.Windows.Forms.Label" & vbCrLf
    s = s & "$lb.Text = '下面这几类零件默认不导出。要哪一类，就在它前面打勾（可多选）：'" & vbCrLf
    s = s & "$lb.SetBounds(18, 16, 450, 22)" & vbCrLf
    s = s & "$f.Controls.Add($lb)" & vbCrLf
    s = s & "$cbH = New-Object System.Windows.Forms.CheckBox" & vbCrLf
    s = s & "$cbH.Text = ('也导出【已隐藏】的零件          共 ' + $nHid + ' 个')" & vbCrLf
    s = s & "$cbH.SetBounds(32, 54, 428, 26)" & vbCrLf
    s = s & "$f.Controls.Add($cbH)" & vbCrLf
    s = s & "$cbS = New-Object System.Windows.Forms.CheckBox" & vbCrLf
    s = s & "$cbS.Text = ('也导出【已压缩】的零件          共 ' + $nSup + ' 个')" & vbCrLf
    s = s & "$cbS.SetBounds(32, 90, 428, 26)" & vbCrLf
    s = s & "$f.Controls.Add($cbS)" & vbCrLf
    s = s & "$cbE = New-Object System.Windows.Forms.CheckBox" & vbCrLf
    s = s & "$cbE.Text = ('也导出【封套】零件              共 ' + $nEnv + ' 个')" & vbCrLf
    s = s & "$cbE.SetBounds(32, 126, 428, 26)" & vbCrLf
    s = s & "$f.Controls.Add($cbE)" & vbCrLf
    s = s & "$lb2 = New-Object System.Windows.Forms.Label" & vbCrLf
    s = s & "$lb2.Text = '打勾的类别会被导出（预置状态来自配置）；点「取消导出」= 本次不导了'" & vbCrLf
    s = s & "$lb2.SetBounds(18, 166, 450, 22)" & vbCrLf
    s = s & "$f.Controls.Add($lb2)" & vbCrLf
    s = s & "$ok = New-Object System.Windows.Forms.Button" & vbCrLf
    s = s & "$ok.Text = '确定'" & vbCrLf
    s = s & "$ok.DialogResult = [System.Windows.Forms.DialogResult]::OK" & vbCrLf
    s = s & "$ok.SetBounds(280, 204, 90, 30)" & vbCrLf
    s = s & "$f.Controls.Add($ok)" & vbCrLf
    s = s & "$no = New-Object System.Windows.Forms.Button" & vbCrLf
    s = s & "$no.Text = '取消导出'" & vbCrLf
    s = s & "$no.DialogResult = [System.Windows.Forms.DialogResult]::Cancel" & vbCrLf
    s = s & "$no.SetBounds(378, 204, 90, 30)" & vbCrLf
    s = s & "$f.Controls.Add($no)" & vbCrLf
    s = s & "$f.AcceptButton = $ok" & vbCrLf
    s = s & "$f.CancelButton = $no" & vbCrLf
    s = s & "if ($dHid -eq 1) { $cbH.Checked = $true }" & vbCrLf
    s = s & "if ($dSup -eq 1) { $cbS.Checked = $true }" & vbCrLf
    s = s & "if ($dEnv -eq 1) { $cbE.Checked = $true }" & vbCrLf
    s = s & "$r = $f.ShowDialog()" & vbCrLf
    s = s & "if ($r -eq [System.Windows.Forms.DialogResult]::OK) {" & vbCrLf
    s = s & "    $v = ''" & vbCrLf
    s = s & "    if ($cbH.Checked) { $v = $v + '1' } else { $v = $v + '0' }" & vbCrLf
    s = s & "    if ($cbS.Checked) { $v = $v + '1' } else { $v = $v + '0' }" & vbCrLf
    s = s & "    if ($cbE.Checked) { $v = $v + '1' } else { $v = $v + '0' }" & vbCrLf
    s = s & "    Set-Content -LiteralPath $Out -Value $v -NoNewline -Encoding ASCII" & vbCrLf
    s = s & "} else {" & vbCrLf
    s = s & "    Set-Content -LiteralPath $Out -Value 'C' -NoNewline -Encoding ASCII" & vbCrLf
    s = s & "}" & vbCrLf

    U_AskSkipScript = s
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
        ' 该接口返回 (名称, 全路径) 成对出现。
        ' 只认【绝对路径 + 文件真实存在】：否则 ".\零件.SLDPRT" 这类名称项会被误当成一条独立路径
        If U_IsModel(sItem) And U_IsAbsPath(sItem) And U_FileExists(sItem) Then
            sKey = U_FileKey(sItem)
            If Not m_IdxByPath.Exists(sKey) Then m_IdxByPath.Add sKey, sDrwPath
            sKey = UCase$(U_FileName(sItem))
            If Not m_IdxByName.Exists(sKey) Then m_IdxByName.Add sKey, sDrwPath
        End If
    Next i
End Sub

Private Function F_Lookup(ByVal sModelPath As String) As String
    Dim sKey As String
    If m_IdxByPath Is Nothing Then Exit Function
    sKey = U_FileKey(sModelPath)                 ' 与 F_IndexOne 用同一种 key 规范化方式
    If m_IdxByPath.Exists(sKey) Then
        F_Lookup = m_IdxByPath(sKey)
        Exit Function
    End If
    sKey = UCase$(U_FileName(sModelPath))
    If m_IdxByName.Exists(sKey) Then F_Lookup = m_IdxByName(sKey)
End Function

' 反过来：一个工程图引用了哪些模型（用于"工程图 -> 模型"模式）
' 【这里是"同一个零件导出两份、文件名带 _2"的源头，改动请慎重】
'   GetDocumentDependencies2 返回的数组是 (名称, 全路径) 成对出现，
'   其中"名称"项在某些保存方式下自带 "\"（如 ".\零件.SLDPRT"），
'   只判断"含不含 \" 就会把它当成第二条独立路径收进来 → 同一个零件被收录两次。
'   所以这里必须同时满足：绝对路径 + 文件真实存在 + 按规范化 key 去重。
Private Function F_ModelsOfDrawing(ByVal sDrwPath As String) As Collection
    Dim vDeps As Variant, i As Long, sItem As String, sKey As String
    Dim col As New Collection
    Dim dic As Object
    On Error Resume Next
    Set F_ModelsOfDrawing = col
    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1
    On Error GoTo 0
    vDeps = m_swApp.GetDocumentDependencies2(sDrwPath, False, False, False)
    If Not IsArray(vDeps) Then Exit Function
    For i = LBound(vDeps) To UBound(vDeps)
        sItem = CStr(vDeps(i))
        If U_IsModel(sItem) And U_IsAbsPath(sItem) And U_FileExists(sItem) Then
            sKey = U_FileKey(sItem)
            If Not dic.Exists(sKey) Then
                dic.Add sKey, True
                col.Add sItem
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
'  第 3.5 节  目标收集（参与状态过滤 + 标准件过滤 + 装配体展开）
'            只服务于 SwAuto_ExportActiveDoc：
'            把"当前激活的文档"解析成 一组要导出的模型(→STEP) + 一组工程图(→PDF)
'
'  装配体展开时会依次过两道筛子（顺序不能反）：
'    ① 参与状态筛（3.5.2）：隐藏 / 压缩 / 封套的组件跳过
'    ② 标准件筛  （3.5.1）：螺钉螺母垫圈等跳过
'  先状态后关键字 —— 被隐藏的零件不必再去猜它是不是标准件。
'==========================================================================================

'---- 3.5.1 标准件判定 ----
' 命中任意一条即算标准件：
'   ① SolidWorks 官方 Toolbox 标记（最准，需要拿到 ModelDoc）
'   ② 文件路径里含 CFG_STD_PATH_HINTS
'   ③ 文件名里含 CFG_STD_KEYWORDS / CFG_STD_STDCODES（词边界匹配，不会误伤）
'   ④ 文件夹名含 Toolbox 标准代号（如 ...\ISO\、...\GB\）
'
' 注意：判断"是不是 Toolbox 零件"的正确 API 是
'       ModelDocExtension::ToolboxPartType，
'       而不是 Component2 上的什么 IsToolboxComponent —— 那个属性并不存在（试过，编译不过）。
'       轻化组件的 GetModelDoc2 可能返回 Nothing，此时这一条自动跳过，交给后面的规则兜底。
Private Function G_IsStdPart(ByVal sPath As String, ByVal mDoc As Object, ByRef sReason As String) As Boolean
    Dim i As Long, v As Variant, sName As String
    Dim vTb As Variant
    sReason = ""
    If Not CFG_EXCLUDE_STD_PARTS Then Exit Function

    ' ① 官方 Toolbox 标记
    '    该属性可能返回 Boolean(-1/0)，也可能返回 swToolboxPartType_e 枚举，统一用 <> 0 判断
    If CFG_STD_USE_API And Not mDoc Is Nothing Then
        vTb = Empty
        On Error Resume Next
        vTb = mDoc.Extension.ToolboxPartType
        Err.Clear
        On Error GoTo 0
        If IsNumeric(vTb) Then
            If vTb <> 0 Then
                sReason = "Toolbox 标准件"
                G_IsStdPart = True
                Exit Function
            End If
        End If
    End If

    ' ② 路径关键字
    v = U_KeywordList(CFG_STD_PATH_HINTS)
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            If InStr(1, sPath, CStr(v(i)), vbTextCompare) > 0 Then
                sReason = "路径含 " & v(i)
                G_IsStdPart = True
                Exit Function
            End If
        Next i
    End If

    ' ③ 文件名关键字
    sName = U_FileName(sPath)
    v = U_KeywordList(CFG_STD_KEYWORDS)
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            If U_HitKeyword(sName, CStr(v(i))) Then
                sReason = "文件名含 " & v(i)
                G_IsStdPart = True
                Exit Function
            End If
        Next i
    End If

    ' ③ 国标 / 国际标准号
    v = U_KeywordList(CFG_STD_STDCODES)
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            If U_HitKeyword(sName, CStr(v(i))) Then
                sReason = "标准号 " & v(i)
                G_IsStdPart = True
                Exit Function
            End If
        Next i
    End If
End Function

' 记一个被过滤的标准件（同一次运行内去重），并写进导出清单
Private Sub G_NoteStd(ByVal sPath As String, ByVal sReason As String)
    If F_Contains(m_SkippedStd, sPath) Then Exit Sub
    m_SkippedStd.Add sPath
    U_AddResult "模型", sPath, "", "跳过", "标准件已过滤（" & sReason & "）"
End Sub

' 取组件对应的模型文档；组件被轻化 / 抑制时可能拿到 Nothing（此时官方 Toolbox 标记用不了，
' 后面的路径与文件名规则仍然有效，不影响整体判定）
Private Function G_DocOfComp(ByVal comp As Object) As Object
    If comp Is Nothing Then Exit Function
    On Error Resume Next
    Set G_DocOfComp = comp.GetModelDoc2
    Err.Clear
    On Error GoTo 0
End Function

Private Sub G_AddModel(ByVal sPath As String, ByVal colModels As Collection, ByRef n As Long)
    If Not F_Contains(colModels, sPath) Then
        colModels.Add sPath
        n = n + 1
    End If
End Sub

' 注：同一个零件在装配体里可能有多个实例，状态各不相同（一个隐藏、一个可见）。
' 处理办法见 G_MergeOne —— 它写"跳过"之前会先看这个路径是不是已经进了待导出列表，
' 是的话就不写，免得清单里出现"一边成功、一边跳过"的自相矛盾。

'---- 3.5.2 参与状态过滤（隐藏 / 压缩 / 封套）----
' 这三种组件在装配体里不参与当前的设计状态，导出它们只会污染交付物：
'   隐藏 —— 视觉上根本看不到（Component2::IsHidden）
'   压缩 —— 不参与重建、也不进 BOM（Component2::IsSuppressed）
'   封套 —— 只是空间占位的参考体，不是真实零件（Component2::IsEnvelope）
'
' 两个必须交代的细节（都是查过官方帮助才敢写的）：
'
'  ① IsHidden 必须传 False。
'     Component2::IsHidden(ConsiderSuppressed) 的语义是：
'       传 True  → "隐藏 或 压缩 或 轻化" 都返回 True
'       传 False → 只按可见性判断（真隐藏才返回 True）
'     官方帮助明确写着"For lightweight components, IsHidden returns True if
'     ConsiderSuppressed is True"。也就是说图省事写 IsHidden(True) 一把梭，
'     会**把轻化组件全部误杀** —— 而轻化只是加载方式，零件本身完全有效。
'     所以隐藏与压缩必须分开判断。
'
'  ② 判定要沿"父级链"往上走。
'     GetComponents(False) 是**扁平**返回所有层级的组件，不带层级信息。
'     某个零件自身是"显示"的，但它所在的子装配体被隐藏/压缩时，实际同样看不见。
'     所以拿到组件后用 GetParent 逐级回溯，任一级不合格就跳过。

' 沿父级链判断：返回 True = 应当跳过。
'   nKind 输出类别（1 封套 / 2 已压缩 / 4 已隐藏），sWhy 输出中文原因（父级命中时带后缀）。
Private Function G_SkipKind(ByVal comp As Object, ByRef nKind As Long, ByRef sWhy As String) As Boolean
    Dim c As Object, p As Object, sSub As String, n As Long
    nKind = 0
    sWhy = ""
    If comp Is Nothing Then Exit Function

    Set c = comp
    Do While Not c Is Nothing
        n = n + 1
        If n > 64 Then Exit Do                 ' 防御：万一父级链出现环，别把 SolidWorks 挂死

        nKind = G_CompBadState(c, sSub)
        If nKind <> 0 Then
            If n > 1 Then sSub = sSub & "（上级装配体）"
            sWhy = sSub
            G_SkipKind = True
            Exit Function
        End If

        Set p = Nothing                        ' 先清空：取父级失败时 p 保持 Nothing，循环自然结束
        On Error Resume Next
        Set p = c.GetParent
        Err.Clear
        On Error GoTo 0
        Set c = p
    Loop
    nKind = 0
End Function

' 单个组件自身的状态判定。**只看状态，不看配置开关** —— 要不要真跳过由调用方定，
' 因为现在有三个决策来源：CFG_SKIP_* 的设定、用户在勾选框里的选择、以及"父级牵连"。
' 返回：0 = 正常；1 = 封套；2 = 已压缩；4 = 已隐藏
' 顺序固定为 封套 → 压缩 → 隐藏：一个组件只归一类，免得同一个被数两遍。
Private Function G_CompBadState(ByVal c As Object, ByRef sWhy As String) As Long
    Dim b As Boolean

    b = False
    On Error Resume Next
    b = c.IsEnvelope
    Err.Clear
    On Error GoTo 0
    If b Then
        G_CompBadState = 1
        sWhy = "封套"
        Exit Function
    End If

    b = False
    On Error Resume Next
    b = c.IsSuppressed
    Err.Clear
    On Error GoTo 0
    If b Then
        G_CompBadState = 2
        sWhy = "已压缩"
        Exit Function
    End If

    b = False
    On Error Resume Next
    b = c.IsHidden(False)                      ' ★ 必须传 False，理由见上面 ①
    Err.Clear
    On Error GoTo 0
    If b Then
        G_CompBadState = 4
        sWhy = "已隐藏"
        Exit Function
    End If
End Function

' 这个类别按 CFG_SKIP_* 的设定该不该跳过（只有"不询问用户"时才用得上）
Private Function G_KindSkipped(ByVal nKind As Long) As Boolean
    Select Case nKind
        Case 1: G_KindSkipped = CFG_SKIP_ENVELOPE
        Case 2: G_KindSkipped = CFG_SKIP_SUPPRESSED
        Case 4: G_KindSkipped = CFG_SKIP_HIDDEN
    End Select
End Function

' 取组件的可读标识：优先磁盘路径；取不到（压缩 / 轻化组件可能返回空）就用特征树里的名字。
' 有了它，被跳过的组件至少能在清单里留下名字，而不是只报一个数字。
'   sPath 单独输出：空表示"拿不到磁盘路径"，这种组件即使被勾中也导不出来，要如实说明。
Private Function G_CompLabel(ByVal comp As Object, ByRef sPath As String) As String
    Dim s As String
    sPath = ""
    If comp Is Nothing Then Exit Function
    On Error Resume Next
    s = comp.GetPathName
    Err.Clear
    On Error GoTo 0
    sPath = s
    If Len(s) = 0 Then
        On Error Resume Next
        s = comp.Name2
        Err.Clear
        On Error GoTo 0
    End If
    G_CompLabel = s
End Function

' 记一个因"参与状态"被跳过的组件（同一次运行内按标识去重）。
' 刻意**不**写进明细表：状态是"实例级"的（同一零件的多个实例状态可以不同），
' 写进去会和它的"成功"行按源文件打架。改为在清单末尾单独成段，原因也一并列出。
Private Sub G_NoteSkip(ByVal sLabel As String, ByVal sReason As String)
    If Len(sLabel) = 0 Then Exit Sub
    If m_SkipWhy Is Nothing Then Exit Sub
    If m_SkipWhy.Exists(sLabel) Then Exit Sub
    m_SkipWhy(sLabel) = sReason
    m_SkippedState.Add sLabel
End Sub

' 一个"状态正常"的组件：取路径 → 标准件判定 → 收进待导出列表。
Private Sub G_CollectOne(ByVal comp As Object, ByVal sAsmPath As String, _
                         ByVal colModels As Collection, ByRef nPart As Long, ByRef nSkipped As Long)
    Dim sLabel As String, sP As String, sR As String, mdoc As Object

    sLabel = G_CompLabel(comp, sP)

    If Len(sP) = 0 Then
        ' 虚拟零件（几何存在装配体内部）没法单独导出成 STEP，记一笔免得用户以为漏了
        If Len(sLabel) > 0 Then
            U_AddResult "模型", sLabel, "", "跳过", "虚拟零件（保存在装配体内部，无法单独导出）"
            nSkipped = nSkipped + 1
        End If
        Exit Sub
    End If

    Select Case U_Ext(sP)
    Case ".SLDPRT"
        ' 标准件：Toolbox 标记 / 路径关键字 / 文件名关键字
        Set mdoc = G_DocOfComp(comp)
        If G_IsStdPart(sP, mdoc, sR) Then
            G_NoteStd sP, sR
            nSkipped = nSkipped + 1
        Else
            G_AddModel sP, colModels, nPart
        End If
    Case ".SLDASM"
        ' 子装配体自身是否也出一份（默认不导：只交付零件）
        If CFG_ASM_SELF Then
            If StrComp(sP, sAsmPath, vbTextCompare) <> 0 Then G_AddModel sP, colModels, nPart
        End If
    End Select
End Sub

' 把"状态不正常"的组件按类别先存着，等用户在勾选框里拍板。
' 为什么不当场记成"跳过"：勾选框里要显示每一类的实际个数（扫完才知道），
' 且用户勾了之后还要把这一类放回来 —— 扫的时候就钉死就来不及了。
Private Sub G_StageSkipped(ByVal nKind As Long, ByVal comp As Object, ByVal sWhy As String)
    Dim sLabel As String, sP As String
    sLabel = G_CompLabel(comp, sP)
    If Len(sLabel) = 0 Then Exit Sub
    Select Case nKind
        Case 1: m_ListEnvelope.Add Array(sLabel, sP, sWhy)
        Case 2: m_ListSuppressed.Add Array(sLabel, sP, sWhy)
        Case 4: m_ListHidden.Add Array(sLabel, sP, sWhy)
    End Select
End Sub

' 用户拍板之后（或者没开询问时）：把勾中的类别放回待导出列表，没勾的才写成"跳过"记录。
' 注意"跳过"记录是**延后到这里**才写的，早写会出两个问题：
'   ① 用户勾了的那一类会被误报成"已跳过"；
'   ② 同一零件的另一个可见实例若已导出，会留下"一边成功、一边跳过"的矛盾记录。
Private Sub G_MergeSkipped(ByVal nMask As Long, ByVal colModels As Collection)
    G_MergeOne m_ListEnvelope, 1, nMask, "封套", colModels
    G_MergeOne m_ListSuppressed, 2, nMask, "已压缩", colModels
    G_MergeOne m_ListHidden, 4, nMask, "已隐藏", colModels
End Sub

Private Sub G_MergeOne(ByVal col As Collection, ByVal nBit As Long, ByVal nMask As Long, _
                       ByVal sName As String, ByVal colModels As Collection)
    Dim i As Long, it As Variant, sLabel As String, sP As String, n As Long
    For i = 1 To col.Count
        it = col(i)
        sLabel = it(0)
        sP = it(1)
        If (nMask And nBit) <> 0 Then
            ' 用户勾了这一类 → 放回待导出列表
            If Len(sP) > 0 Then
                G_AddModel sP, colModels, n
            Else
                ' 取不到磁盘路径（典型：压缩的组件根本没加载）。勾了也导不出来，如实写明。
                U_AddResult "模型", sLabel, "", "跳过", sName & "，但取不到文件路径，无法导出"
            End If
        ElseIf Len(sP) > 0 And F_Contains(colModels, sP) Then
            ' 同一零件的另一个实例可见、已进待导出列表 → 这条跳过记录不写，免得自相矛盾
        Else
            G_NoteSkip sLabel, it(2)
        End If
    Next i
End Sub

'---- 3.5.3 装配体展开 ----
' 遍历装配体（含各级子装配体）把**零件**收进 colModels。
' 两道筛子：① 隐藏/压缩/封套（含父级链） ② 标准件关键字
' 返回：本次新增的零件个数（默认 CFG_ASM_SELF = False，不含装配体自身）
Private Function G_ExpandAssembly(ByVal d As Object, ByVal sAsmPath As String, _
                                  ByVal colModels As Collection) As Long
    Dim vComps As Variant, i As Long, comp As Object
    Dim sP As String, sR As String, sLbl As String
    Dim nPart As Long, nSkipped As Long, nKind As Long
    Dim bStateCheck As Boolean

    ' 三种决策来源任意一个开着，就得去读组件状态；
    ' 全关着时连这一步都省掉 —— 状态判断要逐个组件调 3 个属性再回溯父级，不是免费的。
    bStateCheck = (CFG_SKIP_HIDDEN Or CFG_SKIP_SUPPRESSED Or CFG_SKIP_ENVELOPE Or CFG_ASK_SKIP_OPTIONS)

    ' 轻化组件先还原，否则 GetComponents 可能漏件、也取不到路径
    On Error Resume Next
    d.ResolveAllLightWeightComponents True
    Err.Clear
    On Error GoTo 0

    ' False = 展开所有层级（含子装配体内部的组件）
    On Error Resume Next
    vComps = d.GetComponents(False)
    Err.Clear
    On Error GoTo 0

    If IsArray(vComps) Then
        For i = LBound(vComps) To UBound(vComps)
            Set comp = Nothing
            On Error Resume Next
            Set comp = vComps(i)
            Err.Clear
            On Error GoTo 0

            If Not comp Is Nothing Then
                nKind = 0
                sR = ""
                If bStateCheck Then G_SkipKind comp, nKind, sR

                If nKind <> 0 And CFG_ASK_SKIP_OPTIONS Then
                    ' 交给用户拍板：先按类别存着，一个都不写"跳过"
                    G_StageSkipped nKind, comp, sR
                ElseIf nKind <> 0 And G_KindSkipped(nKind) Then
                    ' 不询问用户，且配置说这一类要跳过
                    sLbl = G_CompLabel(comp, sP)
                    G_NoteSkip sLbl, sR
                    nSkipped = nSkipped + 1
                Else
                    ' 状态正常，或者配置明确说这一类要导（比如 CFG_SKIP_HIDDEN 关掉了）
                    G_CollectOne comp, sAsmPath, colModels, nPart, nSkipped
                End If
            End If
        Next i
    End If

    ' 兜底通道：仅当"一个模型都没收到、且没有任何组件被主动跳过或暂存"时才启用。
    ' 后半句很重要 —— 否则整个装配体恰好全被过滤时，兜底会把它们又捞回来，白过滤一场。
    If nPart = 0 And nSkipped = 0 And G_StagedCount() = 0 Then G_DepsOfAssembly sAsmPath, colModels
    G_ExpandAssembly = nPart
End Function

' 暂存区里一共有多少个"等用户拍板"的组件
Private Function G_StagedCount() As Long
    Dim n As Long
    If Not m_ListHidden Is Nothing Then n = n + m_ListHidden.Count
    If Not m_ListSuppressed Is Nothing Then n = n + m_ListSuppressed.Count
    If Not m_ListEnvelope Is Nothing Then n = n + m_ListEnvelope.Count
    G_StagedCount = n
End Function

' 兜底通道：装配体打开异常 / GetComponents 空返时，读文件依赖表（不需要打开文档）
' ！注意：这条通道拿不到 Component2 对象，所以**做不了**参与状态过滤（隐藏/压缩/封套），
'   只能做标准件的"路径 / 文件名"判定。它只在主通道完全失效的异常情况下才会走到。
Private Function G_DepsOfAssembly(ByVal sAsmPath As String, ByVal colModels As Collection) As Long
    Dim vDeps As Variant, i As Long, sItem As String, n As Long, sR As String
    On Error Resume Next
    vDeps = m_swApp.GetDocumentDependencies2(sAsmPath, True, False, False)
    Err.Clear
    On Error GoTo 0
    If Not IsArray(vDeps) Then Exit Function
    For i = LBound(vDeps) To UBound(vDeps)
        sItem = CStr(vDeps(i))
        If Len(sItem) > 0 And InStr(sItem, "\") > 0 Then
            Select Case U_Ext(sItem)
            Case ".SLDPRT"
                If G_IsStdPart(sItem, Nothing, sR) Then
                    G_NoteStd sItem, sR
                Else
                    G_AddModel sItem, colModels, n
                End If
            Case ".SLDASM"
                If CFG_ASM_SELF Then
                    If StrComp(sItem, sAsmPath, vbTextCompare) <> 0 Then G_AddModel sItem, colModels, n
                End If
            End Select
        End If
    Next i
    G_DepsOfAssembly = n
End Function

'---- 3.5.4 入口：按"当前激活的文档"收集 ----
' 返回 False 表示无法处理或用户取消
Private Function G_CollectActive(ByRef colModels As Collection, _
                                 ByRef colDrawings As Collection, _
                                 ByRef sDesc As String, ByRef sExtra As String) As Boolean
    Dim d As Object, sPath As String, sR As String, n As Long, t As Long
    Dim nMask As Long, nAns As Long

    Set colModels = New Collection
    Set colDrawings = New Collection
    sDesc = ""
    sExtra = ""

    Set d = m_swApp.ActiveDoc
    If d Is Nothing Then
        MsgBox "当前没有打开的文档。" & vbCrLf & vbCrLf & _
               "这个宏要读「当前激活的那个窗口」，请先打开一个零件 / 装配体 / 工程图；" & vbCrLf & _
               "想手动挑文件的话，请改用 SwAuto_PickFiles。", vbExclamation, "SWauto"
        Exit Function
    End If
    sPath = d.GetPathName
    If Len(sPath) = 0 Then
        MsgBox "当前文档还没有保存到磁盘，请先保存后再运行。", vbExclamation, "SWauto"
        Exit Function
    End If

    ' 导出清单的"来源名"：三种文档类型（零件/装配体/工程图）都用当前文档名。
    ' 放在这里一次设好，省得三个分支里各写一遍。
    m_SrcName = U_BaseName(sPath)

    t = d.GetType
    Select Case t

    Case swDocumentTypes_e.swDocPART
        sDesc = "零件：" & U_FileName(sPath)
        If G_IsStdPart(sPath, d, sR) Then
            G_NoteStd sPath, sR
            sExtra = "该零件被判定为标准件（" & sR & "），已跳过。" & vbCrLf & _
                     "如果它其实需要导出，请到配置第 1.6 节删掉对应关键字。"
        Else
            colModels.Add sPath
            sExtra = "本零件 → STEP；它对应的工程图 → PDF。"
        End If

    Case swDocumentTypes_e.swDocASSEMBLY
        sDesc = "装配体：" & U_FileName(sPath)
        If CFG_ASM_EXPAND Then
            If CFG_ASM_SELF Then G_AddModel sPath, colModels, n
            n = G_ExpandAssembly(d, sPath, colModels)

            ' 把"隐藏 / 压缩 / 封套"三类逐行列出来，由用户当场决定要哪几类。
            ' 只有装配体才问 —— 零件和工程图根本没有"隐藏的零件"这回事。
            nMask = 0
            If CFG_ASK_SKIP_OPTIONS And G_StagedCount() > 0 Then
                nAns = U_AskSkipOptions(m_ListHidden.Count, m_ListSuppressed.Count, m_ListEnvelope.Count)
                If nAns = -1 Then
                    G_CollectActive = False              ' 用户点了「取消导出」
                    Exit Function
                End If
                ' nAns = -2 表示对话框弹不出来（比如 PowerShell 被安全策略拦了）。
                ' 这种情况**不能**中断导出，保持 nMask = 0 按"三类都不导"的默认继续，
                ' 免得一个界面问题把整个功能卡死。
                If nAns > 0 Then nMask = nAns
            End If
            G_MergeSkipped nMask, colModels

            sExtra = "装配体已展开，找到 " & colModels.Count & " 个待导出的零件。"
            ' 把筛子各过滤掉多少写清楚 —— 万一用户觉得"怎么少了个零件"，
            ' 光看这个确认框就能自己判断是不是配置问题，不用去翻清单。
            If m_SkippedState.Count > 0 Then
                sExtra = sExtra & vbCrLf & "另有 " & m_SkippedState.Count & " 个组件因【隐藏/压缩/封套】被跳过。"
            End If
            If m_SkippedStd.Count > 0 Then
                sExtra = sExtra & vbCrLf & "另有 " & m_SkippedStd.Count & " 个标准件被跳过。"
            End If
            If Not CFG_ASM_SELF Then
                sExtra = sExtra & vbCrLf & "装配体 / 子装配体本身不导出（配置 1.5）。"
            End If
        Else
            colModels.Add sPath
            sExtra = "装配体未展开（配置 1.5 CFG_ASM_EXPAND = False），只导出装配体本身。"
        End If

    Case swDocumentTypes_e.swDocDRAWING
        sDesc = "工程图：" & U_FileName(sPath)
        colDrawings.Add sPath
        sExtra = "本工程图 → PDF；它引用的模型会一并导出 STEP。"

    Case Else
        MsgBox "当前文档不是 SolidWorks 零件 / 装配体 / 工程图，无法处理。", vbExclamation, "SWauto"
        Exit Function

    End Select

    ' 确认框里标清楚是哪种模式，避免用户点错宏时看不出来
    sDesc = "【当前激活文档】" & sDesc
    G_CollectActive = True
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
    Dim d As Object, bOurs As Boolean, sOut As String, sMsg As String, sKey As String
    If Not U_FileExists(sPath) Then
        U_AddResult "模型", sPath, "", "失败", "源文件不存在"
        Exit Sub
    End If
    ' 同一物理文件本次只导一次。
    ' 这是最后一道保险：即使前面因为路径写法不同（绝对/相对、大小写差异）
    ' 把同一个零件收录了两遍，这里也会挡住，不会再产出 零件.STEP + 零件_2.STEP。
    sKey = U_FileKey(sPath)
    If m_DoneSrc.Exists(sKey) Then
        U_AddResult "模型", sPath, m_DoneSrc(sKey), "跳过", "本次已导出过，跳过重复项"
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
        m_DoneSrc(sKey) = sOut
        U_AddResult "模型", sPath, sOut, "成功", sMsg
    Else
        U_AddResult "模型", sPath, sOut, "失败", sMsg
    End If
    X_Close sPath, bOurs
End Sub

'---- 4.5 处理一个工程图文件：导出 PDF ----
Private Sub X_DoDrawing(ByVal sPath As String)
    Dim d As Object, bOurs As Boolean, sOut As String, sMsg As String, sKey As String
    If Not U_FileExists(sPath) Then
        U_AddResult "图纸", sPath, "", "失败", "源文件不存在"
        Exit Sub
    End If
    sKey = U_FileKey(sPath)
    If m_DoneSrc.Exists(sKey) Then
        U_AddResult "图纸", sPath, m_DoneSrc(sKey), "跳过", "本次已导出过，跳过重复项"
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
        m_DoneSrc(sKey) = sOut
        U_AddResult "图纸", sPath, sOut, "成功", sMsg
    Else
        U_AddResult "图纸", sPath, sOut, "失败", sMsg
    End If
    X_Close sPath, bOurs
End Sub


'==========================================================================================
'  第 5 节  入口程序
'    SwAuto_ExportActiveDoc   【模式 1/3 · 当前文档】主入口，按当前激活的文档导出
'    SwAuto_PickFiles         【模式 2/3 · 自选文件】弹文件对话框，多选文件
'    SwAuto_PickFolder        【模式 3/3 · 文件夹批量】弹文件夹对话框，批量导出
'    SwAuto_Help              【使用说明】宏一览 + 快捷键绑定方法
'------------------------------------------------------------------------------------------
'  每个入口第一件事都是设置 m_ModeTag，之后所有弹窗标题都会带上模式标识。
'  这样"快捷键绑到了别的宏"这种误操作，按下的一瞬间就能看出来（而不是导错了才发现）。
'==========================================================================================

'---- 5.0.1 主入口：当前激活的文档 ----
'   零件   -> 本零件 STEP + 对应工程图 PDF
'   装配体 -> 装配体下所有零件 STEP + 各自工程图 PDF（标准件自动过滤）
'   工程图 -> 本图 PDF + 它引用的模型 STEP
Public Sub SwAuto_ExportActiveDoc()
    Dim colModels As Collection, colDrawings As Collection
    Dim sDesc As String, sExtra As String
    If Not M_Init() Then Exit Sub
    m_ModeTag = "1/3 · 当前文档"
    If Not G_CollectActive(colModels, colDrawings, sDesc, sExtra) Then
        M_Cleanup
        Exit Sub
    End If
    If colModels.Count = 0 And colDrawings.Count = 0 Then
        MsgBox "没有需要导出的内容。" & vbCrLf & vbCrLf & sExtra, vbInformation, M_Caption()
        M_Cleanup
        Exit Sub
    End If
    M_RunPlan colModels, colDrawings, sDesc, sExtra
    M_Cleanup
End Sub

'---- 5.0.2 手动多选文件 ----
Public Sub SwAuto_PickFiles()
    Dim v As Variant, col As Collection, i As Long
    If Not M_Init() Then Exit Sub
    m_ModeTag = "2/3 · 自选文件"
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
        MsgBox "没有选中可处理的 SolidWorks 文件（.sldprt / .sldasm / .slddrw）。", vbExclamation, M_Caption()
        M_Cleanup
        Exit Sub
    End If
    Cfg_PutSetting CFG_KEY_IN, U_Parent(col(1))
    m_SrcName = "多选文件"                       ' 多选没有单一来源，清单名就用这个
    M_Run col, "已选 " & col.Count & " 个文件"
    M_Cleanup
End Sub

'---- 5.0.3 整个文件夹 ----
Public Sub SwAuto_PickFolder()
    Dim sDir As String, col As Collection, d As Object, sAct As String
    If Not M_Init() Then Exit Sub
    m_ModeTag = "3/3 · 文件夹批量"

    ' 防呆：当前正开着文档时先确认一次 ——
    ' 「只想导当前打开的这个模型」却点了文件夹模式，是实际发生过的误操作
    ' （多数是因为快捷键绑到了本宏，而不是 SwAuto_ExportActiveDoc）。
    ' 默认按钮特意设为「否」：回车 = 取消，不会误扫整个文件夹。
    Set d = m_swApp.ActiveDoc
    If Not d Is Nothing Then sAct = d.GetPathName
    If Len(sAct) > 0 Then
        If MsgBox("你现在运行的是【文件夹批量模式】：接下来先选源文件夹，再选输出目录。" & vbCrLf & vbCrLf & _
                  "如果只想导出当前打开的这个 —— " & U_FileName(sAct) & " ——" & vbCrLf & _
                  "请点「否」，然后运行 SwAuto_ExportActiveDoc 这个宏。" & vbCrLf & vbCrLf & _
                  "！如果这个框是「按快捷键」弹出来的，说明快捷键绑错了，" & vbCrLf & _
                  "   正确的绑法见宏 SwAuto_Help。" & vbCrLf & vbCrLf & _
                  "确实要扫描整个文件夹吗？", _
                  vbYesNo + vbQuestion + vbDefaultButton2, M_Caption()) <> vbYes Then
            M_Cleanup
            Exit Sub
        End If
    End If

    sDir = U_PickFolder("【" & m_ModeTag & "】请选择要导出的源文件夹", Cfg_GetSetting(CFG_KEY_IN))
    If Len(sDir) = 0 Then
        M_Cleanup
        Exit Sub
    End If
    Cfg_PutSetting CFG_KEY_IN, sDir
    m_SrcName = U_FileName(sDir)                 ' 文件夹模式用源文件夹名给清单命名
    Set col = M_ScanFolder(sDir)
    If col.Count = 0 Then
        MsgBox "该文件夹里没有找到 .sldprt / .sldasm / .slddrw 文件。" & vbCrLf & _
               "（如需包含子目录，请把 CFG_RECURSE_SUBDIR 改为 True）", vbExclamation, M_Caption()
        M_Cleanup
        Exit Sub
    End If
    M_Run col, "文件夹：" & sDir
    M_Cleanup
End Sub

'---- 5.0.4 使用说明（宏一览 + 快捷键绑定方法）----
'  特意做成"不依赖 SolidWorks 连接"的纯说明，任何时候按下去都能看。
'  定位：一个"忘了哪个宏该绑快捷键"时的自救入口。
Public Sub SwAuto_Help()
    Dim s As String, sAct As String
    On Error Resume Next
    If m_swApp Is Nothing Then Set m_swApp = Application.SldWorks
    sAct = m_swApp.ActiveDoc.GetTitle          ' 没打开文档时会出错，下面兜住
    If Err.Number <> 0 Then
        Err.Clear
        sAct = "（当前没有打开任何文档）"
    End If
    On Error GoTo 0

    s = "SWautoExport " & CFG_VER & "　宏一览与绑定方法" & vbCrLf & String(44, "=") & vbCrLf & vbCrLf
    s = s & "① SwAuto_ExportActiveDoc　【模式 1/3 · 当前文档】← 平时只用这个" & vbCrLf
    s = s & "　　按当前激活的文档导出：" & vbCrLf
    s = s & "　　　零件　 → 本零件 STEP + 对应工程图 PDF" & vbCrLf
    s = s & "　　　装配体 → 展开它下面的所有零件，各出一份 STEP + PDF" & vbCrLf
    s = s & "　　　　　　　装配体 / 子装配体本身不导出（只交付零件）" & vbCrLf
    s = s & "　　　　　　　自动跳过：螺钉螺母等标准件" & vbCrLf
    s = s & "　　　　　　　隐藏 / 压缩 / 封套：展开后弹框逐类问你，你勾哪类就导哪类" & vbCrLf
    s = s & "　　　工程图 → 本图 PDF + 它引用的模型 STEP" & vbCrLf & vbCrLf
    s = s & "② SwAuto_PickFiles　　【模式 2/3 · 自选文件】只导我挑出来的那几个" & vbCrLf
    s = s & "③ SwAuto_PickFolder　【模式 3/3 · 文件夹批量】整个文件夹批量导" & vbCrLf & vbCrLf
    s = s & "你现在打开的是：" & sAct & vbCrLf & vbCrLf
    s = s & String(44, "-") & vbCrLf
    s = s & "【怎么把宏绑到快捷键】" & vbCrLf
    s = s & "　工具 > 自定义 > 键盘 > 类别选「宏」" & vbCrLf
    s = s & "　输入框里按快捷键 > 点「指派」" & vbCrLf & vbCrLf
    s = s & "　！若「命令」列里还没有这个宏：点「新建宏按钮」那一行，" & vbCrLf
    s = s & "　　 在「操作」里打开你的 .swp 宏工程（你这边是 Macro1.swp），" & vbCrLf
    s = s & "　　 然后——" & vbCrLf
    s = s & "　　 ★ 把【方法】下拉框选成 SwAuto_ExportActiveDoc ★" & vbCrLf
    s = s & "　　 这个下拉框最容易被漏掉。漏掉它，SolidWorks 会自己挑一个" & vbCrLf
    s = s & "　　 方法运行，就会变成「按快捷键却弹出文件夹批量模式」。" & vbCrLf & vbCrLf
    s = s & String(44, "-") & vbCrLf
    s = s & "【怎么确认没绑错】按一下快捷键，看弹窗标题：" & vbCrLf
    s = s & "　标题带 1/3 · 当前文档　 → 绑对了" & vbCrLf
    s = s & "　标题带 3/3 · 文件夹批量 → 绑错了，按上面重绑一次" & vbCrLf
    MsgBox s, vbInformation, "SWauto 使用说明"
End Sub

'---- 5.1 路径列表入口（文件多选 / 文件夹）→ 分类后交给 M_RunPlan ----
Private Sub M_Run(ByVal colPaths As Collection, ByVal sDesc As String)
    Dim colModels As New Collection, colDrawings As New Collection
    Dim i As Long, sPath As String
    For i = 1 To colPaths.Count
        sPath = colPaths(i)
        If U_IsModel(sPath) Then
            If Not F_Contains(colModels, sPath) Then colModels.Add sPath
        ElseIf U_IsDrawing(sPath) Then
            If Not F_Contains(colDrawings, sPath) Then colDrawings.Add sPath
        End If
    Next i
    M_RunPlan colModels, colDrawings, sDesc, ""
End Sub

'---- 5.2 统一执行管线（三个入口都走这里）----
'   colModels / colDrawings 由调用方准备好，本过程会就地补充（查找工程图 / 反查模型）
Private Sub M_RunPlan(ByVal colModels As Collection, ByVal colDrawings As Collection, _
                      ByVal sDesc As String, ByVal sExtra As String)
    Dim colDep As Collection
    Dim i As Long, j As Long, sSeed As String, sFind As String, sDrw As String
    Dim dicName As Object

    If colModels.Count = 0 And colDrawings.Count = 0 Then Exit Sub

    ' (1) 输出目录：默认上次使用的目录
    If colModels.Count > 0 Then sSeed = U_Parent(colModels(1)) Else sSeed = U_Parent(colDrawings(1))
    m_OutRoot = M_AskOutDir(sSeed)
    If Len(m_OutRoot) = 0 Then Exit Sub              ' 用户取消
    If Not U_MakeDir(m_OutRoot) Then
        ' 特意把长度也显示出来：路径里混进不可见字符（换行 / BOM）时，
        ' 光看路径本身完全正常，只有长度对不上 —— v2.2.1 就是栽在这上面，
        ' 当时弹窗只给了个看着没问题的路径，无从下手。
        U_Diag "无法创建输出目录。路径=[" & m_OutRoot & "]  长度=" & Len(m_OutRoot)
        MsgBox "无法创建输出目录：" & vbCrLf & m_OutRoot & vbCrLf & vbCrLf & _
               "路径长度：" & Len(m_OutRoot) & " 个字符。" & vbCrLf & _
               "若路径看着完全正常，请把这段连同 %TEMP%\swauto_diag.txt 一起反馈。", _
               vbCritical, "SWauto"
        Exit Sub
    End If
    If CFG_USE_SUBFOLDER Then
        m_StepDir = U_JoinPath(m_OutRoot, "STEP")
        m_PdfDir = U_JoinPath(m_OutRoot, "PDF")
        U_MakeDir m_StepDir
        U_MakeDir m_PdfDir
    Else
        m_StepDir = m_OutRoot                     ' 默认：STEP 与 PDF 放在同一个目录
        m_PdfDir = m_OutRoot
    End If

    ' (2) 确认
    If Not M_Confirm(sDesc, colModels.Count, colDrawings.Count, sExtra) Then Exit Sub

    ' (3) 为每个模型查找对应工程图
    If CFG_FIND_DRAWING Then
        For i = 1 To colModels.Count
            sDrw = F_Find(colModels(i), sFind)          ' sFind 只是承接"未找到"的原因，可忽略
            If Len(sDrw) > 0 Then
                If Not F_Contains(colDrawings, sDrw) Then colDrawings.Add sDrw
            Else
                If Not F_Contains(m_NoDrawing, colModels(i)) Then m_NoDrawing.Add colModels(i)
            End If
        Next i
    End If

    ' (4) 工程图反查它引用的模型
    '     反查属于"顺便带上"，所以如果待导出列表里**已经有同名模型**，就不再追加。
    '     工程图里记录的引用路径有时是旧位置、或是另一种写法，会和扫描结果指到同一个零件；
    '     无脑追加就会出现 零件.STEP + 零件_2.STEP 这种看着像 bug 的重复文件。
    '     （真需要两个不同目录的同名零件时，请用 SwAuto_PickFiles 明确挑出来，那种情况不会被合并）
    If CFG_DRW_ALSO_MODEL Then
        On Error Resume Next
        Set dicName = CreateObject("Scripting.Dictionary")
        dicName.CompareMode = 1
        On Error GoTo 0
        For i = 1 To colModels.Count
            dicName(UCase$(U_FileName(colModels(i)))) = True
        Next i
        For i = 1 To colDrawings.Count
            Set colDep = F_ModelsOfDrawing(colDrawings(i))
            For j = 1 To colDep.Count
                If Not F_Contains(colModels, colDep(j)) Then
                    If Not dicName.Exists(UCase$(U_FileName(colDep(j)))) Then
                        colModels.Add colDep(j)
                        dicName(UCase$(U_FileName(colDep(j)))) = True
                    End If
                End If
            Next j
        Next i
        Set dicName = Nothing
    End If

    ' (5) 执行
    For i = 1 To colModels.Count
        X_DoModel colModels(i)
    Next i
    For i = 1 To colDrawings.Count
        X_DoDrawing colDrawings(i)
    Next i

    ' (6) 清单 + 汇总
    If CFG_WRITE_REPORT Then M_WriteReport
    M_RestoreActive
    M_Summary
End Sub

'---- 5.2.1 弹窗标题：始终带上模式标识 ----
'  同一次运行里的所有弹窗标题都形如「SWauto v2.0.2　1/3 · 当前文档」。
'  意义：快捷键若被绑到了别的入口，按下后第一眼就能发现，而不是导完才发现导错了范围。
Private Function M_Caption() As String
    If Len(m_ModeTag) > 0 Then
        M_Caption = "SWauto " & CFG_VER & "　" & m_ModeTag
    Else
        M_Caption = "SWauto " & CFG_VER
    End If
End Function

'---- 5.3 导出确认 ----
Private Function M_Confirm(ByVal sDesc As String, ByVal nModel As Long, _
                           ByVal nDrw As Long, ByVal sExtra As String) As Boolean
    Dim s As String
    s = "模式：" & m_ModeTag & vbCrLf
    s = s & "来源：" & sDesc & vbCrLf
    If Len(sExtra) > 0 Then s = s & sExtra & vbCrLf
    s = s & vbCrLf
    s = s & "模型 " & nModel & " 个  →  STEP (AP" & CFG_STEP_AP & ")" & vbCrLf
    If nDrw > 0 Then s = s & "工程图 " & nDrw & " 个  →  PDF" & vbCrLf
    If CFG_FIND_DRAWING And nModel > 0 Then s = s & "（每个模型还会自动查找它的工程图，一起导出 PDF）" & vbCrLf
    If m_SkippedStd.Count > 0 Then s = s & "已过滤标准件 " & m_SkippedStd.Count & " 个（螺钉螺母等，详见清单）" & vbCrLf
    If m_SkippedState.Count > 0 Then s = s & "已跳过 隐藏/压缩/封套 组件 " & m_SkippedState.Count & " 个（详见清单）" & vbCrLf
    s = s & vbCrLf & "输出目录：" & m_OutRoot & vbCrLf & vbCrLf & "开始导出？"
    M_Confirm = (MsgBox(s, vbYesNo + vbQuestion, M_Caption()) = vbYes)
End Function

'---- 5.4 输出目录选择 ----
Private Function M_AskOutDir(ByVal sSeed As String) As String
    Dim sDef As String
    sDef = Cfg_GetSetting(CFG_KEY_OUT)
    If Not U_DirExists(sDef) Then sDef = sSeed
    If Not U_DirExists(sDef) Then sDef = ""
    M_AskOutDir = U_PickFolder("【" & m_ModeTag & "】请选择导出文件的输出目录", sDef)
    If Len(M_AskOutDir) > 0 Then Cfg_PutSetting CFG_KEY_OUT, M_AskOutDir
End Function

'---- 5.5 扫描文件夹 ----
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

'---- 5.6 导出清单 ----

' 清单文件的完整路径。
' 默认带上"来源名"，例如：家庭服务机器人-重组_导出清单.csv
' 好处是同一个输出目录里放过好几轮导出时，一眼就能看出每份清单各是哪个文档 / 文件夹导出的，
' 也不会互相覆盖（固定名 _导出清单.csv 是会被后一轮盖掉的）。
Private Function M_ReportFile() As String
    Dim sName As String
    If CFG_REPORT_WITH_SRC And Len(m_SrcName) > 0 Then
        sName = U_SafeName(m_SrcName) & CFG_REPORT_SUFFIX & ".csv"
    Else
        sName = CFG_REPORT_SUFFIX & ".csv"       ' 固定名：_导出清单.csv
    End If
    M_ReportFile = U_JoinPath(m_OutRoot, sName)
End Function

Private Sub M_WriteReport()
    Dim sb As String, i As Long, it As Variant
    sb = "SWauto 导出清单" & vbCrLf
    sb = sb & U_Csv("程序版本") & "," & U_Csv(CFG_VER) & vbCrLf
    sb = sb & U_Csv("运行模式") & "," & U_Csv(m_ModeTag) & vbCrLf
    sb = sb & U_Csv("导出时间") & "," & U_Csv(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & vbCrLf
    sb = sb & U_Csv("输出目录") & "," & U_Csv(m_OutRoot) & vbCrLf
    sb = sb & U_Csv("STEP 协议") & "," & U_Csv("AP" & CStr(CFG_STEP_AP)) & vbCrLf
    sb = sb & U_Csv("模型 STEP 成功") & "," & U_CountBy("模型", "成功") & vbCrLf
    sb = sb & U_Csv("模型 STEP 失败") & "," & U_CountBy("模型", "失败") & vbCrLf
    sb = sb & U_Csv("模型 STEP 跳过") & "," & U_CountBy("模型", "跳过") & vbCrLf
    sb = sb & U_Csv("工程图 PDF 成功") & "," & U_CountBy("图纸", "成功") & vbCrLf
    sb = sb & U_Csv("工程图 PDF 失败") & "," & U_CountBy("图纸", "失败") & vbCrLf
    sb = sb & U_Csv("工程图 PDF 跳过") & "," & U_CountBy("图纸", "跳过") & vbCrLf
    sb = sb & U_Csv("标准件已过滤") & "," & m_SkippedStd.Count & vbCrLf
    sb = sb & U_Csv("隐藏/压缩/封套已跳过") & "," & m_SkippedState.Count & vbCrLf
    sb = sb & U_Csv("未找到工程图的模型") & "," & m_NoDrawing.Count & vbCrLf
    sb = sb & vbCrLf
    sb = sb & "序号,类别,源文件,输出文件,结果,备注" & vbCrLf
    For i = 1 To m_Results.Count
        it = m_Results(i)
        sb = sb & i & "," & U_Csv(it(0)) & "," & U_Csv(it(1)) & "," & _
             U_Csv(it(2)) & "," & U_Csv(it(3)) & "," & U_Csv(it(4)) & vbCrLf
    Next i
    sb = sb & vbCrLf & U_Csv("以下组件因【隐藏 / 压缩 / 封套】被跳过（想导出请改配置第 1.7 节）") & vbCrLf
    If m_SkippedState.Count = 0 Then
        sb = sb & U_Csv("（无）") & vbCrLf
    Else
        For i = 1 To m_SkippedState.Count
            ' 第二列写清走的是哪条规则，误判时一眼能看出该改哪儿
            If m_SkipWhy Is Nothing Then
                sb = sb & U_Csv(m_SkippedState(i)) & vbCrLf
            ElseIf m_SkipWhy.Exists(m_SkippedState(i)) Then
                sb = sb & U_Csv(m_SkippedState(i)) & "," & U_Csv(m_SkipWhy(m_SkippedState(i))) & vbCrLf
            Else
                sb = sb & U_Csv(m_SkippedState(i)) & vbCrLf
            End If
        Next i
    End If
    sb = sb & vbCrLf & U_Csv("以下零件被判定为标准件而跳过（想导出请改配置第 1.6 节）") & vbCrLf
    If m_SkippedStd.Count = 0 Then
        sb = sb & U_Csv("（无）") & vbCrLf
    Else
        For i = 1 To m_SkippedStd.Count
            sb = sb & U_Csv(m_SkippedStd(i)) & vbCrLf
        Next i
    End If
    sb = sb & vbCrLf & U_Csv("以下模型未找到对应工程图") & vbCrLf
    If m_NoDrawing.Count = 0 Then
        sb = sb & U_Csv("（无）") & vbCrLf
    Else
        For i = 1 To m_NoDrawing.Count
            sb = sb & U_Csv(m_NoDrawing(i)) & vbCrLf
        Next i
    End If
    U_WriteText M_ReportFile(), sb
End Sub

'---- 5.7 汇总弹窗 ----
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

    If m_SkippedState.Count > 0 Then
        s = s & vbCrLf & vbCrLf & "已跳过隐藏/压缩/封套组件（" & m_SkippedState.Count & " 个）：" & vbCrLf & _
            U_BriefList(m_SkippedState, 6)
    End If

    If m_SkippedStd.Count > 0 Then
        s = s & vbCrLf & vbCrLf & "已自动过滤标准件（" & m_SkippedStd.Count & " 个）：" & vbCrLf & _
            U_BriefList(m_SkippedStd, 6)
    End If

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
        s = s & vbCrLf & "明细清单：" & M_ReportFile()
    End If
    MsgBox s, IIf(nFail > 0, vbExclamation, vbInformation), M_Caption()
End Sub

'---- 5.8 初始化 / 收尾 ----
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
    Set m_DoneSrc = CreateObject("Scripting.Dictionary")
    m_DoneSrc.CompareMode = 1
    Set m_SkipWhy = CreateObject("Scripting.Dictionary")
    m_SkipWhy.CompareMode = 1                    ' 1 = 不区分大小写
    On Error GoTo 0
    Set m_Results = New Collection
    Set m_NoDrawing = New Collection
    Set m_SkippedStd = New Collection
    Set m_SkippedState = New Collection
    Set m_ListHidden = New Collection
    Set m_ListSuppressed = New Collection
    Set m_ListEnvelope = New Collection
    Set m_IdxByPath = Nothing
    Set m_IdxByName = Nothing
    m_IdxBuilt = False
    m_OutRoot = ""
    m_StepDir = ""
    m_PdfDir = ""
    m_SrcName = ""                               ' 由各入口 / G_CollectActive 按来源设置
    m_OriginalActive = ""
    If Not m_swApp.ActiveDoc Is Nothing Then m_OriginalActive = m_swApp.ActiveDoc.GetTitle
    M_Init = True
End Function

Private Sub M_Cleanup()
    Set m_Results = Nothing
    Set m_NoDrawing = Nothing
    Set m_SkippedStd = Nothing
    Set m_SkippedState = Nothing
    Set m_SkipWhy = Nothing
    Set m_ListHidden = Nothing
    Set m_ListSuppressed = Nothing
    Set m_ListEnvelope = Nothing
    Set m_IdxByPath = Nothing
    Set m_IdxByName = Nothing
    Set m_UsedNames = Nothing
    Set m_DoneSrc = Nothing
    m_SrcName = ""
End Sub

' 还原用户原来的活动文档，避免批量导出后界面停在最后一个文件上
Private Sub M_RestoreActive()
    Dim lErr As Long
    If Len(m_OriginalActive) = 0 Then Exit Sub
    On Error Resume Next
    m_swApp.ActivateDoc3 m_OriginalActive, False, 0, lErr    ' 0 = 激活时不重建
    On Error GoTo 0
End Sub
