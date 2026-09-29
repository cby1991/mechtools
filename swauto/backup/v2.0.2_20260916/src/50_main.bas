
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
    s = s & "　　　　　　　（螺钉螺母等标准件自动过滤）" & vbCrLf
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
        MsgBox "无法创建输出目录：" & m_OutRoot, vbCritical, "SWauto"
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
    sb = sb & U_Csv("未找到工程图的模型") & "," & m_NoDrawing.Count & vbCrLf
    sb = sb & vbCrLf
    sb = sb & "序号,类别,源文件,输出文件,结果,备注" & vbCrLf
    For i = 1 To m_Results.Count
        it = m_Results(i)
        sb = sb & i & "," & U_Csv(it(0)) & "," & U_Csv(it(1)) & "," & _
             U_Csv(it(2)) & "," & U_Csv(it(3)) & "," & U_Csv(it(4)) & vbCrLf
    Next i
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
    U_WriteText U_JoinPath(m_OutRoot, "_导出清单.csv"), sb
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
        s = s & vbCrLf & "明细清单：" & U_JoinPath(m_OutRoot, "_导出清单.csv")
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
    On Error GoTo 0
    Set m_Results = New Collection
    Set m_NoDrawing = New Collection
    Set m_SkippedStd = New Collection
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
    Set m_SkippedStd = Nothing
    Set m_IdxByPath = Nothing
    Set m_IdxByName = Nothing
    Set m_UsedNames = Nothing
    Set m_DoneSrc = Nothing
End Sub

' 还原用户原来的活动文档，避免批量导出后界面停在最后一个文件上
Private Sub M_RestoreActive()
    Dim lErr As Long
    If Len(m_OriginalActive) = 0 Then Exit Sub
    On Error Resume Next
    m_swApp.ActivateDoc3 m_OriginalActive, False, 0, lErr    ' 0 = 激活时不重建
    On Error GoTo 0
End Sub
