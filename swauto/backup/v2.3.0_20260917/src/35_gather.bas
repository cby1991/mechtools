
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
