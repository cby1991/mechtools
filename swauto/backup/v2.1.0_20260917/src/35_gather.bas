
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
    ' 同一个零件在装配体里可能有多个实例，各自状态不同。
    ' 某个实例被隐藏/压缩而记了"跳过"，但另一个可见实例把零件带进来了 ——
    ' 此时零件本身是要导出的，之前那条跳过记录就作废，撤掉，免得清单自相矛盾
    ' （一边写"零件成功"、一边写"零件跳过"）。
    G_UnnoteSkip sPath
    If Not F_Contains(colModels, sPath) Then
        colModels.Add sPath
        n = n + 1
    End If
End Sub

' 撤销某个已记录的"状态跳过"
Private Sub G_UnnoteSkip(ByVal sLabel As String)
    Dim i As Long
    If Len(sLabel) = 0 Then Exit Sub
    If m_SkipWhy Is Nothing Then Exit Sub
    If Not m_SkipWhy.Exists(sLabel) Then Exit Sub
    m_SkipWhy.Remove sLabel
    For i = m_SkippedState.Count To 1 Step -1
        If StrComp(m_SkippedState(i), sLabel, vbTextCompare) = 0 Then
            m_SkippedState.Remove i
            Exit For
        End If
    Next i
End Sub

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

' 沿父级链判断：返回非空字符串 = 应当跳过，字符串内容就是原因
Private Function G_SkipReason(ByVal comp As Object) As String
    Dim c As Object, p As Object, sWhy As String, n As Long
    If comp Is Nothing Then Exit Function
    If Not (CFG_SKIP_HIDDEN Or CFG_SKIP_SUPPRESSED Or CFG_SKIP_ENVELOPE) Then Exit Function

    Set c = comp
    Do While Not c Is Nothing
        n = n + 1
        If n > 64 Then Exit Do                 ' 防御：万一父级链出现环，别把 SolidWorks 挂死

        sWhy = G_CompBadState(c)
        If Len(sWhy) > 0 Then
            If n > 1 Then sWhy = sWhy & "（上级装配体）"
            G_SkipReason = sWhy
            Exit Function
        End If

        Set p = Nothing                        ' 先清空：取父级失败时 p 保持 Nothing，循环自然结束
        On Error Resume Next
        Set p = c.GetParent
        Err.Clear
        On Error GoTo 0
        Set c = p
    Loop
End Function

' 单个组件自身的状态；返回非空 = 该组件本身就不该参与导出
Private Function G_CompBadState(ByVal c As Object) As String
    Dim b As Boolean

    If CFG_SKIP_ENVELOPE Then
        b = False
        On Error Resume Next
        b = c.IsEnvelope
        Err.Clear
        On Error GoTo 0
        If b Then
            G_CompBadState = "封套"
            Exit Function
        End If
    End If

    If CFG_SKIP_SUPPRESSED Then
        b = False
        On Error Resume Next
        b = c.IsSuppressed
        Err.Clear
        On Error GoTo 0
        If b Then
            G_CompBadState = "已压缩"
            Exit Function
        End If
    End If

    If CFG_SKIP_HIDDEN Then
        b = False
        On Error Resume Next
        b = c.IsHidden(False)                  ' ★ 必须传 False，理由见上面 ①
        Err.Clear
        On Error GoTo 0
        If b Then
            G_CompBadState = "已隐藏"
            Exit Function
        End If
    End If
End Function

' 取组件的可读标识：优先磁盘路径；取不到（压缩 / 轻化组件可能返回空）就用特征树里的名字。
' 有了它，被跳过的组件至少能在清单里留下名字，而不是只报一个数字。
Private Function G_CompLabel(ByVal comp As Object) As String
    Dim s As String
    If comp Is Nothing Then Exit Function
    On Error Resume Next
    s = comp.GetPathName
    If Len(s) = 0 Then s = comp.Name2
    Err.Clear
    On Error GoTo 0
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

'---- 3.5.3 装配体展开 ----
' 遍历装配体（含各级子装配体）把**零件**收进 colModels。
' 两道筛子：① 隐藏/压缩/封套（含父级链） ② 标准件关键字
' 返回：本次新增的零件个数（默认 CFG_ASM_SELF = False，不含装配体自身）
Private Function G_ExpandAssembly(ByVal d As Object, ByVal sAsmPath As String, _
                                  ByVal colModels As Collection) As Long
    Dim vComps As Variant, i As Long, comp As Object, mdoc As Object
    Dim sP As String, sR As String, nPart As Long, nSkipped As Long

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

                ' ① 参与状态：隐藏 / 压缩 / 封套（含上级装配体）→ 跳过
                sR = G_SkipReason(comp)
                If Len(sR) > 0 Then
                    G_NoteSkip G_CompLabel(comp), sR
                    nSkipped = nSkipped + 1
                Else
                    sP = ""
                    On Error Resume Next
                    sP = comp.GetPathName
                    Err.Clear
                    On Error GoTo 0

                    If Len(sP) > 0 Then
                        Select Case U_Ext(sP)
                        Case ".SLDPRT"
                            ' ② 标准件：Toolbox 标记 / 路径关键字 / 文件名关键字
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
                    Else
                        ' 虚拟零件（几何存在装配体内部）没法单独导出成 STEP，记一笔免得用户以为漏了
                        sP = G_CompLabel(comp)
                        If Len(sP) > 0 Then
                            U_AddResult "模型", sP, "", "跳过", "虚拟零件（保存在装配体内部，无法单独导出）"
                            nSkipped = nSkipped + 1
                        End If
                    End If
                End If
            End If
        Next i
    End If

    ' 兜底通道：仅当"一个模型都没收到、且没有任何组件被主动跳过"时才启用。
    ' 后半句很重要 —— 否则整个装配体恰好全被过滤时，兜底会把它们又捞回来，白过滤一场。
    If nPart = 0 And nSkipped = 0 Then G_DepsOfAssembly sAsmPath, colModels
    G_ExpandAssembly = nPart
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
            sExtra = "装配体已展开，找到 " & colModels.Count & " 个待导出的零件。"
            ' 把两个筛子各过滤掉多少写清楚 —— 万一用户觉得"怎么少了个零件"，
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
