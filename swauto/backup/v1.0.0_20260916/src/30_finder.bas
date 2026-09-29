
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
