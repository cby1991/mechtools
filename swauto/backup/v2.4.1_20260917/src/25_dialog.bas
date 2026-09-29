
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
