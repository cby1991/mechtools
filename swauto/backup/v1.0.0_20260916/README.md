# SWauto —— SolidWorks 一键导出（模型 → STEP，工程图 → PDF）

一个 SolidWorks VBA 宏：**一键把模型另存为 .stp/.step，把它对应的工程图另存为 .pdf**。
运行时会弹出输出目录选择框，**默认定位到上次使用的目录**。

---

## 1. 文件说明

| 文件 | 用途 |
|---|---|
| `dist/SWautoExport.bas` | **导入用的单文件宏**（把 6 个源码段拼成 1 个模块）。**编码 GBK + CRLF**，给 VBE 的"文件 > 导入文件"用 |
| `dist/SWautoExport_粘贴用.txt` | 同一份代码的 **UTF-8 版**，给"打开→全选→复制→粘贴到 VBE"用（剪贴板是 Unicode，不受代码页影响，最保险） |
| `src/00_header.bas` … `src/50_main.bas` | 分模块源码，按功能分段（头/配置/工具/图纸查找/导出/入口），便于后续维护 |
| `build.py` | 把 `src/*.bas` 拼成上面两个交付文件（GBK + CRLF；顺带检查有没有 GBK 编不出的字符） |
| `check.py` | 结构自检：过程块配平、**引号成对性**、过程重名、类型库枚举清单 |

改动源码后重新生成导入文件：

```
python build.py && python check.py
```

## 2. 安装（一次性）

### 2.1 先删掉旧的
如果之前导入过有问题的版本，先清理：VBE 左侧工程窗口里右键那个模块 → **移除 Module**（不要点"导出"）。

### 2.2 两种导入方式，任选一种

**方式 A：导入文件（推荐）**

1. SolidWorks → 菜单 **工具 > 宏 > 新建**，文件名随便起（如 `SwAutoExport`），保存到固定位置，例如
   `C:\Users\<你>\AppData\Roaming\SolidWorks\SOLIDWORKS <版本>\Macros\`。SolidWorks 会自动打开 VBA 编辑器。
2. VBA 编辑器里：**文件 > 导入文件…** → 选择 `dist/SWautoExport.bas`。
3. **Ctrl+S 保存**。宏工程存成 `.swp`，下次启动 SolidWorks 自动可用。

> ⚠️ `dist/SWautoExport.bas` 是 **GBK 编码**，这是刻意的：VBE 导入 `.bas` 时按**系统 ANSI 代码页**解码
> （简体中文 Windows = GBK）。如果写成 UTF-8，中文会全部乱码，而且中文标点的字节会**吞掉后面的结束引号**，
> 直接报"编译错误：语法错误"。详见第 8 节。

**方式 B：复制粘贴（最保险，不受代码页影响）**

1. 用记事本 / VS Code 打开 `dist/SWautoExport_粘贴用.txt`（UTF-8，中文正常显示）。
2. 全选复制。
3. SolidWorks → 工具 > 宏 > 新建 → VBE 里双击 `模块1` → 全选原有内容后粘贴 → Ctrl+S。

### 2.3 绑一个快捷键
**工具 > 自定义 > 键盘** → 分类选 `宏` → 找到 `SwAuto_ExportActiveDoc` → 设快捷键（比如 `Ctrl+Shift+E`）。
想放工具条：同一对话框的"工具栏"页 → 新建一个工具栏 → 把宏拖上去。

> 若编译报**"用户定义类型未定义"**：VBE → **工具 > 引用** → 勾选
> `SldWorks <版本> Type Library` 与 `SOLIDWORKS <版本> Constant type library`。

## 3. 三个入口宏

| 入口 | 行为 |
|---|---|
| `SwAuto_ExportActiveDoc` | 导出**当前打开的文档**。是模型就出 STEP；是工程图就出 PDF，并把它引用的模型也导出 STEP。 |
| `SwAuto_PickFiles` | 弹原生文件对话框**多选**文件（`.sldprt/.sldasm/.slddrw`）后排入队列逐个导出。 |
| `SwAuto_PickFolder` | 弹文件夹对话框，**批量导出整个文件夹**（是否含子目录由 `CFG_RECURSE_SUBDIR` 控制）。 |

**工程图是怎么找到的？** 复刻 SolidWorks「打开工程图」的查找顺序，三级递进：

1. 模型同目录下的**同名** `.slddrw`；
2. `CFG_EXTRA_DIRS` 里配置的搜索目录中的同名 `.slddrw`；
3. **反查引用关系**：扫描目录里的工程图，用 `GetDocumentDependencies2` 读出它引用了哪些模型
   （这个接口**不需要打开文件**，所以很快），建立 `模型 → 工程图` 索引。批量时索引只建一次。

三级都没找到，不会静默跳过——会写进清单的**"以下模型未找到对应工程图"**段落，并在汇总弹窗里列出来。

## 4. 输出结构

```
<你选的输出目录>\
├─ STEP\                    模型导出的 .STEP
├─ PDF\                     工程图导出的 .PDF
└─ _导出清单.csv            本次运行的完整明细
```

`_导出清单.csv`（UTF-8 BOM，Excel 直接双击可读）开头是统计摘要，然后是明细表：

`序号, 类别, 源文件, 输出文件, 结果, 备注` —— 结果取值 `成功 / 跳过 / 失败`，备注里是错误码或原因；
文件末尾单独列出未找到对应工程图的模型。

## 5. 配置区（文件第 1 节，改完保存即可）

| 常量 | 默认 | 说明 |
|---|---|---|
| `CFG_STEP_AP` | `203` | STEP 协议，可用 `203` 或 `214`（AP214 会带颜色/图层） |
| `CFG_USE_SUBFOLDER` | `True` | 输出目录下自动建 `STEP\`、`PDF\`；`False` 则全部平铺在输出目录 |
| `CFG_OVERWRITE` | `True` | `False` 时同名文件跳过并计入清单 |
| `CFG_FIND_DRAWING` | `True` | 模型导出后是否自动找它的工程图 |
| `CFG_DRW_ALSO_MODEL` | `True` | 处理工程图时是否把它引用的模型也导出 STEP |
| `CFG_DEEP_SEARCH` | `True` | 同名找不到时是否反查引用关系 |
| `CFG_DEEP_LIMIT` | `1500` | 深搜最多扫多少个工程图，防超大目录卡死 |
| `CFG_EXTRA_DIRS` | `""` | 额外搜索目录，多个用 `;` 分隔 |
| `CFG_RECURSE_SUBDIR` | `False` | 文件夹模式是否递归子目录 |
| `CFG_WRITE_REPORT` | `True` | 是否生成 `_导出清单.csv` |
| `CFG_USE_PS_DIALOG` | `True` | 多选文件用 PowerShell 原生对话框 |
| `CFG_QUIET_OPEN` | `True` | 静默打开文档，不弹配置/参考丢失对话框 |

上次用过的目录记在注册表
`HKCU\Software\VB and VBA Program Settings\SWauto\Export`，键 `LastOutputDir` / `LastInputDir`。
想重置就删掉这个键。

## 6. 注意事项与已知限制

- **不做任何源文件的改动**，所有输出都是"另存"，源文件时间戳不变。
- **STEP AP 是全局用户选项**：宏调用 `SetUserPreferenceIntegerValue(swStepAP, 203)` 会改掉
  `工具 > 选项 > 系统选项 > 导出 > STEP` 里的设置（这是 SolidWorks API 的既定行为）。
  如果写入没生效，清单备注里会写明"AP203 未能写入用户选项，已按当前选项导出"。
- **PDF 的线宽、颜色、字体、黑白**遵循 `工具 > 选项 > 系统选项 > 导出 > PDF`；
  多张图纸默认合并成**一个** PDF（`ExportAsSinglePDF = True`）。
- 批量导出时 SolidWorks 界面会短暂无响应，这是宏单线程执行导致的，属正常现象；文件多的目录请耐心等汇总弹窗。
- 深搜反查依赖工程图里记录的引用路径。如果图纸引用的是**已换盘符 / 已改名**的旧路径，反查会落空——
  这种情况请用同名规则，或把图纸所在目录加进 `CFG_EXTRA_DIRS`。
- 多选文件对话框依赖 Windows 自带的 PowerShell + .NET WinForms；若被安全策略拦截，会退化成"没有选中文件"，
  此时改用 `SwAuto_PickFolder` 或 `SwAuto_ExportActiveDoc`。
- 装配体导出前会调用 `ResolveAllLightWeightComponents` 还原轻化零件，避免导出空几何。

## 7. 建议的第一次试跑

先拿一个只有 2～3 个文件的目录跑 `SwAuto_PickFolder`，确认：

1. `STEP\`、`PDF\` 里的文件名和数量符合预期；
2. 打开 `_导出清单.csv`，看"结果"列有没有 `失败`；
3. 看清单末尾的"未找到对应工程图"列表是不是符合实际（这几个模型确实没有工程图）。

确认无误后再对正式目录批量处理。

## 8. 故障排查

### 8.1 导入后满屏乱码 + 编译报"语法错误"

**同一件事的两个症状，根因是文件编码。** VBE 导入 `.bas` 时按**系统 ANSI 代码页**解码。
简体中文 Windows 的代码页是 936(GBK)，所以：

- 把 UTF-8 的文件当 GBK 读 → 所有中文变成 `璇ユ枃浠跺す` 这种乱码；
- 更致命的是：中文标点（例如「。」的 UTF-8 字节是 `E3 80 82`）的收尾字节 `82`、`83`
  在 GBK 里属于**前导字节**，会跟紧随其后的 ASCII 字符（往往就是字符串的结束引号 `"` 或连接符 `&`）
  合成一个汉字 —— **引号被吃掉了**，字符串没有结尾，于是报"语法错误"。

**处理办法**：删掉出错的那个模块，按第 2.2 节重新导入（现在的 `dist/SWautoExport.bas` 已经是 GBK 编码）。
如果还不行，说明这台机器的 ANSI 代码页不是 936，请改用**方式 B 复制粘贴**，那条路不经过代码页，必定正确。

### 8.2 编译报"用户定义类型未定义"

VBE → 工具 > 引用 → 勾选 `SldWorks <版本> Type Library` 和 `SOLIDWORKS <版本> Constant type library`。

### 8.3 运行时报"SldWorks 未定义"或宏列表里找不到入口

入口宏必须是 `Public Sub` 且**不带参数**。列表里只会出现 `SwAuto_` 开头的三个。
如果三个都没有，说明模块没编译通过，先解决编译错误。

### 8.4 清单里某些工程图"失败：无法打开工程图"

工程图引用的模型路径失效（换过盘符 / 手工改过名 / 网络盘未连接）。
先在 SolidWorks 里手工打开一张确认，或者在 `工具 > 选项 > 系统选项 > 文件位置` 里补上搜索路径。

### 8.5 清单备注里出现"AP203 未能写入用户选项"

说明这台机器的 SolidWorks 没能被写进该用户选项（权限或版本差异）。
到 `工具 > 选项 > 系统选项 > 导出 > STEP` 手工把协议设为 AP203 即可，其余功能不受影响。

### 8.6 多选文件对话框点了没反应

`SwAuto_PickFiles` 走的是 PowerShell + .NET WinForms 的原生对话框。
如果被安全策略拦截，会退化成"没有选中文件"。改用 `SwAuto_PickFolder` 或 `SwAuto_ExportActiveDoc`。
