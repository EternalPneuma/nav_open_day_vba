# VBA 模块同步

当 `sync-vba.ps1` 因后台 Excel 进程而无法继续时，先关闭所有 Excel 实例，再执行同步：

```powershell
Get-Process EXCEL | Stop-Process -Force
./sync-vba.ps1
```

> 注意：第一条命令会强制关闭所有 Excel 工作簿。执行前请保存其他正在编辑的文件。

## 使用 `sync-vba.ps1` 同步 VBA 模块

`sync-vba.ps1` 会打开目标 `.xlsm` 工作簿，将 `scripts/vba/` 下的 VBA 源码导入工作簿，并保存结果。默认同步数据、图表、操作面板、工具和周报推荐模块到仓库根目录的 `上层产品净值数据库.xlsm`。

在仓库根目录执行：

```powershell
./sync-vba.ps1
```

可指定其他目标工作簿：

```powershell
./sync-vba.ps1 -WorkbookPath "其他工作簿.xlsm"
```

可通过 `-ModuleGroups` 指定要同步的模块组，多个组用逗号分隔。可用组为 `data`、`chart`、`optional_panel`、`tool`、`weekly`：

```powershell
./sync-vba.ps1 -ModuleGroups "data,chart"
```

脚本会替换工作簿中现有的普通和类 VBA 模块；操作面板的 UserForm 会原地更新代码。同步前请关闭目标工作簿，并在 Excel 信任中心启用“信任对 VBA 工程对象模型的访问”。
