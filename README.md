# AHK_script
## 依赖

- cmd增强，内置了一些写好的高级cmd命令可以直接调用[nircmd](https://www.nirsoft.net/utils/nircmd.html) **需要添加到系统环境中**

- `VirtualDesktopAccessor.dll`,项目内为win10的DLL文件，如果需要win11的DLL请访问 [Ciantic/VirtualDesktopAccessor](https://github.com/Ciantic/VirtualDesktopAccessor)

## TIPS

- 使用快速切换虚拟桌面前，必须提前手动创建虚拟桌面；否则请自行修改源码（有创建虚拟桌面的函数可以调用）
- 在win10环境下读取虚拟桌面名称的接口运行异常，合理怀疑是DLL的问题，修改源码时需要注意


