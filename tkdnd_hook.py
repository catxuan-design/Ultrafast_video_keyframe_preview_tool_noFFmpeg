# -*- mode: python ; coding: utf-8 -*-
import os
import sys
import platform

# 在程序启动时修复 tkdnd 库路径
def patch_tkinterdnd():
    import tkinterdnd2
    from tkinterdnd2 import TkinterDnD
    
    # 保存原始的 _require 函数
    original_require = TkinterDnD._require
    
    def patched_require(tkroot):
        try:
            system = platform.system()
            if system == "Windows":
                machine = os.environ.get('PROCESSOR_ARCHITECTURE', platform.machine())
            else:
                machine = platform.machine()
            
            if system == "Darwin" and machine == "arm64":
                tkdnd_platform_rep = "osx-arm64"
            elif system == "Darwin" and machine == "x86_64":
                tkdnd_platform_rep = "osx-x64"
            elif system == "Linux" and machine == "aarch64":
                tkdnd_platform_rep = "linux-arm64"
            elif system == "Linux" and machine == "x86_64":
                tkdnd_platform_rep = "linux-x64"
            elif system == "Windows" and machine == "ARM64":
                tkdnd_platform_rep = "win-arm64"
            elif system == "Windows" and machine == "AMD64":
                tkdnd_platform_rep = "win-x64"
            elif system == "Windows" and machine == "x86":
                tkdnd_platform_rep = "win-x86"
            else:
                raise RuntimeError('Platform not supported.')
            
            # 获取 tkinterdnd2 的目录
            if getattr(sys, 'frozen', False):
                # PyInstaller 打包模式
                module_path = os.path.join(sys._MEIPASS, 'tkdnd', tkdnd_platform_rep)
            else:
                module_path = os.path.join(os.path.dirname(tkinterdnd2.__file__), 'tkdnd', tkdnd_platform_rep)
            
            tkroot.tk.call('lappend', 'auto_path', module_path)
            import tkinter
            TkinterDnD.TkdndVersion = tkroot.tk.call('package', 'require', 'tkdnd')
            return TkinterDnD.TkdndVersion
        except Exception as e:
            raise RuntimeError(f'Unable to load tkdnd library: {e}')
    
    # 替换原始函数
    TkinterDnD._require = patched_require

# 执行补丁
patch_tkinterdnd()
