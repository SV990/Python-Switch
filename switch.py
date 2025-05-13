import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import winreg
import os
import sys
from pathlib import Path
import json
import subprocess
import ctypes
from typing import List, Tuple
import re

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

class PythonSwitcher:
    def __init__(self, master):
        self.master = master
        master.title("Python 版本切换器")
        
        if not is_admin():
            messagebox.showerror("错误", "请以管理员权限运行此程序！")
            master.destroy()
            return
        
        # 设置窗口图标
        try:
            master.iconbitmap(default="python.ico")
        except:
            pass
        
        # 设置窗口大小和位置
        window_width = 800
        window_height = 500
        screen_width = master.winfo_screenwidth()
        screen_height = master.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        master.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 设置主题
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置样式
        style.configure('TLabel', font=('Microsoft YaHei UI', 10))
        style.configure('TButton', font=('Microsoft YaHei UI', 10))
        style.configure('Treeview', font=('Microsoft YaHei UI', 10))
        style.configure('Treeview.Heading', font=('Microsoft YaHei UI', 10, 'bold'))
        
        # 创建主框架
        self.main_frame = ttk.Frame(master, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题
        title_frame = ttk.Frame(self.main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(title_frame, text="Python 版本管理器", 
                              font=('Microsoft YaHei UI', 16, 'bold'))
        title_label.pack(side=tk.LEFT)
        
        # 版本信息
        version_label = ttk.Label(title_frame, text="v1.0.0", 
                                font=('Microsoft YaHei UI', 8))
        version_label.pack(side=tk.RIGHT, padx=5)

        # 创建内容框架
        content_frame = ttk.Frame(self.main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧框架
        left_frame = ttk.LabelFrame(content_frame, text="已安装的版本", padding=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # 版本列表
        self.version_tree = ttk.Treeview(left_frame, 
                                       columns=('版本', '路径', '状态'),
                                       show='headings',
                                       height=8)
        self.version_tree.heading('版本', text='版本')
        self.version_tree.heading('路径', text='安装路径')
        self.version_tree.heading('状态', text='状态')
        self.version_tree.column('版本', width=100)
        self.version_tree.column('路径', width=300)
        self.version_tree.column('状态', width=100)
        self.version_tree.pack(fill=tk.BOTH, expand=True)
        
        # 添加滚动条
        y_scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, 
                                  command=self.version_tree.yview)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        x_scrollbar = ttk.Scrollbar(left_frame, orient=tk.HORIZONTAL, 
                                  command=self.version_tree.xview)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.version_tree.configure(yscrollcommand=y_scrollbar.set,
                                  xscrollcommand=x_scrollbar.set)
        
        # 右侧框架
        right_frame = ttk.LabelFrame(content_frame, text="操作", padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH)
        
        # 添加按钮
        ttk.Button(right_frame, text="添加 Python 路径", 
                  command=self.add_python_path).pack(pady=5, fill=tk.X)
        ttk.Button(right_frame, text="删除选中路径",
                  command=self.remove_selected_path).pack(pady=5, fill=tk.X)
        ttk.Button(right_frame, text="刷新版本列表", 
                  command=self.refresh_versions).pack(pady=5, fill=tk.X)
        ttk.Button(right_frame, text="切换选中版本", 
                  command=self.switch_selected_version).pack(pady=5, fill=tk.X)
        
        # 分隔线
        ttk.Separator(right_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # 当前版本信息
        current_version_frame = ttk.LabelFrame(right_frame, text="当前系统版本", padding=5)
        current_version_frame.pack(fill=tk.X, pady=5)
        self.current_version_label = ttk.Label(current_version_frame, 
                                             text="检测中...",
                                             wraplength=150)
        self.current_version_label.pack(fill=tk.X)
        
        # 状态标签
        self.status_label = ttk.Label(self.main_frame, text="就绪", 
                                    font=('Microsoft YaHei UI', 9))
        self.status_label.pack(pady=10)
        
        # 保存的路径文件
        self.config_file = Path.home() / '.python_switcher_config.json'
        
        # 加载版本并更新当前版本显示
        self.load_saved_paths()
        self.refresh_versions()
        self.update_current_version()

    def get_python_paths_from_registry(self) -> List[Tuple[str, str]]:
        """从注册表获取所有可能的 Python 安装路径"""
        paths = []
        keys_to_check = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Python\PythonCore"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Python\PythonCore"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Wow6432Node\Python\PythonCore"),
        ]
        
        for hkey, key_path in keys_to_check:
            try:
                with winreg.OpenKey(hkey, key_path) as key:
                    i = 0
                    while True:
                        try:
                            version = winreg.EnumKey(key, i)
                            try:
                                with winreg.OpenKey(key, f"{version}\\InstallPath") as install_key:
                                    path = winreg.QueryValue(install_key, None)
                                    if path and Path(path).exists():
                                        paths.append((version, path))
                            except:
                                # 尝试其他可能的安装路径键
                                for subkey in ['PythonPath', 'InstallDir']:
                                    try:
                                        with winreg.OpenKey(key, f"{version}\\{subkey}") as alt_key:
                                            path = winreg.QueryValue(alt_key, None)
                                            if path and Path(path).exists():
                                                paths.append((version, path))
                                                break
                                    except:
                                        continue
                            i += 1
                        except WindowsError:
                            break
            except:
                continue
        
        return list(set(paths))  # 删除重复项

    def get_version_from_path(self, path: str) -> str:
        """通过运行 python --version 获取版本号"""
        try:
            python_exe = Path(path) / "python.exe"
            if not python_exe.exists():
                python_exe = Path(path).parent / "python.exe"
                if not python_exe.exists():
                    return "未知版本"
            
            result = subprocess.run([str(python_exe), "--version"], 
                                 capture_output=True, text=True)
            version = result.stdout.strip().replace("Python ", "")
            return version
        except:
            return "未知版本"

    def update_current_version(self):
        """更新当前系统 Python 版本显示"""
        try:
            result = subprocess.run(["python", "--version"], 
                                 capture_output=True, text=True)
            version = result.stdout.strip()
            path = subprocess.run(["where", "python"], 
                               capture_output=True, text=True).stdout.strip()
            self.current_version_label.config(
                text=f"{version}\n路径: {path.split()[0]}")
        except:
            self.current_version_label.config(text="未检测到系统 Python")

    def refresh_versions(self):
        """刷新 Python 版本列表"""
        self.status_label.config(text="正在刷新版本列表...")
        
        # 清空现有项
        for item in self.version_tree.get_children():
            self.version_tree.delete(item)
        
        # 获取系统环境变量中的 Python 路径
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                              "SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment",
                              0, winreg.KEY_READ) as key:
                sys_path = winreg.QueryValueEx(key, "Path")[0]
                current_paths = sys_path.split(";")
        except:
            current_paths = []
        
        # 添加注册表中的版本
        for version, path in self.get_python_paths_from_registry():
            status = "系统变量" if any(p in current_paths for p in [path, path.rstrip("\\")])                     else "未使用"
            self.version_tree.insert('', 'end', values=(version, path, status))
        
        # 添加自定义路径
        for path in self.custom_paths:
            if Path(path).exists():
                version = self.get_version_from_path(path)
                status = "系统变量" if any(p in current_paths for p in [path, path.rstrip("\\")])                         else "未使用"
                self.version_tree.insert('', 'end', values=(version, path, status))
        
        self.status_label.config(text="版本列表已更新")

    def remove_selected_path(self):
        """删除选中的自定义路径"""
        selected = self.version_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要删除的路径")
            return
        
        path = self.version_tree.item(selected[0])['values'][1]
        if path in self.custom_paths:
            self.custom_paths.remove(path)
            self.save_paths()
            self.refresh_versions()
            self.status_label.config(text=f"已删除路径: {path}")
        else:
            messagebox.showinfo("提示", "只能删除自定义添加的路径")

    def switch_selected_version(self):
        """切换到选中的 Python 版本"""
        selected = self.version_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个 Python 版本")
            return
        
        version_info = self.version_tree.item(selected[0])['values']
        if not version_info:
            return
        
        python_path = version_info[1]
        script_path = str(Path(python_path) / "Scripts")
        
        try:
            # 获取系统环境变量
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                              "SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment",
                              0, winreg.KEY_ALL_ACCESS) as key:
                path_value = winreg.QueryValueEx(key, "Path")[0]
                path_list = path_value.split(";")
                
                # 移除旧的 Python 路径
                path_list = [p for p in path_list if not any(
                    x in p.lower() for x in ['\\python', '\\scripts'])]
                
                # 将新的 Python 路径添加到最前面
                path_list.insert(0, script_path)
                path_list.insert(0, python_path)
                
                # 清理空路径并确保唯一性
                path_list = list(filter(None, dict.fromkeys(path_list)))
                new_path_value = ";".join(path_list)
                
                # 设置新的环境变量
                winreg.SetValueEx(key, "Path", 0, winreg.REG_EXPAND_SZ, new_path_value)
                
                # 通知其他程序环境变量已更改
                HWND_BROADCAST = 0xFFFF
                WM_SETTINGCHANGE = 0x001A
                SMTO_ABORTIFHUNG = 0x0002
                ctypes.windll.user32.SendMessageTimeoutW(HWND_BROADCAST, WM_SETTINGCHANGE, 
                                                       0, "Environment", SMTO_ABORTIFHUNG, 5000)
                
                self.status_label.config(text=f"成功切换到 Python {version_info[0]}")
                self.refresh_versions()
                self.update_current_version()
                messagebox.showinfo("成功", 
                                  "Python 版本切换成功！\n请重新打开命令行窗口以使更改生效。")
        
        except Exception as e:
            self.status_label.config(text=f"切换失败: {str(e)}")
            messagebox.showerror("错误", f"切换 Python 版本失败: {str(e)}")

    def load_saved_paths(self):
        """加载保存的 Python 路径"""
        self.custom_paths = []
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r') as f:
                    self.custom_paths = json.load(f)
            except:
                pass
    
    def save_paths(self):
        """保存自定义的 Python 路径"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.custom_paths, f)
        except Exception as e:
            messagebox.showerror("错误", f"保存配置文件失败: {str(e)}")
    
    def add_python_path(self):
        """添加自定义 Python 路径"""
        path = filedialog.askdirectory(title="选择 Python 安装目录")
        if path:
            path = Path(path)
            python_exe = path / "python.exe"
            if not python_exe.exists():
                messagebox.showerror("错误", "所选目录不是有效的 Python 安装目录")
                return
            
            if str(path) not in self.custom_paths:
                self.custom_paths.append(str(path))
                self.save_paths()
                self.refresh_versions()
                self.status_label.config(text=f"已添加路径: {path}")
            else:
                messagebox.showinfo("提示", "该路径已存在")

if __name__ == '__main__':
    if not is_admin():
        # 如果不是管理员，则使用管理员权限重新运行
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    else:
        root = tk.Tk()
        app = PythonSwitcher(root)
        root.mainloop()