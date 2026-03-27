import os
import sys
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import tkinterdnd2
import time
from datetime import datetime
import tempfile
import math
from PIL import Image, ImageTk
import json
import atexit
import zipfile
import shutil
import base64

class VideoKeyframeGridApp:
    def __init__(self, root):
        # 使用传入的root实例
        self.root = root
        self.root.title("视频关键帧预览工具")
        self.root.geometry("860x800")  # 修改初始窗口大小为860x800
        
        # 设置窗口大小可调整
        self.root.resizable(True, True)
        
        # 设置样式
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # 初始化变量
        self.file_queue = []  # 处理队列
        self.is_processing = False
        self.current_file_index = -1
        self.processing_start_time = None  # 记录处理开始时间
        
        # 宫格布局配置
        self.grid_options = {
            "3x3": {"total_frames": 9, "rows": 3, "cols": 3},
            "4x4": {"total_frames": 16, "rows": 4, "cols": 4},
            "5x5": {"total_frames": 25, "rows": 5, "cols": 5},
            "6x6": {"total_frames": 36, "rows": 6, "cols": 6},
            "7x7": {"total_frames": 49, "rows": 7, "cols": 7},
            "8x8": {"total_frames": 64, "rows": 8, "cols": 8}
        }
        self.selected_grid_option = tk.StringVar(value="5x5")  # 默认选择5x5
        
        # 分辨率选项配置
        self.resolution_options = {
            "270": {"max_size": 270, "label": "270像素 (低)"},
            "320": {"max_size": 320, "label": "320像素 (低)"},
            "480": {"max_size": 480, "label": "480像素 (中)"},
            "720": {"max_size": 720, "label": "720像素 (高)"},
            "960": {"max_size": 960, "label": "960像素 (超高)"},
            "1024": {"max_size": 1024, "label": "1024像素 (超清)"}
        }
        self.selected_resolution = tk.StringVar(value="320")  # 默认选择320像素
        
        # 结果消息变量
        self.result_message = tk.StringVar(value="")
        self.result_type = tk.StringVar(value="info")  # info, success, error
        
        # 界面状态
        self.interface_state = "drop"  # drop, processing, log
        
        # 注册退出时清理临时文件
        atexit.register(self.cleanup_temp_files)
        
        self.setup_ui()
        self.setup_drag_drop()
        

        
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 标题和管理按钮区域
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # 标题
        self.title_label = ttk.Label(
            title_frame, 
            text="视频关键帧预览工具", 
            font=("Arial", 16, "bold")
        )
        self.title_label.pack(side=tk.LEFT)
        
        # 管理预览按钮
        self.manage_previews_button = ttk.Button(
            title_frame, 
            text="管理预览", 
            command=self.open_manage_window
        )
        self.manage_previews_button.pack(side=tk.RIGHT, padx=(10, 0))
        
        # 布局选择区域
        layout_frame = ttk.LabelFrame(
            main_frame, 
            text="宫格布局选择", 
            padding="10"
        )
        layout_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 布局选项
        ttk.Label(layout_frame, text="选择宫格布局:").pack(side=tk.LEFT, padx=(0, 10))
        
        for layout_key, layout_info in self.grid_options.items():
            rb = ttk.Radiobutton(
                layout_frame,
                text=f"{layout_info['rows']}×{layout_info['cols']} ({layout_info['total_frames']}张截图)",
                variable=self.selected_grid_option,
                value=layout_key
            )
            rb.pack(side=tk.LEFT, padx=(10, 0))
        
        # 分辨率选择区域
        resolution_frame = ttk.LabelFrame(
            main_frame, 
            text="输出分辨率选择", 
            padding="10"
        )
        resolution_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 分辨率选项
        ttk.Label(resolution_frame, text="选择输出分辨率:").pack(side=tk.LEFT, padx=(0, 10))
        
        for res_key, res_info in self.resolution_options.items():
            rb = ttk.Radiobutton(
                resolution_frame,
                text=res_info['label'],
                variable=self.selected_resolution,
                value=res_key
            )
            rb.pack(side=tk.LEFT, padx=(10, 0))
        
        # 拖放区域（初始状态）
        self.drop_frame = ttk.LabelFrame(
            main_frame, 
            text="请拖拽文件至此处", 
            padding="40",
            relief="groove",
            borderwidth=2
        )
        self.drop_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        
        # 拖放提示标签
        self.drop_label = ttk.Label(
            self.drop_frame, 
            text="拖拽视频文件到这里\n\n支持多个文件同时拖拽", 
            font=("Arial", 14),
            foreground="gray",
            justify=tk.CENTER,
            padding=(40, 40)
        )
        self.drop_label.pack(expand=True, fill=tk.BOTH)
        
        # 处理队列和日志区域（初始隐藏）
        self.processing_frame = ttk.LabelFrame(
            main_frame, 
            text="处理队列", 
            padding="10"
        )
        # 初始隐藏
        self.processing_frame.grid_remove()
        
        # 队列列表
        self.queue_frame = ttk.Frame(self.processing_frame)
        self.queue_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 队列列表滚动条
        queue_scrollbar = ttk.Scrollbar(self.queue_frame, orient=tk.VERTICAL)
        queue_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 队列列表树
        self.queue_tree = ttk.Treeview(
            self.queue_frame, 
            columns=("name", "status", "progress", "grid_type", "resolution", "preview"),
            show="headings",
            yscrollcommand=queue_scrollbar.set
        )
        self.queue_tree.heading("name", text="文件名")
        self.queue_tree.heading("status", text="状态")
        self.queue_tree.heading("progress", text="进度")
        self.queue_tree.heading("grid_type", text="宫格类型")
        self.queue_tree.heading("resolution", text="分辨率")
        self.queue_tree.heading("preview", text="预览")
        self.queue_tree.column("name", width=200)
        self.queue_tree.column("status", width=80)
        self.queue_tree.column("progress", width=80)
        self.queue_tree.column("grid_type", width=100)
        self.queue_tree.column("resolution", width=80)
        self.queue_tree.column("preview", width=80)
        self.queue_tree.pack(fill=tk.BOTH, expand=True)
        queue_scrollbar.config(command=self.queue_tree.yview)
        
        # 绑定双击事件，用于预览
        self.queue_tree.bind("<Double-1>", self.on_queue_item_double_click)
        
        # 队列控制说明
        queue_control_frame = ttk.Frame(self.processing_frame)
        queue_control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 添加说明文字
        instruction_label = ttk.Label(
            queue_control_frame,
            text="拖拽更多视频文件到上方区域以添加到队列",
            font=("Arial", 10),
            foreground="gray"
        )
        instruction_label.pack(side=tk.LEFT, padx=(0, 10))
        
        # 处理按钮移到右侧
        self.start_processing_button = ttk.Button(
            queue_control_frame, 
            text="开始处理", 
            command=self.start_processing
        )
        self.start_processing_button.pack(side=tk.RIGHT)
        
        # 日志区域
        self.log_frame = ttk.LabelFrame(
            main_frame, 
            text="处理日志", 
            padding="10"
        )
        # 初始隐藏
        self.log_frame.grid_remove()
        
        # 日志文本
        log_scrollbar = ttk.Scrollbar(self.log_frame, orient=tk.VERTICAL)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(
            self.log_frame, 
            height=10,  # 减少日志显示区域的高度，避免窗口过高
            state=tk.DISABLED,
            wrap=tk.WORD,
            font=("Consolas", 10),
            yscrollcommand=log_scrollbar.set
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        log_scrollbar.config(command=self.log_text.yview)
        
        # 进度条
        self.progress_var = tk.StringVar(value="就绪")
        self.progress_label = ttk.Label(
            main_frame, 
            textvariable=self.progress_var,
            font=("Arial", 10)
        )
        # 初始隐藏
        self.progress_label.grid_remove()
        
        self.progress_bar = ttk.Progressbar(
            main_frame, 
            mode='indeterminate',
            length=400
        )
        # 初始隐藏
        self.progress_bar.grid_remove()
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        # 为不同的行设置合适的权重，确保界面元素合理分布
        main_frame.rowconfigure(4, weight=1)  # 处理队列区域应该获得更多的垂直空间
        main_frame.rowconfigure(5, weight=2)  # 日志区域应该获得最多的垂直空间
        
        # 绑定窗口大小变化事件
        self.root.bind("<Configure>", self.on_window_resize)
        
    def on_window_resize(self, event):
        """窗口大小变化时的处理"""
        pass
        
    def cleanup_temp_files(self):
        """清理临时文件"""
        pass  # 不需要清理任何临时文件
        
    def setup_drag_drop(self):
        """设置拖放功能"""
        # 注册拖放区域
        self.drop_frame.drop_target_register(tkinterdnd2.DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        self.drop_frame.dnd_bind('<<DragEnter>>', self.on_drag_enter)
        self.drop_frame.dnd_bind('<<DragLeave>>', self.on_drag_leave)
        
        # 绑定拖放事件到主窗口
        self.root.drop_target_register(tkinterdnd2.DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop_global)
        
    def on_drag_enter(self, event):
        """鼠标进入拖放区域时的视觉反馈"""
        self.drop_frame.configure(relief="sunken", borderwidth=3)
        self.drop_label.configure(foreground="blue")
        
    def on_drag_leave(self, event):
        """鼠标离开拖放区域时的视觉反馈"""
        self.drop_frame.configure(relief="groove", borderwidth=2)
        self.drop_label.configure(foreground="gray")
        
    def on_drop(self, event):
        """处理拖放事件"""
        # 恢复拖放区域样式
        self.drop_frame.configure(relief="groove", borderwidth=2)
        self.drop_label.configure(foreground="gray")
        
        # 获取拖放的文件路径
        files = self.root.tk.splitlist(event.data)
        if files:
            self.add_files_to_queue(files)
    
    def on_drop_global(self, event):
        """全局拖放事件处理"""
        # 获取拖放的文件路径
        files = self.root.tk.splitlist(event.data)
        if files:
            self.add_files_to_queue(files)
        
    def add_files(self):
        """通过文件对话框添加文件"""
        file_paths = filedialog.askopenfilenames(
            title="选择视频文件",
            filetypes=[
                ("视频文件", "*.mp4 *.avi *.mov *.mkv *.wmv *.flv *.m4v *.3gp *.webm *.mpg *.mpeg *.ts *.mts *.m2ts"),
                ("所有文件", "*.*")
            ]
        )
        
        if file_paths:
            self.add_files_to_queue(file_paths)
            
    def add_files_to_queue(self, file_paths):
        """添加文件到处理队列"""
        # 过滤并验证文件
        valid_files = []
        
        def process_path(path):
            if os.path.isfile(path):
                # 检查文件扩展名
                ext = os.path.splitext(path)[1].lower()
                if ext in ['.mp4', '.avi', '.mov', '.mkv', '.wmv', '.flv', '.m4v', '.3gp', '.webm', '.mpg', '.mpeg', '.ts', '.mts', '.m2ts']:
                    valid_files.append(path)
                else:
                    self.add_log(f"跳过非视频文件: {os.path.basename(path)}", "warning")
            elif os.path.isdir(path):
                # 处理文件夹，递归遍历
                self.add_log(f"处理文件夹: {os.path.basename(path)}", "info")
                for root, dirs, files in os.walk(path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        ext = os.path.splitext(file_path)[1].lower()
                        if ext in ['.mp4', '.avi', '.mov', '.mkv', '.wmv', '.flv', '.m4v', '.3gp', '.webm', '.mpg', '.mpeg', '.ts', '.mts', '.m2ts']:
                            valid_files.append(file_path)
                            self.add_log(f"添加文件: {os.path.basename(file_path)}", "info")
            else:
                self.add_log(f"跳过无效路径: {path}", "error")
        
        for file_path in file_paths:
            process_path(file_path)
        
        if valid_files:
            # 添加到队列
            for file_path in valid_files:
                file_name = os.path.basename(file_path)
                grid_type = self.selected_grid_option.get()
                grid_info = self.grid_options[grid_type]
                resolution = self.selected_resolution.get()
                resolution_info = self.resolution_options[resolution]
                
                self.file_queue.append({
                    'path': file_path,
                    'name': file_name,
                    'status': '等待',
                    'progress': '0%',
                    'grid_type': f"{grid_info['rows']}×{grid_info['cols']}",
                    'resolution': resolution  # 使用实际的分辨率键值，如"320", "480", "720"
                })
                
                # 添加到队列树
                self.queue_tree.insert(
                    '', 
                    tk.END, 
                    values=(file_name, '等待', '0%', f"{grid_info['rows']}×{grid_info['cols']}", resolution_info['label'].split()[0], '')
                )
            
            # 如果当前不是处理状态，切换到处理状态
            if self.interface_state != "processing":
                self.switch_to_processing_state()
            
            # 自动开始处理
            if not self.is_processing:
                self.start_processing()
        
    def switch_to_processing_state(self):
        """切换到处理状态界面"""
        # 隐藏拖放区域
        self.drop_frame.grid_remove()
        
        # 显示处理队列和日志区域
        self.processing_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        self.log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        self.progress_label.grid(row=6, column=0, columnspan=3, pady=(0, 5))
        self.progress_bar.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 更新界面状态
        self.interface_state = "processing"
        
    def switch_to_drop_state(self):
        """切换回拖放状态界面"""
        # 隐藏处理队列和日志区域
        self.processing_frame.grid_remove()
        self.log_frame.grid_remove()
        self.progress_label.grid_remove()
        self.progress_bar.grid_remove()
        
        # 清空队列列表
        for item in self.queue_tree.get_children():
            self.queue_tree.delete(item)
        
        # 清空队列
        self.file_queue = []
        
        # 显示拖放区域
        self.drop_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        
        # 更新界面状态
        self.interface_state = "drop"
        
        # 重置进度条和标签文本
        self.progress_var.set("就绪")
        

            
    def clear_queue(self):
        """清空处理队列"""
        # 清空队列列表
        for item in self.queue_tree.get_children():
            self.queue_tree.delete(item)
        
        # 清空队列
        self.file_queue = []
        
        # 切换回拖放状态
        self.switch_to_drop_state()
        
    def start_processing(self):
        """开始处理队列中的文件"""
        if not self.file_queue or self.is_processing:
            return
            

            
        # 找到第一个未完成的文件索引
        first_incomplete_index = -1
        for i, item in enumerate(self.file_queue):
            if item.get('status') != '完成':
                first_incomplete_index = i
                break
        
        # 如果所有文件都已完成，提示用户
        if first_incomplete_index == -1:
            self.add_log("所有文件都已完成处理", "info")
            return
        
        self.is_processing = True
        self.current_file_index = first_incomplete_index
        self.progress_bar.start()
        
        # 禁用控制按钮
        self.start_processing_button.config(state=tk.DISABLED)
        
        # 开始处理线程
        thread = threading.Thread(target=self.process_queue)
        thread.daemon = True
        thread.start()
        
    def process_queue(self):
        """处理队列中的文件"""
        # 记录处理开始时间
        self.processing_start_time = time.time()
        
        while True:
            # 检查是否还有未完成的文件
            has_incomplete = False
            for i, item in enumerate(self.file_queue):
                # 只处理状态不是'完成'且不是'失败'的文件
                if item.get('status') != '完成' and item.get('status') != '失败':
                    has_incomplete = True
                    # 如果当前索引超过了文件队列长度，或者当前索引指向的是已完成或失败的文件，重置索引
                    if self.current_file_index >= len(self.file_queue) or self.file_queue[self.current_file_index].get('status') == '完成' or self.file_queue[self.current_file_index].get('status') == '失败':
                        self.current_file_index = i
                    break
            
            # 如果所有文件都已完成，退出循环
            if not has_incomplete:
                break
            
            # 确保索引在有效范围内
            if self.current_file_index >= len(self.file_queue):
                self.current_file_index = 0
                continue
            
            current_file = self.file_queue[self.current_file_index]
            
            # 跳过已完成或失败的文件
            if current_file.get('status') == '完成' or current_file.get('status') == '失败':
                self.current_file_index += 1
                continue
            
            file_path = current_file['path']
            file_name = current_file['name']
            grid_type = current_file.get('grid_type', '5×5')  # 默认5x5
            resolution = current_file.get('resolution', '320')  # 默认320像素
            
            # 更新状态
            self.update_file_status(self.current_file_index, '处理中', '处理中...')
            self.progress_var.set(f"正在处理: {file_name} ({grid_type}, {resolution})")
            self.add_log(f"开始处理: {file_name}，宫格类型: {grid_type}，分辨率: {resolution}像素", "info")
            
            # 处理文件 - 生成关键帧宫格
            success = self.generate_keyframe_grid(file_path, grid_type, resolution)
            
            # 更新状态
            if success:
                self.update_file_status(self.current_file_index, '完成', '100%')
                # 计算处理耗时
                elapsed_time = time.time() - self.processing_start_time
                self.add_log(f"处理完成: {file_name} (耗时: {elapsed_time:.2f}秒)", "success")
            else:
                self.update_file_status(self.current_file_index, '失败', '错误')
                self.add_log(f"处理失败: {file_name}", "error")
            
            # 处理下一个文件
            self.current_file_index += 1
        
        # 处理完成
        self.is_processing = False
        self.progress_bar.stop()
        total_elapsed_time = time.time() - self.processing_start_time if self.processing_start_time else 0
        self.progress_var.set(f"处理完成 (总耗时: {total_elapsed_time:.2f}秒)")
        
        # 启用控制按钮
        self.root.after(0, lambda: self.start_processing_button.config(state=tk.NORMAL))
        
    def update_file_status(self, index, status, progress):
        """更新文件状态"""
        # 更新队列数据
        if 0 <= index < len(self.file_queue):
            self.file_queue[index]['status'] = status
            self.file_queue[index]['progress'] = progress
            
            # 计算预览列的值
            preview_text = ''
            if status == '完成':
                preview_text = '点击预览'
            
            # 更新队列树（使用after方法确保在主线程中更新）
            def update_gui():
                if 0 <= index < len(self.file_queue) and index < len(self.queue_tree.get_children()):
                    item = self.queue_tree.get_children()[index]
                    grid_type = self.file_queue[index]['grid_type']
                    resolution = self.file_queue[index].get('resolution', '320')
                    self.queue_tree.item(
                        item, 
                        values=(self.file_queue[index]['name'], status, progress, grid_type, resolution, preview_text)
                    )
            
            self.root.after(0, update_gui)
    
    def get_video_duration(self, video_path):
        """获取视频时长（秒）"""
        import cv2
        try:
            cap = cv2.VideoCapture(video_path)
            if not cap.isOpened():
                self.add_log("无法打开视频文件以获取时长", "error")
                return None
            
            # 获取视频总帧数和FPS
            total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
            fps = cap.get(cv2.CAP_PROP_FPS)
            
            cap.release()
            
            if total_frames <= 0 or fps <= 0:
                self.add_log("无法获取视频信息", "error")
                return None
            
            # 计算视频总时长
            duration = total_frames / fps
            return duration
        except Exception as e:
            self.add_log(f"获取视频时长失败: {str(e)}", "error")
            return None
    
    def extract_keyframes(self, video_path, num_frames, max_size=320):
        """使用截屏方法提取指定数量的关键帧，并调整到目标分辨率"""
        import cv2
        import numpy as np
        
        temp_dir = tempfile.mkdtemp()
        frame_paths = []
        
        try:
            # 打开视频文件
            cap = cv2.VideoCapture(video_path)
            if not cap.isOpened():
                self.add_log("无法打开视频文件", "error")
                return []
            
            # 获取视频总帧数和FPS
            total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
            fps = cap.get(cv2.CAP_PROP_FPS)
            
            if total_frames <= 0 or fps <= 0:
                self.add_log("无法获取视频信息", "error")
                cap.release()
                return []
            
            # 计算视频总时长
            duration = total_frames / fps
            
            # 计算要提取的关键帧位置
            for i in range(num_frames):
                # 计算时间点（避免在开头和结尾提取）
                time_point = (i + 1) * duration / (num_frames + 1)
                frame_number = int(time_point * fps)
                
                # 设置视频位置到指定帧
                cap.set(cv2.CAP_PROP_POS_FRAMES, frame_number)
                
                # 读取帧
                ret, frame = cap.read()
                if not ret:
                    self.add_log(f"无法读取第 {frame_number} 帧", "warning")
                    continue
                
                # 生成输出路径
                frame_path = os.path.join(temp_dir, f"frame_{i+1:03d}.jpg")
                
                # 调整图像大小以适应最大尺寸限制
                height, width = frame.shape[:2]
                
                # 计算缩放比例以适应最大尺寸
                scale = min(max_size / width, max_size / height) if width > height else min(max_size / height, max_size / width)
                
                if scale < 1.0:  # 只需要缩小
                    new_width = int(width * scale)
                    new_height = int(height * scale)
                    frame = cv2.resize(frame, (new_width, new_height), interpolation=cv2.INTER_AREA)
                
                # 保存图像
                success = cv2.imwrite(frame_path, frame, [cv2.IMWRITE_JPEG_QUALITY, 90])
                if success:
                    frame_paths.append(frame_path)
                    self.add_log(f"成功提取第 {frame_number} 帧 (时间: {time_point:.2f}s)", "info")
                else:
                    self.add_log(f"保存第 {frame_number} 帧失败", "error")
            
            cap.release()
            
            self.add_log(f"成功提取 {len(frame_paths)} 个关键帧", "info")
            return frame_paths, temp_dir
            
        except Exception as e:
            self.add_log(f"提取关键帧过程中发生错误: {str(e)}", "error")
            # 清理临时文件
            try:
                import shutil
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                    self.add_log(f"清理临时目录: {temp_dir}", "info")
            except Exception as e:
                self.add_log(f"清理临时文件失败: {str(e)}", "warning")
            return [], None
    
    def resize_image_to_fixed_size(self, image_path, target_size=(200, 150), max_size=320):
        """将图片调整为固定尺寸，确保最大边不超过指定像素"""
        try:
            with Image.open(image_path) as img:
                # 转换为RGB模式
                img = img.convert('RGB')
                
                # 由于FFmpeg已经调整了图片大小，现在只需要居中放置到目标尺寸画布上
                # 计算居中位置
                x = (target_size[0] - img.width) // 2
                y = (target_size[1] - img.height) // 2
                
                # 创建一个固定尺寸的白色背景
                new_img = Image.new('RGB', target_size, (255, 255, 255))
                
                # 将图片粘贴到中心位置
                new_img.paste(img, (x, y))
                
                return new_img
        except Exception as e:
            self.add_log(f"调整图片尺寸失败: {image_path}, 错误: {str(e)}", "error")
            return None
    
    def generate_keyframe_grid(self, video_path, grid_type, resolution="320"):
        """
        生成关键帧宫格图
        :param video_path: 视频文件路径
        :param grid_type: 宫格类型 ('4x4', '5x5', '6x6')
        :param resolution: 输出分辨率 ('320', '480', '720')
        """
        # 检查输入文件是否存在
        if not os.path.exists(video_path):
            self.add_log(f"错误：文件不存在: {video_path}", "error")
            return False

        # 获取宫格配置
        if grid_type == "3×3":
            grid_config = self.grid_options["3x3"]
        elif grid_type == "4×4":
            grid_config = self.grid_options["4x4"]
        elif grid_type == "5×5":
            grid_config = self.grid_options["5x5"]
        elif grid_type == "6×6":
            grid_config = self.grid_options["6x6"]
        elif grid_type == "7×7":
            grid_config = self.grid_options["7x7"]
        elif grid_type == "8×8":
            grid_config = self.grid_options["8x8"]
        else:
            grid_config = self.grid_options["5x5"]  # 默认5x5

        # 获取分辨率配置
        if resolution in self.resolution_options:
            max_size = self.resolution_options[resolution]["max_size"]
        else:
            max_size = 320  # 默认320像素

        total_frames = grid_config["total_frames"]
        rows = grid_config["rows"]
        cols = grid_config["cols"]

        # 根据最大分辨率计算单元格大小，留一些边距空间
        cell_size = (max_size + 20, max_size + 20)  # 留20像素边距

        self.add_log(f"开始生成 {rows}×{cols} 宫格图，共需 {total_frames} 张关键帧", "info")
        self.add_log(f"单个图像容器尺寸: {cell_size[0]}x{cell_size[1]} 像素", "info")
        self.add_log(f"最大输出分辨率: {max_size} 像素", "info")

        # 提取关键帧
        keyframe_paths, temp_dir = self.extract_keyframes(video_path, total_frames, max_size)
        if not keyframe_paths or len(keyframe_paths) < total_frames:
            self.add_log(f"提取关键帧不足，实际提取: {len(keyframe_paths)} 张", "error")
            return False

        # 调整所有图片到统一尺寸
        resized_images = []
        for i, frame_path in enumerate(keyframe_paths):
            if os.path.exists(frame_path):
                # 传递最大尺寸给resize函数
                resized_img = self.resize_image_to_fixed_size(frame_path, cell_size, max_size)  # 使用动态计算的尺寸
                if resized_img:
                    resized_images.append(resized_img)
                    self.add_log(f"已调整第 {i+1} 张图片尺寸", "info")
                else:
                    self.add_log(f"调整第 {i+1} 张图片尺寸失败", "error")
            else:
                self.add_log(f"关键帧文件不存在: {frame_path}", "error")

        if len(resized_images) < total_frames:
            self.add_log(f"成功调整尺寸的图片数量不足: {len(resized_images)}/{total_frames}", "error")
            return False

        # 创建宫格图
        cell_width, cell_height = cell_size
        grid_width = cols * cell_width
        grid_height = rows * cell_height

        # 创建大图
        grid_image = Image.new('RGB', (grid_width, grid_height), (255, 255, 255))

        # 将小图按顺序放置到大图上
        for i, img in enumerate(resized_images):
            row = i // cols
            col = i % cols
            x = col * cell_width
            y = row * cell_height
            grid_image.paste(img, (x, y))

        # 保存宫格图
        input_path = Path(video_path)
        output_dir = input_path.parent
        grid_name = f"{input_path.stem}_keyframes_{rows}x{cols}_{max_size}p.jpg"
        grid_path = output_dir / grid_name

        try:
            # 使用高质量设置保存图像
            grid_image.save(str(grid_path), quality=95, optimize=True, subsampling=0)
            self.add_log(f"宫格图生成成功: {grid_path}", "success")
            self.add_log(f"宫格图尺寸: {grid_width}x{grid_height} 像素", "info")
            
            # 查找并更新对应的队列项，保存宫格图路径
            for i, item in enumerate(self.file_queue):
                if item['path'] == video_path:
                    self.file_queue[i]['grid_path'] = str(grid_path)
                    break
            
            # 清理临时文件
            try:
                import shutil
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                    self.add_log(f"清理临时目录: {temp_dir}", "info")
            except Exception as e:
                self.add_log(f"清理临时文件失败: {str(e)}", "warning")
            
            return True
        except Exception as e:
            # 清理临时文件
            try:
                import shutil
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                    self.add_log(f"清理临时目录: {temp_dir}", "info")
            except Exception as e:
                self.add_log(f"清理临时文件失败: {str(e)}", "warning")
            
            self.add_log(f"保存宫格图失败: {str(e)}", "error")
            return False
    
    def add_log(self, message, level="info"):
        """添加日志"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] [{level.upper()}] {message}"
        
        # 使用after方法确保在主线程中更新GUI
        def update_log():
            # 更新日志文本
            self.log_text.config(state=tk.NORMAL)
            
            # 根据日志级别设置颜色
            if level == "error":
                self.log_text.tag_config("error", foreground="red")
                self.log_text.insert(tk.END, log_entry + "\n", "error")
            elif level == "success":
                self.log_text.tag_config("success", foreground="green")
                self.log_text.insert(tk.END, log_entry + "\n", "success")
            elif level == "warning":
                self.log_text.tag_config("warning", foreground="orange")
                self.log_text.insert(tk.END, log_entry + "\n", "warning")
            else:
                self.log_text.insert(tk.END, log_entry + "\n")
            
            # 限制日志数量，防止内存占用过大
            max_log_lines = 1000  # 最大日志行数
            current_lines = int(self.log_text.index('end-1c').split('.')[0])
            
            if current_lines > max_log_lines:
                # 删除最旧的日志条目
                lines_to_delete = current_lines - max_log_lines
                self.log_text.delete(1.0, f"{lines_to_delete}.0")
            
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
        
        self.root.after(0, update_log)
    
    def filter_logs(self):
        """过滤日志"""
        # 这里可以实现日志过滤功能
        filter_text = self.log_filter_var.get()
        self.add_log(f"过滤日志: {filter_text}", "info")
    
    def on_queue_item_double_click(self, event):
        """处理队列项的双击事件，显示预览窗口"""
        # 获取被双击的项
        item = self.queue_tree.identify_row(event.y)
        if not item:
            return
        
        # 获取项的索引
        index = self.queue_tree.index(item)
        if index < 0 or index >= len(self.file_queue):
            return
        
        # 获取队列项
        queue_item = self.file_queue[index]
        
        # 检查是否已完成处理
        if queue_item.get('status') != '完成':
            return
        
        # 获取宫格图路径
        grid_path = queue_item.get('grid_path')
        if not grid_path or not os.path.exists(grid_path):
            self.add_log("宫格图文件不存在，无法预览", "error")
            return
        
        # 显示预览窗口
        self.show_preview_window(grid_path)
    
    def show_preview_window(self, image_path):
        """显示预览窗口"""
        # 创建预览窗口
        preview_window = tk.Toplevel(self.root)
        preview_window.title("宫格图预览")
        preview_window.geometry("800x600")
        preview_window.transient(self.root)  # 设置为临时窗口
        preview_window.grab_set()  # 模态窗口
        
        # 创建主框架
        main_frame = ttk.Frame(preview_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 加载并显示图片
        img = None
        photo = None
        try:
            img = Image.open(image_path)
            # 调整图片大小以适应窗口
            img.thumbnail((750, 500), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            
            # 图片标签
            img_label = ttk.Label(main_frame, image=photo)
            img_label.image = photo  # 保持引用防止垃圾回收
            img_label.pack(pady=10)
            
        except Exception as e:
            error_label = ttk.Label(main_frame, text=f"无法加载图片: {str(e)}", foreground="red")
            error_label.pack(pady=10)
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        
        # 删除宫格图按钮
        delete_grid_btn = ttk.Button(button_frame, text="删除此宫格图", 
                               command=lambda: self.delete_grid_only(image_path, preview_window))
        delete_grid_btn.pack(side=tk.LEFT, padx=5)
        
        # 删除相关视频按钮
        delete_video_btn = ttk.Button(button_frame, text="删除相关视频", 
                               command=lambda: self.delete_related_video(image_path, preview_window))
        delete_video_btn.pack(side=tk.LEFT, padx=5)
        
        # 删除宫格图及相关视频按钮
        delete_all_btn = ttk.Button(button_frame, text="删除宫格图及相关视频", 
                               command=lambda: self.delete_grid_and_video(image_path, preview_window))
        delete_all_btn.pack(side=tk.LEFT, padx=5)
        
        # 关闭按钮
        def on_close():
            # 释放资源
            if img:
                img.close()
            preview_window.destroy()
        
        close_btn = ttk.Button(button_frame, text="关闭", command=on_close)
        close_btn.pack(side=tk.RIGHT, padx=5)
        
        # 绑定窗口关闭事件
        preview_window.protocol("WM_DELETE_WINDOW", on_close)
    
    def delete_grid_only(self, grid_image_path, preview_window):
        """只删除宫格图"""
        import os
        
        # 显示确认对话框
        result = messagebox.askyesno("确认删除", 
                                   f"确定要删除宫格图吗？\n{grid_image_path}\n\n此操作不可撤销！")
        
        if result:
            try:
                # 删除宫格图文件
                os.remove(grid_image_path)
                
                # 关闭预览窗口
                preview_window.destroy()
                
                # 更新队列显示
                self.refresh_queue_display()
                
                self.add_log(f"已删除宫格图: {grid_image_path}", "success")
                
            except Exception as e:
                self.add_log(f"删除文件失败: {str(e)}", "error")
                messagebox.showerror("删除失败", f"删除文件时出错: {str(e)}")
    
    def delete_related_video(self, grid_image_path, preview_window):
        """删除相关视频"""
        import os
        from pathlib import Path
        
        # 从宫格图文件名中提取视频文件名
        grid_path = Path(grid_image_path)
        grid_name = grid_path.stem
        
        # 解析宫格图文件名，提取原始视频文件名
        # 格式：{视频文件名}_keyframes_{rows}x{cols}_{max_size}p
        parts = grid_name.split('_keyframes_')
        if len(parts) != 2:
            messagebox.showerror("错误", "无法从宫格图文件名中提取视频文件名")
            return
        
        video_name = parts[0]
        video_dir = grid_path.parent
        
        # 查找可能的视频文件
        video_extensions = ['.mp4', '.avi', '.mov', '.mkv', '.wmv', '.flv', '.m4v', '.3gp', '.webm', '.mpg', '.mpeg', '.ts', '.mts', '.m2ts']
        video_file = None
        
        for ext in video_extensions:
            potential_video = video_dir / f"{video_name}{ext}"
            if potential_video.exists():
                video_file = potential_video
                break
        
        if not video_file:
            messagebox.showerror("错误", "未找到相关的视频文件")
            return
        
        # 显示确认对话框
        result = messagebox.askyesno("确认删除", 
                                   f"确定要删除相关视频吗？\n{video_file}\n\n此操作不可撤销！")
        
        if result:
            try:
                # 删除视频文件
                os.remove(video_file)
                
                # 关闭预览窗口
                preview_window.destroy()
                
                # 更新队列显示
                self.refresh_queue_display()
                
                self.add_log(f"已删除相关视频: {video_file}", "success")
                
            except Exception as e:
                self.add_log(f"删除文件失败: {str(e)}", "error")
                messagebox.showerror("删除失败", f"删除文件时出错: {str(e)}")
    
    def delete_grid_and_video(self, grid_image_path, preview_window):
        """删除宫格图及相关视频"""
        import os
        from pathlib import Path
        
        # 从宫格图文件名中提取视频文件名
        grid_path = Path(grid_image_path)
        grid_name = grid_path.stem
        
        # 解析宫格图文件名，提取原始视频文件名
        # 格式：{视频文件名}_keyframes_{rows}x{cols}_{max_size}p
        parts = grid_name.split('_keyframes_')
        if len(parts) != 2:
            messagebox.showerror("错误", "无法从宫格图文件名中提取视频文件名")
            return
        
        video_name = parts[0]
        video_dir = grid_path.parent
        
        # 查找可能的视频文件
        video_extensions = ['.mp4', '.avi', '.mov', '.mkv', '.wmv', '.flv', '.m4v', '.3gp', '.webm', '.mpg', '.mpeg', '.ts', '.mts', '.m2ts']
        video_file = None
        
        for ext in video_extensions:
            potential_video = video_dir / f"{video_name}{ext}"
            if potential_video.exists():
                video_file = potential_video
                break
        
        if not video_file:
            messagebox.showerror("错误", "未找到相关的视频文件")
            return
        
        # 显示确认对话框
        result = messagebox.askyesno("确认删除", 
                                   f"确定要删除宫格图及相关视频吗？\n宫格图: {grid_image_path}\n视频: {video_file}\n\n此操作不可撤销！")
        
        if result:
            try:
                # 删除宫格图文件
                os.remove(grid_image_path)
                
                # 删除视频文件
                os.remove(video_file)
                
                # 关闭预览窗口
                preview_window.destroy()
                
                # 更新队列显示
                self.refresh_queue_display()
                
                self.add_log(f"已删除宫格图: {grid_image_path}", "success")
                self.add_log(f"已删除相关视频: {video_file}", "success")
                
            except Exception as e:
                self.add_log(f"删除文件失败: {str(e)}", "error")
                messagebox.showerror("删除失败", f"删除文件时出错: {str(e)}")
    
    def delete_related_files(self, grid_image_path, preview_window):
        """删除宫格图及相关文件"""
        # 调用新的删除宫格图及相关视频方法
        self.delete_grid_and_video(grid_image_path, preview_window)
    
    def refresh_queue_display(self):
        """刷新队列显示"""
        # 重新加载队列显示，移除已删除的项目
        items_to_delete = []
        indices_to_delete = []
        
        # 遍历队列树中的所有项目
        for i, item in enumerate(self.queue_tree.get_children()):
            if i < len(self.file_queue):
                file_item = self.file_queue[i]
                file_path = file_item['path']
                grid_path = file_item.get('grid_path')
                
                # 检查视频文件和宫格图文件是否都存在
                video_exists = os.path.exists(file_path)
                grid_exists = True
                if grid_path:
                    grid_exists = os.path.exists(grid_path)
                
                # 如果视频文件不存在，或者宫格图文件不存在且状态为完成，则删除该项目
                if not video_exists or (file_item.get('status') == '完成' and not grid_exists):
                    items_to_delete.append(item)
                    indices_to_delete.append(i)
        
        # 删除标记的项目
        for item in items_to_delete:
            self.queue_tree.delete(item)
        
        # 同步更新file_queue，按照索引从大到小删除，避免索引偏移
        for i in sorted(indices_to_delete, reverse=True):
            if i < len(self.file_queue):
                self.file_queue.pop(i)
    
    def populate_preview_tree(self, tree):
        """填充预览列表"""
        # 清空现有项
        for item in tree.get_children():
            tree.delete(item)
        
        # 从处理队列中读取数据
        for file_item in self.file_queue:
            file_path = file_item['path']
            file_name = file_item['name']
            status = file_item['status']
            
            # 生成对应的宫格文件路径
            from pathlib import Path
            input_path = Path(file_path)
            output_dir = input_path.parent
            grid_type = file_item.get('grid_type', '5×5')
            rows, cols = grid_type.split('×')
            resolution = file_item.get('resolution', '320')
            max_size = self.resolution_options.get(resolution, {'max_size': 320})['max_size']
            
            grid_name = f"{input_path.stem}_keyframes_{rows}x{cols}_{max_size}p.jpg"
            grid_path = output_dir / grid_name
            
            # 检查宫格文件是否存在
            if os.path.exists(grid_path):
                file_size = os.path.getsize(grid_path)
                file_size_str = f"{file_size / 1024:.2f} KB"
                file_date = os.path.getmtime(grid_path)
                file_date_str = datetime.fromtimestamp(file_date).strftime("%Y-%m-%d %H:%M:%S")
                
                tree.insert('', tk.END, values=(os.path.basename(grid_path), str(grid_path), file_size_str, file_date_str))
    
    def search_previews(self, search_term, tree):
        """搜索预览"""
        # 清空现有项
        for item in tree.get_children():
            tree.delete(item)
        
        # 搜索所有宫格预览文件
        import glob
        import os
        
        # 搜索当前目录及其子目录
        current_dir = os.getcwd()
        grid_files = []
        
        # 搜索常见的宫格文件命名模式
        patterns = [
            "*_keyframes_*.jpg",
            "*_keyframe_*.jpg",
            "*_preview_*.jpg"
        ]
        
        for pattern in patterns:
            grid_files.extend(glob.glob(os.path.join(current_dir, "**", pattern), recursive=True))
        
        # 去重
        grid_files = list(set(grid_files))
        
        # 过滤搜索结果
        filtered_files = [f for f in grid_files if search_term.lower() in f.lower()]
        
        # 添加到列表
        for file_path in filtered_files:
            if os.path.exists(file_path):
                file_name = os.path.basename(file_path)
                file_size = os.path.getsize(file_path)
                file_size_str = f"{file_size / 1024:.2f} KB"
                file_date = os.path.getmtime(file_path)
                file_date_str = datetime.fromtimestamp(file_date).strftime("%Y-%m-%d %H:%M:%S")
                
                tree.insert('', tk.END, values=(file_name, file_path, file_size_str, file_date_str))
    
    def preview_selected_item(self, tree):
        """预览选中项"""
        selected_items = tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请选择一个预览文件")
            return
        
        item = selected_items[0]
        values = tree.item(item, "values")
        file_path = values[1]
        
        if os.path.exists(file_path):
            self.show_preview_window(file_path)
        else:
            messagebox.showerror("错误", "文件不存在")
    
    def delete_grid_file(self, tree):
        """删除宫格文件"""
        selected_items = tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请选择一个预览文件")
            return
        
        item = selected_items[0]
        values = tree.item(item, "values")
        file_path = values[1]
        file_name = values[0]
        
        # 显示确认对话框
        result = messagebox.askyesno("确认删除", 
                                   f"确定要删除宫格文件吗？\n{file_name}\n\n此操作不可撤销！")
        
        if result:
            try:
                # 删除宫格文件
                os.remove(file_path)
                
                # 从列表中移除
                tree.delete(item)
                
                self.add_log(f"已删除宫格文件: {file_path}", "success")
                
            except Exception as e:
                self.add_log(f"删除文件失败: {str(e)}", "error")
                messagebox.showerror("删除失败", f"删除文件时出错: {str(e)}")
    
    def delete_all_files(self, tree):
        """删除宫格和视频文件"""
        selected_items = tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请选择一个预览文件")
            return
        
        item = selected_items[0]
        values = tree.item(item, "values")
        grid_file_path = values[1]
        grid_file_name = values[0]
        
        # 尝试找到对应的视频文件
        import re
        from pathlib import Path
        
        # 从宫格文件名中提取原始视频文件名
        grid_path = Path(grid_file_path)
        grid_name = grid_path.stem
        
        # 匹配模式：文件名_keyframes_*_*.jpg
        match = re.match(r'(.+)_keyframes_.*', grid_name)
        if match:
            video_name = match.group(1)
            video_dir = grid_path.parent
            
            # 搜索可能的视频文件
            video_extensions = ['.mp4', '.avi', '.mov', '.mkv', '.wmv', '.flv', '.m4v', '.3gp', '.webm', '.mpg', '.mpeg', '.ts', '.mts', '.m2ts']
            video_files = []
            
            for ext in video_extensions:
                video_path = video_dir / f"{video_name}{ext}"
                if video_path.exists():
                    video_files.append(str(video_path))
            
            # 构建删除确认信息
            delete_info = f"确定要删除以下文件吗？\n\n宫格文件: {grid_file_name}\n"
            if video_files:
                delete_info += "\n视频文件: "
                for video_file in video_files:
                    delete_info += f"\n- {os.path.basename(video_file)}"
            delete_info += "\n\n此操作不可撤销！"
            
            # 显示确认对话框
            result = messagebox.askyesno("确认删除", delete_info)
            
            if result:
                try:
                    # 删除宫格文件
                    os.remove(grid_file_path)
                    
                    # 删除视频文件
                    for video_file in video_files:
                        os.remove(video_file)
                        self.add_log(f"已删除视频文件: {video_file}", "success")
                    
                    # 从列表中移除
                    tree.delete(item)
                    
                    self.add_log(f"已删除宫格文件: {grid_file_path}", "success")
                    
                except Exception as e:
                    self.add_log(f"删除文件失败: {str(e)}", "error")
                    messagebox.showerror("删除失败", f"删除文件时出错: {str(e)}")
        else:
            messagebox.showinfo("提示", "无法找到对应的视频文件")
    
    def open_manage_window(self):
        """打开预览管理窗口"""
        # 创建管理窗口
        manage_window = tk.Toplevel(self.root)
        manage_window.title("预览管理")
        manage_window.geometry("900x600")
        manage_window.transient(self.root)  # 设置为临时窗口
        
        # 创建主框架
        main_frame = ttk.Frame(manage_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 搜索和过滤区域
        filter_frame = ttk.LabelFrame(main_frame, text="搜索和过滤", padding="10")
        filter_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 搜索框
        search_var = tk.StringVar()
        ttk.Label(filter_frame, text="搜索:").pack(side=tk.LEFT, padx=(0, 5))
        search_entry = ttk.Entry(filter_frame, textvariable=search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        # 搜索按钮
        search_button = ttk.Button(filter_frame, text="搜索", command=lambda: self.search_previews(search_var.get(), preview_tree))
        search_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # 刷新按钮
        refresh_button = ttk.Button(filter_frame, text="刷新", command=lambda: self.populate_preview_tree(preview_tree))
        refresh_button.pack(side=tk.LEFT)
        
        # 预览列表区域
        preview_frame = ttk.LabelFrame(main_frame, text="宫格预览列表", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 预览列表滚动条
        preview_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL)
        preview_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 预览列表树
        preview_tree = ttk.Treeview(
            preview_frame, 
            columns=("name", "path", "size", "date"),
            show="headings",
            yscrollcommand=preview_scrollbar.set
        )
        preview_tree.heading("name", text="文件名")
        preview_tree.heading("path", text="路径")
        preview_tree.heading("size", text="大小")
        preview_tree.heading("date", text="创建日期")
        preview_tree.column("name", width=200)
        preview_tree.column("path", width=400)
        preview_tree.column("size", width=100)
        preview_tree.column("date", width=150)
        preview_tree.pack(fill=tk.BOTH, expand=True)
        preview_scrollbar.config(command=preview_tree.yview)
        
        # 操作按钮区域
        button_frame = ttk.LabelFrame(main_frame, text="操作", padding="10")
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 预览按钮
        preview_button = ttk.Button(button_frame, text="预览选中项", 
                                   command=lambda: self.preview_selected_item(preview_tree))
        preview_button.pack(side=tk.LEFT, padx=5)
        
        # 删除宫格文件按钮
        delete_grid_button = ttk.Button(button_frame, text="删除宫格文件", 
                                      command=lambda: self.delete_grid_file(preview_tree))
        delete_grid_button.pack(side=tk.LEFT, padx=5)
        
        # 删除宫格和视频按钮
        delete_all_button = ttk.Button(button_frame, text="删除宫格和视频", 
                                     command=lambda: self.delete_all_files(preview_tree))
        delete_all_button.pack(side=tk.LEFT, padx=5)
        
        # 关闭按钮
        close_button = ttk.Button(button_frame, text="关闭", command=manage_window.destroy)
        close_button.pack(side=tk.RIGHT, padx=5)
        
        # 填充预览列表
        self.populate_preview_tree(preview_tree)
        
        # 双击预览
        preview_tree.bind("<Double-1>", lambda event: self.preview_selected_item(preview_tree))


def main():
    # 创建tkinterdnd2根窗口
    root = tkinterdnd2.Tk()
    
    # 创建应用实例，传入root
    app = VideoKeyframeGridApp(root)
    
    # 设置窗口图标（如果有的话）
    try:
        root.iconbitmap('icon.ico')  # 如果有图标文件的话
    except:
        pass  # 忽略图标加载错误
    
    # 居中显示窗口
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()