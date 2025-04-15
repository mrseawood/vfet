import os
import sys
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from pathlib import Path
import threading
import time
import importlib.util
import subprocess

class VideoFrameExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("视频帧提取工具")
        self.root.geometry("600x800")
        self.root.resizable(True, True)
        
        # 变量初始化
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.frame_interval = tk.IntVar(value=1)  # 默认每1帧提取一张
        self.processing = False
        self.total_videos = 0
        self.processed_videos = 0
        self.total_frames = 0
        self.extracted_frames = 0
        self.installing = False
        
        # 检查依赖
        self.check_dependencies()
        
        # 创建UI
        self.create_widgets()
    
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 使用说明
        help_frame = ttk.LabelFrame(main_frame, text="使用说明", padding="5")
        help_frame.pack(fill=tk.X, pady=5)
        
        help_text = """1. 选择包含视频文件的输入文件夹
2. 选择要保存提取帧的输出文件夹
3. 设置每隔多少帧提取一张图片（默认为1，即每帧都提取）
4. 点击"开始提取"按钮开始处理
5. 处理过程中可以通过"停止"按钮中断操作

提取的帧将按照原始文件夹结构保存在输出文件夹中，每个视频会创建单独的子文件夹。
"""
        help_label = ttk.Label(help_frame, text=help_text, justify=tk.LEFT)
        help_label.grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        
        # 输入文件夹选择
        input_frame = ttk.LabelFrame(main_frame, text="输入设置", padding="5")
        input_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_frame, text="输入文件夹:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(input_frame, textvariable=self.input_folder, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(input_frame, text="浏览...", command=self.browse_input).grid(row=0, column=2, padx=5, pady=5)
        
        # 输出文件夹选择
        output_frame = ttk.LabelFrame(main_frame, text="输出设置", padding="5")
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_frame, text="输出文件夹:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(output_frame, textvariable=self.output_folder, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(output_frame, text="浏览...", command=self.browse_output).grid(row=0, column=2, padx=5, pady=5)
        
        # 帧率设置
        settings_frame = ttk.LabelFrame(main_frame, text="提取设置", padding="5")
        settings_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(settings_frame, text="每隔多少帧提取一张:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Spinbox(settings_frame, from_=1, to=1000, textvariable=self.frame_interval, width=10).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 进度显示
        progress_frame = ttk.LabelFrame(main_frame, text="进度", padding="5")
        progress_frame.pack(fill=tk.X, pady=5)
        
        self.status_label = ttk.Label(progress_frame, text="就绪")
        self.status_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        ttk.Label(progress_frame, text="视频进度:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.video_progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=500, mode='determinate')
        self.video_progress.grid(row=1, column=1, sticky=tk.W+tk.E, padx=5, pady=2)
        
        ttk.Label(progress_frame, text="总体进度:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.total_progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=500, mode='determinate')
        self.total_progress.grid(row=2, column=1, sticky=tk.W+tk.E, padx=5, pady=2)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = tk.Text(log_frame, height=10, width=70, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        self.start_button = ttk.Button(button_frame, text="开始提取", command=self.start_extraction)
        self.start_button.pack(side=tk.RIGHT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="停止", command=self.stop_extraction, state=tk.DISABLED)
        self.stop_button.pack(side=tk.RIGHT, padx=5)
        
        # 版权信息
        copyright_frame = ttk.Frame(main_frame)
        copyright_frame.pack(fill=tk.X, pady=5)
        
        copyright_text = "© 2025 一模型Ai (https://jmlovestore.com) - 不会开发软件吗 🙂 Ai会哦"
        copyright_label = ttk.Label(copyright_frame, text=copyright_text, foreground="gray")
        copyright_label.pack(side=tk.RIGHT, padx=5)
    
    def browse_input(self):
        folder = filedialog.askdirectory(title="选择输入文件夹")
        if folder:
            # 使用Path对象处理路径，确保对中文和特殊字符的支持
            folder_path = Path(folder)
            self.input_folder.set(str(folder_path))
            self.log(f"已选择输入文件夹: {folder_path}")
    
    def browse_output(self):
        folder = filedialog.askdirectory(title="选择输出文件夹")
        if folder:
            # 使用Path对象处理路径，确保对中文和特殊字符的支持
            folder_path = Path(folder)
            self.output_folder.set(str(folder_path))
            self.log(f"已选择输出文件夹: {folder_path}")
    
    def log(self, message):
        self.log_text.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see(tk.END)
    
    def update_status(self, message):
        self.status_label.config(text=message)
    
    def start_extraction(self):
        # 检查输入和输出文件夹
        if not self.input_folder.get() or not self.output_folder.get():
            messagebox.showerror("错误", "请选择输入和输出文件夹")
            return
        
        if self.processing:
            return
        
        self.processing = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        
        # 重置进度
        self.total_videos = 0
        self.processed_videos = 0
        self.total_frames = 0
        self.extracted_frames = 0
        self.video_progress['value'] = 0
        self.total_progress['value'] = 0
        
        # 启动处理线程
        threading.Thread(target=self.process_videos, daemon=True).start()
    
    def stop_extraction(self):
        if not self.processing:
            return
        
        self.processing = False
        self.update_status("已停止")
        self.log("提取过程已停止")
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
    
    def process_videos(self):
        try:
            input_dir = Path(self.input_folder.get())
            output_dir = Path(self.output_folder.get())
            
            # 获取所有视频文件
            video_extensions = ('.mp4', '.avi', '.mkv', '.mov', '.wmv', '.flv', '.webm')
            video_files = []
            
            self.update_status("正在扫描视频文件...")
            self.log("开始扫描视频文件...")
            
            # 使用Path对象处理文件路径
            for root, _, files in os.walk(input_dir):
                root_path = Path(root)
                for file in files:
                    if file.lower().endswith(video_extensions):
                        video_files.append(root_path / file)
            
            self.total_videos = len(video_files)
            self.log(f"找到 {self.total_videos} 个视频文件")
            
            if self.total_videos == 0:
                self.update_status("未找到视频文件")
                self.processing = False
                self.start_button.config(state=tk.NORMAL)
                self.stop_button.config(state=tk.DISABLED)
                return
            
            # 处理每个视频
            for video_path in video_files:
                if not self.processing:
                    break
                
                # 计算相对路径，以保持文件夹结构
                video_path_obj = Path(video_path)
                rel_path = video_path_obj.parent.relative_to(input_dir)
                video_name = video_path_obj.stem
                
                # 创建对应的输出文件夹
                output_subdir = output_dir / rel_path / video_name
                output_subdir.mkdir(parents=True, exist_ok=True)
                
                self.extract_frames(video_path, output_subdir)
                self.processed_videos += 1
                self.total_progress['value'] = (self.processed_videos / self.total_videos) * 100
            
            if self.processing:  # 如果没有被中途停止
                self.update_status("提取完成")
                self.log(f"所有视频处理完成，共提取 {self.extracted_frames} 帧")
                messagebox.showinfo("完成", f"所有视频处理完成，共提取 {self.extracted_frames} 帧")
            
            self.processing = False
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            
        except Exception as e:
            self.log(f"错误: {str(e)}")
            self.update_status("处理出错")
            self.processing = False
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
    
    def extract_frames(self, video_path, output_dir):
        try:
            # 使用Path对象处理路径，增强对中文和特殊字符的支持
            video_path = Path(video_path)
            output_dir = Path(output_dir)
            
            video_name = video_path.name
            self.update_status(f"正在处理: {video_name}")
            self.log(f"开始处理视频: {video_path}")
            
            # 确保输出目录存在
            try:
                if not output_dir.exists():
                    output_dir.mkdir(parents=True, exist_ok=True)
                    self.log(f"创建输出目录: {output_dir}")
                
                # 检查目录写入权限
                test_file_path = output_dir / "test_write_permission.tmp"
                try:
                    with open(test_file_path, 'w', encoding='utf-8') as f:
                        f.write("test")
                    if test_file_path.exists():
                        test_file_path.unlink()
                except Exception as perm_error:
                    self.log(f"警告: 输出目录可能没有写入权限: {str(perm_error)}")
                    messagebox.showwarning("权限警告", f"输出目录可能没有写入权限，请选择其他目录或检查权限设置。\n{output_dir}")
                    return
            except Exception as dir_error:
                self.log(f"错误: 无法创建或访问输出目录: {str(dir_error)}")
                messagebox.showerror("目录错误", f"无法创建或访问输出目录，请检查路径是否包含特殊字符或权限设置。\n{output_dir}")
                return
            
            # 打开视频文件
            # 将Path对象转换为字符串，确保cv2.VideoCapture能正确处理中文和特殊字符路径
            video_path_str = str(video_path.resolve())
            self.log(f"尝试打开视频文件: {video_path_str}")
            
            # 对于包含特殊字符的路径，尝试使用绝对路径打开
            cap = cv2.VideoCapture(video_path_str)
            if not cap.isOpened():
                self.log(f"无法打开视频: {video_path}，尝试其他方法...")
                
                # 尝试使用其他方式打开视频
                try:
                    # 在Windows系统上，尝试使用短路径名
                    import ctypes
                    kernel32 = ctypes.WinDLL('kernel32')
                    GetShortPathNameW = kernel32.GetShortPathNameW
                    GetShortPathNameW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_uint]
                    GetShortPathNameW.restype = ctypes.c_uint
                    
                    buffer_size = 1024
                    buffer = ctypes.create_unicode_buffer(buffer_size)
                    result_length = GetShortPathNameW(video_path_str, buffer, buffer_size)
                    
                    if result_length > 0 and result_length < buffer_size:
                        short_path = buffer.value
                        self.log(f"尝试使用短路径名打开: {short_path}")
                        cap = cv2.VideoCapture(short_path)
                except Exception as short_path_error:
                    self.log(f"尝试使用短路径名失败: {str(short_path_error)}")
                
                # 如果仍然无法打开，返回错误
                if not cap.isOpened():
                    self.log(f"所有尝试都失败，无法打开视频: {video_path}")
                    return
            
            # 获取视频信息
            total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
            fps = cap.get(cv2.CAP_PROP_FPS)
            
            self.log(f"视频信息: 总帧数={total_frames}, FPS={fps:.2f}")
            self.video_progress['value'] = 0
            
            # 提取帧
            frame_count = 0
            saved_count = 0
            interval = self.frame_interval.get()
            error_count = 0  # 记录连续错误次数
            max_errors = 5   # 最大允许连续错误次数
            
            while cap.isOpened() and self.processing:
                ret, frame = cap.read()
                if not ret:
                    break
                
                # 按指定间隔提取帧
                if frame_count % interval == 0:
                    try:
                        # 处理文件名，使用Path对象处理路径
                        frame_filename = f"frame_{frame_count:06d}.png"
                        output_path = output_dir / frame_filename
                        
                        # 将Path对象转换为字符串，确保cv2.imwrite能正确处理中文路径
                        output_path_str = str(output_path.resolve())
                        self.log(f"尝试保存图片到: {output_path_str}")
                        
                        # 尝试使用短路径名（如果在Windows系统上）
                        try:
                            import ctypes
                            kernel32 = ctypes.WinDLL('kernel32')
                            GetShortPathNameW = kernel32.GetShortPathNameW
                            GetShortPathNameW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_uint]
                            GetShortPathNameW.restype = ctypes.c_uint
                            
                            buffer_size = 1024
                            buffer = ctypes.create_unicode_buffer(buffer_size)
                            result_length = GetShortPathNameW(output_path_str, buffer, buffer_size)
                            
                            if result_length > 0 and result_length < buffer_size:
                                short_path = buffer.value
                                self.log(f"使用短路径名: {short_path}")
                                output_path_str = short_path
                        except Exception as short_path_error:
                            self.log(f"获取短路径名失败，继续使用原路径: {str(short_path_error)}")
                        
                        # 尝试保存图片
                        # 使用临时目录保存，然后移动到目标位置
                        try:
                            import tempfile
                            import shutil
                            
                            # 创建临时文件
                            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                                temp_path = temp_file.name
                            
                            # 保存到临时文件
                            self.log(f"尝试先保存到临时文件: {temp_path}")
                            temp_success = cv2.imwrite(temp_path, frame, [cv2.IMWRITE_PNG_COMPRESSION, 0])  # 使用无损压缩
                            
                            if temp_success:
                                # 将临时文件复制到目标位置
                                shutil.copy2(temp_path, output_path_str)
                                # 删除临时文件
                                os.remove(temp_path)
                                success = True
                                self.log(f"通过临时文件成功保存图片: {output_path}")
                            else:
                                success = False
                                self.log(f"保存到临时文件失败")
                        except Exception as temp_error:
                            self.log(f"使用临时文件方法失败: {str(temp_error)}，尝试直接保存...")
                            # 如果临时文件方法失败，尝试直接保存
                            success = cv2.imwrite(output_path_str, frame, [cv2.IMWRITE_PNG_COMPRESSION, 0])
                        
                        if success:
                            saved_count += 1
                            self.extracted_frames += 1
                            error_count = 0  # 重置错误计数
                            
                            # 验证文件是否真的被创建
                            if not output_path.exists():
                                self.log(f"警告: 文件似乎未被创建: {output_path}")
                                error_count += 1
                        else:
                            self.log(f"错误: 无法保存图片: {output_path}")
                            error_count += 1
                            
                            # 尝试使用不同的格式保存
                            if error_count <= max_errors:
                                try:
                                    jpg_path = output_dir / f"frame_{frame_count:06d}.jpg"
                                    jpg_path_str = str(jpg_path.resolve())
                                    jpg_success = cv2.imwrite(jpg_path_str, frame, [cv2.IMWRITE_JPEG_QUALITY, 95])
                                    if jpg_success:
                                        self.log(f"成功使用JPG格式保存: {jpg_path}")
                                        saved_count += 1
                                        self.extracted_frames += 1
                                        error_count = 0
                                except Exception as jpg_error:
                                    self.log(f"尝试JPG格式保存也失败: {str(jpg_error)}")
                    except Exception as save_error:
                        self.log(f"保存帧时出错: {str(save_error)}")
                        error_count += 1
                    
                    # 如果连续错误次数过多，提示用户并中断处理
                    if error_count >= max_errors:
                        self.log(f"连续出现{max_errors}次保存错误，可能是路径问题或磁盘空间不足")
                        if messagebox.askyesno("错误", f"连续出现{max_errors}次保存错误，是否继续处理？\n\n可能的原因:\n- 输出路径包含特殊字符\n- 磁盘空间不足\n- 没有写入权限"):
                            error_count = 0  # 重置错误计数
                        else:
                            self.log("用户选择中断处理")
                            break
                
                frame_count += 1
                if frame_count % 10 == 0:  # 每10帧更新一次进度，减少UI更新频率
                    self.video_progress['value'] = (frame_count / total_frames) * 100
                    self.root.update_idletasks()
            
            cap.release()
            self.video_progress['value'] = 100
            
            # 检查是否真的保存了文件
            try:
                actual_files = len([f for f in output_dir.glob("*.png")]) + len([f for f in output_dir.glob("*.jpg")])
                self.log(f"视频 {video_name} 处理完成，预期提取 {saved_count} 帧，实际保存 {actual_files} 个文件")
                
                if actual_files == 0 and saved_count > 0:
                    self.log(f"警告: 没有文件被保存到 {output_dir}，请检查权限或磁盘空间")
                    messagebox.showwarning("警告", f"预期保存了 {saved_count} 个文件，但实际未找到任何文件。\n请检查输出目录的权限或磁盘空间。")
            except Exception as check_error:
                self.log(f"检查保存文件时出错: {str(check_error)}")
                self.log(f"尝试使用绝对路径: {output_dir.resolve()}")
            
        except Exception as e:
            self.log(f"处理视频 {video_path.name} 时出错: {str(e)}")
            import traceback
            self.log(f"错误详情: {traceback.format_exc()}")
            messagebox.showerror("处理错误", f"处理视频时出错: {str(e)}")

    def check_dependencies(self):
        """检查必要的库是否已安装"""
        missing_libs = []
        
        # 检查OpenCV
        if importlib.util.find_spec("cv2") is None:
            missing_libs.append(("opencv-python", "OpenCV"))
        
        # 如果有缺失的库，显示安装对话框
        if missing_libs:
            self.show_dependency_dialog(missing_libs)
            return False
        return True
    
    def show_dependency_dialog(self, missing_libs):
        """显示依赖安装对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("缺少依赖")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="检测到缺少以下必要的库:", font=("Arial", 12)).pack(pady=10)
        
        # 创建一个框架来包含缺失库的列表
        libs_frame = ttk.Frame(dialog)
        libs_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        for package, display_name in missing_libs:
            ttk.Label(libs_frame, text=f"• {display_name} ({package})", font=("Arial", 10)).pack(anchor=tk.W, pady=2)
        
        ttk.Label(dialog, text="是否自动安装这些库？", font=("Arial", 10)).pack(pady=10)
        
        # 按钮框架
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, pady=10)
        
        # 进度条和状态标签（初始隐藏）
        self.install_progress = ttk.Progressbar(dialog, orient=tk.HORIZONTAL, length=350, mode='indeterminate')
        self.install_status = ttk.Label(dialog, text="")
        
        def install_dependencies():
            if self.installing:
                return
                
            self.installing = True
            install_btn.config(state=tk.DISABLED)
            cancel_btn.config(state=tk.DISABLED)
            self.install_status.config(text="正在安装依赖，请稍候...")
            self.install_status.pack(pady=5)
            self.install_progress.pack(pady=5, padx=20)
            self.install_progress.start(10)
            
            # 在新线程中安装依赖
            threading.Thread(target=self._install_dependencies, args=(missing_libs, dialog), daemon=True).start()
        
        def cancel_installation():
            dialog.destroy()
            messagebox.showinfo("提示", "请手动安装缺失的库后再运行程序。")
            self.root.quit()
        
        install_btn = ttk.Button(btn_frame, text="安装", command=install_dependencies)
        install_btn.pack(side=tk.RIGHT, padx=10)
        
        cancel_btn = ttk.Button(btn_frame, text="取消", command=cancel_installation)
        cancel_btn.pack(side=tk.RIGHT, padx=10)
    
    def _install_dependencies(self, missing_libs, dialog):
        """安装缺失的依赖"""
        success = True
        error_msg = ""
        
        for package, display_name in missing_libs:
            try:
                self.install_status.config(text=f"正在安装 {display_name}...")
                # 使用subprocess调用pip安装包
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            except subprocess.CalledProcessError as e:
                success = False
                error_msg += f"安装 {display_name} 失败: {str(e)}\n"
        
        self.install_progress.stop()
        self.install_progress.pack_forget()
        
        if success:
            self.install_status.config(text="所有依赖安装成功！请重启应用。")
            ttk.Button(dialog, text="重启应用", command=lambda: self._restart_app()).pack(pady=10)
        else:
            self.install_status.config(text="安装过程中出现错误")
            error_text = tk.Text(dialog, height=5, width=40, wrap=tk.WORD)
            error_text.insert(tk.END, error_msg)
            error_text.config(state=tk.DISABLED)
            error_text.pack(padx=20, pady=10)
            ttk.Button(dialog, text="关闭", command=dialog.destroy).pack(pady=10)
        
        self.installing = False
    
    def _restart_app(self):
        """重启应用"""
        python = sys.executable
        os.execl(python, python, *sys.argv)

if __name__ == "__main__":
    try:
        import cv2
        root = tk.Tk()
        app = VideoFrameExtractor(root)
        root.mainloop()
    except ImportError:
        # 如果导入cv2失败，仍然启动应用，让依赖检查机制处理
        root = tk.Tk()
        app = VideoFrameExtractor(root)
        root.mainloop()