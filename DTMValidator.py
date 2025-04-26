import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # 设置为非GUI后端
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
from scipy.spatial import Delaunay
import numpy as np
from threading import Thread
from datetime import datetime

matplotlib.rcParams['font.sans-serif'] = ['SimHei']
matplotlib.rcParams['axes.unicode_minus'] = False

class BatchDTMValidatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DTM 批量验证工具")
        self.root.geometry("900x600")

        # 初始化时显示文件格式提示
        self.show_file_format_warning()  # <-- 新增的提示方法

        # 初始化变量
        self.file_list = []
        self.processing = False
        self.current_process = 0
        self.total_files = 0
        
        # 创建界面组件
        self.create_widgets()
        
        # 配置项目路径
        self.PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
        self.OUTPUT_DIR = os.path.join(self.PROJECT_ROOT, 'output')
        os.makedirs(self.OUTPUT_DIR, exist_ok=True)

    def show_file_format_warning(self):
        """显示文件格式要求提示"""
        warning_msg = """
        ⚠️ 请确认Excel文件格式 ⚠️

        必须满足以下要求：
        1. 不含表头（第一行为数据）
        2. 列顺序严格为：
           • 第1列：点名称
           • 第2列：X坐标
           • 第3列：Y坐标 
           • 第4列：Z坐标

        注意：坐标值必须为数字类型
        """
        messagebox.showinfo("文件格式说明", warning_msg)

    def create_widgets(self):
        # 文件表格部分
        table_frame = ttk.LabelFrame(self.root, text="文件列表")
        table_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(table_frame, columns=('file', 'height', 'status'), show='headings')
        self.tree.heading('file', text='文件路径')
        self.tree.heading('height', text='基准高程')
        self.tree.heading('status', text='处理状态')
        self.tree.column('file', width=600)
        self.tree.column('height', width=150)
        self.tree.column('status', width=150)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # 表格操作按钮
        btn_frame = ttk.Frame(table_frame)
        btn_frame.pack(pady=5)
        
        ttk.Button(btn_frame, text="添加文件", command=self.add_files).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="添加文件夹", command=self.add_directory).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="移除选中", command=self.remove_selected).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="清空列表", command=self.clear_list).pack(side=tk.LEFT, padx=3)
        
        # 参数和控制部分
        control_frame = ttk.Frame(self.root)
        control_frame.pack(padx=10, pady=5, fill=tk.X)
        
        # 批量高程设置
        ttk.Label(control_frame, text="批量设置高程:").pack(side=tk.LEFT)
        self.batch_height = tk.DoubleVar(value=10.0)
        ttk.Entry(control_frame, textvariable=self.batch_height, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="应用到选中", command=self.apply_batch_height).pack(side=tk.LEFT, padx=5)
        
        # 在控制区域添加提示按钮
        ttk.Button(control_frame, text="格式说明", 
                 command=self.show_file_format_warning).pack(side=tk.RIGHT, padx=5)

        # 处理控制
        self.process_btn = ttk.Button(control_frame, text="开始批量处理", command=self.start_processing)
        self.process_btn.pack(side=tk.RIGHT, padx=5)
        ttk.Button(control_frame, text="打开输出目录", command=self.open_output_dir).pack(side=tk.RIGHT, padx=5)
        
        # 进度条
        self.progress = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, mode='determinate')
        
        # 在界面底部添加状态栏
        self.status_bar = ttk.Label(self.root, text="准备就绪", relief=tk.SUNKEN)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # 日志窗口
        log_frame = ttk.LabelFrame(self.root, text="处理日志")
        log_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(log_frame, height=8)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def add_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        for f in files:
            self.tree.insert('', 'end', values=(f, self.batch_height.get(), '等待处理'))
    
    def add_directory(self):
        folder = filedialog.askdirectory()
        if folder:
            for f in os.listdir(folder):
                if f.lower().endswith('.xlsx'):
                    path = os.path.join(folder, f)
                    self.tree.insert('', 'end', values=(path, self.batch_height.get(), '等待处理'))
    
    def remove_selected(self):
        for item in self.tree.selection():
            self.tree.delete(item)
    
    def clear_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
    
    def apply_batch_height(self):
        try:
            height = np.float64(self.batch_height.get())  # 获取并转换为float64
            for item in self.tree.selection():
                current_values = list(self.tree.item(item)['values'])
                current_values[1] = height
                self.tree.item(item, values=current_values)
        
        except ValueError:
            messagebox.showerror("输入错误", "高程值必须为数字")
        return
         
    def open_output_dir(self):
        if os.path.exists(self.OUTPUT_DIR):
            os.startfile(self.OUTPUT_DIR)
        else:
            messagebox.showwarning("目录不存在", "输出目录尚未创建，请先处理文件")
    
    def start_processing(self):
        if self.processing:
            return
        
        if not self.tree.get_children():
            messagebox.showerror("错误", "请先添加要处理的文件")
            return
        
        # 初始化进度
        self.total_files = len(self.tree.get_children())
        self.current_process = 0
        self.progress['maximum'] = self.total_files
        
        # 禁用按钮并显示进度条
        self.processing = True
        self.process_btn.config(text="处理中...", state=tk.DISABLED)
        self.progress.pack(fill=tk.X, padx=10, pady=5)
        
        # 创建处理线程
        Thread(target=self.batch_process).start()
    
    def batch_process(self):
        success_count = 0
        error_count = 0
        total_volume = 0.0
        report_data = []
        
        for item in self.tree.get_children():
            if not self.processing:  # 允许中断处理
                break
            
            values = self.tree.item(item)['values']
            file_path = values[0]
            # 确保从treeview获取的是float类型
            try:
                fixed_height = float(values[1])  # 显式类型转换
            except (TypeError, ValueError) as e:
                self.log(f"文件【{os.path.basename(file_path)}】高程值无效：{values[1]}")
                self.update_item_status(item, "✗ 高程错误")
                continue

            try:
                # 更新状态为处理中
                self.update_item_status(item, "处理中...")
                
                # 执行处理
                result = self.process_single_file(file_path, fixed_height, item)
                
                # 更新状态和统计
                if result:
                    self.update_item_status(item, "✓ 完成")
                    success_count += 1
                    total_volume += result['volume']
                    report_data.append(result)
                else:
                    self.update_item_status(item, "✗ 失败")
                    error_count += 1
                
            except Exception as e:
                error_msg = f"处理错误 {os.path.basename(file_path)}:\n{str(e)}"
                self.log(error_msg)
                
                # 显示详细错误位置
                if "发现非数值数据" in str(e):
                    self.log("请检查Excel文件中以下位置：")
                    self.log(str(e).split(":", 1)[-1])
                
                self.update_item_status(item, "✗ 数据错误")
                error_count += 1
            finally:
                self.current_process += 1
                self.update_progress()
        
        # 生成汇总报告
        self.generate_summary_report(report_data)
        
        # 显示处理结果
        self.log("\n处理完成！")
        self.log(f"成功: {success_count} 个文件")
        self.log(f"失败: {error_count} 个文件")
        self.log(f"总体积合计: {total_volume:.2f} 立方米")
        messagebox.showinfo("处理完成", f"处理完成！成功{success_count}个，失败{error_count}个")
        
        self.root.after(0, self.reset_ui)

    def process_single_file(self, file_path, fixed_height, item):
        try:
            raw_vertices = self.read_vertices(file_path)
            triangulated = self.delaunay_triangulation(raw_vertices)
            total_volume = sum(self.calculate_volume(t, fixed_height) for t in triangulated)
            
            num_vertices = raw_vertices.shape[0]
            top_3d_surface_area = sum(self.calculate_3d_area(t) for t in triangulated)
            valid_points = raw_vertices[raw_vertices[:, 2] > fixed_height]
            bottom_projection_area = sum(self.calculate_triangle_area(t) for t in self.delaunay_triangulation(valid_points)) if len(valid_points) >= 3 else 0
            min_z = np.min(raw_vertices[:, 2])
            max_z = np.max(raw_vertices[:, 2])
            
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            output_image = os.path.join(self.OUTPUT_DIR, f'{base_name}_3d.png')
            
            self.plot_result(triangulated, base_name, num_vertices, top_3d_surface_area, 
                            bottom_projection_area, min_z, max_z, total_volume, fixed_height)
            
            return {
                'filename': os.path.basename(file_path),
                'volume': total_volume,
                'vertices': num_vertices,
                '3d_area': top_3d_surface_area,
                'projection_area': bottom_projection_area,
                'min_z': min_z,
                'max_z': max_z
            }
            
        except Exception as e:
            error_msg = f"处理错误 {os.path.basename(file_path)}:\n{str(e)}"
            self.log(error_msg)
            if "发现非数值数据" in str(e):
                self.log("请检查Excel文件中以下位置：")
                self.log(str(e).split(":", 1)[-1])
            self.update_item_status(item, "✗ 数据错误")

    def read_vertices(self, file_path):
        df = pd.read_excel(file_path, header=None)
        
        # 提取需要的列并转换数据类型，互换X和Y轴
        data_subset = df.iloc[:, 1:4].copy()  # 假设XYZ在2-4列
        numeric_data = data_subset.apply(pd.to_numeric, errors='coerce')
        # print(numeric_data)
        
        # 互换X和Y坐标
        numeric_data[[1, 2]] = numeric_data[[2, 1]].values
        
        # 找出包含非数值的行
        invalid_rows = numeric_data.isnull().any(axis=1)
        if invalid_rows.any():
            error_locations = []
            for idx, row in data_subset[invalid_rows].iterrows():
                bad_cols = [i+2 for i, v in enumerate(row.isnull()) if v]  # 列号从2开始
                error_locations.append(f"第{idx+1}行[{','.join(map(str,bad_cols))}列]")
            raise ValueError(
                f"发现非数值数据\n"
                f"错误位置：\n" + "\n".join(error_locations)
            )
        
        vertices = numeric_data.dropna().values
        
        if vertices.size == 0:
            raise ValueError("有效数据为空，请检查文件内容")
        
        return vertices.astype(np.float64)  # 强制转换为浮点数

    def delaunay_triangulation(self, vertices):
        """执行Delaunay三角剖分"""
        if len(vertices) < 3:
            raise ValueError("至少需要3个顶点进行三角剖分")
        
        # 提取XY坐标进行剖分
        xy_points = vertices[:, :2]
        
        # 进行二维Delaunay三角剖分
        tri = Delaunay(xy_points)
        
        # 构建三角形顶点数组
        triangles = []
        for simplex in tri.simplices:
            triangle = np.array([vertices[simplex[0]], vertices[simplex[1]], vertices[simplex[2]]])
            triangles.append(triangle)
        return np.array(triangles)

    def calculate_triangle_area(self, vertices):
        """计算投影到XY平面的三角形面积"""
        # 提取三个顶点的XY坐标
        v0 = vertices[0][:2]  # [x, y]
        v1 = vertices[1][:2]  # [x, y]
        v2 = vertices[2][:2]  # [x, y]
        
        # 计算向量
        vec1 = np.array(v1) - np.array(v0)
        vec2 = np.array(v2) - np.array(v0)
        
        # 二维叉乘计算（行列式绝对值）
        cross_product = vec1[0] * vec2[1] - vec1[1] * vec2[0]
        return abs(cross_product) / 2  # 返回投影面积

    def calculate_volume(self, triangle, fixed_height):
        area = self.calculate_triangle_area(triangle)
        avg_z = triangle[:, 2].mean()
        return abs(area * (avg_z - fixed_height))

    def calculate_3d_area(self, vertices):
        v0, v1, v2 = vertices
        cross = np.cross(v1 - v0, v2 - v0)
        return np.linalg.norm(cross) / 2

    def plot_result(self, triangles, base_name, num_vertices, top_3d_surface_area, 
                   bottom_projection_area, min_z, max_z, total_volume, fixed_height):
        fig = plt.figure(figsize=(12, 8))
        ax = fig.add_subplot(111, projection='3d')
        
        pc = Poly3DCollection(triangles, alpha=0.5, edgecolor='k', facecolor='cyan')
        ax.add_collection3d(pc)
        
        all_points = np.concatenate(triangles, axis=0)
        ax.set(
            xlim=(all_points[:,0].min(), all_points[:,0].max()),
            ylim=(all_points[:,1].min(), all_points[:,1].max()),
            zlim=(all_points[:,2].min(), all_points[:,2].max()),
            xlabel='X坐标', ylabel='Y坐标', zlabel='高程',
            title=f'{base_name} 三维可视化'
        )

        ax.set_box_aspect([
        all_points[:,0].max()-all_points[:,0].min(),
        all_points[:,1].max()-all_points[:,1].min(),
        all_points[:,2].max()-all_points[:,2].min()
        ])
        
        # 添加统计信息文本标注
        stats_text = (
        f"顶点数量: {num_vertices}\n"
        f"三维表面积: {top_3d_surface_area:.4f} 平方米\n"
        f"有效投影面积（基准面上方）: {bottom_projection_area:.4f} 平方米\n"
        f"基准高程: {fixed_height:.4f} 米\n"
        f"顶点高程范围: {min_z:.4f} ~ {max_z:.4f} 米\n"
        f"生成三角形数量: {len(triangles)}\n"
        f"总体积: {total_volume:.4f} 立方米"
        )
        
        # 在图形中添加文本
        ax.text2D(0.05, 0.95, stats_text, transform=ax.transAxes, fontsize=10,
              verticalalignment='top', bbox=dict(boxstyle='round', facecolor='white', alpha=0.5))

        # 在三角形上标注面积和体积
        for triangle in triangles:
            area = self.calculate_3d_area(triangle)
            volume = self.calculate_volume(triangle, fixed_height)
            centroid = triangle.mean(axis=0)
            ax.text(centroid[0], centroid[1], centroid[2], 
                    f"面积: {area:.2f}\n体积: {volume:.2f}", color='red', fontsize=4)

        plt.savefig(os.path.join(self.OUTPUT_DIR, f'{base_name}_3d.png'), dpi=300)
        plt.close()

    def update_item_status(self, item, status):
        values = list(self.tree.item(item)['values'])
        values[2] = status
        self.root.after(0, lambda i=item, v=values: self.tree.item(i, values=v))
    
    def update_progress(self):
        self.root.after(0, self.progress.step)
        self.progress['value'] = self.current_process
    
    def log(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.root.after(0, self.log_text.insert, tk.END, f"[{timestamp}] {message}\n")
        self.root.after(0, self.log_text.see, tk.END)
    
    def generate_summary_report(self, data):
        if not data:
            return
        
        # 创建DataFrame
        df = pd.DataFrame(data)
        df['处理时间'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # 重新排序列
        df = df[['filename', 'volume', 'vertices', '3d_area', 'projection_area', 'min_z', 'max_z', '处理时间']]
        
        # 保存报告
        report_path = os.path.join(self.OUTPUT_DIR, f"处理报告_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
        df.to_excel(report_path, index=False)
        self.log(f"生成汇总报告: {report_path}")
    
    def reset_ui(self):
        self.progress.pack_forget()
        self.process_btn.config(text="开始批量处理", state=tk.NORMAL)
        self.processing = False

if __name__ == "__main__":
    root = tk.Tk()
    app = BatchDTMValidatorApp(root)
    root.mainloop()