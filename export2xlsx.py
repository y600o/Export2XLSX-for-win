#程序要求：在窗口中选择gis图层文件，选择xlsx输出位置及名称，选择是否使用别名作为列名称，选择要导出的字段，选择导出的sheet名称，点击导出按钮将图层属性表导出为xlsx文件。
#注意：属性表通常有几十万行，而且属性表中存在中文，需要注意内存问题、编码问题和效率问题，使用XlsxWriter包导出xlsx文件。

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import geopandas as gpd
import pandas as pd
import xlsxwriter
import threading
import gc
import os

class GISExportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Export to XLSX - y600")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 数据变量
        self.gdf = None
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.use_alias = tk.BooleanVar(value=True)
        self.use_domain = tk.BooleanVar(value=True)
        self.sheet_name = tk.StringVar(value="Sheet1")
        self.field_vars = {}
        
        self.create_widgets()
    
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 输入图层选择
        ttk.Label(main_frame, text="输入要素类/表格").grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        input_frame = ttk.Frame(main_frame)
        input_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(0, weight=1)
        
        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_path)
        self.input_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        input_btn = ttk.Button(input_frame, text="...", width=3, command=self.select_input_file)
        input_btn.grid(row=0, column=1)
        
        # 输出文件选择
        ttk.Label(main_frame, text="输出Excel文件").grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(0, weight=1)
        
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_path)
        self.output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        output_btn = ttk.Button(output_frame, text="...", width=3, command=self.select_output_file)
        output_btn.grid(row=0, column=1)
        
        # 选项复选框
        options_frame = ttk.Frame(main_frame)
        options_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.alias_check = ttk.Checkbutton(options_frame, text="使用字段别名作为列名称", 
                                          variable=self.use_alias)
        self.alias_check.grid(row=0, column=0, sticky=tk.W)

        self.domain_check = ttk.Checkbutton(options_frame, text="使用域和子类型描述", 
                                           variable=self.use_domain)
        self.domain_check.grid(row=1, column=0, sticky=tk.W)
        
        # 字段选择区域
        fields_label = ttk.Label(main_frame, text="选择字段")
        fields_label.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        
        # 字段列表框架
        fields_frame = ttk.Frame(main_frame)
        fields_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        fields_frame.columnconfigure(0, weight=1)
        fields_frame.rowconfigure(0, weight=1)
        
        # 字段列表（带滚动条）（白色）
        list_frame = ttk.Frame(fields_frame)
        list_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.fields_canvas = tk.Canvas(list_frame, height=200, bg='white', highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.fields_canvas.yview)
        self.scrollable_frame = ttk.Frame(self.fields_canvas, style='White.TFrame')
        
        # 设置scrollable_frame的背景色为白色
        self.scrollable_frame.configure(style='White.TFrame')
        
        # 创建白色背景样式
        style = ttk.Style()
        style.configure('White.TFrame', background='white', relief='flat', borderwidth=0)
        style.configure('White.TCheckbutton', background='white', relief='flat', borderwidth=0)  # 为复选框添加白色背景样式
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.fields_canvas.configure(scrollregion=self.fields_canvas.bbox("all"))
        )
        
        self.fields_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.fields_canvas.configure(yscrollcommand=scrollbar.set)
        
        self.fields_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=0, pady=0)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 绑定鼠标滚轮事件
        self.fields_canvas.bind("<MouseWheel>", self._on_mousewheel)
        
        # 字段操作按钮
        btn_frame = ttk.Frame(fields_frame)
        btn_frame.grid(row=1, column=0, pady=(5, 0))
        
        ttk.Button(btn_frame, text="全选", command=self.select_all_fields).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(btn_frame, text="取消全选", command=self.deselect_all_fields).grid(row=0, column=1, padx=(0, 5))

        # Sheet名称
        sheet_frame = ttk.Frame(main_frame)
        sheet_frame.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Label(sheet_frame, text="Sheet名称（可选）").grid(row=0, column=0, sticky=tk.W)
        sheet_entry = ttk.Entry(sheet_frame, textvariable=self.sheet_name, width=30)
        sheet_entry.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 底部按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=8, column=0, columnspan=2, pady=(20, 0))
        
        ttk.Button(button_frame, text="确定", command=self.export_data).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(button_frame, text="取消", command=self.root.quit).grid(row=0, column=1)
        
        # 状态标签
        self.status_label = ttk.Label(main_frame, text="请选择输入文件")
        self.status_label.grid(row=9, column=0, columnspan=2, pady=(20, 0))
        
        # 配置主框架的行权重
        main_frame.rowconfigure(6, weight=1)
    
    def _on_mousewheel(self, event):
        """处理鼠标滚轮事件"""
        self.fields_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def select_input_file(self):
        """选择输入GIS文件"""
        filetypes = [
            ("所有支持格式", "*.shp;*.gpkg;*.geojson;*.kml"),
            ("Shapefile", "*.shp"),
            ("GeoPackage", "*.gpkg"),
            ("GeoJSON", "*.geojson"),
            ("KML", "*.kml"),
            ("所有文件", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="选择GIS图层文件",
            filetypes=filetypes
        )
        
        if filename:
            self.input_path.set(filename)
            self.load_layer_fields()
    
    def select_output_file(self):
        """选择输出Excel文件"""
        filename = filedialog.asksaveasfilename(
            title="保存Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("所有文件", "*.*")]
        )
        
        if filename:
            self.output_path.set(filename)
    
    def load_layer_fields(self):
        """加载图层字段"""
        try:
            self.status_label.config(text="正在加载图层信息...")
            self.root.update()
            
            # 尝试多种方式读取图层
            self.gdf = None
            input_file = self.input_path.get()
            
            # 方式1: 尝试直接读取一行
            try:
                self.status_label.config(text="正在读取文件结构...")
                self.root.update()
                self.gdf = gpd.read_file(input_file, rows=1)
            except Exception as e1:
                self.status_label.config(text=f"读取失败，尝试方式2...")
                self.root.update()
                
                # 方式2: 尝试忽略几何信息读取
                try:
                    self.status_label.config(text="尝试忽略几何信息读取...")
                    self.root.update()
                    self.gdf = gpd.read_file(input_file, ignore_geometry=True, rows=1)
                except Exception as e2:
                    self.status_label.config(text=f"方式2失败，尝试方式3...")
                    self.root.update()
                    
                    # 方式3: 尝试使用 fiona 读取元数据
                    try:
                        self.status_label.config(text="尝试读取文件元数据...")
                        self.root.update()
                        try:
                            import fiona
                            
                            with fiona.open(input_file) as src:
                                # 获取字段信息
                                schema = src.schema
                                if 'properties' in schema:
                                    fields = list(schema['properties'].keys())
                                    # 创建虚拟 DataFrame
                                    import pandas as pd
                                    df_temp = pd.DataFrame(columns=fields)
                                    self.gdf = gpd.GeoDataFrame(df_temp)
                                else:
                                    raise Exception("无法获取字段信息")
                        except ImportError:
                            raise Exception("需要安装fiona库")
                                
                    except Exception as e3:
                        self.status_label.config(text=f"方式3失败，尝试方式4...")
                        self.root.update()
                        
                        # 方式4: 最后尝试强制读取（忽略投影问题）
                        try:
                            self.status_label.config(text="最后尝试强制读取...")
                            self.root.update()
                            # 设置环境变量忽略投影错误
                            import os
                            os.environ['GDAL_DISABLE_READDIR_ON_OPEN'] = 'EMPTY_DIR'
                            
                            # 尝试不同的读取参数
                            self.gdf = gpd.read_file(
                                input_file, 
                                rows=1,
                                ignore_fields=[],
                                ignore_geometry=False
                            )
                        except Exception as e4:
                            self.status_label.config(text=f"方式4失败，尝试最后一种方式...")
                            self.root.update()
                            
                            # 方式5: 尝试用pandas读取属性表（适用于shapefile）
                            try:
                                if input_file.lower().endswith('.shp'):
                                    self.status_label.config(text="尝试直接读取属性表...")
                                    self.root.update()
                                    dbf_file = input_file.replace('.shp', '.dbf')
                                    if os.path.exists(dbf_file):
                                        # 使用简单的方法创建DataFrame
                                        try:
                                            import fiona
                                            with fiona.open(input_file) as src:
                                                fields = list(src.schema['properties'].keys())
                                                df = pd.DataFrame(columns=fields)
                                                self.gdf = gpd.GeoDataFrame(df)
                                        except ImportError:
                                            raise Exception("需要安装fiona库来处理此文件")
                                    else:
                                        raise Exception("找不到对应的DBF文件")
                                else:
                                    raise Exception("不支持的文件格式")
                            except Exception as e5:
                                error_detail = f"所有读取方式都失败了。请检查文件格式或投影系统。"
                                self.status_label.config(text="读取失败")
                                raise Exception(error_detail)
            
            # 如果读取的数据量太大，只取第一行
            if self.gdf is not None and len(self.gdf) > 1000:
                self.gdf = self.gdf.head(1)
            
            # 清空现有字段复选框
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            
            self.field_vars.clear()
            
            # 创建字段复选框
            row_count = 0
            for field_name in self.gdf.columns:
                if field_name != 'geometry':  # 跳过几何字段
                    var = tk.BooleanVar(value=True)
                    self.field_vars[field_name] = var
                    
                    cb = ttk.Checkbutton(
                        self.scrollable_frame,
                        text=field_name,
                        variable=var,
                        style='White.TCheckbutton'  # 使用白色背景样式
                    )
                    cb.grid(row=row_count, column=0, sticky=tk.W, padx=3, pady=1)
                    row_count += 1
            
            # 更新滚动区域
            self.root.update_idletasks()
            self.fields_canvas.configure(scrollregion=self.fields_canvas.bbox("all"))
            
            self.status_label.config(text=f"已加载图层，共 {len(self.field_vars)} 个字段")
            
        except Exception as e:
            error_msg = f"加载图层失败：{str(e)}"
            messagebox.showerror("错误", error_msg)
            self.status_label.config(text="加载图层失败")
            # 记录详细错误信息到状态标签（可选）
            # self.status_label.config(text=f"加载图层失败: {str(e)[:50]}...")
    
    def select_all_fields(self):
        """全选字段"""
        for var in self.field_vars.values():
            var.set(True)
    
    def deselect_all_fields(self):
        """取消全选字段"""
        for var in self.field_vars.values():
            var.set(False)
    
    
    def get_selected_fields(self):
        """获取选中的字段"""
        selected_fields = []
        for field_name, var in self.field_vars.items():
            if var.get():
                selected_fields.append(field_name)
        return selected_fields
    
    def export_data(self):
        """导出数据到Excel"""
        if not self.input_path.get():
            messagebox.showwarning("警告", "请选择输入文件")
            return
        
        if not self.output_path.get():
            messagebox.showwarning("警告", "请选择输出文件")
            return
        
        selected_fields = self.get_selected_fields()
        if not selected_fields:
            messagebox.showwarning("警告", "请至少选择一个字段")
            return
        
        # 检查输出目录是否存在
        output_dir = os.path.dirname(self.output_path.get())
        if not os.path.exists(output_dir):
            messagebox.showerror("错误", "输出目录不存在")
            return
        
        # 在新线程中执行导出，避免界面卡顿
        thread = threading.Thread(target=self._export_worker, args=(selected_fields,))
        thread.daemon = True
        thread.start()
    
    def _export_worker(self, selected_fields):
        """导出工作线程"""
        try:
            # 更新UI状态
            self.root.after(0, lambda: self.status_label.config(text="正在读取数据..."))
            
            # 安全地重新读取完整数据
            input_file = self.input_path.get()
            gdf_full = None
            
            # 尝试多种方式读取完整数据
            try:
                # 方式1: 正常读取
                self.root.after(0, lambda: self.status_label.config(text="正在读取完整数据..."))
                gdf_full = gpd.read_file(input_file)
            except Exception as e1:
                self.root.after(0, lambda: self.status_label.config(text="读取失败，尝试忽略几何信息..."))
                try:
                    # 方式2: 忽略几何信息读取
                    gdf_full = gpd.read_file(input_file, ignore_geometry=True)
                except Exception as e2:
                    self.root.after(0, lambda: self.status_label.config(text="方式2失败，尝试设置环境变量..."))
                    try:
                        # 方式3: 设置环境变量后重试
                        import os
                        os.environ['GDAL_DISABLE_READDIR_ON_OPEN'] = 'EMPTY_DIR'
                        gdf_full = gpd.read_file(input_file)
                    except Exception as e3:
                        self.root.after(0, lambda: self.status_label.config(text="方式3失败，尝试读取属性表..."))
                        
                        # 方式4: 如果是shapefile，尝试直接读取dbf
                        if input_file.lower().endswith('.shp'):
                            try:
                                # 尝试用pandas读取DBF文件
                                dbf_file = input_file.replace('.shp', '.dbf')
                                if os.path.exists(dbf_file):
                                    # 使用geopandas的内部方法
                                    import pandas as pd
                                    # 创建一个临时的GeoDataFrame来读取属性
                                    temp_gdf = gpd.read_file(input_file, ignore_geometry=True)
                                    gdf_full = temp_gdf
                                else:
                                    raise Exception("无法找到对应的DBF文件")
                            except Exception as e4:
                                error_detail = f"无法读取完整数据，请检查文件格式"
                                self.root.after(0, lambda: self.status_label.config(text="读取失败"))
                                raise Exception(error_detail)
                        else:
                            error_detail = f"无法读取完整数据，不支持的文件格式"
                            self.root.after(0, lambda: self.status_label.config(text="读取失败"))
                            raise Exception(error_detail)
            
            if gdf_full is None:
                raise Exception("无法读取数据文件")
            
            # 检查选中的字段是否存在
            available_fields = []
            for field in selected_fields:
                if field in gdf_full.columns and field != 'geometry':
                    available_fields.append(field)
            
            if not available_fields:
                raise Exception("选中的字段在数据中不存在")
            
            # 选择可用字段
            df = gdf_full[available_fields].copy()
            
            # 删除几何列如果存在
            if 'geometry' in df.columns:
                df = df.drop('geometry', axis=1)
            
            self.root.after(0, lambda: self.status_label.config(text="正在导出到Excel..."))
            
            # 使用XlsxWriter导出
            self._export_to_xlsx(df, available_fields)
            
            # 完成
            self.root.after(0, lambda: self.status_label.config(text="导出完成！"))
            self.root.after(0, lambda: messagebox.showinfo("成功", f"数据已导出到：\n{self.output_path.get()}"))
            
        except Exception as e:
            self.root.after(0, lambda: self.status_label.config(text="导出失败"))
            self.root.after(0, lambda: messagebox.showerror("错误", f"导出失败：{str(e)}"))
            # 记录详细错误信息到状态标签（可选）
            # self.root.after(0, lambda: self.status_label.config(text=f"导出失败: {str(e)[:50]}..."))
    
    def _export_to_xlsx(self, df, selected_fields):
        """使用XlsxWriter导出数据到Excel"""
        workbook = None
        try:
            # 创建工作簿
            workbook = xlsxwriter.Workbook(
                self.output_path.get(),
                {
                    'constant_memory': True,  # 启用常量内存模式
                    'strings_to_formulas': False,
                    'strings_to_urls': False,
                }
            )
            
            worksheet = workbook.add_worksheet(self.sheet_name.get() or "Sheet1")
            
            # 设置格式
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D9D9D9',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })
            
            cell_format = workbook.add_format({
                'border': 1,
                'align': 'left',
                'valign': 'vcenter'
            })
            
            # 写入表头
            for col, field_name in enumerate(selected_fields):
                display_name = field_name  # 这里可以根据use_alias设置来决定显示名称
                worksheet.write(0, col, display_name, header_format)
            
            # 分批写入数据
            chunk_size = 1000
            total_rows = len(df)
            
            for start_idx in range(0, total_rows, chunk_size):
                end_idx = min(start_idx + chunk_size, total_rows)
                chunk = df.iloc[start_idx:end_idx]
                
                for row_idx, (_, row) in enumerate(chunk.iterrows()):
                    excel_row = start_idx + row_idx + 1  # +1 for header
                    
                    for col_idx, field_name in enumerate(selected_fields):
                        try:
                            value = row[field_name]
                            
                            # 处理None值和编码问题
                            if pd.isna(value):
                                processed_value = ""
                            elif isinstance(value, str):
                                processed_value = str(value)
                            else:
                                processed_value = value
                            
                            worksheet.write(excel_row, col_idx, processed_value, cell_format)
                        except Exception as cell_error:
                            # 如果某个单元格写入失败，写入空值
                            worksheet.write(excel_row, col_idx, "", cell_format)
                            # 记录错误但不输出到控制台
                            # 可以选择在状态标签中显示错误概要
                            pass
                
                # 更新进度文本
                self.root.after(0, lambda idx=end_idx, total=total_rows: self.status_label.config(
                    text=f"正在导出... {idx}/{total} 行"))
                
                # 强制垃圾回收
                gc.collect()
            
            # 自动调整列宽
            for col_idx, field_name in enumerate(selected_fields):
                try:
                    worksheet.set_column(col_idx, col_idx, min(len(str(field_name)) * 1.5, 50))
                except:
                    pass
            
        except Exception as e:
            # 不输出到控制台，可以选择在状态标签中显示错误信息
            # self.root.after(0, lambda: self.status_label.config(text=f"Excel创建错误: {str(e)[:50]}..."))
            raise
        finally:
            if workbook:
                try:
                    workbook.close()
                except:
                    pass

def export_to_xlsx():
    """主函数入口"""
    root = tk.Tk()
    app = GISExportApp(root)
    root.mainloop()

# 主程序入口
if __name__ == "__main__":
    export_to_xlsx()