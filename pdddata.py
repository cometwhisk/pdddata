import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

class SalesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("电商销量异动分析助手 V1.0")
        self.root.geometry("600x700")

        # 存储数据的变量
        self.df = None
        self.blacklist_list = []
        
        # --- 界面组件 ---
        # 1. 文件选择
        tk.Label(root, text="第一步：选择订单表格", font=('Arial', 10, 'bold')).pack(pady=5)
        self.file_btn = tk.Button(root, text="打开 CSV 文件", command=self.load_file)
        self.file_btn.pack()
        self.file_label = tk.Label(root, text="未选择文件", fg="gray")
        self.file_label.pack()

        # 2. 黑名单时间管理
        tk.Frame(root, height=2, bd=1, relief="sunken").pack(fill="x", padx=5, pady=10)
        tk.Label(root, text="第二步：设置黑名单时间段 (格式: YYYY-MM-DD HH:MM:SS)", font=('Arial', 10, 'bold')).pack()
        
        time_frame = tk.Frame(root)
        time_frame.pack(pady=5)
        tk.Label(time_frame, text="开始:").grid(row=0, column=0)
        self.start_entry = tk.Entry(time_frame, width=20)
        self.start_entry.grid(row=0, column=1, padx=5)
        
        tk.Label(time_frame, text="结束:").grid(row=1, column=0)
        self.end_entry = tk.Entry(time_frame, width=20)
        self.end_entry.grid(row=1, column=1, padx=5)
        
        tk.Button(time_frame, text="添加至排除列表", command=self.add_blacklist).grid(row=2, column=0, columnspan=2, pady=5)
        
        self.blacklist_box = tk.Listbox(root, height=4, width=60)
        self.blacklist_box.pack(padx=10)
        tk.Button(root, text="删除选中时段", command=self.remove_blacklist).pack(pady=2)

        # 3. 日期对比选择
        tk.Frame(root, height=2, bd=1, relief="sunken").pack(fill="x", padx=5, pady=10)
        tk.Label(root, text="第三步：选择对比日期", font=('Arial', 10, 'bold')).pack()
        
        date_frame = tk.Frame(root)
        date_frame.pack(pady=5)
        tk.Label(date_frame, text="基准日期(前):").grid(row=0, column=0)
        self.date_a_combo = ttk.Combobox(date_frame, state="readonly")
        self.date_a_combo.grid(row=0, column=1, padx=5)
        
        tk.Label(date_frame, text="对比日期(后):").grid(row=1, column=0)
        self.date_b_combo = ttk.Combobox(date_frame, state="readonly")
        self.date_b_combo.grid(row=1, column=1, padx=5)

        # 4. 执行按钮
        tk.Button(root, text="开始分析并导出结果", bg="green", fg="white", font=('Arial', 12, 'bold'), 
                  command=self.process_data, height=2, width=20).pack(pady=20)

    # --- 功能逻辑 ---
    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            try:
                # 预读检查表头
                temp_df = pd.read_csv(file_path, nrows=0)
                required = ['订单状态', '订单成交时间', '商品id', '商品', '商品规格', '商家编码-商品维度']
                if not all(col in temp_df.columns for col in required):
                    messagebox.showerror("错误", "表格格式不正确，请检查表头！")
                    return
                
                self.file_path = file_path
                self.file_label.config(text=file_path.split('/')[-1], fg="black")
                
                # 读取日期以供选择
                full_df = pd.read_csv(file_path)
                full_df['订单成交时间'] = pd.to_datetime(full_df['订单成交时间'].str.strip())
                dates = sorted(full_df['订单成交时间'].dt.date.unique().astype(str))
                self.date_a_combo['values'] = dates
                self.date_b_combo['values'] = dates
                messagebox.showinfo("成功", "文件读取成功，请选择日期。")
            except Exception as e:
                messagebox.showerror("错误", f"读取失败: {e}")

    def add_blacklist(self):
        start = self.start_entry.get().strip()
        end = self.end_entry.get().strip()
        if start and end:
            self.blacklist_list.append((start, end))
            self.blacklist_box.insert(tk.END, f"排除: {start} 至 {end}")
            self.start_entry.delete(0, tk.END)
            self.end_entry.delete(0, tk.END)

    def remove_blacklist(self):
        selection = self.blacklist_box.curselection()
        if selection:
            index = selection[0]
            self.blacklist_box.delete(index)
            self.blacklist_list.pop(index)

    def process_data(self):
        if not hasattr(self, 'file_path'):
            messagebox.showwarning("提醒", "请先选择文件")
            return
        
        date_a = self.date_a_combo.get()
        date_b = self.date_b_combo.get()
        if not date_a or not date_b:
            messagebox.showwarning("提醒", "请选择要对比的两个日期")
            return

        try:
            # 1. 读取并清理
            df = pd.read_csv(self.file_path)
            # 这里的 .str.strip() 解决了你表格里可能存在的 \t 制表符问题
            df = df[~df['订单状态'].isin(['待付款', '已取消'])]
            df['订单成交时间'] = pd.to_datetime(df['订单成交时间'].str.strip())
            df['成交日期'] = df['订单成交时间'].dt.date.astype(str)

            # 2. 黑名单过滤
            for s, e in self.blacklist_list:
                df = df[~((df['订单成交时间'] >= pd.to_datetime(s)) & (df['订单成交时间'] <= pd.to_datetime(e)))]

            # 3. 汇总
            # 这里加入了“商家编码-商品维度”
            group_cols = ['商品id', '商品', '商品规格', '商家编码-商品维度']
            # 先统一去掉字符串两端的空格或制表符
            for col in group_cols:
                if df[col].dtype == 'object':
                    df[col] = df[col].str.strip()

            daily = df.groupby(['成交日期'] + group_cols).size().reset_index(name='销量')

            # --- 修正后的第 4 步：对比 ---
            # 提取日期 A 的销量，只保留必要列，并重命名销量列
            sales_a = daily[daily['成交日期'] == date_a][group_cols + ['销量']].rename(columns={'销量': f'销量_{date_a}'})
            # 提取日期 B 的销量，同上
            sales_b = daily[daily['成交日期'] == date_b][group_cols + ['销量']].rename(columns={'销量': f'销量_{date_b}'})
            
            # 合并时只根据 group_cols 合并，这样就不会产生多余的日期列了
            comparison = pd.merge(sales_a, sales_b, on=group_cols, how='outer')
            
            # 填充缺失值为 0
            comparison = comparison.fillna(0)
            
            col_a = f'销量_{date_a}'
            col_b = f'销量_{date_b}'
            comparison['单量差'] = comparison[col_b] - comparison[col_a]
            
            # --- 修正后的第 5 步：最终字段筛选与排序 ---
            # 只保留你想要的这几个表头
            final_columns = ['商品id', '商品', '商品规格', '商家编码-商品维度', col_a, col_b, '单量差']
            
            # 检查字段是否都存在（保险起见）
            existing_cols = [c for c in final_columns if c in comparison.columns]
            result = comparison[existing_cols].sort_values(by='单量差', ascending=True)

            # --- 导出结果 ---
            # (后续导出代码保持不变)

            # 6. 导出
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if save_path:
                result.to_excel(save_path, index=False)
                messagebox.showinfo("大功告成", f"分析完成！结果已保存至：\n{save_path}")

        except Exception as e:
            messagebox.showerror("运行错误", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = SalesApp(root)
    root.mainloop()