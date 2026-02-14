import customtkinter as ctk
import tkinter.filedialog as fd
import threading
import pandas as pd
import numpy as np
import xlwt
import math
import os
import time

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")


class DataAugmentApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("肌电数据批量扩容工具 v2.8 - XLS大文件修正版")
        self.geometry("620x620")
        self.resizable(False, False)

        self.input_dir = ctk.StringVar()
        self.output_dir = ctk.StringVar()
        self.is_paused = False
        self.is_cancelled = False
        self.pause_condition = threading.Condition()
        self.setup_ui()

    def select_input(self):
        folder = fd.askdirectory(title="选择源文件夹")
        if folder: self.input_dir.set(folder)

    def select_output(self):
        folder = fd.askdirectory(title="选择保存文件夹")
        if folder: self.output_dir.set(folder)

    def check_warning(self, event):
        val = self.entry_mult.get()
        if val.isdigit() and int(val) > 50:
            self.lbl_warning.pack(side="left")
        else:
            self.lbl_warning.pack_forget()

    def setup_ui(self):
        lbl_title = ctk.CTkLabel(self, text="EMG Data Augmentation System", font=ctk.CTkFont(size=22, weight="bold"),
                                 text_color="black")
        lbl_title.pack(pady=(20, 15))
        self.create_path_row("输入路径:", self.input_dir, self.select_input)
        self.create_path_row("输出路径:", self.output_dir, self.select_output)

        frame_mult = ctk.CTkFrame(self, fg_color="transparent")
        frame_mult.pack(fill="x", padx=40, pady=10)
        ctk.CTkLabel(frame_mult, text="克隆倍数:", text_color="black", font=("Arial", 14)).pack(side="left")
        self.entry_mult = ctk.CTkEntry(frame_mult, width=80, text_color="black")
        self.entry_mult.insert(0, "20")
        self.entry_mult.pack(side="left", padx=10)
        self.entry_mult.bind("<KeyRelease>", self.check_warning)
        self.lbl_warning = ctk.CTkLabel(frame_mult, text="⚠️ 警告：倍数过高易内存溢出", text_color="red",
                                        font=("Arial", 12))
        self.lbl_warning.pack_forget()

        self.lbl_total = ctk.CTkLabel(self, text="总体任务进度: 0% (0/0)", text_color="black",
                                      font=("Arial", 13, "bold"))
        self.lbl_total.pack(pady=(20, 0), padx=45, anchor="w")
        self.prog_total = ctk.CTkProgressBar(self, width=530)
        self.prog_total.pack(pady=5)
        self.prog_total.set(0)

        self.lbl_single = ctk.CTkLabel(self, text="单文件处理进度: 0% (0/0)", text_color="black", font=("Arial", 13))
        self.lbl_single.pack(pady=(10, 0), padx=45, anchor="w")
        self.prog_single = ctk.CTkProgressBar(self, width=530, progress_color="#2FA572")
        self.prog_single.pack(pady=5)
        self.prog_single.set(0)

        self.lbl_status = ctk.CTkLabel(self, text="就绪：请选择路径后开始", text_color="#555555", font=("Arial", 12))
        self.lbl_status.pack(pady=5)

        frame_btns = ctk.CTkFrame(self, fg_color="transparent")
        frame_btns.pack(pady=20)
        self.btn_start = ctk.CTkButton(frame_btns, text="开始处理", command=self.start_task)
        self.btn_start.pack(side="left", padx=10)
        self.btn_pause = ctk.CTkButton(frame_btns, text="暂停", fg_color="#E59500", command=self.toggle_pause,
                                       state="disabled")
        self.btn_pause.pack(side="left", padx=10)
        self.btn_cancel = ctk.CTkButton(frame_btns, text="取消", fg_color="#D32F2F", command=self.cancel_task,
                                        state="disabled")
        self.btn_cancel.pack(side="left", padx=10)

    def create_path_row(self, label, var, cmd):
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="x", padx=40, pady=5)
        ctk.CTkLabel(frame, text=label, text_color="black", width=70, anchor="e").pack(side="left")
        entry = ctk.CTkEntry(frame, textvariable=var, width=350, text_color="black")
        entry.pack(side="left", padx=10)
        ctk.CTkButton(frame, text="...", width=40, command=cmd).pack(side="left")

    def toggle_pause(self):
        with self.pause_condition:
            self.is_paused = not self.is_paused
            self.btn_pause.configure(text="恢复" if self.is_paused else "暂停")
            self.lbl_status.configure(text="⏸ 已暂停运行" if self.is_paused else "处理中...",
                                      text_color="orange" if self.is_paused else "black")
            if not self.is_paused: self.pause_condition.notify_all()

    def cancel_task(self):
        self.is_cancelled = True
        if self.is_paused: self.toggle_pause()
        self.lbl_status.configure(text="🛑 正在取消任务...", text_color="red")

    def start_task(self):
        if not self.input_dir.get() or not self.output_dir.get():
            self.lbl_status.configure(text="❌ 请先选择路径", text_color="red")
            return
        self.is_cancelled = False
        self.is_paused = False
        self.btn_start.configure(state="disabled")
        self.btn_pause.configure(state="normal")
        self.btn_cancel.configure(state="normal")
        threading.Thread(target=self.process, daemon=True).start()

    def process(self):
        in_dir, out_dir = self.input_dir.get(), self.output_dir.get()
        mult = int(self.entry_mult.get()) if self.entry_mult.get().isdigit() else 20
        files = [os.path.join(r, f) for r, d, fs in os.walk(in_dir) for f in fs if f.endswith('.csv')]

        for idx, path in enumerate(files, 1):
            if self.is_cancelled: break
            self.lbl_total.configure(text=f"总体进度: {int((idx - 1) / len(files) * 100)}% ({idx}/{len(files)})")

            rel = os.path.relpath(os.path.dirname(path), in_dir)
            target_dir = os.path.join(out_dir, rel)
            os.makedirs(target_dir, exist_ok=True)

            self.augment_file(path, target_dir, mult)
            self.prog_total.set(idx / len(files))

        self.lbl_status.configure(text="✅ 处理完成！" if not self.is_cancelled else "❌ 已取消",
                                  text_color="green" if not self.is_cancelled else "red")
        self.btn_start.configure(state="normal")
        self.btn_pause.configure(state="disabled")
        self.btn_cancel.configure(state="disabled")

    def augment_file(self, path, out_dir, mult):
        try:
            file_name = os.path.basename(path)
            df = pd.read_csv(path)
            num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            time_col = 'Time_s' if 'Time_s' in df.columns else df.columns[0]
            dt = (df[time_col].iloc[-1] - df[time_col].iloc[0]) / (len(df) - 1) if len(df) > 1 else 0.001
            noise_cols = [c for c in num_cols if c != time_col]

            wb = xlwt.Workbook()
            MAX_XLS = 65535  # XLS单表上限

            def write_large_df(base_name, target_df):
                """自动分片写入逻辑"""
                chunks = math.ceil(len(target_df) / MAX_XLS)
                for c_idx in range(chunks):
                    # 命名规则：如果只有一片则不带后缀，多片则带 _p1, _p2
                    s_name = f"{base_name}_p{c_idx + 1}" if chunks > 1 else base_name
                    ws = wb.add_sheet(s_name)
                    # 写入表头
                    for col_i, col_n in enumerate(target_df.columns): ws.write(0, col_i, col_n)
                    # 写入切片数据
                    chunk_data = target_df.iloc[c_idx * MAX_XLS: (c_idx + 1) * MAX_XLS].values
                    for row_i, row_v in enumerate(chunk_data):
                        for col_i, val in enumerate(row_v): ws.write(row_i + 1, col_i, val)

            # 1. 写入 subject0
            df_s0 = df.copy()
            df_s0[time_col] = [i * dt for i in range(len(df_s0))]
            write_large_df('subject0', df_s0)

            # 2. 写入克隆数据
            for i in range(1, mult + 1):
                with self.pause_condition:
                    while self.is_paused: self.pause_condition.wait()
                if self.is_cancelled: return

                c_df = df.copy()
                if noise_cols: c_df[noise_cols] += np.random.normal(0, 0.01, c_df[noise_cols].shape)
                c_df[time_col] = [j * dt for j in range(len(c_df))]

                write_large_df(f'subject{i}', c_df)
                self.lbl_single.configure(text=f"单文件进度: {int(i / mult * 100)}% ({i}/{mult})")
                self.prog_single.set(i / mult)

            # 3. 保存
            if not self.is_cancelled:
                self.lbl_status.configure(text="📂 正在写入磁盘，请稍候...", text_color="#1F6AA5")
                self.update_idletasks()
                wb.save(os.path.join(out_dir, file_name.replace('.csv', '.xls')))
                self.lbl_status.configure(text=f"✅ 已成功保存: {file_name.replace('.csv', '.xls')}",
                                          text_color="#2FA572")
                time.sleep(0.2)
        except Exception as e:
            print(f"Error {path}: {e}")


if __name__ == "__main__":
    DataAugmentApp().mainloop()