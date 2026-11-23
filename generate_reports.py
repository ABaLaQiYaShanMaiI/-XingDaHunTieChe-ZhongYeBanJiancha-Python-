#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
混铁车中夜班安全消防检查情况报告自动生成工具 GUI v1.0.1
"""

import os
import re
import sys
from datetime import datetime

from report_core import (
    BI_NAME, read_duty_roster, read_minesweeper, generate_docx,
    resolve_duty, get_sweep_area_key, get_sweep_area_label,
)

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
from tkinterdnd2 import DND_FILES, TkinterDnD


APP_TITLE = '混铁车中夜班安全消防检查情况报告自动生成工具 v1.0.1'


class ReportGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1000x820")
        self.root.resizable(True, True)

        self.roster_data = {}
        self.sweep_items = []
        self.format_name = ''
        self.month_var = tk.StringVar(value=str(datetime.now().month))
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        self.item_dates = {}
        self.item_areas = {}
        self.roster_days = []
        self.sel_info_var = tk.StringVar(value='尚未加载隐患登记表')
        self._build_ui()

    def _build_ui(self):
        mf = ttk.Frame(self.root, padding=10)
        mf.pack(fill=tk.BOTH, expand=True)

        # ---- 文件导入区域 ----
        ff = ttk.LabelFrame(mf, text="文件导入（拖入或点击选择）", padding=10)
        ff.pack(fill=tk.X, pady=(0, 8))

        df = ttk.Frame(ff); df.pack(fill=tk.X, pady=3)
        ttk.Label(df, text="值班表(.xls/.xlsx):", width=18).pack(side=tk.LEFT)
        self.duty_label = tk.Label(df, text="拖入值班表 或 点击选择",
            bg='#E8F0FE', fg='#555', relief=tk.GROOVE, height=2, cursor='hand2')
        self.duty_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.duty_label.bind('<Button-1>', lambda e: self._sel_duty())
        self.duty_label.drop_target_register(DND_FILES)
        self.duty_label.dnd_bind('<<Drop>>', self._drop_duty)
        ttk.Button(df, text="选择", width=8, command=self._sel_duty).pack(side=tk.LEFT)

        sf = ttk.Frame(ff); sf.pack(fill=tk.X, pady=3)
        ttk.Label(sf, text="隐患登记表(.xlsx):", width=18).pack(side=tk.LEFT)
        self.sweep_label = tk.Label(sf, text="拖入事故隐患排查整改登记表 或 点击选择",
            bg='#FCE4D6', fg='#555', relief=tk.GROOVE, height=2, cursor='hand2')
        self.sweep_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.sweep_label.bind('<Button-1>', lambda e: self._sel_sweep())
        self.sweep_label.drop_target_register(DND_FILES)
        self.sweep_label.dnd_bind('<<Drop>>', self._drop_sweep)
        ttk.Button(sf, text="选择", width=8, command=self._sel_sweep).pack(side=tk.LEFT)

        # ---- 值班表概况 ----
        self.roster_info_var = tk.StringVar(value="尚未加载值班表")
        ttk.Label(mf, textvariable=self.roster_info_var, foreground='#555').pack(fill=tk.X, pady=(0, 8))

        # ---- 参数设置 ----
        pf = ttk.LabelFrame(mf, text="参数设置", padding=10)
        pf.pack(fill=tk.X, pady=(0, 8))
        r1 = ttk.Frame(pf); r1.pack(fill=tk.X, pady=3)
        ttk.Label(r1, text="年份:").pack(side=tk.LEFT)
        ttk.Entry(r1, textvariable=self.year_var, width=8).pack(side=tk.LEFT, padx=(5,20))
        ttk.Label(r1, text="月份:").pack(side=tk.LEFT)
        ttk.Entry(r1, textvariable=self.month_var, width=8).pack(side=tk.LEFT, padx=(5,20))
        ttk.Label(r1, text="日期范围(日):").pack(side=tk.LEFT)
        self.ds = tk.StringVar(value="1"); self.de = tk.StringVar(value="31")
        ttk.Entry(r1, textvariable=self.ds, width=6).pack(side=tk.LEFT, padx=(5,2))
        ttk.Label(r1, text="-").pack(side=tk.LEFT)
        ttk.Entry(r1, textvariable=self.de, width=6).pack(side=tk.LEFT, padx=(2,20))
        ttk.Label(r1, text="输出目录:").pack(side=tk.LEFT)
        # 默认输出到桌面，避免 exe 运行时 cwd 指向 C:\Windows\system32
        default_out = os.path.join(os.path.expanduser('~'), 'Desktop')
        if not os.path.isdir(default_out):
            default_out = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.getcwd()
        self.out_dir = tk.StringVar(value=default_out)
        ttk.Entry(r1, textvariable=self.out_dir, width=35).pack(side=tk.LEFT, padx=(5,5))
        ttk.Button(r1, text="浏览", width=8, command=self._sel_outdir).pack(side=tk.LEFT)

        # ---- 隐患项目列表（支持 Ctrl+A 多选） ----
        lf = ttk.LabelFrame(mf, text="隐患项目及日期分配（Ctrl+A 全选 | 双击编辑 | 右键快捷操作）", padding=10)
        lf.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        cols = ('#', '作业区', '部位', '问题描述', '登记日期', '分配日期', '当日值班人员')
        self.tree = ttk.Treeview(lf, columns=cols, show='headings', height=9,
                                 selectmode='extended')
        col_widths = [('#',40), ('作业区',80), ('部位',70), ('问题描述',280),
                      ('登记日期',80), ('分配日期',70), ('当日值班人员',150)]
        for c, w in col_widths:
            self.tree.heading(c, text=c); self.tree.column(c, width=w, minwidth=w)
        ts_v = ttk.Scrollbar(lf, orient=tk.VERTICAL, command=self.tree.yview)
        ts_h = ttk.Scrollbar(lf, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=ts_v.set, xscrollcommand=ts_h.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ts_v.pack(side=tk.RIGHT, fill=tk.Y)
        ts_h.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.bind('<Double-1>', self._dbl_click)
        self.tree.bind('<Button-3>', self._right_click)
        self.tree.bind('<Control-a>', self._select_all)
        self.tree.bind('<Control-A>', self._select_all)
        self.tree.bind('<<TreeviewSelect>>', self._update_sel_info)

        ttk.Label(lf, textvariable=self.sel_info_var,
                  foreground='#0A5', padding=(0, 4)).pack(fill=tk.X)

        bf = ttk.Frame(lf); bf.pack(fill=tk.X, pady=(5,0))
        ttk.Button(bf, text="为选中项设置日期...",
                   command=self._set_date_for_selection).pack(side=tk.LEFT, padx=(0,8))
        ttk.Button(bf, text="自动分配日期（选中项优先）",
                   command=self._auto).pack(side=tk.LEFT, padx=(0,8))
        ttk.Button(bf, text="从登记表提取日期",
                   command=self._auto_from_sweep).pack(side=tk.LEFT, padx=(0,8))
        ttk.Button(bf, text="清空选中项日期",
                   command=self._clear_selected).pack(side=tk.LEFT, padx=(0,8))
        ttk.Button(bf, text="全部清空日期",
                   command=self._clear_all).pack(side=tk.LEFT, padx=(0,8))
        ttk.Button(bf, text="全选(Ctrl+A)",
                   command=self._select_all).pack(side=tk.LEFT, padx=(0,8))
        ttk.Button(bf, text="刷新列表", command=self._refresh).pack(side=tk.LEFT)

        # ---- 生成按钮 + 进度条 ----
        gf = ttk.Frame(mf); gf.pack(fill=tk.X, pady=(0, 5))
        self.prog = ttk.Progressbar(gf, mode='determinate', length=300)
        self.prog.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(10,0))
        ttk.Button(gf, text="生成报告文件", command=self._gen).pack(side=tk.RIGHT, padx=5)
        ttk.Button(gf, text="生成选中报告文件",
                   command=self._gen_selected).pack(side=tk.RIGHT, padx=5)

        # ---- 日志 ----
        logf = ttk.LabelFrame(mf, text="日志", padding=5)
        logf.pack(fill=tk.BOTH, expand=True)
        self.log_text = scrolledtext.ScrolledText(logf, height=5, wrap=tk.WORD, font=('Consolas',9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self._log("欢迎使用！请拖入值班表(.xls/.xlsx)和隐患登记表(.xlsx)。")
        self._log("提示：列表中按 Ctrl+A 可全选项目，再统一设置/清空日期。")
        self._log("提示：选中项目后点击“生成选中报告文件”，可只生成选中项。")


    # ==================== 日志与提示 ====================
    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{ts}] {msg}\n")
        self.log_text.see(tk.END)
        self.root.update()

    # ==================== 文件选择/加载 ====================
    def _sel_duty(self):
        p = filedialog.askopenfilename(title="选择值班表",
                                       filetypes=[("Excel", "*.xls *.xlsx"), ("All", "*.*")])
        if p:
            self._load_duty(p)

    def _sel_sweep(self):
        p = filedialog.askopenfilename(title="选择隐患登记表",
                                       filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        if p:
            self._load_sweep(p)

    def _drop_duty(self, ev):
        self._load_duty(ev.data.strip('{}').strip())

    def _drop_sweep(self, ev):
        self._load_sweep(ev.data.strip('{}').strip())

    def _load_duty(self, p):
        try:
            self._log(f"加载值班表: {os.path.basename(p)}")
            self.roster_data = read_duty_roster(p)
            self.duty_label.config(text=os.path.basename(p), fg='#060', bg='#D4EDDA')
            self.roster_days = sorted(
                (k for k in self.roster_data if k.count('.') == 2),
                key=lambda x: tuple(int(part) for part in x.split('.'))
            )
            n_people = len(set(v.get('leader', '') for v in self.roster_data.values() if v.get('leader')))
            d_min, d_max = (self.roster_days[0], self.roster_days[-1]) if self.roster_days else ('?', '?')
            self.roster_info_var.set(
                f"值班表: {len(self.roster_days)}天 ({d_min} ~ {d_max})  |  "
                f"{n_people}位值班领导  |  双击隐患项编辑日期后可见值班人员预览"
            )
            self._log(f"值班表加载成功! {len(self.roster_days)}天, {n_people}位领导")
            self._refresh()
        except Exception as e:
            self._log(f"失败: {e}")
            messagebox.showerror("错误", str(e))

    def _load_sweep(self, p):
        try:
            self._log(f"加载隐患登记表: {os.path.basename(p)}")
            self.sweep_items, self.format_name = read_minesweeper(p)
            self.sweep_label.config(text=os.path.basename(p), fg='#930', bg='#F8D7DA')
            self._log(f"隐患登记表加载成功! [{self.format_name}] 共 {len(self.sweep_items)} 项")
            self.item_dates = {}
            self.item_areas = {}
            area_counts = {}
            for i, item in enumerate(self.sweep_items):
                area_label = get_sweep_area_label(item)
                self.item_areas[i] = area_label
                area_counts[area_label] = area_counts.get(area_label, 0) + 1
            area_summary = ' | '.join(f'{k}:{v}项' for k, v in area_counts.items())
            self._log(f"  作业区分布: {area_summary}")
            if any(item.get('sweep_date') for item in self.sweep_items):
                self._auto_from_sweep()
            self._refresh()
        except Exception as e:
            self._log(f"失败: {e}")
            messagebox.showerror("错误", str(e))

    # ==================== 列表刷新与多选 ====================
    def _update_sel_info(self, event=None):
        selected = len(self.tree.selection())
        total = len(self.sweep_items)
        if total == 0:
            text = '尚未加载隐患登记表'
        elif selected == 0:
            text = f'共 {total} 项  |  未选择（Ctrl+A 全选，Ctrl/Shift+单击多选）'
        else:
            text = f'共 {total} 项  |  已选 {selected} 项（Ctrl+A 全选，Ctrl/Shift+单击多选）'
        self.sel_info_var.set(text)

    def _selected_indices(self):
        return sorted(int(iid) for iid in self.tree.selection())

    def _select_all(self, event=None):
        iids = self.tree.get_children()
        if iids:
            self.tree.selection_set(iids)
            self.tree.focus(iids[0])
        self._update_sel_info()
        self._log(f"已全选 {len(iids)} 项")
        return 'break'

    def _refresh(self):
        keep = set(self.tree.selection())
        for it in self.tree.get_children():
            self.tree.delete(it)
        for i, item in enumerate(self.sweep_items):
            dv = self.item_dates.get(i, '')
            sd = item.get('sweep_date', '')
            area = self.item_areas.get(i, get_sweep_area_label(item))
            desc = item['description'][:55] + ('...' if len(item['description']) > 55 else '')
            loc = item.get('location', item.get('area', ''))[:15]
            preview = ''
            if dv and self.roster_data:
                leader, inspector = resolve_duty(self.roster_data,
                                                 self.month_var.get().strip(), dv,
                                                 get_sweep_area_key(item))
                preview = f"{leader or '?'} / {inspector or '?'}"
            self.tree.insert('', tk.END, iid=str(i),
                values=(item['seq'], area, loc, desc, sd, dv, preview))
        restore = [iid for iid in keep if self.tree.exists(iid)]
        if restore:
            self.tree.selection_set(restore)
        self._update_sel_info()


    # ==================== 双击 / 右键 / 日期编辑 ====================
    def _dbl_click(self, ev):
        row = self.tree.identify_row(ev.y)
        if row and row not in self.tree.selection():
            self.tree.selection_set(row)
        idxs = self._selected_indices()
        if not idxs:
            return
        if len(idxs) == 1:
            self._edit_date(idxs[0])
        else:
            self._batch_set_date(idxs)

    def _right_click(self, ev):
        row = self.tree.identify_row(ev.y)
        if row and row not in self.tree.selection():
            self.tree.selection_set(row)
        idxs = self._selected_indices()
        if not idxs:
            return
        n = len(idxs)
        menu = tk.Menu(self.root, tearoff=0)
        if n == 1:
            menu.add_command(label="编辑日期...", command=lambda: self._edit_date(idxs[0]))
            menu.add_command(label="清空此项日期", command=lambda: self._clear_many(idxs))
        else:
            menu.add_command(label=f"为选中的 {n} 项批量设置日期...",
                             command=lambda: self._batch_set_date(idxs))
            menu.add_command(label=f"清空选中 {n} 项日期",
                             command=lambda: self._clear_many(idxs))
        days = self._roster_days_for_month()
        if days:
            sub = tk.Menu(menu, tearoff=0)
            for day_part in days:
                sub.add_command(label=f"{day_part}日",
                                command=lambda d=day_part: self._set_dates(idxs, d))
            label = "从值班表快速选择日期" if n == 1 else f"为选中 {n} 项快速选择日期"
            menu.add_cascade(label=label, menu=sub)
        menu.post(ev.x_root, ev.y_root)

    def _roster_days_for_month(self):
        """返回值班表中当前月份的日期（日），按顺序排列。"""
        month_str = self.month_var.get().strip()
        days = []
        for d in self.roster_days:
            parts = d.split('.')
            if len(parts) >= 2:
                m_part = parts[-2] if len(parts) == 2 else parts[1]
                day_part = parts[-1]
                if m_part.lstrip('0') == month_str:
                    days.append(day_part)
        return days

    def _set_dates(self, idxs, day):
        day = str(int(day))
        for i in idxs:
            self.item_dates[i] = day
        self._refresh()
        self._log(f"已为选中的 {len(idxs)} 项设置日期 {self.month_var.get().strip()}.{day}")

    def _clear_many(self, idxs):
        for i in idxs:
            self.item_dates.pop(i, None)
        self._refresh()
        self._log(f"已清空选中的 {len(idxs)} 项日期")

    def _clear_selected(self):
        idxs = self._selected_indices()
        if not idxs:
            messagebox.showinfo("提示", "请先在列表中选中项目（Ctrl+A 可全选）")
            return
        self._clear_many(idxs)

    def _clear_all(self):
        self.item_dates = {}
        self._refresh()
        self._log("已清空所有日期")

    def _set_date_for_selection(self):
        idxs = self._selected_indices()
        if not idxs:
            messagebox.showinfo("提示", "请先在列表中选中项目（Ctrl+A 可全选）")
            return
        if len(idxs) == 1:
            self._edit_date(idxs[0])
        else:
            self._batch_set_date(idxs)

    def _batch_set_date(self, idxs):
        idxs = [i for i in idxs if 0 <= i < len(self.sweep_items)]
        if not idxs:
            return
        dlg = tk.Toplevel(self.root)
        dlg.title(f"批量设置日期 - 已选 {len(idxs)} 项")
        dlg.geometry("500x300")
        dlg.transient(self.root)
        dlg.grab_set()
        ttk.Label(dlg, text=f"已选中 {len(idxs)} 个项目", font=('', 10, 'bold')).pack(pady=(12, 4))
        names = '、'.join(f"第{self.sweep_items[i]['seq']}项" for i in idxs[:8])
        if len(idxs) > 8:
            names += f" 等 {len(idxs)} 项"
        ttk.Label(dlg, text=names, wraplength=460, foreground='#666').pack(pady=2, padx=15)
        df2 = ttk.Frame(dlg); df2.pack(pady=12)
        ttk.Label(df2, text=f"{self.month_var.get().strip()} 月").pack(side=tk.LEFT)
        dv = tk.StringVar(value='')
        ttk.Entry(df2, textvariable=dv, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(df2, text="日").pack(side=tk.LEFT)
        days = self._roster_days_for_month()
        if days:
            ttk.Label(dlg, text="快捷选日（值班表中该月日期）：",
                      foreground='#666').pack(pady=(2, 0))
            qf = ttk.Frame(dlg); qf.pack(pady=6)
            for day_part in days[:31]:
                ttk.Button(qf, text=day_part, width=3,
                           command=lambda d=day_part: dv.set(d)).pack(side=tk.LEFT, padx=2)

        def save():
            try:
                v = dv.get().strip()
                if not v or not 1 <= int(v) <= 31:
                    raise ValueError
                v = str(int(v))
            except ValueError:
                messagebox.showwarning("警告", "请输入 1-31 之间的有效日期", parent=dlg)
                return
            for i in idxs:
                self.item_dates[i] = v
            self._refresh()
            self._log(f"已为选中的 {len(idxs)} 项设置日期 {self.month_var.get().strip()}.{v}")
            dlg.destroy()

        ttk.Button(dlg, text="确定", command=save).pack(pady=10)


    def _edit_date(self, idx):
        if idx >= len(self.sweep_items):
            return
        item = self.sweep_items[idx]
        area = self.item_areas.get(idx, get_sweep_area_label(item))
        dlg = tk.Toplevel(self.root)
        dlg.title(f"设置日期 - 第{item['seq']}项 [{area}]")
        dlg.geometry("500x260")
        dlg.transient(self.root)
        dlg.grab_set()
        ttk.Label(dlg, text=f"作业区: {area}", font=('', 10, 'bold')).pack(pady=(10,2))
        ttk.Label(dlg, text=f"问题: {item['description'][:80]}...",
                  wraplength=460).pack(pady=5, padx=15)
        df2 = ttk.Frame(dlg); df2.pack(pady=10)
        ttk.Label(df2, text=f"{self.month_var.get().strip()} 月  ").pack(side=tk.LEFT)
        dv = tk.StringVar(value=str(self.item_dates.get(idx, '')))
        ttk.Entry(df2, textvariable=dv, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(df2, text="日").pack(side=tk.LEFT)
        days = self._roster_days_for_month()
        if days:
            ttk.Label(dlg, text="快捷选日（值班表中该月日期）：",
                      foreground='#666').pack(pady=(5,0))
            qf = ttk.Frame(dlg); qf.pack(pady=5)
            for day_part in days[:31]:
                ttk.Button(qf, text=day_part, width=3,
                           command=lambda d=day_part: dv.set(d)).pack(side=tk.LEFT, padx=2)

        def save():
            try:
                v = dv.get().strip()
                if v and not 1 <= int(v) <= 31:
                    raise ValueError
                if v:
                    v = str(int(v))
                self.item_dates[idx] = v
                self._refresh()
                self._log(f"第{item['seq']}项日期已设为 {self.month_var.get().strip()}.{v or '未设置'}")
                dlg.destroy()
            except ValueError:
                messagebox.showwarning("警告", "请输入 1-31 之间的有效日期", parent=dlg)

        ttk.Button(dlg, text="确定", command=save).pack(pady=10)


    # ==================== 自动分配 ====================
    def _target_indices(self):
        """有选中项时只处理选中项，否则处理全部。"""
        sel = self._selected_indices()
        return sel if sel else list(range(len(self.sweep_items)))

    def _auto(self):
        try:
            s = int(self.ds.get()); e = int(self.de.get())
            if e < s:
                raise ValueError
        except ValueError:
            messagebox.showwarning("警告", "请输入有效日期范围")
            return
        n = len(self.sweep_items)
        targets = self._target_indices()
        if n == 0 or not targets:
            return
        nd = e - s + 1
        for pos, i in enumerate(targets):
            self.item_dates[i] = str(s + (pos % nd))
        scope = f"选中的 {len(targets)} 项" if len(targets) < n else f"全部 {n} 项"
        self._log(f"已按顺序为{scope}分配 {s}~{e} 日")
        self._refresh()

    def _auto_from_sweep(self):
        month_str = self.month_var.get().strip()
        targets = self._target_indices()
        count = 0
        for i in targets:
            sd = self.sweep_items[i].get('sweep_date', '')
            if not sd:
                continue
            m = re.search(r'(\d{4})?[.\-\\/]?(\d{1,2})[.\-\\/](\d{1,2})', sd)
            if not m:
                continue
            month_part, day_part = m.group(2), m.group(3)
            if month_part == month_str or month_part.lstrip('0') == month_str:
                self.item_dates[i] = str(int(day_part))
                count += 1
        if count > 0:
            self._log(f"从登记表自动提取 {count} 个日期")
            self._refresh()
        else:
            self._log("未能从登记表提取到匹配的日期")

    def _sel_outdir(self):
        p = filedialog.askdirectory(title="选择输出目录")
        if p:
            self.out_dir.set(p)


    # ==================== 生成报告 ====================
    def _gen(self):
        self._generate(only_selected=False)

    def _gen_selected(self):
        if not self._selected_indices():
            messagebox.showinfo("提示", "请先在列表中选中项目（Ctrl+A 可全选），再点击“生成选中报告文件”。")
            return
        self._generate(only_selected=True)

    def _generate(self, only_selected=False):
        if not self.roster_data:
            messagebox.showwarning("警告", "请先加载值班表！")
            return
        if not self.sweep_items:
            messagebox.showwarning("警告", "请先加载隐患登记表！")
            return
        month = self.month_var.get().strip()
        out = self.out_dir.get()
        if only_selected:
            pairs = [(i, self.sweep_items[i]) for i in self._selected_indices()]
        else:
            pairs = list(enumerate(self.sweep_items))
            undated = [i for i, v in self.item_dates.items() if not v]
            if undated:
                if not messagebox.askyesno("确认", f"有 {len(undated)} 个项目未分配日期，将被跳过。\n是否继续？"):
                    return
        items = [(i, item, d) for i, item in pairs
                 if (d := self.item_dates.get(i, '').strip())]
        skipped = len(pairs) - len(items)
        if not items:
            messagebox.showwarning("警告", "选中的项目都未分配日期！" if only_selected
                                   else "没有可生成的项目！")
            return
        if only_selected and skipped:
            self._log(f"选中项中有 {skipped} 项未分配日期，已跳过")
        if not only_selected:
            preview_lines = []
            for idx, item, day in items[:5]:
                ds = f"{month}.{day}"
                leader, inspector = resolve_duty(self.roster_data, month, day,
                                                 get_sweep_area_key(item))
                area = get_sweep_area_label(item)
                preview_lines.append(f"  {ds}日 [{area}] 值班领导:{leader or '?'}  检查人:{inspector or '?'}")
            preview_text = "\n".join(preview_lines)
            if len(items) > 5:
                preview_text += f"\n  ... 共 {len(items)} 项"
            if not messagebox.askyesno("生成预览确认",
                f"即将生成 {len(items)} 个文件：\n\n{preview_text}\n\n输出目录：{out}\n\n确认生成？"):
                return
        os.makedirs(out, exist_ok=True)
        total = len(items)
        self.prog['maximum'] = total
        self.prog['value'] = 0
        ok = err = 0
        scope = "选中的" if only_selected else "全部"
        self._log(f"开始生成{scope} {total} 个文件 -> {out}")
        for idx, item, day in items:
            try:
                ds = f"{month}.{day}"
                fn = f"中夜班（{ds}）安全消防检查情况({BI_NAME}).docx"
                fp = os.path.join(out, fn)
                c = 1
                base, ext = os.path.splitext(fn)
                while os.path.exists(fp):
                    fp = os.path.join(out, f"{base}_{c}{ext}")
                    c += 1
                generate_docx(self.roster_data, item, fp, month, ds)
                ok += 1
                self._log(f"  OK {os.path.basename(fp)}")
            except Exception as e:
                err += 1
                self._log(f"  FAIL 第{item['seq']}项: {e}")
            self.prog['value'] = ok + err
            self.root.update()
        self._log(f"完成! 成功:{ok}, 失败:{err}")
        messagebox.showinfo("完成", f"生成完成！\n\n成功: {ok}\n失败: {err}\n\n输出: {out}")


def main():
    root = TkinterDnD.Tk()
    ReportGeneratorApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()

