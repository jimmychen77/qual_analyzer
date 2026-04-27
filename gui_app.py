#!/usr/bin/env python3
"""
QualCoder Pro - 通用质性分析工具 GUI
=====================================
参考 NVivo / ATLAS.ti / QualCoder / Taguette 设计的通用文本分析桌面应用。
完全通用，不包含任何行业特定内容。

启动:
    cd <项目目录>
    python3 hotel_analyzer/gui_app.py
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import json
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import traceback
from pathlib import Path
from typing import Dict

# ── 样式常量 ──────────────────────────────────────────────
BG         = '#f0f2f5'
ACCENT     = '#1565c0'
ACCENT2    = '#e3f2fd'
SUCCESS    = '#2e7d32'
WARN       = '#ef6c00'
ERROR      = '#c62828'
TEXT_BG    = '#ffffff'
TOOLBAR_BG = '#e8e8e8'
FONT       = ('Segoe UI', 10)
FONT_B     = ('Segoe UI', 10, 'bold')


# ══════════════════════════════════════════════════════════════════
# 工具函数
# ══════════════════════════════════════════════════════════════════

def auto_detect_columns(df):
    """
    自动推断文本列和名称列。
    返回 (text_col, name_col)。
    策略：跳过数值/日期/id类列，选平均字符串长度最长的列作为文本列。
    """
    skip_patterns = [
        'id', 'ID', 'date', 'time', '时间', 'url', 'link', '来源', 'source',
    ]
    text_col = None
    best_avg = 0
    for col in df.columns:
        col_l = col.lower()
        if any(p.lower() in col_l for p in skip_patterns):
            continue
        non_null = df[col].dropna()
        if len(non_null) == 0:
            continue
        # 跳过纯数值列（如ID、编号）
        sample = non_null.iloc[:5].astype(str)
        if all(s.replace('.', '').replace('-', '').isdigit() for s in sample):
            continue
        avg = non_null.apply(lambda x: len(str(x))).mean()
        if avg > best_avg:
            best_avg = avg
            text_col = col

    name_col = None
    for kw in ['标题', '名称', 'name', 'title', '文档名', '文件名', 'subject', '主题']:
        for c in df.columns:
            if kw in c:
                name_col = c
                break
        if name_col:
            break
    return text_col, name_col


# ══════════════════════════════════════════════════════════════════
# 主窗口类
# ══════════════════════════════════════════════════════════════════

class QDAGUI:
    """QualCoder Pro 主窗口。"""

    NAV_TABS = [
        ('📄', '文档',  'docs'),
        ('🏷', '编码',  'coding'),
        ('🔍', '检索',  'search'),
        ('📊', '矩阵',  'matrix'),
        ('📈', '图表',  'charts'),
        ('📝', '备忘',  'memos'),
        ('💾', '导出',  'export'),
        ('✅', '研究质量', 'quality'),
        ('🤝', '一致性',  'reliability'),
        ('📊', '词频',   'wordfreq'),
    ]

    def __init__(self, root: tk.Tk):
        self.root = root
        root.title('QualCoder Pro - 质性分析工具')
        root.geometry('1280x820')
        root.minsize(960, 640)

        self._setup_styles()
        self._build_ui()

        # 状态
        self._app         = None
        self._current_tab = None
        self._tab_pages  = {}   # {tab_id: page_frame}

        # 段落切分结果（文档Tab用）
        self._current_doc    = None
        self._current_paras  = []   # [(start, end, text)]

        # 搜索结果缓存
        self._search_hits = []

        # 图表预览窗口引用
        self._chart_windows = []

        # 高亮图例标志
        self._legend_inserted = False

        # ── 最近项目 ──
        self._recent_projects = self._load_recent_projects()

        # ── 撤销/重做栈 ──
        self._undo_stack: list = []
        self._redo_stack: list = []

        # ── 键盘快捷键 ──
        root.bind('<Control-o>', lambda e: self._open_project())
        root.bind('<Control-s>', lambda e: self._save_project())
        root.bind('<Control-f>', lambda e: self._on_ctrl_f())
        root.bind('<Control-n>', lambda e: self._new_project())
        root.bind('<Control-Shift-E>', lambda e: self._export_word())
        root.bind('<Delete>', lambda e: self._on_delete_shortcut())
        root.bind('<Control-z>', lambda e: self._undo_coding())
        root.bind('<Control-y>', lambda e: self._redo_coding())
        root.bind('<Control-Shift-Z>', lambda e: self._redo_coding())

        # ── 欢迎对话框 ──
        self._show_welcome_dialog()

    # ── 样式 ──────────────────────────────────────────────────

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('.', font=FONT, background=BG)
        style.configure('TFrame',  background=BG)
        style.configure('TLabelframe', background=BG, font=FONT_B)
        style.configure('TLabelframe.Label', background=BG, font=FONT_B)
        style.configure('TButton', font=FONT)
        style.configure('TEntry',  font=FONT)
        style.configure('TLabel', font=FONT, background=BG)
        style.configure('Treeview', font=FONT, rowheight=26)
        style.configure('Treeview.Heading', font=FONT_B)
        style.configure('Nav.TButton', font=FONT_B, padding=(16, 10))
        style.configure('NavActive.TButton', font=FONT_B, padding=(16, 10),
                        background=ACCENT, foreground='white')
        style.map('NavActive.TButton',
                  background=[('active', ACCENT), ('!active', ACCENT)],
                  foreground=[('active', 'white'), ('!active', 'white')])

    # ── 整体布局 ──────────────────────────────────────────────

    def _build_ui(self):
        # 顶部横幅
        banner = tk.Frame(self.root, bg=ACCENT, height=48)
        banner.pack(fill='x')
        banner.pack_propagate(False)
        tk.Label(banner, text='QualCoder Pro',
                font=('Segoe UI', 16, 'bold'),
                fg='white', bg=ACCENT).pack(side='left', padx=16)
        self._status_label = tk.Label(banner, text='就绪',
                                      fg='#bbdefb', bg=ACCENT,
                                      font=('Segoe UI', 9))
        self._status_label.pack(side='right', padx=12)

        # 主工作区
        work = tk.Frame(self.root, bg=BG)
        work.pack(fill='both', expand=True, padx=4, pady=4)

        # 导航
        nav = tk.Frame(work, bg='#dde', width=130)
        nav.pack(side='left', fill='y', padx=(0, 4))
        nav.pack_propagate(False)

        self._nav_buttons = {}
        for emoji, label, tab_id in self.NAV_TABS:
            btn = tk.Button(nav, text=f'{emoji}  {label}', font=FONT_B,
                          relief='flat', anchor='w', padx=8, pady=8,
                          command=lambda t=tab_id: self._show_tab(t),
                          bg='#dde', activebackground=ACCENT,
                          activeforeground='white', fg='#222', bd=0,
                          cursor='hand2')
            btn.pack(fill='x', padx=2, pady=1)
            self._nav_buttons[tab_id] = btn

        # 内容区
        self._content = tk.Frame(work, bg=BG)
        self._content.pack(side='left', fill='both', expand=True)

        # 底部状态栏（加强版）
        status_frame = tk.Frame(self.root, bg=TOOLBAR_BG, bd=1, relief='sunken')
        status_frame.pack(fill='x')
        self._status_bar = tk.Label(status_frame, text='就绪',
                                    font=('Segoe UI', 8), bg=TOOLBAR_BG, anchor='w')
        self._status_bar.pack(side='left', padx=4, fill='x', expand=True)
        self._right_stats = tk.Label(status_frame, text='📄 0 docs | 🏷️ 0 codes | 🔗 0 instances',
                                     font=('Segoe UI', 8), bg=TOOLBAR_BG, fg='#555', anchor='e')
        self._right_stats.pack(side='right', padx=8)

    # ── Tab 切换（缓存机制）────────────────────────────────

    def _show_tab(self, tab_id: str):
        if self._current_tab == tab_id:
            # 仍然刷新数据
            self._refresh_tab(tab_id)
            return

        # 隐藏当前
        if self._current_tab and self._current_tab in self._tab_pages:
            self._tab_pages[self._current_tab].pack_forget()

        self._current_tab = tab_id

        # 高亮导航
        for tid, btn in self._nav_buttons.items():
            if tid == tab_id:
                btn.configure(bg=ACCENT, fg='white')
            else:
                btn.configure(bg='#dde', fg='#222')

        # 按需构建
        if tab_id not in self._tab_pages:
            self._build_tab(tab_id)

        self._tab_pages[tab_id].pack(fill='both', expand=True)

        # 刷新数据
        self._refresh_tab(tab_id)

    def _refresh_tab(self, tab_id: str):
        """Tab 切换时刷新该 Tab 的数据"""
        if tab_id == 'docs':
            self._refresh_docs_list()
        elif tab_id == 'coding':
            self._refresh_code_tree()
            self._refresh_coded_segments()
        elif tab_id == 'search':
            pass  # 不自动刷新搜索
        elif tab_id == 'matrix':
            self._update_matrix_dims()
        elif tab_id == 'charts':
            self._refresh_charts_list()
        elif tab_id == 'memos':
            self._refresh_memos_list()
        elif tab_id == 'export':
            self._refresh_export_preview()
        elif tab_id == 'quality':
            self._refresh_quality_tab()
        elif tab_id == 'reliability':
            self._refresh_reliability_tab()
        elif tab_id == 'wordfreq':
            self._refresh_wordfreq_tab()

    # ── Tab 构建 ─────────────────────────────────────────────

    def _build_tab(self, tab_id: str):
        method = getattr(self, f'_build_{tab_id}_tab', None)
        if method:
            method()

    # ══════════════════════════════════════════════════════════
    # 文档 Tab
    # ══════════════════════════════════════════════════════════

    def _build_docs_tab(self):
        f = tk.Frame(self._content, bg=BG)
        self._tab_pages['docs'] = f

        # 工具栏
        tb = tk.Frame(f, bg=TOOLBAR_BG, height=44)
        tb.pack(fill='x')
        tk.Label(tb, text='文档管理', font=FONT_B, bg=TOOLBAR_BG).pack(side='left', padx=8)
        tk.Button(tb, text='📂  加载文件', font=FONT, command=self._on_load_files,
                 bg=ACCENT, fg='white', relief='flat', padx=10, cursor='hand2'
                 ).pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='🔄  刷新', font=FONT, command=self._refresh_docs_list,
                 relief='flat', padx=8, cursor='hand2'
                 ).pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='🏷 属性管理', font=FONT, command=self._attribute_manager,
                 relief='flat', padx=8, cursor='hand2'
                 ).pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='📄 导入 PDF', font=FONT,
                 command=self._on_import_pdf,
                 relief='flat', padx=8, cursor='hand2'
                 ).pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='📝 粘贴文本', font=FONT,
                 command=self._on_paste_text,
                 relief='flat', padx=8, cursor='hand2'
                 ).pack(side='left', padx=4, pady=4)
        self._doc_count_lbl = tk.Label(tb, text='未加载文档',
                                       font=FONT, bg=TOOLBAR_BG, fg='#555')
        self._doc_count_lbl.pack(side='left', padx=12)

        # 主区：左右分栏
        paned = ttk.PanedWindow(f, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=4, pady=4)

        # 左：文档列表
        lf = ttk.Labelframe(paned, text='文档列表', padding=4)
        paned.add(lf, weight=1)

        cols = ('名称', '字数')
        self._doc_tree = ttk.Treeview(lf, columns=cols, show='tree headings',
                                       height=28, selectmode='browse')
        for col, w in zip(cols, (300, 70)):
            self._doc_tree.heading(col, text=col)
            self._doc_tree.column(col, width=w, stretch=False)
        sy = ttk.Scrollbar(lf, orient='vertical', command=self._doc_tree.yview)
        self._doc_tree.configure(yscrollcommand=sy.set)
        self._doc_tree.pack(side='left', fill='both', expand=True)
        sy.pack(side='right', fill='y')
        self._doc_tree.bind('<<TreeviewSelect>>', lambda e: self._on_doc_select())

        # 右：预览 + 段落
        rf = ttk.Labelframe(paned, text='内容预览', padding=4)
        paned.add(rf, weight=3)

        # 预览文本（可拖蓝选中）
        self._preview = tk.Text(rf, wrap='word', font=('Courier New', 9),
                               bg=TEXT_BG, relief='flat', state='normal')
        sp = ttk.Scrollbar(rf, command=self._preview.yview)
        self._preview.configure(yscrollcommand=sp.set)
        sp.pack(side='right', fill='y')
        self._preview.pack(fill='both', expand=True, pady=(0, 4))
        self._preview.bind('<ButtonRelease-1>', lambda e: self._on_preview_sel_change())
        self._preview.tag_config('selected', background='#bbdefb')

        # 编码操作条（核心交互区）
        coding_bar = tk.Frame(rf, bg='#e3f2fd', height=64)
        coding_bar.pack(fill='x', pady=(0, 4))
        coding_bar.pack_propagate(False)

        # 提示文字
        self._coding_hint = tk.Label(coding_bar,
            text='💡 在上方文本中拖蓝选中任意文字，点击「编码选中文字」为选中内容分配编码',
            font=FONT, bg='#e3f2fd', fg='#1565c0', anchor='w')
        self._coding_hint.pack(side='top', fill='x', padx=8, pady=(4, 0))

        btn_row = tk.Frame(coding_bar, bg='#e3f2fd')
        btn_row.pack(fill='x', padx=6, pady=(2, 4))
        self._btn_code_selected = tk.Button(btn_row, text='🏷 编码选中文字',
            font=FONT_B, bg=ACCENT, fg='white', relief='flat',
            padx=12, cursor='hand2', state='disabled',
            command=self._on_code_selected)
        self._btn_code_selected.pack(side='left', padx=2)
        tk.Button(btn_row, text='📑 切分段落', font=FONT,
                 command=self._split_paragraphs,
                 relief='flat', padx=8, cursor='hand2').pack(side='left', padx=2)
        tk.Button(btn_row, text='🔍 检索选中', font=FONT,
                 command=self._search_selected,
                 relief='flat', padx=8, cursor='hand2').pack(side='left', padx=2)
        tk.Button(btn_row, text='↩ 取消选中', font=FONT,
                 command=lambda: self._preview.tag_remove('sel', '1.0', 'end'),
                 relief='flat', padx=8, cursor='hand2').pack(side='right', padx=2)

        # 段落列表
        plf = ttk.Labelframe(rf, text='段落列表（双击段落直接编码）', padding=4)
        plf.pack(fill='both', expand=True)
        self._para_tree = ttk.Treeview(plf, columns=('内容', '编码'), show='headings',
                                        height=14)
        self._para_tree.heading('内容', text='段落内容（双击直接打开编码对话框）')
        self._para_tree.heading('编码',  text='已分配编码')
        self._para_tree.column('内容',  width=520)
        self._para_tree.column('编码',   width=160)
        ps = ttk.Scrollbar(plf, orient='vertical', command=self._para_tree.yview)
        self._para_tree.configure(yscrollcommand=ps.set)
        self._para_tree.pack(side='left', fill='both', expand=True)
        ps.pack(side='right', fill='y')
        self._para_tree.bind('<Double-Button-1>', lambda e: self._on_para_dbl_click())

    def _on_load_files(self):
        paths = filedialog.askopenfiles(
            title='加载数据文件',
            filetypes=[
                ('数据文件', '*.xlsx *.xls *.csv *.json *.txt *.pdf'),
                ('Excel', '*.xlsx *.xls'),
                ('CSV', '*.csv'),
                ('PDF', '*.pdf'),
                ('全部', '*.*'),
            ])
        if not paths:
            return
        file_paths = [p.name for p in paths]
        self._set_status(f'正在加载 {len(file_paths)} 个文件...')
        self._doc_count_lbl.config(text='加载中...')
        self.root.update()

        def do_load():
            from hotel_analyzer import load_documents, QDAApplication

            import pandas as pd
            first_df = pd.read_excel(file_paths[0])
            text_col, name_col = auto_detect_columns(first_df)

            if not text_col:
                raise ValueError(
                    '无法自动识别文本列。\n请确保文件包含文字内容列（如"评论内容"、"正文"等）。')

            docs = load_documents(*file_paths, text_col=text_col, name_col=name_col)
            app = QDAApplication()
            app.documents = docs
            self._app = app

        try:
            self._with_progress(f'正在加载 {len(file_paths)} 个文件...', do_load)
            self._refresh_docs_list()
            n = len(self._app.documents)
            self._doc_count_lbl.config(text=f'📄 {n} 条文档')
            self._set_status(f'加载完成：{n} 条文档')
            self._update_stats()
            messagebox.showinfo('完成', f'成功加载 {n} 条文档')
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror('加载错误', str(e))
            self._doc_count_lbl.config(text='加载失败')
            self._set_status('加载失败')

    def _on_import_pdf(self):
        path = filedialog.askopenfilename(
            title='选择 PDF 文件',
            filetypes=[('PDF', '*.pdf'), ('全部', '*.*')])
        if not path:
            return
        if not self._app:
            from hotel_analyzer import QDAApplication
            self._app = QDAApplication()
        try:
            self._set_status('正在提取 PDF 文本...')
            self.root.update()
            from hotel_analyzer.data_processor import extract_text_from_pdf
            pages = extract_text_from_pdf(path)
            if not pages:
                messagebox.showwarning('提示', 'PDF 中未提取到文字（可能是扫描件）')
                return
            for i, text in enumerate(pages):
                self._app.add_document(text=text, name=f'PDF第{i+1}页',
                                      source=str(path), page=str(i+1))
            self._refresh_docs_list()
            n = len(pages)
            self._doc_count_lbl.config(text=f'📄 {len(self._app.documents)} 条文档')
            self._set_status(f'从 PDF 导入了 {n} 页')
            messagebox.showinfo('完成', f'已从 PDF 提取 {n} 页')
        except ImportError as e:
            messagebox.showerror('缺少依赖', str(e))
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror('导入错误', str(e))

    def _on_paste_text(self):
        """粘贴文本创建新文档"""
        dlg = tk.Toplevel(self.root)
        dlg.title('粘贴文本')
        dlg.geometry('600x420')
        dlg.transient(self.root)
        dlg.grab_set()
        tk.Label(dlg, text='输入文档名称：', font=FONT_B).pack(anchor='w', padx=10, pady=(10, 4))
        name_entry = tk.Entry(dlg, font=FONT, width=40)
        name_entry.insert(0, '粘贴文档')
        name_entry.pack(fill='x', padx=10, pady=(0, 6))
        tk.Label(dlg, text='粘贴文本内容：', font=FONT_B).pack(anchor='w', padx=10)
        txt = tk.Text(dlg, font=('Courier New', 9), bg=TEXT_BG)
        txt.pack(fill='both', expand=True, padx=10, pady=4)

        def do_add():
            name = name_entry.get().strip() or f'粘贴文档_{len(self._app.documents) if self._app else 1}'
            text = txt.get('1.0', 'end').strip()
            if not text:
                messagebox.showwarning('提示', '文本内容不能为空')
                return
            if not self._app:
                from hotel_analyzer import QDAApplication
                self._app = QDAApplication()
            self._app.add_document(text=text, name=name)
            self._refresh_docs_list()
            self._doc_count_lbl.config(
                text=f'📄 {len(self._app.documents)} 条文档')
            self._set_status(f'已添加文档：{name}')
            dlg.destroy()

        btn_row = tk.Frame(dlg, bg=BG)
        btn_row.pack(fill='x', padx=10, pady=6)
        tk.Button(btn_row, text='✅ 添加', font=FONT_B, bg=ACCENT, fg='white',
                 relief='flat', padx=14, command=do_add,
                 cursor='hand2').pack(side='left', padx=6)
        tk.Button(btn_row, text='取消', font=FONT, relief='flat',
                 padx=12, command=dlg.destroy).pack(side='left', padx=6)

    def _refresh_docs_list(self):
        if not hasattr(self, '_doc_tree') or not self._doc_tree.winfo_exists():
            return
        self._doc_tree.delete(*self._doc_tree.get_children())
        if not self._app or not self._app.documents:
            return
        for doc in self._app.documents:
            self._doc_tree.insert('', 'end', text=doc.id,
                                  values=(doc.name, len(doc.text)))

    def _on_doc_select(self):
        sel = self._doc_tree.selection()
        if not sel:
            return
        doc_id = self._doc_tree.item(sel[0], 'text')
        doc = self._app.get_document(doc_id) if self._app else None
        self._current_doc = doc
        self._legend_inserted = False  # 重置图例标志
        if doc:
            self._show_doc_preview(doc)
            self._split_paragraphs()

    def _show_doc_preview(self, doc):
        """显示文档内容到预览区（保持可选择状态）"""
        self._preview.config(state='normal')
        self._preview.delete('1.0', 'end')
        self._preview.insert('end', doc.text)
        self._preview.config(state='normal')  # 保持可选择，不设为 disabled
        self._update_preview_highlights(doc)
        # 禁止键盘输入（保持只读）
        self._preview.unbind('<Key>')
        self._preview.bind('<Key>', lambda e: 'break')

    def _update_preview_highlights(self, doc=None):
        """用编码颜色高亮预览文本中的编码片段"""
        if doc is None:
            doc = self._current_doc
        if not doc or not self._app:
            return
        # 清除旧标签（不改变 state）
        for tag in self._preview.tag_names():
            if tag != 'sel':  # 保留选中高亮
                self._preview.tag_delete(tag)
        # 为每个编码的实例添加高亮
        legend_items = []
        for code in self._app.code_system.all_codes.values():
            color = code.color or '#cccccc'
            for inst in code.instances:
                if inst.get('doc_id') == doc.id:
                    seg = inst.get('segment', '')
                    start = inst.get('start', -1)
                    if start >= 0:
                        end = inst.get('end', start + len(seg))
                        try:
                            tag_name = f'code_{code.id}'
                            self._preview.tag_add(tag_name, f'1.0+{start}c', f'1.0+{end}c')
                            self._preview.tag_config(tag_name, background=color + '44',
                                                     foreground='#000000')
                        except Exception:
                            pass
            if code.instances:
                legend_items.append((code.name, color))
        # 添加图例
        if legend_items and (not hasattr(self, '_legend_inserted') or not self._legend_inserted):
            legend_text = '\n\n' + '─' * 40 + '\n编码颜色图例：\n'
            for name, color in legend_items:
                legend_text += f'  ██ {name}\n'
            self._preview.insert('end', legend_text)
            self._legend_inserted = True

    def _split_paragraphs(self):
        if not self._current_doc:
            return
        self._para_tree.delete(*self._para_tree.get_children())
        from hotel_analyzer import ParagraphTagger
        tagger = ParagraphTagger()
        segs = tagger.tag(self._current_doc)
        self._current_paras = []
        for seg in segs:
            if isinstance(seg, dict):
                self._current_paras.append((seg['start'], seg['end'], seg['segment']))
            else:
                self._current_paras.append(seg)
        for i, item in enumerate(self._current_paras):
            if len(item) == 3:
                s, e, txt = item
            else:
                txt = item
                s = e = -1
            display = txt[:80] + ('...' if len(txt) > 80 else '')
            coded = self._get_para_codes(self._current_doc.id, txt)
            self._para_tree.insert('', 'end', iid=str(i),
                                  values=(display, ', '.join(coded) if coded else '—'))

    def _get_para_codes(self, doc_id, segment_text):
        """根据片段文本查找已分配给该段的编码名称列表"""
        if not self._app:
            return []
        found = []
        seg_start = segment_text[:20]
        for code in self._app.code_system.all_codes.values():
            for inst in code.instances:
                if (inst.get('doc_id') == doc_id and
                    inst.get('segment', '').startswith(seg_start)):
                    found.append(code.name)
                    break
        return found

    def _on_para_dbl_click(self):
        """段落列表双击 → 打开编码对话框"""
        if not self._current_doc:
            messagebox.showinfo('提示', '请先在左侧选择一个文档')
            return
        sel = self._para_tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if idx >= len(self._current_paras):
            return
        item = self._current_paras[idx]
        if isinstance(item, dict):
            txt = item.get('segment', '')
            s = item.get('start', -1)
            e = item.get('end', -1)
        elif len(item) == 3:
            s, e, txt = item
        else:
            txt = str(item)
            s, e = -1, -1
        self._show_assign_dialog(self._current_doc, txt, s, e)

    def _show_assign_dialog(self, doc, segment, start, end):
        """
        编码对话框——核心交互界面。

        布局（上下三区）：
        1. 片段预览（高亮显示待编码内容）
        2. 现有编码列表 + 新建编码（可同时使用）
        3. 操作按钮
        """
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        if not doc:
            messagebox.showinfo('提示', '请先选择文档')
            return

        dlg = tk.Toplevel(self.root)
        dlg.title('🏷 分配编码')
        dlg.geometry('600x580')
        dlg.transient(self.root)
        dlg.grab_set()

        # ── 顶部提示 ────────────────────────────────────
        hint = tk.Label(dlg,
            text='← 从列表选择已有编码  |  或在下方输入新编码名称 →',
            font=FONT, fg='#1565c0', bg='#e3f2fd', pady=4)
        hint.pack(fill='x')

        # ── 第一区：片段预览 ────────────────────────────────
        frag_frame = tk.LabelFrame(dlg, text='📝 待编码片段', font=FONT_B,
                                   bg=BG, padx=8, pady=6)
        frag_frame.pack(fill='x', padx=8, pady=(8, 4))
        frag_txt = tk.Text(frag_frame, font=('Courier New', 10),
                           height=3, bg='#fff8e1', relief='flat', wrap='word')
        frag_txt.insert('end', segment[:300] + ('...' if len(segment) > 300 else ''))
        frag_txt.config(state='disabled')
        frag_txt.pack(fill='x')

        # ── 第二区：左侧编码列表 + 右侧新建编码 ──────────
        mid = tk.Frame(dlg, bg=BG)
        mid.pack(fill='both', expand=True, padx=8, pady=4)

        # 左侧：已有编码列表
        left_f = tk.LabelFrame(mid, text='🏷 已有编码（单击选中）', font=FONT_B,
                              bg=BG, padx=4, pady=4)
        left_f.pack(side='left', fill='both', expand=True)

        all_codes = self._app.code_system.all_codes
        code_names = sorted(c.name for c in all_codes.values())
        code_lb_var = tk.StringVar()

        code_lb = tk.Listbox(left_f, listvariable=code_lb_var,
                             font=FONT, selectmode='single',
                             height=10, relief='groove', borderwidth=1)
        code_lb.pack(fill='both', expand=True, side='left')
        sc_lb = tk.Scrollbar(left_f, orient='vertical', command=code_lb.yview)
        code_lb.configure(yscrollcommand=sc_lb.set)
        sc_lb.pack(side='right', fill='y')

        if code_names:
            for cn in code_names:
                code_lb.insert('end', cn)
        else:
            code_lb.insert('end', '（暂无已有编码）')
            code_lb.config(state='disabled', fg='#aaa')

        # 右侧：新建编码
        right_f = tk.LabelFrame(mid, text='🆕 新建编码', font=FONT_B,
                                bg=BG, padx=8, pady=4)
        right_f.pack(side='left', fill='y', padx=(8, 0))

        tk.Label(right_f, text='编码名称：', font=FONT, bg=BG).pack(anchor='w')
        new_entry = tk.Entry(right_f, font=FONT, width=16)
        new_entry.pack(fill='x', pady=4)
        new_entry.focus()

        tk.Label(right_f, text='选择颜色：', font=FONT, bg=BG).pack(anchor='w', pady=(6, 2))
        colors = ['#e53935', '#d81b60', '#8e24aa', '#5e35b1',
                  '#1e88e5', '#039be5', '#00897b', '#43a047',
                  '#7cb342', '#fdd835', '#fb8c00', '#f4511e']
        color_var = tk.StringVar(value=colors[4])
        for row_items in [colors[:6], colors[6:]]:
            row = tk.Frame(right_f, bg=BG)
            row.pack(anchor='w', pady=1)
            for col in row_items:
                b = tk.Button(row, text=' ', bg=col, width=2, height=1,
                             relief='raised', cursor='hand2',
                             command=lambda _, c=col: color_var.set(c))
                b.pack(side='left', padx=1)

        # ── 第三区：操作按钮 ─────────────────────────────
        btn_fr = tk.Frame(dlg, bg=BG)
        btn_fr.pack(fill='x', padx=8, pady=(6, 8))

        def do_assign():
            """执行编码分配"""
            self._push_undo()
            new_name = new_entry.get().strip()
            sel_curs = code_lb.curselection()

            if new_name:
                # 新建编码 + 分配
                chosen = new_name
                self._app.create_code(chosen, color=color_var.get())
            elif sel_curs:
                # 选择已有编码
                chosen = code_lb.get(sel_curs[0])
            else:
                messagebox.showwarning('提示', '请从左侧列表选择编码，或在右侧输入新编码名称')
                return

            if not doc or not chosen:
                return
            self._app.assign_code(doc.id, segment, chosen)
            self._refresh_coded_segments()
            self._split_paragraphs()
            self._update_stats()
            n = len(chosen)
            self._set_status(f'✅ 已分配编码「{chosen}」到片段')
            dlg.destroy()

        # 主操作按钮
        assign_btn = tk.Button(btn_fr, text='✅ 确认分配', font=FONT_B,
                             bg=ACCENT, fg='white', relief='flat',
                             padx=18, cursor='hand2', command=do_assign)
        assign_btn.pack(side='left', padx=6)

        # 快速提示
        tip = tk.Label(btn_fr,
            text='← 直接输入新名称即可创建编码并分配！',
            font=('Segoe UI', 8), fg='#888', bg=BG)
        tip.pack(side='left', padx=12)

        tk.Button(btn_fr, text='取消', font=FONT, relief='flat',
                 padx=12, command=dlg.destroy).pack(side='right', padx=6)

        # 快捷键
        new_entry.bind('<Return>', lambda e: do_assign())
        code_lb.bind('<Double-Button-1>', lambda e: do_assign())

    def _on_preview_sel_change(self):
        """预览区文本选中状态变化时，启用/禁用编码按钮"""
        try:
            if not hasattr(self, '_btn_code_selected'):
                return
            sel = self._preview.tag_ranges('sel')
            if sel:
                selected_text = self._preview.get(*sel)
                if selected_text and selected_text.strip():
                    self._btn_code_selected.config(state='normal')
                    n = len(selected_text.strip())
                    display = selected_text.strip()[:20]
                    self._coding_hint.config(
                        fg='#1b5e20',
                        text=f'✅ 已选中 {n} 字「{display}」→ 点击下方「编码选中文字」按钮')
                    return
        except Exception:
            pass
        # 无选中文字
        if hasattr(self, '_btn_code_selected'):
            self._btn_code_selected.config(state='disabled')
        if hasattr(self, '_coding_hint'):
            self._coding_hint.config(
                fg='#1565c0',
                text='💡 在上方文本中拖蓝选中任意文字，再点击「编码选中文字」')

    def _on_code_selected(self):
        """对预览区选中的文字进行编码（核心操作）"""
        if not hasattr(self, '_preview'):
            return
        if not self._app or not self._current_doc:
            messagebox.showinfo('提示', '请先在左侧选择一个文档')
            return
        try:
            sel = self._preview.tag_ranges('sel')
            if not sel:
                messagebox.showinfo('提示', '请先在预览区拖蓝选中要编码的文字')
                return
            text = self._preview.get(*sel).strip()
            if not text:
                return
            doc_text = self._current_doc.text
            s = doc_text.find(text)
            e = s + len(text) if s >= 0 else -1
            self._show_assign_dialog(self._current_doc, text, s, e)
        except Exception as e:
            messagebox.showerror('错误', str(e))

    def _search_selected(self):
        try:
            sel = self._preview.tag_ranges('sel')
            if sel:
                text = self._preview.get(*sel)
                self._show_tab('search')
                if hasattr(self, '_search_entry'):
                    self._search_entry.delete(0, 'end')
                    self._search_entry.insert(0, text)
                    self._on_do_search()
        except Exception:
            pass

    # ══════════════════════════════════════════════════════════
    # 编码 Tab
    # ══════════════════════════════════════════════════════════

    def _build_coding_tab(self):
        f = tk.Frame(self._content, bg=BG)
        self._tab_pages['coding'] = f

        tb = tk.Frame(f, bg=TOOLBAR_BG, height=44)
        tb.pack(fill='x')

        # 工具栏按钮组
        btn_bar = tk.Frame(tb, bg=TOOLBAR_BG)
        btn_bar.pack(side='left')
        for text, cmd in [
            ('➕ 添加编码',    self._add_code_dialog),
            ('🗑 删除选中',    self._delete_selected_code),
            ('🏗 合并编码',    self._merge_codes_dialog),
            ('🔍 编码查询',    self._query_builder),
            ('⚡ 自动编码',    self._on_toggle_auto_code),
        ]:
            tk.Button(btn_bar, text=text, font=FONT, command=cmd,
                     bg=ACCENT if '添加' in text else TOOLBAR_BG,
                     fg='white' if '添加' in text else '#222',
                     relief='flat', padx=8, cursor='hand2').pack(side='left', padx=3, pady=4)

        # 扎根理论阶段选择器
        tk.Label(tb, text='编码阶段：', font=FONT, bg=TOOLBAR_BG,
                fg='#555').pack(side='left', padx=(12, 4))
        self._gt_stage_var = tk.StringVar(value='开放编码')
        gt_combo = ttk.Combobox(tb, textvariable=self._gt_stage_var,
            values=['开放编码', '轴心编码', '选择性编码'],
            state='readonly', font=FONT, width=10)
        gt_combo.pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='📍 设置阶段', font=FONT,
                 command=self._set_gt_stage_for_code,
                 relief='flat', padx=6, cursor='hand2').pack(side='left', pady=4)
        tk.Button(tb, text='🔄 刷新', font=FONT, command=lambda: (
            self._refresh_code_tree(), self._refresh_coded_segments()),
                 relief='flat', padx=8, cursor='hand2'
                 ).pack(side='right', padx=8, pady=4)

        # 主区（先创建，让 paned window 先被 pack）
        self._coding_paned = ttk.PanedWindow(f, orient='horizontal')
        self._coding_paned.pack(fill='both', expand=True, padx=4, pady=4)

        # 自动编码面板（放在最后 pack，这样 show/hide 不影响 paned 的位置）
        self._auto_panel = tk.Frame(f, bg='#fff8e1')
        self._auto_panel.pack(fill='x', padx=4, pady=(0, 4))
        self._auto_panel.pack_forget()
        self._auto_panel_visible = False  # 用布尔标志跟踪状态

        # 代码树
        lf = ttk.Labelframe(self._coding_paned, text='代码系统', padding=4)
        self._coding_paned.add(lf, weight=1)
        self._code_tree = ttk.Treeview(lf, show='tree', height=28)
        self._code_tree.pack(fill='both', expand=True, side='left')
        cs = ttk.Scrollbar(lf, orient='vertical', command=self._code_tree.yview)
        self._code_tree.configure(yscrollcommand=cs.set)
        cs.pack(side='right', fill='y')
        self._code_tree.bind('<<TreeviewSelect>>', lambda e: self._on_code_tree_select())
        self._code_tree.bind('<Double-Button-1>', lambda e: self._show_code_memo())
        # ── Drag & Drop 重排代码树 ──
        self._code_tree._drag_data = {'item': None, 'x': 0, 'y': 0}
        self._code_tree.bind('<ButtonPress-1>', self._on_code_tree_drag_start)
        self._code_tree.bind('<ButtonRelease-1>', self._on_code_tree_drag_end)

        # 片段列表
        rf = ttk.Labelframe(self._coding_paned, text='已编码片段', padding=4)
        self._coding_paned.add(rf, weight=3)
        cols = ('文档', '片段内容', '编码')
        self._seg_tree = ttk.Treeview(rf, columns=cols, show='headings', height=28)
        for col, w in zip(cols, (130, 520, 110)):
            self._seg_tree.heading(col, text=col)
            self._seg_tree.column(col, width=w)
        ss = ttk.Scrollbar(rf, orient='vertical', command=self._seg_tree.yview)
        self._seg_tree.configure(yscrollcommand=ss.set)
        self._seg_tree.pack(side='left', fill='both', expand=True)
        ss.pack(side='right', fill='y')
        self._seg_tree.bind('<Double-Button-1>', lambda e: self._on_seg_dbl_click())

    def _refresh_code_tree(self):
        if not hasattr(self, '_code_tree') or not self._code_tree.winfo_exists():
            return
        self._code_tree.delete(*self._code_tree.get_children())
        if not self._app:
            return
        for code in self._app.code_system.root_codes:
            self._insert_code_node('', code)

    # ── 代码树拖放 ──────────────────────────────────────────

    def _on_code_tree_drag_start(self, event):
        """记录拖拽起始位置和选中的项目"""
        tree = event.widget
        item = tree.identify_row(event.y)
        if item:
            tree._drag_data = {'item': item, 'x': event.x, 'y': event.y}
        else:
            tree._drag_data = {'item': None, 'x': event.x, 'y': event.y}

    def _on_code_tree_drag_end(self, event):
        """拖拽结束：尝试将源代码重设父级到目标代码"""
        tree = event.widget
        src_item = tree._drag_data.get('item')
        if not src_item:
            return
        # 找目标项目
        target_item = tree.identify_row(event.y)
        if not target_item or target_item == src_item:
            return
        # 获取编码名称
        src_name = tree.item(src_item, 'text')
        tgt_name = tree.item(target_item, 'text')
        if not self._app or not src_name or not tgt_name:
            return
        src_code = self._app.code_system._find_code_by_name(src_name)
        tgt_code = self._app.code_system._find_code_by_name(tgt_name)
        if not src_code or not tgt_code:
            return
        # 不能将自己作为自己的父级
        if src_code == tgt_code:
            return
        # 不能将父级移到子级下（循环引用检查）
        parent = tgt_code
        while parent:
            if parent == src_code:
                self._set_status('⚠️ 不能创建循环引用')
                return
            parent = parent.parent
        try:
            # 从原父级移除
            if src_code.parent:
                src_code.parent.children.remove(src_code)
            elif src_code in self._app.code_system.root_codes:
                self._app.code_system.root_codes.remove(src_code)
            # 添加到新父级
            src_code.parent = tgt_code
            tgt_code.children.append(src_code)
            self._refresh_code_tree()
            self._set_status(f'已将编码「{src_name}」移动到「{tgt_name}」下')
        except Exception as e:
            self._set_status(f'重排失败：{e}')

    def _set_gt_stage_for_code(self):
        """为选中的编码设置扎根理论阶段"""
        if not self._app:
            return
        sel = self._code_tree.selection()
        if not sel:
            messagebox.showinfo('提示', '请先在代码树中选中一个编码')
            return
        # 去掉末尾的阶段徽标，获取真实编码名
        raw_name = self._code_tree.item(sel[0], 'text')
        real_name = raw_name.split(' [')[0].strip()
        code = self._app.code_system._find_code_by_name(real_name)
        if not code:
            return
        code.grounded_stage = self._gt_stage_var.get()
        self._refresh_code_tree()
        self._set_status(f'「{code.name}」阶段 → {code.grounded_stage}')

    def _insert_code_node(self, parent, code):
        display_name = code.name
        if code.grounded_stage:
            badges = {'开放编码': '[○]', '轴心编码': '[◉]', '选择性编码': '[●]'}
            badge = badges.get(code.grounded_stage, '')
            display_name = f'{code.name} {badge}'
        tag = self._code_tree.insert(parent, 'end', text=display_name,
                                     values=(len(code.instances),),
                                     tags=('code',))
        for child in code.children:
            self._insert_code_node(tag, child)

    def _refresh_coded_segments(self, code_name=None):
        """刷新右侧片段列表（code_name=None 时显示全部）"""
        if not hasattr(self, '_seg_tree') or not self._seg_tree.winfo_exists():
            return
        self._seg_tree.delete(*self._seg_tree.get_children())
        if not self._app:
            return
        codes_to_show = (
            [self._app.code_system.all_codes[c].name
             for c in self._app.code_system.all_codes
             if self._app.code_system.all_codes[c].name == code_name]
            if code_name else
            list(self._app.code_system.all_codes.values())
        )
        for code in codes_to_show:
            if isinstance(code, str):
                code = self._app.code_system.all_codes.get(
                    next((k for k, v in self._app.code_system.all_codes.items()
                          if v.name == code), None))
                if not code:
                    continue
            for inst in code.instances:
                doc_name = '?'
                if self._app.documents:
                    doc = self._app.get_document(inst.get('doc_id', ''))
                    if doc:
                        doc_name = doc.name[:15]
                seg = str(inst.get('segment', ''))[:80]
                self._seg_tree.insert('', 'end',
                    values=(doc_name, seg, code.name))

    def _on_code_tree_select(self):
        sel = self._code_tree.selection()
        if not sel:
            return
        # 去掉末尾的阶段徽标，获取真实编码名
        raw_name = self._code_tree.item(sel[0], 'text')
        code_name = raw_name.split(' [')[0].strip()
        self._seg_tree.delete(*self._seg_tree.get_children())
        if not self._app:
            return
        for code in self._app.code_system.all_codes.values():
            if code.name == code_name:
                for inst in code.instances:
                    doc_name = '?'
                    if self._app.documents:
                        doc = self._app.get_document(inst.get('doc_id', ''))
                        if doc:
                            doc_name = doc.name[:15]
                    seg = str(inst.get('segment', ''))[:80]
                    self._seg_tree.insert('', 'end',
                        values=(doc_name, seg, code.name))

    def _on_seg_dbl_click(self):
        sel = self._seg_tree.selection()
        if not sel:
            return
        vals = self._seg_tree.item(sel[0], 'values')
        doc_name, seg_text, code_name = vals[0], vals[1], vals[2]
        self._show_segment_editor(doc_name, seg_text, code_name)

    def _show_segment_editor(self, doc_name, seg_text, code_name):
        """编码片段编辑器：可查看/编辑片段详情"""
        dlg = tk.Toplevel(self.root)
        dlg.title('编码片段详情')
        dlg.geometry('620x320')
        dlg.transient(self.root)
        dlg.grab_set()

        # 查找该片段的实例信息
        inst = None
        if self._app:
            for code in self._app.code_system.all_codes.values():
                for i in code.instances:
                    if i.get('segment', '')[:30] == seg_text[:30]:
                        inst = i
                        break
                if inst:
                    break

        tk.Label(dlg, text=f'编码：{code_name}  |  文档：{doc_name}',
                font=FONT_B).pack(anchor='w', padx=10, pady=(10, 4))

        row = tk.Frame(dlg, bg=BG)
        row.pack(anchor='w', padx=10, pady=4)
        tk.Label(row, text='起始位置：', font=FONT, bg=BG).pack(side='left')
        start_var = tk.IntVar(value=max(inst.get('start', 0), 0) if inst else 0)
        tk.Entry(row, textvariable=start_var, width=8, font=FONT).pack(side='left', padx=4)
        tk.Label(row, text='结束位置：', font=FONT, bg=BG).pack(side='left', padx=(8, 0))
        end_var = tk.IntVar(value=inst.get('end', -1) if inst else -1)
        tk.Entry(row, textvariable=end_var, width=8, font=FONT).pack(side='left', padx=4)
        if inst:
            tk.Label(row, text=f'  编码者备注：{inst.get("memo", "")[:20]}',
                    font=('Segoe UI', 8), fg='#888', bg=BG).pack(side='left', padx=(8, 0))

        tk.Label(dlg, text='片段内容（只读）：', font=FONT_B, bg=BG).pack(anchor='w', padx=10)
        txt = tk.Text(dlg, font=('Courier New', 9), height=10,
                      bg='#f5f5f5', relief='flat', wrap='word')
        txt.insert('end', seg_text)
        txt.config(state='disabled')
        txt.pack(fill='both', expand=True, padx=10, pady=4)

        btn_row = tk.Frame(dlg, bg=BG)
        btn_row.pack(pady=6)
        tk.Button(btn_row, text='💾 保存位置', font=FONT_B, bg=ACCENT, fg='white',
                 relief='flat', padx=12, command=lambda: self._save_segment_position(
                     inst, start_var.get(), end_var.get(), dlg)
                 ).pack(side='left', padx=6)
        tk.Button(btn_row, text='🗑 删除该片段', font=FONT, fg='#c62828',
                 relief='flat', padx=10,
                 command=lambda: self._delete_segment_instance(
                     code_name, seg_text, dlg)
                 ).pack(side='left', padx=6)
        tk.Button(btn_row, text='关闭', font=FONT, relief='flat',
                 padx=10, command=dlg.destroy).pack(side='left', padx=6)

    def _save_segment_position(self, inst, start, end, dlg):
        if inst is not None:
            inst['start'] = start
            inst['end'] = end
        self._refresh_coded_segments()
        self._set_status('片段位置已更新')
        dlg.destroy()

    def _delete_segment_instance(self, code_name, seg_text, dlg):
        if not self._app:
            return
        if messagebox.askyesno('确认', f'从编码「{code_name}」中删除该片段？'):
            code = self._app.code_system._find_code_by_name(code_name)
            if code:
                code.instances = [i for i in code.instances
                                  if i.get('segment', '')[:30] != seg_text[:30]]
            self._refresh_coded_segments()
            self._set_status(f'已从「{code_name}」删除片段')
            dlg.destroy()

    def _add_code_dialog(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        dlg = tk.Toplevel(self.root)
        dlg.title('添加编码')
        dlg.geometry('440x320')
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg, text='编码名称：', font=FONT_B).place(x=10, y=16)
        name_entry = tk.Entry(dlg, font=FONT, width=28)
        name_entry.place(x=120, y=14)
        name_entry.focus()

        tk.Label(dlg, text='上级编码（可选）：', font=FONT_B).place(x=10, y=52)
        parent_var = tk.StringVar(value='（顶级编码）')
        parents = ['（顶级编码）'] + [c.name for c in self._app.code_system.root_codes]
        parent_combo = ttk.Combobox(dlg, textvariable=parent_var, values=parents,
                                    state='readonly', font=FONT, width=26)
        parent_combo.place(x=160, y=50)

        tk.Label(dlg, text='颜色：', font=FONT_B).place(x=10, y=92)
        colors = ['#e53935', '#d81b60', '#8e24aa', '#5e35b1',
                  '#1e88e5', '#039be5', '#00897b', '#43a047',
                  '#7cb342', '#fdd835', '#fb8c00', '#f4511e']
        color_var = tk.StringVar(value=colors[4])
        for i, c in enumerate(colors):
            btn = tk.Button(dlg, text=' ', bg=c, width=3, height=1, relief='raised',
                           command=lambda _, col=c: color_var.set(col))
            btn.place(x=70 + (i % 6) * 40, y=88 + (i // 6) * 26)

        tk.Label(dlg, text='描述（可选）：', font=FONT_B).place(x=10, y=148)
        desc_entry = tk.Entry(dlg, font=FONT, width=40)
        desc_entry.place(x=10, y=170, width=410)

        def do_add():
            name = name_entry.get().strip()
            if not name:
                messagebox.showwarning('提示', '请输入编码名称')
                return
            parent = (None if parent_var.get() == '（顶级编码）'
                      else parent_var.get())
            self._app.create_code(name, color=color_var.get(),
                                  description=desc_entry.get().strip(),
                                  parent_name=parent)
            self._refresh_code_tree()
            self._set_status(f'已添加编码：{name}')
            dlg.destroy()

        tk.Button(dlg, text='✅ 添加', font=FONT_B, bg=ACCENT, fg='white',
                 relief='flat', padx=16, command=do_add).place(x=150, y=205)
        tk.Button(dlg, text='取消', font=FONT, relief='flat',
                 padx=12, command=dlg.destroy).place(x=250, y=205)

    def _delete_selected_code(self):
        sel = self._code_tree.selection()
        if not sel:
            messagebox.showinfo('提示', '请先选中要删除的编码')
            return
        code_name = self._code_tree.item(sel[0], 'text')
        real_name = code_name.split(' [')[0].strip()  # 去掉阶段徽标
        if messagebox.askyesno('确认删除',
                               f'删除编码「{real_name}」？\n其所有片段实例也会移除。'):
            self._app.code_system.remove_code(real_name)
            self._refresh_code_tree()
            self._refresh_coded_segments()
            self._set_status(f'已删除：{real_name}')

    def _on_toggle_auto_code(self):
        """切换自动编码面板的显示/隐藏（使用布尔标志 _auto_panel_visible）。

        根因：ttk.PanedWindow(expand=True, fill='both') 会消耗父容器的全部剩余空间，
        当 pack_propagate=True 时子 Frame 高度被压缩为 1px。
        解决：先设置 height=220 + pack_propagate(False)，再 pack。
        """
        if not hasattr(self, '_auto_panel'):
            self._show_tab('coding')
        panel = self._auto_panel
        if self._auto_panel_visible:
            panel.pack_forget()
            self._auto_panel_visible = False
            return

        # 先设置高度 + 禁用 pack_propagate，防止被 paned expand=True 挤压到 1px
        panel['height'] = 220
        panel.pack_propagate(False)

        # 构建内容
        for w in panel.winfo_children():
            w.destroy()

        tk.Label(panel, text='⚡ 自动编码 - 关键词模式匹配（JSON格式）',
                font=FONT_B, bg='#fff8e1').pack(anchor='w', padx=8, pady=(6, 2))

        hint = tk.Label(panel,
                       text='格式: {"编码名": ["关键词1", "关键词2"], ...}',
                       font=('Segoe UI', 8), fg='#888', bg='#fff8e1')
        hint.pack(anchor='w', padx=8)

        txt = tk.Text(panel, font=('Courier New', 9), height=8,
                     bg=TEXT_BG, relief='groove')
        default_json = (
            '{\n'
            '    "正向情感": ["好", "优秀", "满意", "积极", "正面"],\n'
            '    "负向情感": ["差", "问题", "消极", "负面", "困难"],\n'
            '    "原因分析": ["因为", "由于", "所以", "导致", "造成"],\n'
            '    "建议意见": ["建议", "希望", "应该", "可以", "需要"],\n'
            '    "事实描述": ["是", "有", "包括", "属于", "表明"],\n'
            '    "程度强调": ["非常", "特别", "极其", "很", "相当"]\n'
            '}'
        )
        txt.insert('end', default_json)
        txt.pack(fill='x', padx=8, pady=4)

        self._auto_result_lbl = tk.Label(panel, text='',
                                         font=FONT, bg='#fff8e1')
        self._auto_result_lbl.pack(anchor='w', padx=8, pady=2)

        def do_auto():
            content = txt.get('1.0', 'end').strip()
            try:
                kw_dict = json.loads(content)
                if not isinstance(kw_dict, dict):
                    raise ValueError('必须是 JSON 对象')
                results = self._app.auto_code_from_keywords(kw_dict)
                total = sum(results.values())
                detail = ', '.join(f'{k}: {v}条' for k, v in results.items())
                self._auto_result_lbl.config(
                    text=f'✅ 完成！共 {total} 个片段，分配到 {len(results)} 个代码。{detail}',
                    fg=SUCCESS)
                self._refresh_code_tree()
                self._refresh_coded_segments()
            except json.JSONDecodeError as e:
                self._auto_result_lbl.config(text=f'❌ JSON 格式错误：{e}', fg=ERROR)

        btn_row = tk.Frame(panel, bg='#fff8e1')
        btn_row.pack(fill='x', padx=8, pady=4)
        tk.Button(btn_row, text='▶ 执行自动编码', font=FONT_B, bg=WARN, fg='white',
                 relief='flat', padx=12, command=do_auto).pack(side='left', padx=2)
        tk.Button(btn_row, text='关闭', font=FONT, relief='flat',
                 padx=10, command=self._on_toggle_auto_code).pack(side='left', padx=2)

        panel.pack(fill='x', padx=4, pady=(0, 4), before=self._coding_paned)
        self._auto_panel_visible = True

    def _show_code_memo(self):
        sel = self._code_tree.selection()
        if not sel:
            return
        code_name = self._code_tree.item(sel[0], 'text')
        real_name = code_name.split(' [')[0].strip()  # 去掉阶段徽标
        code = (self._app.code_system._find_code_by_name(real_name)
                if self._app else None)
        if not code:
            return
        dlg = tk.Toplevel(self.root)
        dlg.title(f'编码备注：{real_name}')
        dlg.geometry('500x300')
        dlg.transient(self.root)
        tk.Label(dlg, text=f'编码「{code_name}」- {len(code.instances)} 个片段',
                font=FONT_B).pack(anchor='w', padx=10, pady=8)
        memo_txt = tk.Text(dlg, font=('Courier New', 9), height=12,
                          bg=TEXT_BG, relief='flat')
        memo_txt.pack(fill='both', expand=True, padx=10, pady=4)
        memo_txt.insert('end', code.description or '(无备注)')
        tk.Button(dlg, text='💾 保存', font=FONT_B, bg=ACCENT, fg='white',
                 relief='flat', padx=12,
                 command=lambda: (
                     setattr(code, 'description', memo_txt.get('1.0', 'end').strip()),
                     self._set_status('备注已保存')
                 )).pack(pady=4)
        tk.Button(dlg, text='关闭', font=FONT, relief='flat',
                 padx=12, command=dlg.destroy).pack(pady=4)

    # ══════════════════════════════════════════════════════════
    # 检索 Tab
    # ══════════════════════════════════════════════════════════

    def _build_search_tab(self):
        f = tk.Frame(self._content, bg=BG)
        self._tab_pages['search'] = f

        tb = tk.Frame(f, bg=TOOLBAR_BG, height=44)
        tb.pack(fill='x')
        tk.Label(tb, text='全文检索', font=FONT_B, bg=TOOLBAR_BG).pack(side='left', padx=8)
        self._search_entry = tk.Entry(tb, font=('Segoe UI', 11), width=32)
        self._search_entry.pack(side='left', padx=4, pady=4)
        self._search_entry.bind('<Return>', lambda e: self._on_do_search())
        self._regex_var = tk.BooleanVar(value=False)
        tk.Checkbutton(tb, text='正则', variable=self._regex_var,
                       font=FONT, bg=TOOLBAR_BG).pack(side='left', padx=6)
        tk.Button(tb, text='🔍 搜索', font=FONT_B, bg=ACCENT, fg='white',
                 relief='flat', padx=12, command=self._on_do_search,
                 cursor='hand2').pack(side='left', padx=4, pady=4)
        self._search_count_lbl = tk.Label(tb, text='', font=FONT,
                                          bg=TOOLBAR_BG, fg='#555')
        self._search_count_lbl.pack(side='left', padx=8)

        # 快速词
        quick_frame = tk.Frame(f, bg=BG)
        quick_frame.pack(fill='x', padx=4, pady=(2, 0))
        tk.Label(quick_frame, text='快速：', font=FONT, bg=BG, fg='#888').pack(side='left')
        quick_words = ['现象', '原因', '影响', '观点', '建议', '描述', '结论', '数据', '理论', '方法']
        for w in quick_words:
            tk.Button(quick_frame, text=w, font=FONT, relief='flat', padx=6,
                     command=lambda _, word=w: (
                         self._search_entry.delete(0, 'end'),
                         self._search_entry.insert(0, word),
                         self._on_do_search()),
                     cursor='hand2').pack(side='left', padx=2)

        # 结果
        res_lf = ttk.Labelframe(f, text='搜索结果（双击查看详情）', padding=4)
        res_lf.pack(fill='both', expand=True, padx=4, pady=4)
        cols = ('文档', '片段内容', '编码')
        self._search_tree = ttk.Treeview(res_lf, columns=cols, show='headings', height=30)
        for col, w in zip(cols, (150, 530, 100)):
            self._search_tree.heading(col, text=col)
            self._search_tree.column(col, width=w)
        rs = ttk.Scrollbar(res_lf, orient='vertical', command=self._search_tree.yview)
        self._search_tree.configure(yscrollcommand=rs.set)
        self._search_tree.pack(side='left', fill='both', expand=True)
        rs.pack(side='right', fill='y')
        self._search_tree.bind('<Double-Button-1>', lambda e: self._on_search_result_dbl())

        btn_row = tk.Frame(f, bg=BG)
        btn_row.pack(fill='x', padx=4, pady=(0, 4))
        tk.Button(btn_row, text='📋 复制片段', font=FONT, relief='flat',
                 command=self._copy_search_result, padx=8).pack(side='left', padx=2)
        tk.Button(btn_row, text='📄 编码所选片段', font=FONT, relief='flat',
                 command=self._code_search_result, padx=8).pack(side='left', padx=2)
        tk.Button(btn_row, text='🔍 在文档中定位', font=FONT, relief='flat',
                 command=self._locate_in_doc, padx=8).pack(side='left', padx=2)

        self._search_hits = []

    def _on_do_search(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        query    = self._search_entry.get().strip()
        if not query:
            return
        use_regex = self._regex_var.get()

        # 优先用 AdvancedSearch（支持正则），否则用 search_fulltext
        if use_regex:
            from hotel_analyzer import AdvancedSearch
            analyzer = AdvancedSearch()
            docs = list(self._app.documents)
            raw_hits = analyzer.search_documents(docs, query, use_regex=True)
            self._search_hits = [
                {'doc_id': h['doc_id'], 'doc_name': h['doc_name'],
                 'segment': h.get('context', h.get('match', '')),
                 'text': h.get('match', '')}
                for h in raw_hits
            ]
        else:
            raw_hits = self._app.search_fulltext(query)
            self._search_hits = raw_hits

        self._search_tree.delete(*self._search_tree.get_children())
        for hit in self._search_hits[:500]:
            doc_id   = hit.get('doc_id', '')
            doc_name = (hit.get('doc_name') or
                        (self._app.get_document(doc_id).name
                         if self._app.get_document(doc_id) else doc_id))[:20]
            seg = str(hit.get('segment') or hit.get('text') or '')[:100]
            coded = ', '.join(
                c.name for c in self._app.code_system.all_codes.values()
                if any(i.get('doc_id') == doc_id for i in c.instances)
            )
            self._search_tree.insert('', 'end',
                values=(doc_name, seg, coded or '—'))
        n = len(self._search_hits)
        self._search_count_lbl.config(text=f'找到 {n} 条结果')
        self._set_status(f'搜索「{query}」：{n} 条结果')

    def _on_search_result_dbl(self):
        sel = self._search_tree.selection()
        if not sel:
            return
        vals = self._search_tree.item(sel[0], 'values')
        messagebox.showinfo('搜索结果', f"文档：{vals[0]}\n\n片段：\n{vals[1]}")

    def _copy_search_result(self):
        sel = self._search_tree.selection()
        if sel:
            seg = self._search_tree.item(sel[0], 'values')[1]
            self.root.clipboard_clear()
            self.root.clipboard_append(seg)
            self._set_status('已复制到剪贴板')

    def _code_search_result(self):
        sel = self._search_tree.selection()
        if not sel:
            messagebox.showinfo('提示', '请先选中搜索结果')
            return
        vals    = self._search_tree.item(sel[0], 'values')
        doc_name = vals[0]
        seg     = vals[1]
        doc = None
        if self._app and self._app.documents:
            for d in self._app.documents:
                if d.name.startswith(doc_name):
                    doc = d
                    break
        if doc:
            self._show_assign_dialog(doc, seg, -1, -1)
        else:
            messagebox.showinfo('提示', '未找到对应文档')

    def _locate_in_doc(self):
        """在文档Tab中定位当前选中的搜索结果"""
        sel = self._search_tree.selection()
        if not sel:
            return
        vals     = self._search_tree.item(sel[0], 'values')
        doc_name = vals[0]
        if self._app and self._app.documents:
            for d in self._app.documents:
                if d.name.startswith(doc_name):
                    self._show_tab('docs')
                    # 选中文档
                    for child in self._doc_tree.get_children(''):
                        if self._doc_tree.item(child, 'text') == d.id:
                            self._doc_tree.selection_set(child)
                            self._doc_tree.see(child)
                            self._on_doc_select()
                    return

    # ══════════════════════════════════════════════════════════
    # 矩阵 Tab
    # ══════════════════════════════════════════════════════════

    def _build_matrix_tab(self):
        f = tk.Frame(self._content, bg=BG)
        self._tab_pages['matrix'] = f

        tb = tk.Frame(f, bg=TOOLBAR_BG, height=44)
        tb.pack(fill='x')
        tk.Label(tb, text='矩阵分析', font=FONT_B, bg=TOOLBAR_BG).pack(side='left', padx=8)
        tk.Label(tb, text='分析维度：', font=FONT, bg=TOOLBAR_BG).pack(side='left', padx=(0, 4))
        self._dim_var  = tk.StringVar()
        self._dim_combo = ttk.Combobox(tb, textvariable=self._dim_var,
                                        state='readonly', font=FONT, width=16)
        self._dim_combo.pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='📊 生成交叉表', font=FONT_B, bg=ACCENT, fg='white',
                 relief='flat', padx=10, command=self._on_cross_tab,
                 cursor='hand2').pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='🔗 共现矩阵', font=FONT, relief='flat',
                 command=self._on_cooccurrence, padx=8).pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='📋 导出为 CSV', font=FONT, relief='flat',
                 command=self._export_matrix_csv, padx=8).pack(side='right', padx=4, pady=4)

        # 矩阵显示
        matrix_lf = ttk.Labelframe(f, text='结果', padding=4)
        matrix_lf.pack(fill='both', expand=True, padx=4, pady=4)
        self._matrix_tree = ttk.Treeview(matrix_lf, show='headings', height=22)
        msy = ttk.Scrollbar(matrix_lf, orient='vertical', command=self._matrix_tree.yview)
        self._matrix_tree.configure(yscrollcommand=msy.set)
        self._matrix_tree.pack(side='left', fill='both', expand=True)
        msy.pack(side='right', fill='y')

        self._current_matrix = None

    def _update_matrix_dims(self):
        """更新维度下拉框"""
        if not hasattr(self, '_dim_combo') or not self._dim_combo.winfo_exists():
            return
        if not self._app or not self._app.documents:
            return
        cols = list(self._app.documents.df.columns)
        skip = {'_text', '_name', 'text', '内容', 'content', '文本'}
        dims = [c for c in cols if c not in skip]
        self._dim_combo['values'] = dims
        if dims:
            self._dim_combo.set(dims[0])

    def _on_cross_tab(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        dim = self._dim_var.get()
        if not dim:
            messagebox.showinfo('提示', '请选择分析维度')
            return
        def do_analyze():
            from hotel_analyzer import CrossTabAnalysis
            analyzer = CrossTabAnalysis()
            matrix = analyzer.build_matrix(
                list(self._app.documents),
                self._app.code_system, dim)
            self._current_matrix = matrix
            self._display_matrix(matrix, f'交叉分析：{dim}')
        try:
            self._with_progress(f'正在生成交叉分析「{dim}」...', do_analyze)
            self._set_status(f'交叉分析「{dim}」完成')
        except Exception as e:
            messagebox.showerror('错误', str(e))

    def _on_cooccurrence(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        def do_analyze():
            from hotel_analyzer import CooccurrenceMatrix
            analyzer = CooccurrenceMatrix()
            matrix = analyzer.build_matrix(
                list(self._app.documents),
                self._app.code_system)
            self._current_matrix = matrix
            self._display_matrix(matrix, '共现矩阵')
        try:
            self._with_progress('正在生成共现矩阵...', do_analyze)
            self._set_status('共现矩阵生成完成')
        except Exception as e:
            messagebox.showerror('错误', str(e))

    def _display_matrix(self, df, title='结果'):
        self._matrix_tree.delete(*self._matrix_tree.get_children())
        if df is None or df.empty:
            messagebox.showinfo('提示', '没有数据（请先创建编码）')
            return
        cols = list(df.columns)
        self._matrix_tree['columns'] = cols
        self._matrix_tree.heading('#0', text='')
        self._matrix_tree.column('#0', width=130, stretch=False)
        for col in cols:
            self._matrix_tree.heading(col, text=str(col)[:15])
            self._matrix_tree.column(col, width=90)
        for row_idx, row in df.iterrows():
            values = [str(row[c])[:15] for c in cols]
            self._matrix_tree.insert('', 'end', text=str(row_idx)[:15], values=values)

    def _export_matrix_csv(self):
        if self._current_matrix is None:
            messagebox.showinfo('提示', '先生成矩阵后再导出')
            return
        path = filedialog.asksaveasfilename(
            title='导出矩阵', defaultextension='.csv',
            filetypes=[('CSV', '*.csv'), ('全部', '*.*')])
        if not path:
            return
        try:
            self._current_matrix.to_csv(path, encoding='utf-8-sig')
            messagebox.showinfo('完成', f'已导出：\n{path}')
        except Exception as e:
            messagebox.showerror('导出失败', str(e))

    # ══════════════════════════════════════════════════════════
    # 图表 Tab（新增）
    # ══════════════════════════════════════════════════════════

    def _build_charts_tab(self):
        f = tk.Frame(self._content, bg=BG)
        self._tab_pages['charts'] = f

        tb = tk.Frame(f, bg=TOOLBAR_BG, height=44)
        tb.pack(fill='x')
        tk.Label(tb, text='可视化图表', font=FONT_B, bg=TOOLBAR_BG).pack(side='left', padx=8)

        tk.Button(tb, text='🥧 编码分布饼图', font=FONT, relief='flat', padx=8,
                 command=self._chart_code_distribution,
                 cursor='hand2').pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='🔥 共现热力图', font=FONT, relief='flat', padx=8,
                 command=self._chart_cooccurrence_heatmap,
                 cursor='hand2').pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='📊 情感分布', font=FONT, relief='flat', padx=8,
                 command=self._chart_sentiment_bar,
                 cursor='hand2').pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='📈 文档词频柱状图', font=FONT, relief='flat', padx=8,
                 command=self._chart_word_freq,
                 cursor='hand2').pack(side='left', padx=4, pady=4)

        # 图表显示区
        self._chart_frame = tk.Frame(f, bg=BG)
        self._chart_frame.pack(fill='both', expand=True, padx=4, pady=4)
        self._chart_canvas = None
        self._chart_image   = None

        # 说明
        info = tk.Label(self._chart_frame,
                        text='点击上方按钮生成图表。图表将在新窗口中打开。',
                        font=FONT, fg='#888', bg=BG)
        info.pack(pady=40)

    def _refresh_charts_list(self):
        pass  # 图表Tab不需要刷新列表

    def _get_chart_output_path(self, name):
        import tempfile, os
        tmp = tempfile.gettempdir()
        return os.path.join(tmp, f'qda_chart_{name}.png')

    def _open_chart_window(self, title, img_path):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry('900x700')

        from PIL import Image, ImageTk
        try:
            img = Image.open(img_path)
        except Exception as e:
            tk.Label(win, text=f'图片加载失败：{e}', font=FONT).pack()
            return

        canvas = tk.Canvas(win, bg='white')
        canvas.pack(fill='both', expand=True)

        self._chart_canvas = canvas
        self._chart_image   = ImageTk.PhotoImage(img)

        # 缩放以适应窗口
        w_img, h_img = img.size
        screen_w = win.winfo_screenwidth()
        scale = min(1.0, (screen_w - 100) / w_img)
        new_w = int(w_img * scale)
        new_h = int(h_img * scale)

        self._chart_image = ImageTk.PhotoImage(img.resize((new_w, new_h)))
        canvas.create_image(0, 0, anchor='nw', image=self._chart_image)
        canvas.configure(scrollregion=canvas.bbox('all'))

        # 绑定关闭事件
        win.protocol('WM_DELETE_WINDOW', win.destroy)
        self._chart_windows.append(win)

    def _chart_code_distribution(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档或创建编码')
            return
        try:
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt
            from matplotlib import font_manager as fm

            # 尝试设置中文字体
            font_paths = [
                '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
                '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
                '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
                '/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf',
            ]
            for fp in font_paths:
                if Path(fp).exists():
                    prop = fm.FontProperties(fname=fp)
                    plt.rcParams['font.family'] = prop.get_name()
                    break

            labels = []
            sizes  = []
            colors = []
            for code in self._app.code_system.root_codes:
                labels.append(code.name)
                sizes.append(len(code.instances))
                colors.append(code.color)
            for code in self._app.code_system.all_codes.values():
                if code.parent is not None:
                    labels.append(f"  {code.name}")
                    sizes.append(len(code.instances))
                    colors.append(code.color)

            if not labels:
                messagebox.showinfo('提示', '尚无编码，请先创建或运行自动编码')
                return

            fig, ax = plt.subplots(figsize=(10, 8), dpi=120)
            wedges, texts, autotexts = ax.pie(
                sizes, labels=labels, autopct='%1.1f%%',
                colors=colors, startangle=90,
                pctdistance=0.75)
            for t in texts:
                t.set_fontsize(10)
            for at in autotexts:
                at.set_fontsize(9)
            ax.set_title('编码分布饼图', size=14)
            plt.tight_layout()

            path = self._get_chart_output_path('code_dist')
            fig.savefig(path, dpi=120, bbox_inches='tight', facecolor='white')
            plt.close(fig)
            self._open_chart_window('编码分布饼图', path)
            self._set_status('编码分布饼图已生成')
        except Exception as e:
            messagebox.showerror('图表生成失败', str(e))

    def _chart_cooccurrence_heatmap(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        try:
            from hotel_analyzer import CooccurrenceMatrix
            analyzer = CooccurrenceMatrix()
            matrix = analyzer.build_matrix(
                list(self._app.documents), self._app.code_system)
            if matrix.empty or len(matrix) < 2:
                messagebox.showinfo('提示', '共现矩阵为空（需要至少2个编码）')
                return

            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt
            from matplotlib import font_manager as fm

            font_paths = [
                '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
                '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
                '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
            ]
            for fp in font_paths:
                if Path(fp).exists():
                    prop = fm.FontProperties(fname=fp)
                    plt.rcParams['font.family'] = prop.get_name()
                    break

            fig, ax = plt.subplots(figsize=(12, 10), dpi=120)
            import numpy as np
            data = matrix.values.astype(float)
            np.fill_diagonal(data, 0)

            im = ax.imshow(data, cmap='Blues', aspect='auto')
            ax.set_xticks(range(len(matrix.columns)))
            ax.set_yticks(range(len(matrix.index)))
            ax.set_xticklabels(matrix.columns, rotation=45, ha='right', fontsize=9)
            ax.set_yticklabels(matrix.index, fontsize=9)
            ax.set_title('编码共现热力图', size=14)
            plt.colorbar(im, ax=ax, label='共现次数')
            plt.tight_layout()

            path = self._get_chart_output_path('cooccurrence')
            fig.savefig(path, dpi=120, bbox_inches='tight', facecolor='white')
            plt.close(fig)
            self._open_chart_window('编码共现热力图', path)
            self._set_status('共现热力图已生成')
        except Exception as e:
            messagebox.showerror('图表生成失败', str(e))

    def _chart_sentiment_bar(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        try:
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt
            from matplotlib import font_manager as fm

            for fp in ['/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
                       '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc']:
                if Path(fp).exists():
                    prop = fm.FontProperties(fname=fp)
                    plt.rcParams['font.family'] = prop.get_name()
                    break

            analyzer = self._app.sentiment_analyzer
            levels   = {'很强': 0, '正向': 0, '中性': 0, '负向': 0, '很负': 0}
            for doc in self._app.documents:
                lvl, lbl, _ = analyzer.classify(doc.text)
                key = lbl if lbl in levels else '中性'
                levels[key] = levels.get(key, 0) + 1

            categories = list(levels.keys())
            counts     = list(levels.values())
            bar_colors = ['#1b5e20', '#66bb6a', '#9e9e9e', '#ef5350', '#b71c1c']

            fig, ax = plt.subplots(figsize=(9, 6), dpi=120)
            bars = ax.bar(categories, counts, color=bar_colors, alpha=0.85)
            for bar, val in zip(bars, counts):
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.5,
                       str(val), ha='center', va='bottom', fontsize=11)
            ax.set_xlabel('情感强度')
            ax.set_ylabel('文档数量')
            ax.set_title('情感分布柱状图', size=14)
            ax.grid(axis='y', alpha=0.3)
            plt.tight_layout()

            path = self._get_chart_output_path('sentiment')
            fig.savefig(path, dpi=120, bbox_inches='tight', facecolor='white')
            plt.close(fig)
            self._open_chart_window('情感分布', path)
            self._set_status('情感分布图已生成')
        except Exception as e:
            messagebox.showerror('图表生成失败', str(e))

    def _chart_word_freq(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        try:
            import jieba
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt
            from matplotlib import font_manager as fm

            for fp in ['/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
                       '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc']:
                if Path(fp).exists():
                    prop = fm.FontProperties(fname=fp)
                    plt.rcParams['font.family'] = prop.get_name()
                    break

            # 合并所有文档文本
            all_text = ' '.join(doc.text for doc in self._app.documents)
            words    = jieba.lcut(all_text)
            # 过滤停用词（简单处理）
            stopwords = {'的', '了', '是', '在', '我', '有', '和', '就', '不', '人',
                        '都', '一', '一个', '上', '也', '很', '到', '说', '要', '去',
                        '你', '会', '着', '没有', '看', '好', '自己', '这', '那', '他'}
            words  = [w for w in words if len(w) >= 2 and w not in stopwords]
            from collections import Counter
            top20  = Counter(words).most_common(20)
            if not top20:
                messagebox.showinfo('提示', '没有足够的词汇生成词频图')
                return

            labels, values = zip(*top20)

            fig, ax = plt.subplots(figsize=(11, 8), dpi=120)
            bars = ax.barh(range(len(labels)), values,
                          color='#1976d2', alpha=0.8)
            ax.set_yticks(range(len(labels)))
            ax.set_yticklabels(labels, fontsize=10)
            ax.invert_yaxis()
            ax.set_xlabel('出现次数')
            ax.set_title('词频 Top 20', size=14)
            ax.grid(axis='x', alpha=0.3)
            plt.tight_layout()

            path = self._get_chart_output_path('wordfreq')
            fig.savefig(path, dpi=120, bbox_inches='tight', facecolor='white')
            plt.close(fig)
            self._open_chart_window('词频统计', path)
            self._set_status('词频图已生成')
        except Exception as e:
            messagebox.showerror('图表生成失败', str(e))

    # ══════════════════════════════════════════════════════════
    # 备忘 Tab（新增）
    # ══════════════════════════════════════════════════════════

    def _build_memos_tab(self):
        f = tk.Frame(self._content, bg=BG)
        self._tab_pages['memos'] = f

        tb = tk.Frame(f, bg=TOOLBAR_BG, height=44)
        tb.pack(fill='x')
        tk.Label(tb, text='备忘录与研究设计', font=FONT_B, bg=TOOLBAR_BG).pack(side='left', padx=8)
        tk.Button(tb, text='📝 研究笔记', font=FONT, command=self._add_project_memo,
                 bg=ACCENT, fg='white', relief='flat', padx=10, cursor='hand2'
                 ).pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='🏛 理论框架', font=FONT, command=self._add_theoretical_framework,
                 relief='flat', padx=8, cursor='hand2'
                 ).pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='🔬 方法论', font=FONT, command=self._add_methodology_memo,
                 relief='flat', padx=8, cursor='hand2'
                 ).pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='📄 文档批注', font=FONT, command=self._add_doc_memo,
                 relief='flat', padx=8, cursor='hand2'
                 ).pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='🗑 删除选中', font=FONT, command=self._delete_memo,
                 relief='flat', padx=8, cursor='hand2'
                 ).pack(side='left', padx=4, pady=4)

        # 主区
        paned = ttk.PanedWindow(f, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=4, pady=4)

        # 左侧：备忘录列表
        lf = ttk.Labelframe(paned, text='备忘录列表', padding=4)
        paned.add(lf, weight=1)
        cols = ('类型', '关联', '内容预览')
        self._memo_tree = ttk.Treeview(lf, columns=cols, show='headings', height=30)
        for col, w in zip(cols, (90, 130, 300)):
            self._memo_tree.heading(col, text=col)
            self._memo_tree.column(col, width=w)
        ms = ttk.Scrollbar(lf, orient='vertical', command=self._memo_tree.yview)
        self._memo_tree.configure(yscrollcommand=ms.set)
        self._memo_tree.pack(side='left', fill='both', expand=True)
        ms.pack(side='right', fill='y')
        self._memo_tree.bind('<<TreeviewSelect>>', lambda e: self._on_memo_select())
        self._memo_tree.bind('<Double-Button-1>', lambda e: self._edit_memo())

        # 右侧：备忘录内容
        rf = ttk.Labelframe(paned, text='内容', padding=4)
        paned.add(rf, weight=2)
        self._memo_text = tk.Text(rf, wrap='word', font=('Courier New', 10),
                                  bg=TEXT_BG, relief='flat')
        ms2 = ttk.Scrollbar(rf, command=self._memo_text.yview)
        self._memo_text.configure(yscrollcommand=ms2.set)
        ms2.pack(side='right', fill='y')
        self._memo_text.pack(fill='both', expand=True)
        tk.Button(rf, text='💾 保存修改', font=FONT_B, bg=SUCCESS, fg='white',
                 relief='flat', padx=12,
                 command=self._save_memo_edit).pack(pady=4)

        self._current_memo_id = None

    def _get_all_memos(self):
        """获取所有备忘录的扁平列表，返回 list of dict"""
        if not self._app:
            return []
        mm = self._app.memo_manager
        results = []
        # 项目级
        for m in mm.project_memos:
            results.append({'id': id(m), 'type': m.type,
                           'linked_id': '(项目)', 'text': m.text})
        # 文档级
        for doc_id, memos in mm.doc_memos.items():
            for m in memos:
                results.append({'id': id(m), 'type': m.type,
                               'linked_id': f'文档:{doc_id}', 'text': m.text})
        # 代码级
        for code_id, memos in mm.code_memos.items():
            for m in memos:
                results.append({'id': id(m), 'type': m.type,
                               'linked_id': f'代码:{code_id}', 'text': m.text})
        return results

    def _refresh_memos_list(self):
        if not hasattr(self, '_memo_tree') or not self._memo_tree.winfo_exists():
            return
        self._memo_tree.delete(*self._memo_tree.get_children())
        if not self._app:
            return
        for memo_d in self._get_all_memos():
            memo_type = memo_d.get('type', 'general')
            linked_id = str(memo_d.get('linked_id', ''))[:15]
            content   = str(memo_d.get('text', ''))[:50]
            self._memo_tree.insert('', 'end', iid=str(memo_d['id']),
                                  values=(memo_type, linked_id, content))

    def _on_memo_select(self):
        sel = self._memo_tree.selection()
        if not sel:
            return
        memo_id = int(sel[0])
        self._current_memo_id = memo_id
        if not self._app:
            return
        # Find memo object in the right store
        mm = self._app.memo_manager
        for m in mm.project_memos:
            if id(m) == memo_id:
                self._memo_text.delete('1.0', 'end')
                self._memo_text.insert('end', m.text)
                return
        for doc_id, memos in mm.doc_memos.items():
            for m in memos:
                if id(m) == memo_id:
                    self._memo_text.delete('1.0', 'end')
                    self._memo_text.insert('end', m.text)
                    return
        for code_id, memos in mm.code_memos.items():
            for m in memos:
                if id(m) == memo_id:
                    self._memo_text.delete('1.0', 'end')
                    self._memo_text.insert('end', m.text)
                    return

    def _edit_memo(self):
        pass  # 双击和选择都通过 _on_memo_select 处理

    def _save_memo_edit(self):
        if not self._current_memo_id or not self._app:
            return
        new_text = self._memo_text.get('1.0', 'end').strip()
        memo_id = self._current_memo_id
        mm = self._app.memo_manager
        found = False
        for m in mm.project_memos:
            if id(m) == memo_id:
                mm.update_memo(m, new_text)
                found = True
                break
        if not found:
            for doc_id, memos in list(mm.doc_memos.items()):
                for m in memos:
                    if id(m) == memo_id:
                        mm.update_memo(m, new_text)
                        found = True
                        break
                if found:
                    break
        if not found:
            for code_id, memos in list(mm.code_memos.items()):
                for m in memos:
                    if id(m) == memo_id:
                        mm.update_memo(m, new_text)
                        found = True
                        break
                if found:
                    break
        self._refresh_memos_list()
        self._set_status('备忘录已保存')

    def _add_project_memo(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        dlg = tk.Toplevel(self.root)
        dlg.title('添加项目备忘录')
        dlg.geometry('500x300')
        dlg.transient(self.root)
        dlg.grab_set()
        tk.Label(dlg, text='备忘录内容：', font=FONT_B).pack(anchor='w', padx=10, pady=6)
        txt = tk.Text(dlg, font=('Courier New', 10), bg=TEXT_BG)
        txt.pack(fill='both', expand=True, padx=10, pady=4)
        def do_add():
            text = txt.get('1.0', 'end').strip()
            if not text:
                return
            self._app.add_project_memo(text)
            self._refresh_memos_list()
            self._set_status('项目备忘录已添加')
            dlg.destroy()
        tk.Button(dlg, text='✅ 添加', font=FONT_B, bg=ACCENT, fg='white',
                 relief='flat', padx=12, command=do_add).pack(pady=6)
        tk.Button(dlg, text='取消', font=FONT, relief='flat',
                 padx=12, command=dlg.destroy).pack(pady=4)

    def _add_doc_memo(self):
        if not self._app or not self._app.documents:
            messagebox.showinfo('提示', '请先加载文档')
            return
        # 选择文档对话框
        dlg = tk.Toplevel(self.root)
        dlg.title('选择文档')
        dlg.geometry('400x300')
        dlg.transient(self.root)
        tk.Label(dlg, text='选择要添加备忘录的文档：', font=FONT_B).pack(pady=6)
        listbox = tk.Listbox(dlg, font=FONT)
        listbox.pack(fill='both', expand=True, padx=10)
        for doc in self._app.documents:
            listbox.insert('end', doc.name)
        selected_doc = [None]

        def do_select():
            idx = listbox.curselection()
            if not idx:
                return
            selected_doc[0] = self._app.documents[listbox.index(idx[0])]
            dlg.destroy()
            # 弹出内容输入
            dlg2 = tk.Toplevel(self.root)
            dlg2.title('文档备忘录')
            dlg2.geometry('500x300')
            dlg2.transient(self.root)
            tk.Label(dlg2, text=f'文档：{selected_doc[0].name}',
                    font=FONT_B).pack(anchor='w', padx=10, pady=6)
            txt2 = tk.Text(dlg2, font=('Courier New', 10), bg=TEXT_BG)
            txt2.pack(fill='both', expand=True, padx=10, pady=4)
            def do_add():
                text = txt2.get('1.0', 'end').strip()
                if text:
                    self._app.add_document_memo(selected_doc[0].id, text)
                    self._refresh_memos_list()
                    self._set_status('文档备忘录已添加')
                dlg2.destroy()
            tk.Button(dlg2, text='✅ 添加', font=FONT_B, bg=ACCENT, fg='white',
                     relief='flat', padx=12, command=do_add).pack(pady=6)

        tk.Button(dlg, text='选择', font=FONT_B, command=do_select).pack(pady=4)

    def _add_theoretical_framework(self):
        """添加理论框架备忘录"""
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        dlg = tk.Toplevel(self.root)
        dlg.title('理论框架')
        dlg.geometry('600x400')
        dlg.transient(self.root)
        dlg.grab_set()
        
        tk.Label(dlg, text='理论框架说明', font=FONT_B).pack(anchor='w', padx=10, pady=6)
        tk.Label(dlg, text='记录研究的理论基础、概念框架和理论假设',
                 font=('Segoe UI', 8), fg='#666').pack(anchor='w', padx=10)
        
        hint_text = """示例框架结构：
        
理论名称：[如：扎根理论、计划行为理论等]

核心概念：
- 概念1：定义
- 概念2：定义

理论假设：
- 假设1：...
- 假设2：...

与本研究的关联：
[说明该理论如何指导本研究]"""
        
        txt = tk.Text(dlg, font=('Courier New', 10), bg=TEXT_BG)
        txt.pack(fill='both', expand=True, padx=10, pady=4)
        txt.insert('end', hint_text)
        
        def do_add():
            text = txt.get('1.0', 'end').strip()
            if text and text != hint_text:
                self._app.add_project_memo(text, memo_type='theory')
                self._refresh_memos_list()
                self._set_status('理论框架已添加')
            dlg.destroy()
        
        tk.Button(dlg, text='✅ 保存', font=FONT_B, bg=ACCENT, fg='white',
                 relief='flat', padx=12, command=do_add).pack(pady=6)
        tk.Button(dlg, text='取消', font=FONT, relief='flat',
                 padx=12, command=dlg.destroy).pack(pady=4)

    def _add_methodology_memo(self):
        """添加方法论备忘录"""
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        dlg = tk.Toplevel(self.root)
        dlg.title('方法论说明')
        dlg.geometry('600x450')
        dlg.transient(self.root)
        dlg.grab_set()
        
        tk.Label(dlg, text='研究方法论', font=FONT_B).pack(anchor='w', padx=10, pady=6)
        tk.Label(dlg, text='记录研究设计、数据收集和分析方法',
                 font=('Segoe UI', 8), fg='#666').pack(anchor='w', padx=10)
        
        hint_text = """方法论说明示例：

研究方法：[如：质性研究、混合方法、案例研究等]

数据收集：
- 来源：...
- 时间：...
- 样本量：...

分析方法：
- 编码方式：[如：开放编码、轴心编码、选择性编码]
- 分析工具：[如：NVivo、ATLAS.ti、手工编码]
- 验证方法：[如：三角验证、成员检验、同行审议]

伦理考量：
[数据匿名化、保密性等说明]"""
        
        txt = tk.Text(dlg, font=('Courier New', 10), bg=TEXT_BG)
        txt.pack(fill='both', expand=True, padx=10, pady=4)
        txt.insert('end', hint_text)
        
        def do_add():
            text = txt.get('1.0', 'end').strip()
            if text and text != hint_text:
                self._app.add_project_memo(text, memo_type='methodology')
                self._refresh_memos_list()
                self._set_status('方法论说明已添加')
            dlg.destroy()
        
        tk.Button(dlg, text='✅ 保存', font=FONT_B, bg=ACCENT, fg='white',
                 relief='flat', padx=12, command=do_add).pack(pady=6)
        tk.Button(dlg, text='取消', font=FONT, relief='flat',
                 padx=12, command=dlg.destroy).pack(pady=4)

    def _delete_memo(self):
        sel = self._memo_tree.selection()
        if not sel:
            messagebox.showinfo('提示', '请先选中要删除的备忘录')
            return
        memo_id = int(sel[0])
        if messagebox.askyesno('确认', '删除这条备忘录？'):
            if self._app:
                mm = self._app.memo_manager
                # 项目级
                for i, m in enumerate(mm.project_memos):
                    if id(m) == memo_id:
                        mm.project_memos.pop(i)
                        break
                else:
                    # 文档级
                    for doc_id, memos in list(mm.doc_memos.items()):
                        for i, m in enumerate(memos):
                            if id(m) == memo_id:
                                mm.doc_memos[doc_id].pop(i)
                                if not mm.doc_memos[doc_id]:
                                    del mm.doc_memos[doc_id]
                                break
                        else:
                            continue
                        break
                    else:
                        # 代码级
                        for code_id, memos in list(mm.code_memos.items()):
                            for i, m in enumerate(memos):
                                if id(m) == memo_id:
                                    mm.code_memos[code_id].pop(i)
                                    if not mm.code_memos[code_id]:
                                        del mm.code_memos[code_id]
                                    break
                            else:
                                continue
                            break
            self._refresh_memos_list()
            self._memo_text.delete('1.0', 'end')
            self._current_memo_id = None
            self._set_status('备忘录已删除')

    # ══════════════════════════════════════════════════════════
    # 导出 Tab
    # ══════════════════════════════════════════════════════════

    def _build_export_tab(self):
        f = tk.Frame(self._content, bg=BG)
        self._tab_pages['export'] = f

        tb = tk.Frame(f, bg=TOOLBAR_BG, height=44)
        tb.pack(fill='x')
        tk.Label(tb, text='导出报告', font=FONT_B, bg=TOOLBAR_BG).pack(side='left', padx=8)

        row1 = tk.Frame(f, bg=BG)
        row1.pack(fill='x', padx=8, pady=8)

        # 报告卡片
        report_card = tk.LabelFrame(row1, text='📄 分析报告', font=FONT_B,
                                    bg=BG, padx=12, pady=8)
        report_card.pack(side='left', padx=(0, 8))
        for fmt, icon, cmd in [
            ('Word 报告 (.docx)', '📄', self._export_word),
            ('Excel 报告 (.xlsx)', '📊', self._export_excel),
            ('Markdown 报告 (.md)', '📝', self._export_markdown),
        ]:
            tk.Button(report_card, text=f'{icon}  {fmt}', font=FONT,
                     relief='groove', padx=10, pady=5, command=cmd,
                     cursor='hand2').pack(fill='x', pady=2)

        # 数据卡片
        data_card = tk.LabelFrame(row1, text='💾 编码数据', font=FONT_B,
                                  bg=BG, padx=12, pady=8)
        data_card.pack(side='left', padx=(0, 8))
        for label, icon, cmd in [
            ('编码片段 (CSV)', '📋', self._export_segments_csv),
            ('编码本 (JSON)', '📖', self._export_codebook_json),
            ('导出代码书 (Markdown)', '📖', self._export_codebook_md),
            ('编码频率统计', '📈', self._export_code_frequency),
            ('备忘录导出', '📝', self._export_memos),
            ('保存项目', '💾', self._save_project),
        ]:
            tk.Button(data_card, text=f'{icon}  {label}', font=FONT,
                     relief='groove', padx=10, pady=5, command=cmd,
                     cursor='hand2').pack(fill='x', pady=2)

        # 项目管理卡片
        proj_card = tk.LabelFrame(row1, text='💼 项目', font=FONT_B,
                                  bg=BG, padx=12, pady=8)
        proj_card.pack(side='left')
        for label, icon, cmd in [
            ('🆕 新建项目', None, self._new_project),
            ('📂 打开项目', None, self._open_project),
            ('🔁 打开最近', None, self._open_recent_dialog),
        ]:
            tk.Button(proj_card, text=label, font=FONT,
                     relief='groove', padx=10, pady=5, command=cmd,
                     cursor='hand2').pack(fill='x', pady=2)

        # 预览
        preview_lf = ttk.Labelframe(f, text='导出预览', padding=4)
        preview_lf.pack(fill='both', expand=True, padx=8, pady=(0, 8))
        self._export_preview = tk.Text(preview_lf, font=('Courier New', 8),
                                        bg=TEXT_BG, relief='flat', state='disabled')
        sp = ttk.Scrollbar(preview_lf, command=self._export_preview.yview)
        self._export_preview.configure(yscrollcommand=sp.set)
        sp.pack(side='right', fill='y')
        self._export_preview.pack(side='left', fill='both', expand=True)

    def _refresh_export_preview(self):
        if not hasattr(self, '_export_preview') or not self._export_preview.winfo_exists():
            return
        self._export_preview.config(state='normal')
        self._export_preview.delete('1.0', 'end')
        if not self._app:
            self._export_preview.insert('end', '(尚未加载项目或数据)')
        else:
            s = self._app.summary()
            preview = [
                f"项目: {s.get('project_name', '未命名')}",
                f"文档数: {s.get('total_documents', 0)}",
                f"代码数: {s.get('total_codes', 0)}",
                f"总编码片段: {s.get('total_instances', s.get('code_instances', 0))}",
                f"备忘录: {s.get('total_memos', 0)}",
                '',
                '--- 代码列表 ---',
            ]
            if self._app.code_system:
                for code in self._app.code_system.root_codes:
                    preview.append(
                        f"  🏷 {code.name} ({len(code.instances)} 片段)")
                    for child in code.children:
                        preview.append(
                            f"    └ {child.name} ({len(child.instances)} 片段)")
            self._export_preview.insert('end', '\n'.join(preview))
        self._export_preview.config(state='disabled')

    def _do_export_file(self, title, ext, generate_fn):
        path = filedialog.asksaveasfilename(title=title, defaultextension=ext,
                                           filetypes=[(title, f'*{ext}'), ('全部', '*.*')])
        if not path:
            return
        try:
            generate_fn(path)
            messagebox.showinfo('完成', f'已导出：\n{path}')
            self._set_status(f'已导出：{path}')
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror('导出失败', str(e))

    def _export_word(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        self._do_export_file('导出 Word', '.docx',
                             lambda p: self._with_progress(
                                 '正在生成 Word 报告...',
                                 lambda: self._app.generate_report(p, format='docx')))

    def _export_excel(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        self._do_export_file('导出 Excel', '.xlsx',
                             lambda p: self._with_progress(
                                 '正在生成 Excel 报告...',
                                 lambda: self._app.generate_report(p, format='xlsx')))

    def _export_markdown(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        self._do_export_file('导出 Markdown', '.md',
                             lambda p: self._with_progress(
                                 '正在生成 Markdown 报告...',
                                 lambda: self._app.generate_report(p, format='md')))

    def _export_segments_csv(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        path = filedialog.asksaveasfilename(title='导出编码片段',
                                           defaultextension='.csv',
                                           filetypes=[('CSV', '*.csv'), ('全部', '*.*')])
        if not path:
            return
        try:
            import csv
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                w = csv.writer(f)
                w.writerow(['doc_id', 'doc_name', 'code', 'segment',
                            'start', 'end', 'memo'])
                for code in self._app.code_system.all_codes.values():
                    for inst in code.instances:
                        doc_name = ''
                        if self._app.documents:
                            doc = self._app.get_document(inst.get('doc_id', ''))
                            if doc:
                                doc_name = doc.name
                        w.writerow([inst.get('doc_id', ''), doc_name, code.name,
                                   inst.get('segment', ''), inst.get('start', ''),
                                   inst.get('end', ''), inst.get('memo', '')])
            messagebox.showinfo('完成', f'已导出：\n{path}')
        except Exception as e:
            messagebox.showerror('导出失败', str(e))

    def _export_codebook_json(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        path = filedialog.asksaveasfilename(title='导出编码本',
                                           defaultextension='.json',
                                           filetypes=[('JSON', '*.json'),
                                                     ('全部', '*.*')])
        if not path:
            return
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(self._app.code_system.to_dict(), f,
                         ensure_ascii=False, indent=2)
            messagebox.showinfo('完成', f'已导出：\n{path}')
        except Exception as e:
            messagebox.showerror('导出失败', str(e))

    def _export_codebook_md(self):
        """导出代码书为 Markdown"""
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        path = filedialog.asksaveasfilename(title='导出代码书 (Markdown)',
                                           defaultextension='.md',
                                           filetypes=[('Markdown', '*.md'),
                                                     ('全部', '*.*')])
        if not path:
            return
        try:
            self._app.export_codebook(path)
            messagebox.showinfo('完成', f'已导出代码书：\n{path}')
            self._set_status(f'已导出代码书：{path}')
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror('导出失败', str(e))

    def _export_code_frequency(self):
        """导出编码频率统计表（用于学术论文）"""
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        path = filedialog.asksaveasfilename(title='导出编码频率统计',
                                           defaultextension='.csv',
                                           filetypes=[('CSV', '*.csv'), ('全部', '*.*')])
        if not path:
            return
        try:
            import csv
            codes = self._app.code_system.all_codes
            total_instances = sum(len(c.instances) for c in codes.values())
            
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                w = csv.writer(f)
                w.writerow(['编码名称', '片段数量', '占比(%)', '父级编码', '描述'])
                
                root_codes = self._app.code_system.root_codes
                for code in root_codes:
                    count = len(code.instances)
                    pct = (count / total_instances * 100) if total_instances > 0 else 0
                    desc = code.description or ''
                    w.writerow([code.name, count, f'{pct:.1f}', '-', desc])
                    
                    for child in code.children:
                        child_count = len(child.instances)
                        child_pct = (child_count / total_instances * 100) if total_instances > 0 else 0
                        child_desc = child.description or ''
                        w.writerow([f'  └ {child.name}', child_count, f'{child_pct:.1f}', code.name, child_desc])
                
                w.writerow(['', '', '', '', ''])
                w.writerow(['总计', total_instances, '100.0', '', ''])
            
            messagebox.showinfo('完成', f'已导出：\n{path}')
            self._set_status(f'已导出编码频率统计：{path}')
        except Exception as e:
            messagebox.showerror('导出失败', str(e))

    def _export_memos(self):
        """导出所有备忘录为Markdown格式"""
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        path = filedialog.asksaveasfilename(title='导出备忘录',
                                           defaultextension='.md',
                                           filetypes=[('Markdown', '*.md'), ('全部', '*.*')])
        if not path:
            return
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write('# 研究备忘录\n\n')
                
                mm = self._app.memo_manager
                
                f.write('## 项目级备忘录\n\n')
                for m in mm.project_memos:
                    f.write(f'### {m.type}\n\n')
                    f.write(f'{m.text}\n\n')
                    f.write(f'_创建时间: {m.created_at}_\n\n---\n\n')
                
                f.write('## 文档级备忘录\n\n')
                for doc_id, memos in mm.doc_memos.items():
                    doc = self._app.get_document(doc_id)
                    doc_name = doc.name if doc else doc_id
                    f.write(f'### 文档: {doc_name}\n\n')
                    for m in memos:
                        f.write(f'**类型: {m.type}**\n\n')
                        f.write(f'{m.text}\n\n')
                        f.write(f'_创建时间: {m.created_at}_\n\n---\n\n')
                
                f.write('## 代码级备忘录\n\n')
                for code_id, memos in mm.code_memos.items():
                    code = self._app.code_system._find_code_by_name(code_id)
                    code_name = code.name if code else code_id
                    f.write(f'### 代码: {code_name}\n\n')
                    for m in memos:
                        f.write(f'**类型: {m.type}**\n\n')
                        f.write(f'{m.text}\n\n')
                        f.write(f'_创建时间: {m.created_at}_\n\n---\n\n')
            
            messagebox.showinfo('完成', f'已导出备忘录：\n{path}')
            self._set_status(f'已导出备忘录：{path}')
        except Exception as e:
            messagebox.showerror('导出失败', str(e))

    def _save_project(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        path = filedialog.asksaveasfilename(title='保存项目',
                                           defaultextension='.qda',
                                           filetypes=[('QualCoder项目', '*.qda'),
                                                     ('全部', '*.*')])
        if not path:
            return
        def do_save():
            self._app.save_project(path)
            self._record_recent_project(path)
        self._with_progress('正在保存项目...', do_save)
        messagebox.showinfo('完成', f'项目已保存：\n{path}')
        self._set_status(f'项目已保存：{path}')

    def _new_project(self):
        if messagebox.askyesno('新建项目',
                               '确定要新建项目吗？当前未保存的数据将丢失。'):
            from hotel_analyzer import QDAApplication
            self._app = QDAApplication()
            self._refresh_export_preview()
            self._set_status('新项目已创建')

    def _open_project(self):
        path = filedialog.askopenfilename(title='打开项目',
                                          filetypes=[('QualCoder项目', '*.qda'),
                                                    ('全部', '*.*')])
        if not path:
            return
        self._open_project_path(path)

    # ── 状态栏 ────────────────────────────────────────────────

    def _set_status(self, msg: str):
        self._status_bar.config(text=msg)
        self._status_label.config(text=msg)
        self.root.update_idletasks()

    def _update_stats(self):
        """更新右侧统计栏"""
        if not self._app:
            self._right_stats.config(text='📄 0 docs | 🏷️ 0 codes | 🔗 0 instances')
            return
        n_docs = len(self._app.documents) if hasattr(self._app, 'documents') and self._app.documents else 0
        codes = self._app.code_system.all_codes if hasattr(self._app, 'code_system') else {}
        n_codes = len(codes)
        n_instances = sum(len(c.instances) for c in codes.values()) if codes else 0
        self._right_stats.config(
            text=f'📄 {n_docs} docs | 🏷️ {n_codes} codes | 🔗 {n_instances} instances')

    # ── 键盘快捷键辅助 ─────────────────────────────────────────

    def _on_ctrl_f(self):
        """Ctrl+F: focus search entry, switch to search tab"""
        self._show_tab('search')
        if hasattr(self, '_search_entry') and self._search_entry.winfo_exists():
            self._search_entry.focus_set()
            self._search_entry.selection_range(0, 'end')

    def _on_delete_shortcut(self):
        """Delete key: if coding tab active, delete selected code"""
        if self._current_tab == 'coding':
            self._delete_selected_code()

    # ── 最近项目 ──────────────────────────────────────────────

    def _recent_projects_path(self) -> str:
        return os.path.join(os.path.expanduser('~'), '.qda_recent.json')

    def _load_recent_projects(self) -> list:
        path = self._recent_projects_path()
        try:
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return data if isinstance(data, list) else []
        except Exception:
            pass
        return []

    def _save_recent_projects(self):
        path = self._recent_projects_path()
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(self._recent_projects, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _record_recent_project(self, project_path: str):
        """记录最近项目"""
        if project_path in self._recent_projects:
            self._recent_projects.remove(project_path)
        self._recent_projects.insert(0, project_path)
        self._recent_projects = self._recent_projects[:10]  # keep max 10
        self._save_recent_projects()

    def _open_recent_dialog(self):
        """显示最近项目对话框"""
        if not self._recent_projects:
            messagebox.showinfo('提示', '没有最近打开的项目')
            return
        dlg = tk.Toplevel(self.root)
        dlg.title('打开最近项目')
        dlg.geometry('500x320')
        dlg.transient(self.root)
        dlg.grab_set()
        tk.Label(dlg, text='最近打开的项目：', font=FONT_B).pack(anchor='w', padx=10, pady=6)
        lb = tk.Listbox(dlg, font=FONT)
        lb.pack(fill='both', expand=True, padx=10, pady=4)
        for p in self._recent_projects:
            lb.insert('end', p)
        def do_open():
            sel = lb.curselection()
            if not sel:
                return
            path = self._recent_projects[sel[0]]
            dlg.destroy()
            if os.path.exists(path):
                self._open_project_path(path)
            else:
                messagebox.showwarning('提示', f'项目文件不存在：\n{path}')
        def do_clear():
            self._recent_projects = []
            self._save_recent_projects()
            dlg.destroy()
        btn_fr = tk.Frame(dlg, bg=BG)
        btn_fr.pack(fill='x', padx=10, pady=6)
        tk.Button(btn_fr, text='📂 打开', font=FONT_B, bg=ACCENT, fg='white',
                 relief='flat', padx=12, command=do_open).pack(side='left', padx=4)
        tk.Button(btn_fr, text='🗑 清空记录', font=FONT, relief='flat',
                 padx=8, command=do_clear).pack(side='left', padx=4)
        tk.Button(btn_fr, text='取消', font=FONT, relief='flat',
                 padx=8, command=dlg.destroy).pack(side='right', padx=4)

    def _open_project_path(self, path: str):
        """通过路径打开项目"""
        try:
            from hotel_analyzer import QDAApplication
            app = QDAApplication.load_project(path)
            self._app = app
            self._record_recent_project(path)
            self._refresh_export_preview()
            self._refresh_docs_list()
            self._update_stats()
            self._set_status(f'已打开项目：{path}')
            messagebox.showinfo('完成', f'项目已加载：\n{path}')
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror('打开失败', str(e))

    # ── 进度对话框 ────────────────────────────────────────────

    class ProgressDialog:
        """简单的模态进度对话框（无取消按钮）"""
        def __init__(self, parent, label_text='处理中...'):
            self.dlg = tk.Toplevel(parent)
            self.dlg.title('请稍候')
            self.dlg.geometry('360x100')
            self.dlg.transient(parent)
            self.dlg.grab_set()
            self.dlg.resizable(False, False)
            # 居中
            self.dlg.update_idletasks()
            x = parent.winfo_x() + (parent.winfo_width() - 360) // 2
            y = parent.winfo_y() + (parent.winfo_height() - 100) // 2
            self.dlg.geometry(f'+{x}+{y}')
            tk.Label(self.dlg, text=label_text, font=FONT,
                    wraplength=320).pack(pady=(14, 6))
            self.progress = ttk.Progressbar(self.dlg, mode='indeterminate', length=300)
            self.progress.pack(pady=4)
            self.progress.start(20)
            self.dlg.update()

        def close(self):
            try:
                self.progress.stop()
                self.dlg.grab_release()
                self.dlg.destroy()
            except Exception:
                pass

    def _with_progress(self, label_text, callback):
        """显示进度对话框，执行回调，关闭对话框"""
        dlg = self.ProgressDialog(self.root, label_text)
        try:
            callback()
        finally:
            dlg.close()
            self._update_stats()
            self.root.update_idletasks()

    # ── 代码合并对话框 ────────────────────────────────────────

    def _merge_codes_dialog(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        all_names = sorted(c.name for c in self._app.code_system.all_codes.values())
        if len(all_names) < 2:
            messagebox.showinfo('提示', '需要至少2个编码才能合并')
            return
        dlg = tk.Toplevel(self.root)
        dlg.title('合并编码')
        dlg.geometry('520x440')
        dlg.transient(self.root)
        dlg.grab_set()
        # Source codes
        tk.Label(dlg, text='源编码（将被合并，可多选）：', font=FONT_B).pack(anchor='w', padx=10, pady=(8, 2))
        src_frame = tk.Frame(dlg, bg=BG)
        src_frame.pack(fill='both', expand=True, padx=10)
        src_lb = tk.Listbox(src_frame, font=FONT, selectmode='extended')
        src_sc = ttk.Scrollbar(src_frame, orient='vertical', command=src_lb.yview)
        src_lb.configure(yscrollcommand=src_sc.set)
        src_lb.pack(side='left', fill='both', expand=True)
        src_sc.pack(side='right', fill='y')
        for n in all_names:
            src_lb.insert('end', n)
        # Target code
        tk.Label(dlg, text='目标编码（合并到）：', font=FONT_B).pack(anchor='w', padx=10, pady=(6, 2))
        tgt_var = tk.StringVar(value=all_names[0] if all_names else '')
        tgt_combo = ttk.Combobox(dlg, textvariable=tgt_var, values=all_names,
                                 state='readonly', font=FONT, width=30)
        tgt_combo.pack(anchor='w', padx=10, pady=(0, 6))
        def do_merge():
            sel = src_lb.curselection()
            if not sel:
                messagebox.showwarning('提示', '请选择要合并的源编码')
                return
            sources = [src_lb.get(i) for i in sel]
            target = tgt_var.get()
            if not target:
                messagebox.showwarning('提示', '请选择目标编码')
                return
            if target in sources:
                messagebox.showwarning('提示', '目标编码不能同时是源编码')
                return
            try:
                self._app.merge_codes(sources, target)
                self._refresh_code_tree()
                self._refresh_coded_segments()
                self._update_stats()
                self._set_status(f'已合并 {len(sources)} 个编码到「{target}」')
                dlg.destroy()
            except Exception as e:
                messagebox.showerror('合并失败', str(e))
        btn_fr = tk.Frame(dlg, bg=BG)
        btn_fr.pack(fill='x', padx=10, pady=8)
        tk.Button(btn_fr, text='✅ 合并', font=FONT_B, bg=ACCENT, fg='white',
                 relief='flat', padx=14, command=do_merge).pack(side='left', padx=4)
        tk.Button(btn_fr, text='取消', font=FONT, relief='flat',
                 padx=10, command=dlg.destroy).pack(side='left', padx=4)

    # ── 属性管理器 ────────────────────────────────────────────

    def _attribute_manager(self):
        if not self._app or not self._app.documents:
            messagebox.showinfo('提示', '请先加载文档')
            return
        dlg = tk.Toplevel(self.root)
        dlg.title('文档属性管理')
        dlg.geometry('800x480')
        dlg.transient(self.root)
        dlg.grab_set()
        # Collect all attribute keys
        docs = list(self._app.documents)
        all_keys = []
        for doc in docs:
            attrs = getattr(doc, 'attributes', None) or {}
            for k in attrs:
                if k not in all_keys:
                    all_keys.append(k)
        # Treeview
        frame = tk.Frame(dlg, bg=BG)
        frame.pack(fill='both', expand=True, padx=8, pady=4)
        cols = ['文档名称'] + all_keys
        tree = ttk.Treeview(frame, columns=cols, show='headings', height=20)
        tree.heading('#0', text='')
        tree.column('#0', width=0, stretch=False)
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=120 if col != '文档名称' else 180)
        for doc in docs:
            attrs = getattr(doc, 'attributes', None) or {}
            vals = [doc.name] + [str(attrs.get(k, '')) for k in all_keys]
            tree.insert('', 'end', values=vals)
        sc = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=sc.set)
        tree.pack(side='left', fill='both', expand=True)
        sc.pack(side='right', fill='y')
        op_frame = tk.Frame(dlg, bg=BG)
        op_frame.pack(fill='x', padx=8, pady=4)
        def add_attr():
            name = tk.simpledialog.askstring('添加属性', '属性名称：', parent=dlg)
            if name and name.strip():
                all_keys.append(name.strip())
                # Rebuild columns
                cols = ['文档名称'] + all_keys
                tree['columns'] = cols
                for col in cols:
                    tree.heading(col, text=col)
                    tree.column(col, width=120 if col != '文档名称' else 180)
                # Re-populate
                for item in tree.get_children():
                    tree.delete(item)
                for doc in docs:
                    attrs = getattr(doc, 'attributes', None) or {}
                    vals = [doc.name] + [str(attrs.get(k, '')) for k in all_keys]
                    tree.insert('', 'end', values=vals)
        def edit_cell():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo('提示', '请先选中一个单元格所在的行')
                return
            # Ask for column and value
            col_names = ['文档名称'] + all_keys
            col_sel = tk.simpledialog.askstring('编辑属性',
                f'可用列：{", ".join(col_names[1:])}\n请输入列名称：', parent=dlg)
            if not col_sel or col_sel.strip() not in col_names:
                return
            col_name = col_sel.strip()
            new_val = tk.simpledialog.askstring('编辑属性',
                f'输入「{col_name}」的新值：', parent=dlg)
            if new_val is None:
                return
            col_idx = col_names.index(col_name)
            item_id = sel[0]
            cur_vals = list(tree.item(item_id, 'values'))
            doc_name = cur_vals[0]
            # Find doc
            target_doc = None
            for d in docs:
                if d.name == doc_name:
                    target_doc = d
                    break
            if target_doc:
                if not hasattr(target_doc, 'attributes') or target_doc.attributes is None:
                    target_doc.attributes = {}
                target_doc.attributes[col_name] = new_val
                cur_vals[col_idx] = new_val
                tree.item(item_id, values=cur_vals)
        def delete_attr():
            col_name = tk.simpledialog.askstring('删除属性',
                f'可用属性：{", ".join(all_keys)}\n请输入要删除的属性名称：', parent=dlg)
            if not col_name or col_name.strip() not in all_keys:
                return
            col_name = col_name.strip()
            all_keys.remove(col_name)
            # Remove from all documents
            for d in docs:
                attrs = getattr(d, 'attributes', None)
                if attrs and col_name in attrs:
                    del attrs[col_name]
            # Rebuild tree
            cols = ['文档名称'] + all_keys
            tree['columns'] = cols
            for col in cols:
                tree.heading(col, text=col)
                tree.column(col, width=120 if col != '文档名称' else 180)
            for item in tree.get_children():
                tree.delete(item)
            for doc in docs:
                attrs = getattr(doc, 'attributes', None) or {}
                vals = [doc.name] + [str(attrs.get(k, '')) for k in all_keys]
                tree.insert('', 'end', values=vals)
        tk.Button(op_frame, text='➕ 添加属性', font=FONT, bg=ACCENT, fg='white',
                 relief='flat', padx=8, command=add_attr).pack(side='left', padx=4)
        tk.Button(op_frame, text='✏️ 编辑单元格', font=FONT, relief='flat',
                 padx=8, command=edit_cell).pack(side='left', padx=4)
        tk.Button(op_frame, text='🗑 删除属性', font=FONT, relief='flat',
                 padx=8, fg='red', command=delete_attr).pack(side='left', padx=4)
        tk.Button(op_frame, text='关闭', font=FONT, relief='flat',
                 padx=8, command=dlg.destroy).pack(side='right', padx=4)

    # ── 查询构建器 ────────────────────────────────────────────

    def _query_builder(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        all_names = sorted(c.name for c in self._app.code_system.all_codes.values())
        if not all_names:
            messagebox.showinfo('提示', '尚无编码，请先创建编码')
            return
        dlg = tk.Toplevel(self.root)
        dlg.title('编码查询')
        dlg.geometry('700x560')
        dlg.transient(self.root)
        dlg.grab_set()
        # Include
        inc_frame = tk.LabelFrame(dlg, text='包含编码（OR - 可多选）', font=FONT_B, bg=BG, padx=6, pady=4)
        inc_frame.pack(fill='both', expand=True, padx=8, pady=(6, 2))
        inc_lb = tk.Listbox(inc_frame, font=FONT, selectmode='extended')
        inc_sc = ttk.Scrollbar(inc_frame, orient='vertical', command=inc_lb.yview)
        inc_lb.configure(yscrollcommand=inc_sc.set)
        inc_lb.pack(side='left', fill='both', expand=True)
        inc_sc.pack(side='right', fill='y')
        for n in all_names:
            inc_lb.insert('end', n)
        # Require all
        and_var = tk.BooleanVar(value=False)
        tk.Checkbutton(dlg, text='要求包含所有编码（AND）', variable=and_var,
                      font=FONT, bg=BG).pack(anchor='w', padx=12, pady=2)
        # Exclude
        exc_frame = tk.LabelFrame(dlg, text='排除编码（NOT - 可多选）', font=FONT_B, bg=BG, padx=6, pady=4)
        exc_frame.pack(fill='both', expand=True, padx=8, pady=2)
        exc_lb = tk.Listbox(exc_frame, font=FONT, selectmode='extended')
        exc_sc = ttk.Scrollbar(exc_frame, orient='vertical', command=exc_lb.yview)
        exc_lb.configure(yscrollcommand=exc_sc.set)
        exc_lb.pack(side='left', fill='both', expand=True)
        exc_sc.pack(side='right', fill='y')
        for n in all_names:
            exc_lb.insert('end', n)
        # Result tree
        res_frame = tk.LabelFrame(dlg, text='查询结果（双击浏览文档）', font=FONT_B, bg=BG, padx=6, pady=4)
        res_frame.pack(fill='both', expand=True, padx=8, pady=2)
        res_cols = ('文档ID', '文档名称', '匹配编码', '片段数')
        res_tree = ttk.Treeview(res_frame, columns=res_cols, show='headings', height=6)
        for col, w in zip(res_cols, (80, 180, 200, 70)):
            res_tree.heading(col, text=col)
            res_tree.column(col, width=w)
        res_tree.pack(fill='both', expand=True, side='left')
        res_sc = ttk.Scrollbar(res_frame, orient='vertical', command=res_tree.yview)
        res_tree.configure(yscrollcommand=res_sc.set)
        res_sc.pack(side='right', fill='y')
        def run_query():
            inc_sel = [inc_lb.get(i) for i in inc_lb.curselection()]
            exc_sel = [exc_lb.get(i) for i in exc_lb.curselection()]
            include = inc_sel if inc_sel else None
            exclude = exc_sel if exc_sel else None
            try:
                results = self._app.code_query(
                    include_codes=include, exclude_codes=exclude,
                    require_all=and_var.get())
                res_tree.delete(*res_tree.get_children())
                for r in results:
                    res_tree.insert('', 'end', values=(
                        r.get('doc_id', ''),
                        r.get('doc_name', '')[:30],
                        ', '.join(r.get('codes', [])),
                        r.get('segment_count', 0),
                    ))
                self._set_status(f'查询完成：{len(results)} 条结果')
            except Exception as e:
                messagebox.showerror('查询失败', str(e))
        def browse_doc():
            sel = res_tree.selection()
            if not sel:
                return
            vals = res_tree.item(sel[0], 'values')
            doc_id = vals[0]
            self._show_tab('docs')
            for child in self._doc_tree.get_children(''):
                if self._doc_tree.item(child, 'text') == doc_id:
                    self._doc_tree.selection_set(child)
                    self._doc_tree.see(child)
                    self._on_doc_select()
                    break
            dlg.destroy()
        res_tree.bind('<Double-Button-1>', lambda e: browse_doc())
        btn_frame = tk.Frame(dlg, bg=BG)
        btn_frame.pack(fill='x', padx=8, pady=4)
        tk.Button(btn_frame, text='▶ 运行查询', font=FONT_B, bg=ACCENT, fg='white',
                 relief='flat', padx=14, command=run_query).pack(side='left', padx=4)
        tk.Button(btn_frame, text='关闭', font=FONT, relief='flat',
                 padx=10, command=dlg.destroy).pack(side='right', padx=4)

    # ── 欢迎对话框 ──────────────────────────────────────────

    def _show_welcome_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title('QualCoder Pro')
        dlg.geometry('520x340')
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)
        dlg.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 520) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 340) // 2
        dlg.geometry(f'+{x}+{y}')

        tk.Label(dlg, text='QualCoder Pro', font=('Segoe UI', 22, 'bold'),
                fg=ACCENT).pack(pady=(28, 4))
        tk.Label(dlg, text='通用质性分析工具 · 学术研究版',
                font=FONT, fg='#555').pack()
        tk.Label(dlg, text='参考 NVivo · ATLAS.ti · MAXQDA 设计理念',
                font=('Segoe UI', 8), fg='#aaa').pack(pady=(0, 16))

        for label_text, icon_cmd in [
            ('🆕  新建项目',           self._new_project),
            ('📂  打开现有项目',       self._open_project),
            ('📁  打开最近',           self._open_recent_dialog),
        ]:
            tk.Button(dlg, text=label_text, font=FONT_B, width=22,
                     relief='groove', pady=6, cursor='hand2',
                     command=lambda c=icon_cmd: (dlg.destroy(), c())
                     ).pack(pady=5)

        tk.Label(dlg, text='支持文本/CSV/Excel/PDF 导入 · 支持扎根理论 · 支持 Cohen\'s Kappa',
                font=('Segoe UI', 7), fg='#bbb').pack(side='bottom', pady=10)

    # ── 撤销/重做 ──────────────────────────────────────────

    def _push_undo(self):
        """将当前编码系统快照压入撤销栈"""
        if not self._app or not self._app.code_system:
            return
        self._undo_stack.append({
            'codes': self._app.code_system.to_dict(),
        })
        self._redo_stack.clear()
        if len(self._undo_stack) > 50:
            self._undo_stack.pop(0)

    def _undo_coding(self):
        if not self._undo_stack:
            return
        if not self._app:
            return
        from hotel_analyzer.coding_browser import CodeSystem
        self._redo_stack.append({'codes': self._app.code_system.to_dict()})
        snap = self._undo_stack.pop()
        self._app.code_system = CodeSystem.from_dict(snap['codes'])
        self._refresh_code_tree()
        self._refresh_coded_segments()
        self._set_status('已撤销（Ctrl+Z）')

    def _redo_coding(self):
        if not self._redo_stack:
            return
        if not self._app:
            return
        from hotel_analyzer.coding_browser import CodeSystem
        self._undo_stack.append({'codes': self._app.code_system.to_dict()})
        snap = self._redo_stack.pop()
        self._app.code_system = CodeSystem.from_dict(snap['codes'])
        self._refresh_code_tree()
        self._refresh_coded_segments()
        self._set_status('已重做（Ctrl+Y）')

    # ══════════════════════════════════════════════════════════
    # 质性研究质量 Tab
    # ══════════════════════════════════════════════════════════

    def _build_quality_tab(self):
        f = tk.Frame(self._content, bg=BG)
        self._tab_pages['quality'] = f

        tb = tk.Frame(f, bg=TOOLBAR_BG, height=44)
        tb.pack(fill='x')
        tk.Label(tb, text='质性研究质量评估', font=FONT_B,
                bg=TOOLBAR_BG).pack(side='left', padx=8)
        tk.Button(tb, text='🔄 刷新', font=FONT, relief='flat',
                 padx=8, cursor='hand2',
                 command=self._refresh_quality_tab).pack(side='right', padx=8, pady=4)

        paned = ttk.PanedWindow(f, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=4, pady=4)

        # 左：四准则面板
        left_f = tk.Frame(paned, bg=BG)
        paned.add(left_f, weight=1)

        def criteria_card(parent, title, metrics: list, accent='#1565c0'):
            """单个准则卡片"""
            card = tk.LabelFrame(parent, text=title, font=FONT_B,
                               bg=BG, padx=8, pady=4)
            card.pack(fill='x', padx=4, pady=4)
            for label, value in metrics:
                row = tk.Frame(card, bg=BG)
                row.pack(fill='x')
                tk.Label(row, text=f'{label}：', font=FONT, bg=BG,
                        fg='#555', width=14, anchor='e').pack(side='left')
                tk.Label(row, text=str(value), font=FONT_B, bg=BG,
                        fg=accent).pack(side='left', padx=(4, 0))

        # 可信度
        self._cred_card_frame = tk.LabelFrame(left_f, text='可信度 Credibility',
                                              font=FONT_B, bg=BG, padx=8, pady=4)
        self._cred_card_frame.pack(fill='x', padx=4, pady=4)

        # 可迁移性
        self._trans_card_frame = tk.LabelFrame(left_f, text='可迁移性 Transferability',
                                               font=FONT_B, bg=BG, padx=8, pady=4)
        self._trans_card_frame.pack(fill='x', padx=4, pady=4)

        # 可靠性
        self._dep_card_frame = tk.LabelFrame(left_f, text='可靠性 Dependability',
                                              font=FONT_B, bg=BG, padx=8, pady=4)
        self._dep_card_frame.pack(fill='x', padx=4, pady=4)

        # 可确认性
        self._conf_card_frame = tk.LabelFrame(left_f, text='可确认性 Confirmability',
                                               font=FONT_B, bg=BG, padx=8, pady=4)
        self._conf_card_frame.pack(fill='x', padx=4, pady=4)

        # 饱和度按钮
        sat_btn = tk.Button(left_f, text='📈 编码饱和度图表',
                           font=FONT, bg=ACCENT, fg='white',
                           relief='flat', padx=10, cursor='hand2',
                           command=self._show_saturation_chart)
        sat_btn.pack(fill='x', padx=4, pady=(8, 4))

        # 右：编码密度表
        right_f = ttk.Labelframe(paned, text='编码密度（每文档）', padding=4)
        paned.add(right_f, weight=2)
        cols = ('文档', '编码数', '实例数', '密度')
        self._density_tree = ttk.Treeview(right_f, columns=cols,
                                          show='headings', height=28)
        for col, w in zip(cols, (220, 70, 70, 70)):
            self._density_tree.heading(col, text=col)
            self._density_tree.column(col, width=w)
        ds = ttk.Scrollbar(right_f, orient='vertical',
                           command=self._density_tree.yview)
        self._density_tree.configure(yscrollcommand=ds.set)
        self._density_tree.pack(side='left', fill='both', expand=True)
        ds.pack(side='right', fill='y')

    def _refresh_quality_tab(self):
        if not hasattr(self, '_density_tree') or not self._density_tree.winfo_exists():
            return
        if not self._app:
            return

        app = self._app
        density_df = app.get_coding_density()
        uncoded = app.get_uncoded_documents()
        memo_sum = app.memo_manager.summary()
        audit = app.get_audit_trail()

        # 填充密度表
        self._density_tree.delete(*self._density_tree.get_children())
        for _, row in density_df.iterrows():
            self._density_tree.insert('', 'end', values=(
                row.get('doc_name', '')[:30],
                int(row.get('total_codes', 0)),
                int(row.get('total_instances', 0)),
                f"{row.get('density', 0):.4f}",
            ))

        # 可信度
        self._update_criteria_card(self._cred_card_frame, [
            ('未编码文档', len(uncoded)),
            ('总备忘录', memo_sum.get('总备忘录数', 0)),
            ('编码实例', sum(len(c.instances) for c in app.code_system.all_codes.values())),
        ])

        # 可迁移性
        n_docs = len(app.documents)
        avg_codes = density_df['total_codes'].mean() if n_docs else 0
        self._update_criteria_card(self._trans_card_frame, [
            ('文档总数', n_docs),
            ('文档覆盖率', f"{(n_docs - len(uncoded)) / max(n_docs, 1) * 100:.1f}%"),
            ('平均编码/文档', f'{avg_codes:.1f}'),
        ])

        # 可靠性
        self._update_criteria_card(self._dep_card_frame, [
            ('代码总数', len(app.code_system.all_codes)),
            ('审计追踪条目', len(audit)),
            ('顶级代码', len(app.code_system.root_codes)),
        ])

        # 可确认性
        kappa_note = '尚未导入'
        if hasattr(self, '_last_kappa') and self._last_kappa is not None:
            kappa_note = f'κ = {self._last_kappa:.4f}'
        self._update_criteria_card(self._conf_card_frame, [
            ('Kappa系数', kappa_note),
            ('研究时间', '见备忘录'),
            ('方法论', '见备忘-方法论'),
        ], accent='#6a1b9a')

    def _update_criteria_card(self, frame, metrics: list, accent='#1565c0'):
        for w in frame.pack_slaves():
            w.destroy()
        for label, value in metrics:
            row = tk.Frame(frame, bg=BG)
            row.pack(fill='x')
            tk.Label(row, text=f'{label}：', font=FONT, bg=BG,
                    fg='#555', width=14, anchor='e').pack(side='left')
            tk.Label(row, text=str(value), font=FONT_B, bg=BG,
                    fg=accent).pack(side='left', padx=(4, 0))

    def _show_saturation_chart(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        try:
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt
            from matplotlib import font_manager as fm

            for fp in ['/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
                       '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc']:
                if Path(fp).exists():
                    prop = fm.FontProperties(fname=fp)
                    plt.rcParams['font.family'] = prop.get_name()
                    break

            df = self._app.get_coding_saturation()
            if df.empty:
                messagebox.showinfo('提示', '尚无编码数据')
                return

            fig, ax = plt.subplots(figsize=(12, max(6, len(df) * 0.4)), dpi=120)
            labels = [str(n)[:20] for n in df['doc_name']]
            x = range(len(df))
            ax.barh(x, df['new_codes'], color='#1e88e5', alpha=0.8, label='新编码数')
            ax.barh(x, df['total_codes'], color='#90caf9', alpha=0.5, label='累计编码数')
            ax.set_yticks(x)
            ax.set_yticklabels(labels, fontsize=8)
            ax.set_xlabel('编码数量')
            ax.set_title('编码饱和度（按文档）', size=13)
            ax.legend()
            ax.grid(axis='x', alpha=0.3)
            plt.tight_layout()

            path = self._get_chart_output_path('saturation')
            fig.savefig(path, dpi=120, bbox_inches='tight', facecolor='white')
            plt.close(fig)
            self._open_chart_window('编码饱和度图表', path)
            self._set_status('饱和度图表已生成')
        except Exception as e:
            messagebox.showerror('图表生成失败', str(e))

    # ══════════════════════════════════════════════════════════
    # 编码者一致性 Tab
    # ══════════════════════════════════════════════════════════

    def _build_reliability_tab(self):
        f = tk.Frame(self._content, bg=BG)
        self._tab_pages['reliability'] = f

        tb = tk.Frame(f, bg=TOOLBAR_BG, height=44)
        tb.pack(fill='x')
        tk.Label(tb, text='编码者一致性检验', font=FONT_B,
                bg=TOOLBAR_BG).pack(side='left', padx=8)
        tk.Button(tb, text='📂 导入 CSV', font=FONT,
                 bg=ACCENT, fg='white', relief='flat',
                 padx=10, cursor='hand2',
                 command=self._import_reliability_csv
                 ).pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='📊 计算 Kappa', font=FONT,
                 relief='flat', padx=8, cursor='hand2',
                 command=self._compute_kappa
                 ).pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='📋 导出报告', font=FONT,
                 relief='flat', padx=8, cursor='hand2',
                 command=self._export_reliability_report
                 ).pack(side='right', padx=8, pady=4)

        # CSV 预览
        preview_lf = ttk.Labelframe(f, text='第二位编码者数据（CSV 预览）', padding=4)
        preview_lf.pack(fill='x', padx=4, pady=(4, 0))
        self._rel_csv_path = tk.StringVar(value='（尚未导入 CSV）')
        tk.Label(preview_lf, textvariable=self._rel_csv_path,
                font=('Segoe UI', 8), fg='#555', bg=BG).pack(side='left', padx=4)
        self._rel_preview_tree = ttk.Treeview(preview_lf, columns=('doc_id', 'code'),
                                               show='headings', height=6)
        self._rel_preview_tree.heading('doc_id', text='doc_id')
        self._rel_preview_tree.heading('code', text='code_name')
        self._rel_preview_tree.column('doc_id', width=150)
        self._rel_preview_tree.column('code', width=200)
        self._rel_preview_tree.pack(fill='x', padx=4, pady=4)
        sp = ttk.Scrollbar(preview_lf, orient='vertical',
                           command=self._rel_preview_tree.yview)
        self._rel_preview_tree.configure(yscrollcommand=sp.set)
        self._rel_preview_tree.pack(side='left', fill='x', expand=True)
        sp.pack(side='right', fill='y')

        # Kappa 结果区
        result_lf = tk.Frame(f, bg=BG)
        result_lf.pack(fill='x', padx=4, pady=4)

        self._kappa_label = tk.Label(result_lf, text='κ = —',
                                    font=('Segoe UI', 28, 'bold'), bg=BG, fg='#555')
        self._kappa_label.pack(pady=(8, 0))

        self._agreement_label = tk.Label(result_lf, text='一致率：—',
                                        font=FONT, bg=BG, fg='#666')
        self._agreement_label.pack()

        confusion_frame = tk.Frame(result_lf, bg=BG)
        confusion_frame.pack(pady=6)
        self._confusion_labels = {}
        for label_text, key in [('TP（真阳性）', 'tp'), ('TN（真阴性）', 'tn'),
                                  ('FP（假阳性）', 'fp'), ('FN（假阴性）', 'fn')]:
            lbl = tk.Label(confusion_frame, text=f'{label_text}: —',
                          font=FONT, bg='#f5f5f5', relief='groove',
                          padx=10, pady=2, width=16)
            lbl.pack(side='left', padx=4)
            self._confusion_labels[key] = lbl

        # 一致性矩阵
        matrix_lf = ttk.Labelframe(f, text='编码一致性矩阵（双编码者）', padding=4)
        matrix_lf.pack(fill='both', expand=True, padx=4, pady=(0, 4))
        self._reliability_matrix_tree = ttk.Treeview(matrix_lf,
                                                      columns=('文档', '本软件编码', '第二位编码者'),
                                                      show='headings', height=22)
        for col, w in zip(('文档', '本软件编码', '第二位编码者'), (200, 200, 200)):
            self._reliability_matrix_tree.heading(col, text=col)
            self._reliability_matrix_tree.column(col, width=w)
        rm_s = ttk.Scrollbar(matrix_lf, orient='vertical',
                             command=self._reliability_matrix_tree.yview)
        self._reliability_matrix_tree.configure(yscrollcommand=rm_s.set)
        self._reliability_matrix_tree.pack(side='left', fill='both', expand=True)
        rm_s.pack(side='right', fill='y')

        self._other_codes_data: Dict[str, List[str]] = {}

    def _refresh_reliability_tab(self):
        pass  # 手动刷新，不需要自动刷新

    def _import_reliability_csv(self):
        path = filedialog.askopenfilename(
            title='导入第二位编码者的编码结果',
            filetypes=[('CSV', '*.csv'), ('全部', '*.*')])
        if not path:
            return
        try:
            import pandas as pd
            df = pd.read_csv(path, encoding='utf-8-sig')
            self._rel_csv_path.set(f'文件：{path}')
            self._other_codes_data = self._app.load_intercoder_codes(path)
            # 预览前10行
            self._rel_preview_tree.delete(*self._rel_preview_tree.get_children())
            for _, row in df.head(10).iterrows():
                self._rel_preview_tree.insert('', 'end', values=(
                    str(row.get('doc_id', '')), str(row.get('code_name', ''))))
            self._set_status(f'已导入 {len(df)} 条编码记录')
        except Exception as e:
            messagebox.showerror('导入错误', str(e))

    def _compute_kappa(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        if not self._other_codes_data:
            messagebox.showinfo('提示', '请先导入第二位编码者的 CSV')
            return
        try:
            result = self._app.intercoder_reliability(self._other_codes_data)
            kappa = result['kappa']
            self._last_kappa = kappa
            agreement_pct = result['agreement_pct']
            ct = result['confusion_table']

            # Kappa 颜色
            if kappa < 0:
                color = '#c62828'
            elif kappa < 0.4:
                color = '#ef6c00'
            elif kappa < 0.7:
                color = '#2e7d32'
            else:
                color = '#1b5e20'

            self._kappa_label.config(text=f'κ = {kappa:.4f}', fg=color)
            self._agreement_label.config(text=f'一致率：{agreement_pct:.2f}%')
            self._confusion_labels['tp'].config(text=f'TP: {ct.get("true_positive", 0)}')
            self._confusion_labels['tn'].config(text=f'TN: {ct.get("true_negative", 0)}')
            self._confusion_labels['fp'].config(text=f'FP: {ct.get("false_positive", 0)}')
            self._confusion_labels['fn'].config(text=f'FN: {ct.get("false_negative", 0)}')

            # 矩阵
            self._reliability_matrix_tree.delete(
                *self._reliability_matrix_tree.get_children())
            my_codes: Dict[str, set] = {}
            for code in self._app.code_system.all_codes.values():
                for inst in code.instances:
                    did = inst.get('doc_id')
                    if did:
                        my_codes.setdefault(did, set()).add(code.name)

            all_docs = sorted(set(list(my_codes.keys()) +
                                  list(self._other_codes_data.keys())))
            for did in all_docs:
                my = ', '.join(sorted(my_codes.get(did, set())))
                other = ', '.join(sorted(self._other_codes_data.get(did, [])))
                self._reliability_matrix_tree.insert('', 'end', values=(
                    did[:25], my[:30] or '—', other[:30] or '—'))

            self._set_status(
                f"Cohen's Kappa: {kappa:.4f} | 一致率: {agreement_pct:.2f}%")
        except Exception as e:
            messagebox.showerror('计算错误', str(e))

    def _export_reliability_report(self):
        if not hasattr(self, '_last_kappa') or self._last_kappa is None:
            messagebox.showinfo('提示', '请先计算 Kappa')
            return
        path = filedialog.asksaveasfilename(
            title='导出一致性报告', defaultextension='.csv',
            filetypes=[('CSV', '*.csv'), ('全部', '*.*')])
        if not path:
            return
        try:
            import csv
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                w = csv.writer(f)
                w.writerow(['Cohen\'s Kappa 系数', self._last_kappa])
                w.writerow(['一致率(%)', self._app.intercoder_reliability(
                    self._other_codes_data)['agreement_pct']])
            messagebox.showinfo('完成', f'已导出：\n{path}')
        except Exception as e:
            messagebox.showerror('导出失败', str(e))

    # ══════════════════════════════════════════════════════════
    # 词频分析 Tab
    # ══════════════════════════════════════════════════════════

    def _build_wordfreq_tab(self):
        f = tk.Frame(self._content, bg=BG)
        self._tab_pages['wordfreq'] = f

        # ── 工具栏 ──
        tb = tk.Frame(f, bg=TOOLBAR_BG, height=44)
        tb.pack(fill='x')
        tk.Label(tb, text='词频与 TF-IDF 分析', font=FONT_B,
                bg=TOOLBAR_BG).pack(side='left', padx=8)
        tk.Button(tb, text='▶ 执行词频分析', font=FONT_B,
                 bg=ACCENT, fg='white', relief='flat',
                 padx=12, cursor='hand2',
                 command=self._run_wordfreq_analysis
                 ).pack(side='left', padx=4, pady=4)
        tk.Button(tb, text='🔄 刷新', font=FONT,
                 relief='flat', padx=8, cursor='hand2',
                 command=self._refresh_wordfreq_tab
                 ).pack(side='left', padx=4, pady=4)

        tk.Button(tb, text='📊 导出词频表', font=FONT,
                 relief='flat', padx=8, cursor='hand2',
                 command=self._export_tfidf_csv
                 ).pack(side='right', padx=4, pady=4)
        tk.Button(tb, text='📋 导出词-文档矩阵', font=FONT,
                 relief='flat', padx=8, cursor='hand2',
                 command=self._export_word_doc_matrix
                 ).pack(side='right', padx=4, pady=4)
        tk.Button(tb, text='☁️ 导出词云图', font=FONT,
                 relief='flat', padx=8, cursor='hand2',
                 command=self._save_wordcloud_image
                 ).pack(side='right', padx=4, pady=4)

        # ── 控制行 ──
        ctrl = tk.Frame(f, bg=BG)
        ctrl.pack(fill='x', padx=4, pady=(2, 0))

        tk.Label(ctrl, text='分析类型：', font=FONT, bg=BG).pack(side='left', padx=(0, 4))
        self._wf_mode_var = tk.StringVar(value='raw')
        for label, val in [('原始词频', 'raw'), ('TF-IDF', 'tfidf')]:
            tk.Radiobutton(ctrl, text=label, variable=self._wf_mode_var,
                          value=val, font=FONT, bg=BG,
                          command=self._refresh_wordfreq_table
                          ).pack(side='left', padx=4)

        tk.Label(ctrl, text='N-gram：', font=FONT, bg=BG).pack(side='left', padx=(16, 4))
        self._ngram_var = tk.StringVar(value='1')
        for label, val in [('单词', '1'), ('二元组', '2'), ('三元组', '3')]:
            tk.Radiobutton(ctrl, text=label, variable=self._ngram_var,
                          value=val, font=FONT, bg=BG,
                          command=self._run_wordfreq_analysis
                          ).pack(side='left', padx=4)

        tk.Label(ctrl, text='词汇量上限：', font=FONT, bg=BG).pack(side='left', padx=(16, 4))
        self._max_feat_var = tk.IntVar(value=500)
        tk.Spinbox(ctrl, from_=10, to=5000, textvariable=self._max_feat_var,
                  width=6, font=FONT).pack(side='left', padx=4)

        tk.Label(ctrl, text='停用词文件：', font=FONT, bg=BG).pack(side='left', padx=(16, 4))
        self._stopword_var = tk.StringVar()
        tk.Entry(ctrl, textvariable=self._stopword_var,
                font=FONT, width=16).pack(side='left', padx=4)
        tk.Button(ctrl, text='浏览', font=FONT, relief='flat',
                 padx=6, cursor='hand2',
                 command=self._browse_stopwords).pack(side='left', padx=4)

        # ── 主区域：左右分栏 ──
        paned = ttk.PanedWindow(f, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=4, pady=4)

        # ── 左：词项统计表 ──
        left_f = ttk.Labelframe(paned, text='词项统计表', padding=4)
        paned.add(left_f, weight=3)

        # 顶部工具条（每列统计）
        stats_bar = tk.Frame(left_f, bg='#e8eaf6', height=28)
        stats_bar.pack(fill='x', pady=(0, 4))
        self._wf_stats_var = tk.StringVar(value='未加载')
        tk.Label(stats_bar, textvariable=self._wf_stats_var,
                font=('Segoe UI', 9), bg='#e8eaf6', fg='#333'
                ).pack(side='left', padx=8)

        # 表头（可点击排序）
        hdr = tk.Frame(left_f, bg='#f5f5f5')
        hdr.pack(fill='x')

        self._wf_sort_col = tk.StringVar(value='词频')
        self._wf_sort_dir = tk.StringVar(value='desc')  # asc / desc

        hdr_cols = [
            ('#',          40),
            ('词项',      180),
            ('词频',       90),
            ('频率%',      75),
            ('TF-IDF最高', 100),
            ('文档数',     70),
            ('字符数',     70),
        ]
        self._wf_hdr_labels = {}
        for col_text, col_w in hdr_cols:
            lbl = tk.Label(hdr, text=col_text, font=FONT_B,
                          bg='#f5f5f5', fg='#1565c0',
                          relief='flat', padx=6, pady=3,
                          cursor='hand2')
            lbl.pack(side='left')
            lbl.bind('<Button-1>', lambda e, c=col_text: self._wf_sort_by_column(c))
            if col_text in ('词频', 'TF-IDF最高', '文档数', '字符数'):
                arrow = ' ▼' if self._wf_sort_col.get() == col_text else ''
                lbl.config(text=col_text + arrow)
                self._wf_hdr_labels[col_text] = lbl
            if col_text != '#':
                sep = tk.Frame(hdr, bg='#ddd', width=1)
                sep.pack(side='left', fill='y')

        # 表格主体（Canvas + Scrollbar，替代 Treeview）
        table_outer = tk.Frame(left_f)
        table_outer.pack(fill='both', expand=True)

        wf_vscroll = tk.Scrollbar(table_outer, orient='vertical')
        wf_vscroll.pack(side='right', fill='y')
        wf_hscroll = tk.Scrollbar(table_outer, orient='horizontal')
        wf_hscroll.pack(side='bottom', fill='x')

        canvas_w, canvas_h = 620, 520
        self._wf_canvas = tk.Canvas(table_outer,
                                   bg='white',
                                   yscrollcommand=wf_vscroll.set,
                                   xscrollcommand=wf_hscroll.set,
                                   highlightthickness=0,
                                   height=canvas_h)
        self._wf_canvas.pack(side='left', fill='both', expand=True)
        wf_vscroll.config(command=self._wf_canvas.yview)
        wf_hscroll.config(command=self._wf_canvas.xview)

        self._wf_table_frame = tk.Frame(self._wf_canvas, bg='white')
        self._wf_canvas.create_window((0, 0), window=self._wf_table_frame,
                                       anchor='nw')
        self._wf_table_frame.bind('<Configure>',
            lambda e: self._wf_canvas.configure(
                scrollregion=self._wf_canvas.bbox('all')))

        # 绑定鼠标滚轮
        self._wf_canvas.bind_all('<MouseWheel>',
            lambda e: self._wf_canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')
            if e.delta else None)
        self._wf_canvas.bind('<Enter>',
            lambda e: self._wf_canvas.bind_all('<MouseWheel>',
                lambda ev: self._wf_canvas.yview_scroll(int(-1 * (ev.delta / 120)), 'units')))
        self._wf_canvas.bind('<Leave>',
            lambda e: self._wf_canvas.unbind_all('<MouseWheel>'))

        self._wf_rows = []          # 当前显示的所有行widget
        self._wf_all_data = []      # 所有词项数据 [(term, count, pct, tfidf_max, doc_freq, char_len), ...]
        self._wf_raw_freq = {}      # {term: count}
        self._wf_tfidf_df = None    # TF-IDF DataFrame

        # ── 右：图表按钮面板 ──
        right_f = ttk.Labelframe(paned, text='可视化与导出', padding=8)
        paned.add(right_f, weight=1)

        btn_cfg = dict(font=FONT, relief='groove',
                      cursor='hand2', wraplength=120)

        chart_btns = [
            ('📈 词频柱状图（Top 30）',   self._chart_word_frequency),
            ('📈 词频柱状图（Top 50）',   lambda: self._chart_word_frequency(top_n=50)),
            ('📈 TF-IDF 柱状图',         self._chart_tfidf_bars),
            ('☁️ 词云图（蓝色）',         lambda: self._show_wordcloud(colormap='Blues')),
            ('☁️ 词云图（彩虹）',         lambda: self._show_wordcloud(colormap='rainbow')),
            ('☁️ 词云图（绿色）',         lambda: self._show_wordcloud(colormap='Greens')),
            ('📊 N-gram 词组分布',       self._chart_ngram_distribution),
            ('📉 文档覆盖率分布',         self._chart_doc_coverage),
        ]

        for i, (label, cmd) in enumerate(chart_btns):
            bt = tk.Button(right_f, text=label, command=cmd, **btn_cfg)
            bt.pack(fill='x', padx=4, pady=3)
            if i == 3:  # after first wordcloud button
                tk.Label(right_f, text='配色方案', font=('Segoe UI', 8),
                        fg='#888').pack(pady=(4, 1))
            elif i == 6:  # before ngram
                sep = ttk.Separator(right_f, orient='horizontal')
                sep.pack(fill='x', padx=4, pady=6)
                tk.Label(right_f, text='分布分析', font=FONT_B,
                        fg=ACCENT).pack(pady=(0, 4))

        sep2 = ttk.Separator(right_f, orient='horizontal')
        sep2.pack(fill='x', padx=4, pady=6)
        tk.Label(right_f, text='导出数据', font=FONT_B,
                fg=ACCENT).pack(pady=(0, 4))

        export_btns = [
            ('📥 导出词频表（全部）',  self._export_tfidf_csv),
            ('📥 导出词-文档矩阵',   self._export_word_doc_matrix),
            ('🖼️  保存词云为图片',   self._save_wordcloud_image),
            ('📄  保存柱状图为 PDF',  lambda: self._save_bar_chart_pdf(top_n=30)),
        ]
        for label, cmd in export_btns:
            tk.Button(right_f, text=label, font=FONT,
                     relief='groove', cursor='hand2',
                     command=cmd
                     ).pack(fill='x', padx=4, pady=3)

        # 底部提示
        tip = tk.Label(f,
                       text='提示：点击列标题可排序 | 词频分析基于 jieba 分词',
                       font=('Segoe UI', 8), fg='#888', bg=BG)
        tip.pack(fill='x', padx=8, pady=(0, 2))

    def _refresh_wordfreq_tab(self):
        """刷新词频分析结果"""
        if self._wf_all_data:
            self._refresh_wordfreq_table()

    def _browse_stopwords(self):
        path = filedialog.askopenfilename(
            title='选择停用词文件（UTF-8 编码，一行一个词）',
            filetypes=[('文本文件', '*.txt'), ('全部', '*.*')])
        if path:
            self._stopword_var.set(path)

    def _load_custom_stopwords(self):
        sw_path = self._stopword_var.get()
        if not sw_path or not Path(sw_path).exists():
            return None
        try:
            with open(sw_path, 'r', encoding='utf-8') as f:
                return set(line.strip() for line in f if line.strip())
        except Exception:
            return None

    # ── 词频分析核心 ──────────────────────────────────────────

    def _run_wordfreq_analysis(self):
        """执行完整词频分析（原始词频 + TF-IDF）"""
        if not self._app or not self._app.documents:
            messagebox.showinfo('提示', '请先在【文档】标签页加载文档')
            return

        try:
            self._set_status('正在分词与统计词频…')
            self.root.update()

            from hotel_analyzer.data_processor import compute_tfidf, compute_word_doc_matrix
            import jieba
            jieba.setLogLevel(jieba.logging.INFO)
            from collections import Counter
            import numpy as np

            sw = self._load_custom_stopwords()
            docs = list(self._app.documents)
            ngram_val = int(self._ngram_var.get())
            max_f = self._max_feat_var.get()

            # ── 1. 原始词频统计 ──
            word_counter = Counter()
            for doc in docs:
                words = [w for w in jieba.cut(doc.text)
                        if w.strip() and (sw is None or w not in sw) and len(w) > 1]
                word_counter.update(words)

            all_words = list(self._app.documents)
            total_chars = sum(len(d.text) for d in docs)
            total_words = sum(word_counter.values())

            # ── 2. TF-IDF 矩阵 ──
            ngram_range = (1, ngram_val)
            tfidf_df = compute_tfidf(docs, stopwords=sw,
                                    max_features=max_f,
                                    ngram_range=ngram_range)
            self._wf_tfidf_df = tfidf_df

            # ── 3. 合成完整词项表 ──
            # 词汇来源：TF-IDF 词表（已按 n-gram 扩展）
            tfidf_terms = set(tfidf_df.index)

            # 如果词频上限比 TF-IDF 结果多，从词频计数器补足
            extra_needed = max_f - len(tfidf_terms)
            if extra_needed > 0:
                remaining = [(w, c) for w, c in word_counter.most_common(max_f * 2)
                            if w not in tfidf_terms][:extra_needed]
                for w, c in remaining:
                    tfidf_terms.add(w)

            wf_data = []
            for term in tfidf_terms:
                raw_count = word_counter.get(term, 0)
                if raw_count == 0 and term not in word_counter:
                    raw_count = sum(1 for d in docs if term in d.text)

                tfidf_max = float(tfidf_df.loc[term].max()) if term in tfidf_df.index else 0.0
                doc_freq = int((tfidf_df.loc[term] > 0).sum()) if term in tfidf_df.index else 0

                pct = (raw_count / total_words * 100) if total_words > 0 else 0.0
                char_len = len(term)

                wf_data.append((term, raw_count, pct, tfidf_max, doc_freq, char_len))

            # 降序排列
            wf_data.sort(key=lambda x: x[1], reverse=True)
            self._wf_all_data = wf_data
            self._wf_raw_freq = word_counter

            # 统计摘要
            self._wf_stats_var.set(
                f'词项总数：{len(wf_data):,}  |  总词次：{total_words:,}  |  '
                f'文档数：{len(docs)}  |  平均词长：{total_chars/max(total_words,1):.1f}字符')

            self._refresh_wordfreq_table()
            self._set_status(f'词频分析完成：{len(wf_data):,} 个词项')

        except ImportError as e:
            messagebox.showerror('缺少依赖', str(e))
        except Exception as e:
            import traceback
            messagebox.showerror('分析失败', f'{e}\n{traceback.format_exc()}')

    def _refresh_wordfreq_table(self):
        """按当前排序刷新表格显示"""
        if not self._wf_all_data:
            return

        mode = self._wf_mode_var.get()
        sort_col = self._wf_sort_col.get()
        sort_dir = self._wf_sort_dir.get()

        data = list(self._wf_all_data)

        # 根据模式排序
        col_idx = {
            '词项': 0, '词频': 1, '频率%': 2,
            'TF-IDF最高': 3, '文档数': 4, '字符数': 5
        }.get(sort_col, 1)
        data.sort(key=lambda x: x[col_idx], reverse=(sort_dir == 'desc'))

        # 筛选模式
        if mode == 'tfidf':
            data = [(t, c, p, tf, df, cl) for t, c, p, tf, df, cl in data if tf > 0]

        self._render_wf_table(data)

    def _wf_sort_by_column(self, col_name: str):
        """点击列标题切换排序"""
        if self._wf_sort_col.get() == col_name:
            self._wf_sort_dir.set('desc' if self._wf_sort_dir.get() == 'asc' else 'asc')
        else:
            self._wf_sort_col.set(col_name)
            self._wf_sort_dir.set('desc')
        # 更新标题箭头
        for col, lbl in self._wf_hdr_labels.items():
            arrow = ' ▼' if col == col_name and self._wf_sort_dir.get() == 'desc' else \
                    ' ▲' if col == col_name and self._wf_sort_dir.get() == 'asc' else ''
            lbl.config(text=col + arrow)
        self._refresh_wordfreq_table()

    def _render_wf_table(self, data):
        """将数据渲染到 Canvas 表格"""
        for w in self._wf_rows:
            w.destroy()
        self._wf_rows = []

        # 列宽配置
        col_widths = [40, 180, 90, 75, 100, 70, 70]
        hdr_height = 28
        row_height = 22
        ROW_BG_ODD = '#f8f9fa'
        ROW_BG_EVN = '#ffffff'

        # 更新 canvas 滚动区域（宽高都设置）
        canvas_w = sum(col_widths)
        canvas_h = (len(data) + 1) * row_height + hdr_height
        self._wf_canvas.configure(scrollregion=(0, 0, canvas_w, canvas_h))

        # 设置内部 frame 尺寸，使 Canvas 滚动生效
        self._wf_table_frame.configure(width=canvas_w, height=canvas_h)

        # 绘制表头
        hdr_f = self._wf_table_frame
        x_positions = []
        x = 0
        for cw in col_widths:
            x_positions.append(x)
            x += cw

        hdr_bg = '#e8eaf6'
        hdr_sep = '#c5cae9'
        for col_i, (xpos, cw) in enumerate(zip(x_positions, col_widths)):
            lbl = tk.Label(hdr_f, text=['#', '词项', '词频', '频率%', 'TF-IDF最高', '文档数', '字符数'][col_i],
                          font=FONT_B, bg=hdr_bg, fg='#1565c0',
                          relief='flat', anchor='w',
                          width=cw // 8, height=1)
            lbl.place(x=xpos, y=0, width=cw, height=hdr_height)
            self._wf_rows.append(lbl)

        sep_y = hdr_height - 1
        sep = tk.Frame(hdr_f, bg=hdr_sep, height=1)
        sep.place(x=0, y=sep_y, width=canvas_w, height=1)
        self._wf_rows.append(sep)

        # 绘制数据行
        for row_i, row in enumerate(data):
            term, count, pct, tfidf_max, doc_freq, char_len = row
            y = hdr_height + row_i * row_height
            bg = ROW_BG_ODD if row_i % 2 == 0 else ROW_BG_EVN

            row_data = [str(row_i + 1), term,
                       f'{count:,}',
                       f'{pct:.3f}%',
                       f'{tfidf_max:.4f}',
                       str(doc_freq),
                       str(char_len)]

            for col_i, (xpos, cw, val) in enumerate(zip(x_positions, col_widths, row_data)):
                fg = '#1565c0' if col_i == 1 else '#333'
                font_obj = FONT if col_i != 1 else ('Segoe UI', 10)
                align = 'center' if col_i != 1 else 'w'

                lbl = tk.Label(hdr_f, text=val,
                              font=font_obj, bg=bg, fg=fg,
                              relief='flat', anchor=align,
                              height=1)
                lbl.place(x=xpos, y=y, width=cw, height=row_height)
                self._wf_rows.append(lbl)

                # 列分隔线
                if col_i > 0:
                    sep_v = tk.Frame(hdr_f, bg='#e0e0e0', width=1)
                    sep_v.place(x=xpos, y=y, width=1, height=row_height)
                    self._wf_rows.append(sep_v)

            # 行底部分隔线
            row_sep = tk.Frame(hdr_f, bg='#f0f0f0', height=1)
            row_sep.place(x=0, y=y + row_height - 1, width=canvas_w, height=1)
            self._wf_rows.append(row_sep)

        self._wf_stats_var.set(
            f'共 {len(data):,} 个词项  |  显示全部  |  '
            f'总词次：{sum(d[1] for d in data):,}')

    # ── 图表方法 ─────────────────────────────────────────────

    def _get_font_path(self):
        for fp in ['/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
                   '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc']:
            if Path(fp).exists():
                return fp
        return None

    def _chart_word_frequency(self, top_n=30):
        if not self._wf_all_data:
            messagebox.showinfo('提示', '请先执行词频分析')
            return
        try:
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt
            from matplotlib import font_manager as fm
            font_path = self._get_font_path()
            if font_path:
                plt.rcParams['font.family'] = fm.FontProperties(fname=font_path).get_name()
            plt.rcParams['axes.unicode_minus'] = False

            data = sorted(self._wf_all_data, key=lambda x: x[1], reverse=True)[:top_n]
            labels = [d[0] for d in data]
            values = [d[1] for d in data]

            fig, ax = plt.subplots(figsize=(11, max(6, top_n * 0.38)), dpi=120)
            bars = ax.barh(range(len(labels)), values, color='#1976d2', alpha=0.85)
            ax.set_yticks(range(len(labels)))
            ax.set_yticklabels(labels, fontsize=9)
            ax.invert_yaxis()
            ax.set_xlabel('词频（次）')
            ax.set_title(f'词频统计 Top {top_n}', size=14)
            for bar, val in zip(bars, values):
                ax.text(bar.get_width() + max(values) * 0.005,
                        bar.get_y() + bar.get_height() / 2,
                        f'{int(val):,}', va='center', fontsize=8)
            ax.grid(axis='x', alpha=0.3)
            plt.tight_layout()

            path = self._get_chart_output_path(f'wordfreq_top{top_n}')
            fig.savefig(path, dpi=120, bbox_inches='tight', facecolor='white')
            plt.close(fig)
            self._open_chart_window(f'词频统计 Top {top_n}', path)
            self._set_status(f'词频柱状图已生成（Top {top_n}）')
        except Exception as e:
            messagebox.showerror('图表生成失败', str(e))

    def _chart_tfidf_bars(self):
        if not self._wf_all_data:
            messagebox.showinfo('提示', '请先执行词频分析')
            return
        try:
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt
            from matplotlib import font_manager as fm
            font_path = self._get_font_path()
            if font_path:
                plt.rcParams['font.family'] = fm.FontProperties(fname=font_path).get_name()
            plt.rcParams['axes.unicode_minus'] = False

            top_n = 30
            data = sorted(self._wf_all_data, key=lambda x: x[3], reverse=True)[:top_n]
            labels = [d[0] for d in data]
            values = [d[3] for d in data]

            fig, ax = plt.subplots(figsize=(11, max(6, top_n * 0.38)), dpi=120)
            bars = ax.barh(range(len(labels)), values, color='#7b1fa2', alpha=0.85)
            ax.set_yticks(range(len(labels)))
            ax.set_yticklabels(labels, fontsize=9)
            ax.invert_yaxis()
            ax.set_xlabel('TF-IDF 最高值')
            ax.set_title(f'TF-IDF 词项 Top {top_n}', size=14)
            for bar, val in zip(bars, values):
                ax.text(bar.get_width() + max(values) * 0.005,
                        bar.get_y() + bar.get_height() / 2,
                        f'{val:.3f}', va='center', fontsize=8)
            ax.grid(axis='x', alpha=0.3)
            plt.tight_layout()

            path = self._get_chart_output_path(f'tfidf_top{top_n}')
            fig.savefig(path, dpi=120, bbox_inches='tight', facecolor='white')
            plt.close(fig)
            self._open_chart_window(f'TF-IDF 词项 Top {top_n}', path)
            self._set_status(f'TF-IDF 柱状图已生成')
        except Exception as e:
            messagebox.showerror('图表生成失败', str(e))

    def _show_wordcloud(self, colormap='Blues'):
        if not self._wf_raw_freq:
            messagebox.showinfo('提示', '请先执行词频分析')
            return
        try:
            from hotel_analyzer.visualizer import plot_wordcloud
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt

            font_path = self._get_font_path()
            freq = {term: count for term, count, *_ in self._wf_all_data if count > 0}

            chart = plot_wordcloud(freq, title='词云图',
                                  max_words=200,
                                  colormap=colormap,
                                  font_path=font_path)

            path = self._get_chart_output_path(f'wordcloud_{colormap}')
            chart.fig.savefig(path, dpi=150, bbox_inches='tight', facecolor='white')
            plt.close(chart.fig)

            self._open_chart_window('词云图', path)
            self._set_status(f'词云图已生成（配色：{colormap}）')
        except ImportError as e:
            messagebox.showerror('缺少依赖', str(e))
        except Exception as e:
            messagebox.showerror('词云生成失败', str(e))

    def _chart_ngram_distribution(self):
        if not self._app or not self._app.documents:
            messagebox.showinfo('提示', '请先加载文档')
            return
        try:
            from collections import Counter
            import jieba
            jieba.setLogLevel(jieba.logging.INFO)
            sw = self._load_custom_stopwords()
            ngram_val = int(self._ngram_var.get())

            docs = list(self._app.documents)
            ngram_counter = Counter()

            for doc in docs:
                if ngram_val == 1:
                    words = [w for w in jieba.cut(doc.text)
                            if w.strip() and (sw is None or w not in sw) and len(w) > 1]
                    for w in words:
                        ngram_counter[w] += 1
                elif ngram_val == 2:
                    words = [w for w in jieba.cut(doc.text)
                            if w.strip() and (sw is None or w not in sw) and len(w) > 1]
                    for i in range(len(words) - 1):
                        ng = f'{words[i]}_{words[i+1]}'
                        ngram_counter[ng] += 1
                else:
                    words = [w for w in jieba.cut(doc.text)
                            if w.strip() and (sw is None or w not in sw) and len(w) > 1]
                    for i in range(len(words) - 2):
                        ng = f'{words[i]}_{words[i+1]}_{words[i+2]}'
                        ngram_counter[ng] += 1

            from hotel_analyzer.visualizer import plot_ngram_distribution
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt

            chart = plot_ngram_distribution(dict(ngram_counter.most_common(50)),
                                           title=f'{ngram_val}-gram 词组分布')
            path = self._get_chart_output_path(f'{ngram_val}gram_dist')
            chart.fig.savefig(path, dpi=120, bbox_inches='tight', facecolor='white')
            plt.close(chart.fig)

            self._open_chart_window(f'{ngram_val}-gram 词组分布', path)
            self._set_status(f'{ngram_val}-gram 分布图已生成')
        except Exception as e:
            messagebox.showerror('图表生成失败', str(e))

    def _chart_doc_coverage(self):
        if self._wf_tfidf_df is None:
            messagebox.showinfo('提示', '请先执行词频分析')
            return
        try:
            from hotel_analyzer.visualizer import plot_document_coverage
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt

            chart = plot_document_coverage(self._wf_tfidf_df,
                                          title='词项文档覆盖率（Top 30）')
            path = self._get_chart_output_path('doc_coverage')
            chart.fig.savefig(path, dpi=120, bbox_inches='tight', facecolor='white')
            plt.close(chart.fig)

            self._open_chart_window('文档覆盖率分布', path)
            self._set_status('文档覆盖率图已生成')
        except Exception as e:
            messagebox.showerror('图表生成失败', str(e))

    def _save_wordcloud_image(self):
        if not self._wf_raw_freq:
            messagebox.showinfo('提示', '请先执行词频分析')
            return
        path = filedialog.asksaveasfilename(
            title='保存词云图片',
            defaultextension='.png',
            filetypes=[('PNG 图片', '*.png'), ('JPEG', '*.jpg'), ('全部', '*.*')])
        if not path:
            return
        try:
            from hotel_analyzer.visualizer import plot_wordcloud
            import matplotlib
            matplotlib.use('Agg')

            font_path = self._get_font_path()
            freq = {term: count for term, count, *_ in self._wf_all_data if count > 0}

            chart = plot_wordcloud(freq, title='词云图',
                                  max_words=200,
                                  colormap='Blues',
                                  font_path=font_path)
            chart.fig.savefig(path, dpi=150, bbox_inches='tight')
            plt.close(chart.fig)
            messagebox.showinfo('完成', f'词云图已保存：\n{path}')
        except Exception as e:
            messagebox.showerror('保存失败', str(e))

    def _save_bar_chart_pdf(self, top_n=30):
        if not self._wf_all_data:
            messagebox.showinfo('提示', '请先执行词频分析')
            return
        path = filedialog.asksaveasfilename(
            title='保存柱状图',
            defaultextension='.pdf',
            filetypes=[('PDF', '*.pdf'), ('PNG', '*.png'), ('全部', '*.*')])
        if not path:
            return
        try:
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt
            from matplotlib import font_manager as fm
            font_path = self._get_font_path()
            if font_path:
                plt.rcParams['font.family'] = fm.FontProperties(fname=font_path).get_name()
            plt.rcParams['axes.unicode_minus'] = False

            data = sorted(self._wf_all_data, key=lambda x: x[1], reverse=True)[:top_n]
            labels = [d[0] for d in data]
            values = [d[1] for d in data]

            fig, ax = plt.subplots(figsize=(14, max(8, top_n * 0.4)), dpi=150)
            bars = ax.barh(range(len(labels)), values, color='#1976d2', alpha=0.85)
            ax.set_yticks(range(len(labels)))
            ax.set_yticklabels(labels, fontsize=10)
            ax.invert_yaxis()
            ax.set_xlabel('词频（次）', fontsize=12)
            ax.set_title(f'词频统计 Top {top_n}', size=15)
            for bar, val in zip(bars, values):
                ax.text(bar.get_width() + max(values) * 0.005,
                        bar.get_y() + bar.get_height() / 2,
                        f'{int(val):,}', va='center', fontsize=9)
            ax.grid(axis='x', alpha=0.3)
            plt.tight_layout()
            fig.savefig(path, bbox_inches='tight', facecolor='white')
            plt.close(fig)
            messagebox.showinfo('完成', f'柱状图已保存：\n{path}')
        except Exception as e:
            messagebox.showerror('保存失败', str(e))

    # ── 导出 ──────────────────────────────────────────────────

    def _export_tfidf_csv(self):
        if not self._wf_all_data:
            messagebox.showinfo('提示', '请先执行词频分析')
            return
        path = filedialog.asksaveasfilename(
            title='导出词频表', defaultextension='.csv',
            filetypes=[('CSV', '*.csv'), ('Excel', '*.xlsx'), ('全部', '*.*')])
        if not path:
            return
        try:
            import pandas as pd
            rows = [{'词项': t, '词频': c, '频率%': f'{p:.4f}',
                    'TF-IDF最高': tf, '文档数': df, '字符数': cl}
                   for t, c, p, tf, df, cl in self._wf_all_data]
            pd.DataFrame(rows).to_csv(path, index=False, encoding='utf-8-sig')
            messagebox.showinfo('完成', f'已导出：\n{path}')
        except Exception as e:
            messagebox.showerror('导出失败', str(e))

    def _export_word_doc_matrix(self):
        if not self._app:
            messagebox.showinfo('提示', '请先加载文档')
            return
        path = filedialog.asksaveasfilename(
            title='导出词-文档矩阵', defaultextension='.csv',
            filetypes=[('CSV', '*.csv'), ('Excel', '*.xlsx'), ('全部', '*.*')])
        if not path:
            return
        try:
            from hotel_analyzer.data_processor import compute_word_doc_matrix
            sw = self._load_custom_stopwords()
            df = compute_word_doc_matrix(list(self._app.documents),
                                         stopwords=sw, top_n=500)
            df.to_csv(path, encoding='utf-8-sig')
            messagebox.showinfo('完成', f'已导出：\n{path}')
        except Exception as e:
            messagebox.showerror('导出失败', str(e))

    # ── 运行 ────────────────────────────────────────────────

    def run(self):
        self.root.mainloop()


# ══════════════════════════════════════════════════════════════════
# 入口
# ══════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    root = tk.Tk()
    app  = QDAGUI(root)
    app.run()
