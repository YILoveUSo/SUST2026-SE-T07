import json
import os
from collections import deque
from dataclasses import dataclass, field
from typing import Dict, List, Tuple
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.patches import Patch
import numpy as np

# 设置中文字体支持
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# ------------------------------------------------------------
# 数据结构定义
# ------------------------------------------------------------

@dataclass
class Task:
    """表示一个任务节点"""
    name: str
    optimistic: float
    pessimistic: float
    most_likely: float
    expected: float = field(init=False)
    variance: float = field(init=False)
    
    es: float = 0
    ef: float = 0
    ls: float = float('inf')
    lf: float = float('inf')
    slack: float = 0
    
    predecessors: List[str] = field(default_factory=list)
    successors: List[str] = field(default_factory=list)
    
    def __post_init__(self):
        self.expected = (self.optimistic + 4 * self.most_likely + self.pessimistic) / 6.0
        self.variance = ((self.pessimistic - self.optimistic) / 6.0) ** 2


class ProjectNetwork:
    """项目管理网络图"""
    
    def __init__(self):
        self.tasks: Dict[str, Task] = {}
        self.start_nodes: List[str] = []
        self.end_nodes: List[str] = []
        self.critical_path: List[str] = []
        self.project_duration: float = 0.0
        
    def add_task(self, task: Task):
        self.tasks[task.name] = task
        
    def remove_task(self, name: str):
        if name in self.tasks:
            del self.tasks[name]
            for task in self.tasks.values():
                if name in task.predecessors:
                    task.predecessors.remove(name)
    
    def clear(self):
        self.tasks.clear()
        self.start_nodes.clear()
        self.end_nodes.clear()
        self.critical_path.clear()
        self.project_duration = 0.0
    
    def _build_successors(self):
        for task in self.tasks.values():
            task.successors.clear()
        for task in self.tasks.values():
            for pred in task.predecessors:
                if pred in self.tasks:
                    self.tasks[pred].successors.append(task.name)
    
    def _identify_start_end_nodes(self):
        self.start_nodes = [name for name, t in self.tasks.items() if not t.predecessors]
        self.end_nodes = [name for name, t in self.tasks.items() if not t.successors]
    
    def forward_pass(self):
        for task in self.tasks.values():
            task.es = 0.0
            task.ef = task.expected
        
        indegree = {name: len(t.predecessors) for name, t in self.tasks.items()}
        queue = deque([name for name, deg in indegree.items() if deg == 0])
        topo_order = []
        
        while queue:
            u = queue.popleft()
            topo_order.append(u)
            for v in self.tasks[u].successors:
                indegree[v] -= 1
                if indegree[v] == 0:
                    queue.append(v)
        
        if len(topo_order) != len(self.tasks):
            raise ValueError("检测到循环依赖，请检查任务前置关系")
        
        for u in topo_order:
            for v in self.tasks[u].successors:
                if self.tasks[v].es < self.tasks[u].ef:
                    self.tasks[v].es = self.tasks[u].ef
                    self.tasks[v].ef = self.tasks[v].es + self.tasks[v].expected
        
        self.project_duration = max(self.tasks[e].ef for e in self.end_nodes) if self.end_nodes else 0
    
    def backward_pass(self):
        for e in self.end_nodes:
            self.tasks[e].lf = self.project_duration
            self.tasks[e].ls = self.tasks[e].lf - self.tasks[e].expected
        
        indegree = {name: len(t.successors) for name, t in self.tasks.items()}
        queue = deque([name for name, deg in indegree.items() if deg == 0])
        reverse_topo = []
        
        while queue:
            u = queue.popleft()
            reverse_topo.append(u)
            for v in self.tasks[u].predecessors:
                indegree[v] -= 1
                if indegree[v] == 0:
                    queue.append(v)
        
        for u in reverse_topo:
            for v in self.tasks[u].predecessors:
                if self.tasks[v].lf > self.tasks[u].ls:
                    self.tasks[v].lf = self.tasks[u].ls
                    self.tasks[v].ls = self.tasks[v].lf - self.tasks[v].expected
        
        for task in self.tasks.values():
            task.slack = task.ls - task.es
    
    def find_critical_path(self):
        self.critical_path = []
        start_candidates = [n for n in self.start_nodes if abs(self.tasks[n].slack) < 1e-6]
        if not start_candidates and self.start_nodes:
            start_candidates = self.start_nodes[:1]
        
        def dfs(node, path):
            nonlocal best_path
            path.append(node)
            if node in self.end_nodes and abs(self.tasks[node].slack) < 1e-6:
                if len(path) > len(best_path):
                    best_path = path.copy()
            else:
                for succ in self.tasks[node].successors:
                    if abs(self.tasks[succ].slack) < 1e-6 and succ not in path:
                        dfs(succ, path)
            path.pop()
        
        best_path = []
        for start in start_candidates:
            dfs(start, [])
        self.critical_path = best_path
    
    def compute_all(self):
        if not self.tasks:
            return
        self._build_successors()
        self._identify_start_end_nodes()
        self.forward_pass()
        self.backward_pass()
        self.find_critical_path()


class ResourceScheduler:
    def __init__(self, network: ProjectNetwork):
        self.network = network
        
    def compute_min_staff(self) -> int:
        events = []
        for task in self.network.tasks.values():
            events.append((task.es, 0, task.name))
            events.append((task.ef, 1, task.name))
        events.sort(key=lambda x: (x[0], x[1]))
        
        current_tasks = set()
        max_concurrent = 0
        for _, etype, _ in events:
            if etype == 0:
                current_tasks.add(None)
                max_concurrent = max(max_concurrent, len(current_tasks))
            else:
                current_tasks.discard(None)
        return max(1, max_concurrent)
    
    def generate_assignment_table(self) -> List[Dict]:
        tasks_list = []
        for task in self.network.tasks.values():
            tasks_list.append({
                'name': task.name,
                'start': task.es,
                'end': task.ef,
                'duration': task.expected
            })
        tasks_list.sort(key=lambda x: x['start'])
        
        min_people = self.compute_min_staff()
        workers = [0.0] * min_people
        assignment = []
        
        for t in tasks_list:
            earliest_worker_idx = min(range(min_people), key=lambda i: workers[i])
            available_time = workers[earliest_worker_idx]
            actual_start = max(t['start'], available_time)
            actual_end = actual_start + t['duration']
            workers[earliest_worker_idx] = actual_end
            assignment.append({
                'task': t['name'],
                'assigned_to': f"员工 {earliest_worker_idx + 1}",
                'start': actual_start,
                'end': actual_end,
                'duration': t['duration']
            })
        assignment.sort(key=lambda x: x['start'])
        return assignment


# ------------------------------------------------------------
# 示例数据管理器
# ------------------------------------------------------------

class ExampleManager:
    """管理多个示例项目"""
    
    @staticmethod
    def get_examples():
        return {
            "网络计划图示例 (A-H)": {
                "description": "经典网络计划图示例，包含任务A到H，展示复杂依赖关系",
                "tasks": [
                    ("A", 8, 8, 8, []),
                    ("B", 3, 3, 3, ["A"]),
                    ("C", 4, 4, 4, ["A"]),
                    ("D", 10, 10, 10, ["B", "C"]),
                    ("E", 9, 9, 9, ["B"]),
                    ("F", 2, 2, 2, ["C"]),
                    ("G", 5, 5, 5, ["D", "E"]),
                    ("H", 1, 1, 1, ["E", "F", "G"])
                ]
            },
            "软件开发项目": {
                "description": "典型的软件开发流程，包含需求分析、设计、编码、测试、部署",
                "tasks": [
                    ("需求分析", 2, 3, 5, []),
                    ("系统设计", 3, 4, 6, ["需求分析"]),
                    ("数据库设计", 2, 3, 4, ["需求分析"]),
                    ("前端开发", 4, 6, 9, ["系统设计"]),
                    ("后端开发", 5, 7, 10, ["系统设计", "数据库设计"]),
                    ("集成测试", 3, 4, 6, ["前端开发", "后端开发"]),
                    ("用户验收", 2, 3, 5, ["集成测试"]),
                    ("部署上线", 1, 2, 3, ["用户验收"])
                ]
            },
            "建筑工程项目": {
                "description": "建筑工程示例：场地准备、基础、结构、装修",
                "tasks": [
                    ("场地清理", 2, 3, 5, []),
                    ("地基开挖", 3, 4, 6, ["场地清理"]),
                    ("混凝土基础", 4, 5, 7, ["地基开挖"]),
                    ("主体结构", 7, 10, 14, ["混凝土基础"]),
                    ("外墙施工", 5, 6, 9, ["主体结构"]),
                    ("内部装修", 6, 8, 12, ["主体结构"]),
                    ("设备安装", 4, 5, 7, ["内部装修"]),
                    ("竣工验收", 1, 2, 3, ["外墙施工", "设备安装"])
                ]
            }
        }


# ------------------------------------------------------------
# 使用流程引导对话框
# ------------------------------------------------------------

class UsageGuideDialog:
    def __init__(self, parent_app):
        self.app = parent_app
        self.dialog = tk.Toplevel(self.app.root)
        self.dialog.title("使用流程引导")
        self.dialog.geometry("600x650")
        self.dialog.transient(self.app.root)
        self.dialog.grab_set()
        self.dialog.configure(bg='#f0f0f0')
        
        self.dialog.update_idletasks()
        x = self.app.root.winfo_x() + (self.app.root.winfo_width() - 600) // 2
        y = self.app.root.winfo_y() + (self.app.root.winfo_height() - 650) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        self.setup_ui()
        
    def setup_ui(self):
        main_frame = tk.Frame(self.dialog, bg='#f0f0f0')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        title_label = tk.Label(main_frame, text="PERT 项目进度规划工具", 
                               font=("微软雅黑", 16, "bold"), bg='#f0f0f0', fg='#2c3e50')
        title_label.pack(pady=(0, 5))
        
        subtitle_label = tk.Label(main_frame, text="使用流程与快速入门", 
                                  font=("微软雅黑", 11), bg='#f0f0f0', fg='#7f8c8d')
        subtitle_label.pack(pady=(0, 20))
        
        text_frame = tk.Frame(main_frame, bg='#f0f0f0')
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        steps_text = tk.Text(text_frame, wrap=tk.WORD, font=("微软雅黑", 10), 
                             bg='white', fg='#333333', padx=15, pady=15,
                             relief=tk.GROOVE, borderwidth=1)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=steps_text.yview)
        steps_text.configure(yscrollcommand=scrollbar.set)
        steps_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        steps_content = """
【快速开始】

第一步：添加任务
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• 点击左侧「添加任务」按钮
• 输入任务名称（如：A, B, C...）
• 填写三点时间估算：
  - 乐观时间(a)：最顺利情况
  - 最可能时间(m)：正常情况  
  - 悲观时间(b)：最不顺利情况
• 如有前置任务，用逗号分隔填写
• 点击「确定」完成

第二步：加载示例（推荐）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• 点击「加载示例」按钮
• 选择「网络计划图示例 (A-H)」
• 自动加载并计算，立即查看结果

第三步：计算分析
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• 添加完任务后，点击「计算关键路径」
• 系统自动计算：
  ✓ 每个任务的期望工期
  ✓ 总时差（浮动时间）
  ✓ 关键路径识别
  ✓ 项目总完成时间
  ✓ 所需最少人数

第四步：查看结果
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• 分析结果：详细计算数据表格
• 甘特图：可视化时间安排
• 依赖关系图：任务依赖网络
• 任务分配表：人员分配方案

【网络计划图示例说明】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
任务依赖关系：
A → B, C
B → D, E
C → D, F
D → G
E → G, H
F → H
G → H

【小贴士】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
时间单位保持一致（天/周/月）
至少需要一个没有前置任务的起始任务
避免循环依赖
双击任务列表可快速编辑任务
        """
        
        steps_text.insert(1.0, steps_content)
        steps_text.config(state=tk.DISABLED)
        
        btn_frame = tk.Frame(main_frame, bg='#f0f0f0')
        btn_frame.pack(fill=tk.X, pady=(20, 0))
        
        ttk.Button(btn_frame, text="开始使用", command=self.dialog.destroy,
                  width=12).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="加载示例", command=self.show_examples,
                  width=12).pack(side=tk.RIGHT, padx=5)
        
    def show_examples(self):
        self.dialog.destroy()
        self.app.show_example_selector()


# ------------------------------------------------------------
# 示例选择对话框
# ------------------------------------------------------------

class ExampleSelectorDialog:
    def __init__(self, parent_app):
        self.app = parent_app
        self.dialog = tk.Toplevel(self.app.root)
        self.dialog.title("选择示例项目")
        self.dialog.geometry("550x500")
        self.dialog.transient(self.app.root)
        self.dialog.grab_set()
        self.dialog.configure(bg='#f0f0f0')
        
        self.dialog.update_idletasks()
        x = self.app.root.winfo_x() + (self.app.root.winfo_width() - 550) // 2
        y = self.app.root.winfo_y() + (self.app.root.winfo_height() - 500) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        self.setup_ui()
        
    def setup_ui(self):
        main_frame = tk.Frame(self.dialog, bg='#f0f0f0')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        title_label = tk.Label(main_frame, text="选择示例项目", 
                               font=("微软雅黑", 14, "bold"), bg='#f0f0f0', fg='#2c3e50')
        title_label.pack(pady=(0, 10))
        
        examples = ExampleManager.get_examples()
        
        # 创建按钮列表
        for name, data in examples.items():
            self.create_example_card(main_frame, name, data["description"], data["tasks"])
        
        btn_frame = tk.Frame(main_frame, bg='#f0f0f0')
        btn_frame.pack(fill=tk.X, pady=(20, 0))
        ttk.Button(btn_frame, text="关闭", command=self.dialog.destroy, width=10).pack()
        
    def create_example_card(self, parent, name, description, tasks):
        """创建示例卡片"""
        card = tk.Frame(parent, bg='white', relief=tk.RAISED, borderwidth=1)
        card.pack(fill=tk.X, pady=5, padx=5)
        card.configure(height=80)
        card.pack_propagate(False)
        
        title_label = tk.Label(card, text=f"{name}", font=("微软雅黑", 11, "bold"),
                               bg='white', fg='#2c3e50')
        title_label.place(x=10, y=8)
        
        desc_label = tk.Label(card, text=description, font=("微软雅黑", 9),
                              bg='white', fg='#7f8c8d')
        desc_label.place(x=10, y=35)
        
        task_count = len(tasks)
        count_label = tk.Label(card, text=f"{task_count} 个任务", font=("微软雅黑", 8),
                               bg='white', fg='#3498db')
        count_label.place(x=10, y=55)
        
        select_btn = tk.Button(card, text="选择", font=("微软雅黑", 9),
                               bg='#3498db', fg='white', relief=tk.FLAT,
                               activebackground='#2980b9', activeforeground='white',
                               cursor='hand2', width=8,
                               command=lambda n=name: self.load_example(n))
        select_btn.place(x=470, y=28)
        
    def load_example(self, example_name):
        """加载选中的示例"""
        examples = ExampleManager.get_examples()
        example = examples.get(example_name)
        if not example:
            messagebox.showerror("错误", f"找不到示例: {example_name}")
            return
            
        # 清空当前项目
        self.app.network = ProjectNetwork()
        
        # 添加示例任务
        for task_data in example["tasks"]:
            name, opt, ml, pess, preds = task_data
            task = Task(name=name, optimistic=opt, most_likely=ml,
                       pessimistic=pess, predecessors=preds)
            self.app.network.add_task(task)
        
        # 刷新界面
        self.app.refresh_task_list()
        self.app.clear_graphs()
        self.app.results_text.delete(1.0, tk.END)
        
        # 自动计算
        try:
            self.app.network.compute_all()
            self.app.display_results()
            self.app.draw_gantt_chart()
            self.app.draw_dependency_graph()
            self.app.display_assignment_table()
            self.app.status_bar.config(
                text=f"已加载示例「{example_name}」| 项目总工期: {self.app.network.project_duration:.2f} 时间单位"
            )
            messagebox.showinfo("加载成功", 
                               f"已加载示例项目「{example_name}」\n\n"
                               f"包含 {len(self.app.network.tasks)} 个任务\n"
                               f"项目总工期: {self.app.network.project_duration:.2f} 时间单位\n\n"
                               f"关键路径: {' → '.join(self.app.network.critical_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"加载示例失败: {str(e)}")
        
        self.dialog.destroy()


# ------------------------------------------------------------
# GUI 主界面
# ------------------------------------------------------------

class PERTApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PERT 项目进度规划工具 - 关键路径分析")
        self.root.geometry("1400x900")
        self.root.configure(bg='#f0f0f0')
        
        self.network = ProjectNetwork()
        self.current_file = None
        
        self.setup_ui()
        self.setup_menu()
        
        # 显示欢迎信息
        self.root.after(500, self.show_welcome)
        
    def show_welcome(self):
        """显示欢迎对话框"""
        welcome_dialog = tk.Toplevel(self.root)
        welcome_dialog.title("欢迎")
        welcome_dialog.geometry("450x300")
        welcome_dialog.transient(self.root)
        welcome_dialog.grab_set()
        welcome_dialog.configure(bg='#f0f0f0')
        
        welcome_dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 450) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 300) // 2
        welcome_dialog.geometry(f"+{x}+{y}")
        
        frame = tk.Frame(welcome_dialog, bg='#f0f0f0')
        frame.pack(expand=True, fill=tk.BOTH, padx=30, pady=30)
        
        title = tk.Label(frame, text="欢迎使用 PERT 项目规划工具", 
                        font=("微软雅黑", 14, "bold"), bg='#f0f0f0', fg='#2c3e50')
        title.pack(pady=(0, 20))
        
        desc = tk.Label(frame, text="基于PERT技术的项目进度规划系统\n\n"
                      "• 三点估算法计算期望工期\n"
                      "• 自动识别关键路径\n"
                      "• 可视化甘特图和依赖关系\n"
                      "• 资源需求估算和任务分配",
                      font=("微软雅黑", 10), bg='#f0f0f0', fg='#555555',
                      justify=tk.CENTER)
        desc.pack(pady=(0, 30))
        
        btn_frame = tk.Frame(frame, bg='#f0f0f0')
        btn_frame.pack()
        
        ttk.Button(btn_frame, text="开始使用", command=welcome_dialog.destroy,
                  width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="查看教程", 
                  command=lambda: [welcome_dialog.destroy(), self.show_guide()],
                  width=12).pack(side=tk.LEFT, padx=5)
        
    def show_guide(self):
        UsageGuideDialog(self)
        
    def show_example_selector(self):
        ExampleSelectorDialog(self)
        
    def setup_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="新建项目", command=self.new_project, accelerator="Ctrl+N")
        file_menu.add_command(label="保存项目", command=self.save_project, accelerator="Ctrl+S")
        file_menu.add_command(label="加载项目", command=self.load_project, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        
        example_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="示例", menu=example_menu)
        example_menu.add_command(label="加载示例项目", command=self.show_example_selector)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="使用流程", command=self.show_guide)
        help_menu.add_command(label="关于", command=self.show_about)
        
        self.root.bind('<Control-n>', lambda e: self.new_project())
        self.root.bind('<Control-s>', lambda e: self.save_project())
        self.root.bind('<Control-o>', lambda e: self.load_project())
        
    def show_about(self):
        about_text = """PERT 项目进度规划工具

版本 2.0
基于 PERT (Program Evaluation and Review Technique) 技术

主要功能：
• 三点估算法计算任务期望工期
• 关键路径法(CPM)识别关键任务
• 自动计算项目总工期和时差
• 资源需求估算和任务分配
• 甘特图和依赖关系图可视化"""
        
        messagebox.showinfo("关于", about_text)
        
    def setup_ui(self):
        # 主框架
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左侧面板
        left_frame = tk.Frame(main_frame, bg='#ffffff', relief=tk.GROOVE, borderwidth=1)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 10))
        left_frame.configure(width=320)
        left_frame.pack_propagate(False)
        
        # 标题栏
        title_bar = tk.Frame(left_frame, bg='#3498db', height=40)
        title_bar.pack(fill=tk.X)
        title_bar.pack_propagate(False)
        tk.Label(title_bar, text="任务管理", font=("微软雅黑", 12, "bold"),
                bg='#3498db', fg='white').pack(pady=8)
        
        # 内容区域
        content_frame = tk.Frame(left_frame, bg='white', padx=10, pady=10)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 统计信息
        stats_frame = tk.Frame(content_frame, bg='#ecf0f1', relief=tk.GROOVE, borderwidth=1)
        stats_frame.pack(fill=tk.X, pady=(0, 10))
        self.task_count_label = tk.Label(stats_frame, text="任务数量: 0", 
                                         font=("微软雅黑", 10), bg='#ecf0f1', fg='#2c3e50')
        self.task_count_label.pack(pady=5)
        
        # 任务列表
        tk.Label(content_frame, text="任务列表 (双击编辑):", font=("微软雅黑", 9),
                bg='white', fg='#555555', anchor=tk.W).pack(anchor=tk.W, pady=(0, 5))
        
        list_frame = tk.Frame(content_frame, bg='white')
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.task_listbox = tk.Listbox(list_frame, height=12, font=("微软雅黑", 10),
                                       yscrollcommand=scrollbar.set,
                                       selectbackground='#3498db', selectforeground='white',
                                       bg='#f8f9fa', fg='#333333', relief=tk.FLAT)
        self.task_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.task_listbox.yview)
        self.task_listbox.bind('<Double-Button-1>', lambda e: self.edit_task_dialog())
        
        # 操作按钮
        btn_frame = tk.Frame(content_frame, bg='white')
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="添加任务", command=self.add_task_dialog,
                  width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="编辑", command=self.edit_task_dialog,
                  width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="删除", command=self.delete_task,
                  width=8).pack(side=tk.LEFT, padx=2)
        
        # 文件操作
        file_frame = tk.LabelFrame(content_frame, text="文件操作", font=("微软雅黑", 9),
                                   bg='white', fg='#2c3e50')
        file_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(file_frame, text="新建项目", command=self.new_project).pack(fill=tk.X, pady=2, padx=5)
        ttk.Button(file_frame, text="保存项目", command=self.save_project).pack(fill=tk.X, pady=2, padx=5)
        ttk.Button(file_frame, text="加载项目", command=self.load_project).pack(fill=tk.X, pady=2, padx=5)
        
        # 示例按钮
        ttk.Button(content_frame, text="加载示例", command=self.show_example_selector,
                  style="Success.TButton").pack(fill=tk.X, pady=5)
        
        # 计算按钮
        ttk.Button(content_frame, text="计算关键路径", command=self.calculate,
                  style="Primary.TButton").pack(fill=tk.X, pady=10)
        
        # 右侧面板
        right_frame = tk.Frame(main_frame, bg='#f0f0f0')
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 记事本
        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 结果标签页
        self.results_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.results_frame, text="分析结果")
        
        self.results_text = tk.Text(self.results_frame, wrap=tk.WORD, font=("Consolas", 10),
                                    bg='white', fg='#333333', padx=10, pady=10)
        scrollbar_r = ttk.Scrollbar(self.results_frame, orient=tk.VERTICAL,
                                     command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=scrollbar_r.set)
        self.results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_r.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 甘特图标签页
        self.gantt_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.gantt_frame, text="甘特图")
        
        # 依赖图标签页
        self.graph_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.graph_frame, text="依赖关系图")
        
        # 分配表标签页
        self.assignment_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.assignment_frame, text="任务分配表")
        
        # 分配表 Treeview
        columns = ("task", "worker", "start", "end", "duration")
        self.assignment_tree = ttk.Treeview(self.assignment_frame, columns=columns,
                                             show="headings", height=20)
        self.assignment_tree.heading("task", text="任务名称")
        self.assignment_tree.heading("worker", text="分配给")
        self.assignment_tree.heading("start", text="开始时间")
        self.assignment_tree.heading("end", text="结束时间")
        self.assignment_tree.heading("duration", text="工期")
        
        self.assignment_tree.column("task", width=150)
        self.assignment_tree.column("worker", width=80)
        self.assignment_tree.column("start", width=100)
        self.assignment_tree.column("end", width=100)
        self.assignment_tree.column("duration", width=80)
        
        scrollbar_a = ttk.Scrollbar(self.assignment_frame, orient=tk.VERTICAL,
                                     command=self.assignment_tree.yview)
        self.assignment_tree.configure(yscrollcommand=scrollbar_a.set)
        self.assignment_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_a.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 状态栏
        self.status_bar = tk.Label(self.root, text="就绪 | 点击「加载示例」选择示例项目快速体验",
                                   relief=tk.SUNKEN, anchor=tk.W, font=("微软雅黑", 9),
                                   bg='#ecf0f1', fg='#555555')
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def update_task_count(self):
        count = len(self.network.tasks)
        self.task_count_label.config(text=f"任务数量: {count}")
        
    def add_task_dialog(self):
        """打开添加任务对话框"""
        dialog = TaskDialog(self.root, self.network.tasks.keys())
        self.root.wait_window(dialog.dialog)
        if dialog.result:
            task = Task(
                name=dialog.result['name'],
                optimistic=dialog.result['optimistic'],
                most_likely=dialog.result['most_likely'],
                pessimistic=dialog.result['pessimistic'],
                predecessors=dialog.result['predecessors']
            )
            self.network.add_task(task)
            self.refresh_task_list()
            self.status_bar.config(text=f"已添加任务: {task.name}")
            
    def edit_task_dialog(self):
        selection = self.task_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要编辑的任务")
            return
        task_name = self.task_listbox.get(selection[0])
        task = self.network.tasks[task_name]
        
        dialog = TaskDialog(self.root, self.network.tasks.keys(), task)
        self.root.wait_window(dialog.dialog)
        if dialog.result:
            self.network.remove_task(task_name)
            new_task = Task(
                name=dialog.result['name'],
                optimistic=dialog.result['optimistic'],
                most_likely=dialog.result['most_likely'],
                pessimistic=dialog.result['pessimistic'],
                predecessors=dialog.result['predecessors']
            )
            self.network.add_task(new_task)
            self.refresh_task_list()
            self.status_bar.config(text=f"已更新任务: {new_task.name}")
            
    def delete_task(self):
        selection = self.task_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要删除的任务")
            return
        task_name = self.task_listbox.get(selection[0])
        if messagebox.askyesno("确认", f"确定删除任务 '{task_name}' 吗？"):
            self.network.remove_task(task_name)
            self.refresh_task_list()
            self.status_bar.config(text=f"已删除任务: {task_name}")
            
    def refresh_task_list(self):
        self.task_listbox.delete(0, tk.END)
        for name in sorted(self.network.tasks.keys()):
            self.task_listbox.insert(tk.END, name)
        self.update_task_count()
        
    def new_project(self):
        if self.network.tasks and messagebox.askyesno("确认", "当前项目未保存，是否新建项目？"):
            self.network = ProjectNetwork()
            self.current_file = None
            self.refresh_task_list()
            self.results_text.delete(1.0, tk.END)
            self.clear_graphs()
            self.status_bar.config(text="已创建新项目")
            
    def save_project(self):
        if not self.network.tasks:
            messagebox.showwarning("警告", "没有任务可保存")
            return
        filename = filedialog.asksaveasfilename(defaultextension=".json",
                                                  filetypes=[("JSON files", "*.json")])
        if filename:
            data = {"tasks": []}
            for task in self.network.tasks.values():
                data["tasks"].append({
                    "name": task.name,
                    "optimistic": task.optimistic,
                    "most_likely": task.most_likely,
                    "pessimistic": task.pessimistic,
                    "predecessors": task.predecessors
                })
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            self.current_file = filename
            self.status_bar.config(text=f"已保存到: {filename}")
            messagebox.showinfo("成功", "项目已保存")
            
    def load_project(self):
        filename = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.network = ProjectNetwork()
                for tdata in data["tasks"]:
                    task = Task(
                        name=tdata["name"],
                        optimistic=tdata["optimistic"],
                        most_likely=tdata["most_likely"],
                        pessimistic=tdata["pessimistic"],
                        predecessors=tdata["predecessors"]
                    )
                    self.network.add_task(task)
                self.refresh_task_list()
                self.current_file = filename
                self.status_bar.config(text=f"已加载: {filename} ({len(self.network.tasks)}个任务)")
                messagebox.showinfo("成功", f"已加载 {len(self.network.tasks)} 个任务")
            except Exception as e:
                messagebox.showerror("错误", f"加载失败: {e}")
                
    def calculate(self):
        if not self.network.tasks:
            messagebox.showwarning("警告", "请先添加任务\n点击「添加任务」按钮开始")
            return
            
        try:
            self.status_bar.config(text="正在计算...")
            self.root.update()
            
            self.network.compute_all()
            self.display_results()
            self.draw_gantt_chart()
            self.draw_dependency_graph()
            self.display_assignment_table()
            
            self.status_bar.config(
                text=f"计算完成 | 项目总工期: {self.network.project_duration:.2f} 时间单位 | "
                     f"关键任务: {len(self.network.critical_path)} 个"
            )
            messagebox.showinfo("计算完成",
                               f"项目分析完成！\n\n"
                               f"项目总工期: {self.network.project_duration:.2f} 时间单位\n"
                               f"关键路径: {' → '.join(self.network.critical_path) if self.network.critical_path else '无'}\n"
                               f"总任务数: {len(self.network.tasks)} 个")
        except ValueError as e:
            messagebox.showerror("错误", str(e))
            self.status_bar.config(text="计算失败")
            
    def display_results(self):
        self.results_text.delete(1.0, tk.END)
        
        results = []
        results.append("=" * 80)
        results.append(" " * 25 + "项目进度分析结果 (PERT)")
        results.append("=" * 80)
        
        results.append("\n【任务期望工期与浮动时间】")
        results.append("-" * 80)
        results.append(f"{'任务名称':<18} {'期望工期':<12} {'方差':<14} {'总时差':<12} {'关键任务':<10}")
        results.append("-" * 80)
        for task in self.network.tasks.values():
            critical = "是" if abs(task.slack) < 1e-6 else "否"
            results.append(f"{task.name:<18} {task.expected:<12.2f} {task.variance:<14.4f} {task.slack:<12.2f} {critical:<10}")
        
        results.append("\n【关键路径】")
        results.append("-" * 80)
        if self.network.critical_path:
            results.append(" → ".join(self.network.critical_path))
            results.append(f"\n关键路径长度: {len(self.network.critical_path)} 个任务")
        else:
            results.append("无法确定关键路径")
        
        results.append(f"\n【项目总完成时间】: {self.network.project_duration:.2f} 时间单位")
        
        scheduler = ResourceScheduler(self.network)
        min_people = scheduler.compute_min_staff()
        results.append(f"\n【最快完成项目所需最少人数】: {min_people} 人")
        
        results.append("\n" + "=" * 80)
        
        self.results_text.insert(1.0, "\n".join(results))
        
    def draw_gantt_chart(self):
        for widget in self.gantt_frame.winfo_children():
            widget.destroy()
            
        fig = Figure(figsize=(12, 7), dpi=100, facecolor='white')
        ax = fig.add_subplot(111, facecolor='white')
        
        tasks = list(self.network.tasks.values())
        tasks.sort(key=lambda t: t.es)
        
        colors = ['#2ecc71' if abs(t.slack) < 1e-6 else '#3498db' for t in tasks]
        
        y_pos = range(len(tasks))
        for i, task in enumerate(tasks):
            ax.barh(i, task.expected, left=task.es, color=colors[i],
                    edgecolor='black', linewidth=0.5, height=0.6)
            ax.text(task.es + task.expected/2, i, task.name,
                    ha='center', va='center', fontsize=9, fontweight='bold')
        
        ax.set_yticks(y_pos)
        ax.set_yticklabels([t.name for t in tasks], fontsize=9)
        ax.set_xlabel('时间', fontsize=11, fontweight='bold')
        ax.set_title('项目甘特图', fontsize=12, fontweight='bold', pad=15)
        ax.grid(True, axis='x', alpha=0.3, linestyle='--')
        
        legend_elements = [Patch(facecolor='#2ecc71', label='关键任务 (时差=0)'),
                          Patch(facecolor='#3498db', label='非关键任务')]
        ax.legend(handles=legend_elements, loc='lower right', frameon=True)
        
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=self.gantt_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
    def draw_dependency_graph(self):
        """绘制从左到右的依赖关系图"""
        for widget in self.graph_frame.winfo_children():
            widget.destroy()
            
        try:
            import networkx as nx
            fig = Figure(figsize=(14, 8), dpi=100, facecolor='white')
            ax = fig.add_subplot(111, facecolor='white')
            
            G = nx.DiGraph()
            for task in self.network.tasks.values():
                G.add_node(task.name)
                for pred in task.predecessors:
                    G.add_edge(pred, task.name)
            
            # 计算节点层级（用于从左到右布局）
            try:
                layers = {}
                for node in nx.topological_sort(G):
                    if G.in_degree(node) == 0:
                        layers[node] = 0
                    else:
                        max_pred_layer = max(layers[pred] for pred in G.predecessors(node))
                        layers[node] = max_pred_layer + 1
                
                # 根据层级设置位置
                pos = {}
                layer_nodes = {}
                for node, layer in layers.items():
                    if layer not in layer_nodes:
                        layer_nodes[layer] = []
                    layer_nodes[layer].append(node)
                
                for layer, nodes in layer_nodes.items():
                    y_positions = np.linspace(-len(nodes)/2, len(nodes)/2, len(nodes))
                    for i, node in enumerate(nodes):
                        pos[node] = (layer * 2, y_positions[i] if len(nodes) > 1 else 0)
            except:
                # 如果拓扑排序失败，使用spring布局
                pos = nx.spring_layout(G, k=2, iterations=50, seed=42)
            
            node_colors = ['#2ecc71' if name in self.network.critical_path else '#3498db'
                          for name in G.nodes()]
            
            nx.draw(G, pos, ax=ax, with_labels=True, node_color=node_colors,
                   node_size=3000, font_size=10, font_weight='bold',
                   arrows=True, arrowstyle='->', arrowsize=20,
                   edge_color='#95a5a6', width=1.5, alpha=0.8)
            
            ax.set_title('任务依赖关系图（从左到右布局）', fontsize=12, fontweight='bold', pad=15)
            ax.axis('off')
            
            fig.tight_layout()
            canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        except ImportError:
            label = tk.Label(self.graph_frame,
                            text="需要安装 networkx 库来显示依赖图\n\n请运行: pip install networkx",
                            font=("微软雅黑", 12), fg='red', bg='white')
            label.pack(expand=True)
            
    def display_assignment_table(self):
        for item in self.assignment_tree.get_children():
            self.assignment_tree.delete(item)
            
        scheduler = ResourceScheduler(self.network)
        assignment = scheduler.generate_assignment_table()
        
        for a in assignment:
            self.assignment_tree.insert("", tk.END, values=(
                a['task'], a['assigned_to'], f"{a['start']:.2f}",
                f"{a['end']:.2f}", f"{a['duration']:.2f}"
            ))
            
    def clear_graphs(self):
        for widget in self.gantt_frame.winfo_children():
            widget.destroy()
        for widget in self.graph_frame.winfo_children():
            widget.destroy()
        for item in self.assignment_tree.get_children():
            self.assignment_tree.delete(item)


class TaskDialog:
    """任务添加/编辑对话框"""
    
    def __init__(self, parent, existing_names, task=None):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("编辑任务" if task else "添加任务")
        self.dialog.geometry("450x550")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.configure(bg='#f0f0f0')
        
        self.parent = parent
        self.existing_names = list(existing_names)
        self.result = None
        
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 450) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 550) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        self.setup_ui(task)
        
    def setup_ui(self, task):
        main_frame = tk.Frame(self.dialog, bg='#f0f0f0', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        title_label = tk.Label(main_frame, text="任务信息", font=("微软雅黑", 14, "bold"),
                              bg='#f0f0f0', fg='#2c3e50')
        title_label.pack(pady=(0, 20))
        
        # 表单框架
        form_frame = tk.Frame(main_frame, bg='white', relief=tk.GROOVE, borderwidth=1)
        form_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        inner_frame = tk.Frame(form_frame, bg='white', padx=15, pady=15)
        inner_frame.pack(fill=tk.BOTH, expand=True)
        
        # 任务名称
        tk.Label(inner_frame, text="任务名称 *", font=("微软雅黑", 10),
                bg='white', fg='#333333', anchor=tk.W).pack(anchor=tk.W, pady=(0, 3))
        self.name_entry = ttk.Entry(inner_frame, width=35, font=("微软雅黑", 10))
        self.name_entry.pack(fill=tk.X, pady=(0, 12))
        
        # 乐观时间
        tk.Label(inner_frame, text="乐观时间 (a) - 最顺利情况", font=("微软雅黑", 10),
                bg='white', fg='#333333', anchor=tk.W).pack(anchor=tk.W, pady=(0, 3))
        self.opt_entry = ttk.Entry(inner_frame, width=35, font=("微软雅黑", 10))
        self.opt_entry.pack(fill=tk.X, pady=(0, 12))
        
        # 最可能时间
        tk.Label(inner_frame, text="最可能时间 (m) - 正常情况", font=("微软雅黑", 10),
                bg='white', fg='#333333', anchor=tk.W).pack(anchor=tk.W, pady=(0, 3))
        self.ml_entry = ttk.Entry(inner_frame, width=35, font=("微软雅黑", 10))
        self.ml_entry.pack(fill=tk.X, pady=(0, 12))
        
        # 悲观时间
        tk.Label(inner_frame, text="悲观时间 (b) - 最不顺利情况", font=("微软雅黑", 10),
                bg='white', fg='#333333', anchor=tk.W).pack(anchor=tk.W, pady=(0, 3))
        self.pess_entry = ttk.Entry(inner_frame, width=35, font=("微软雅黑", 10))
        self.pess_entry.pack(fill=tk.X, pady=(0, 12))
        
        # 前置任务
        tk.Label(inner_frame, text="前置任务 (多个用逗号分隔)", font=("微软雅黑", 10),
                bg='white', fg='#333333', anchor=tk.W).pack(anchor=tk.W, pady=(0, 3))
        self.pred_entry = ttk.Entry(inner_frame, width=35, font=("微软雅黑", 10))
        self.pred_entry.pack(fill=tk.X, pady=(0, 5))
        
        tip_label = tk.Label(inner_frame, text="提示：前置任务必须是已存在的任务名称",
                            font=("微软雅黑", 8), bg='white', fg='#7f8c8d')
        tip_label.pack(anchor=tk.W)
        
        # 按钮
        btn_frame = tk.Frame(main_frame, bg='#f0f0f0')
        btn_frame.pack(fill=tk.X, pady=(20, 0))
        
        ttk.Button(btn_frame, text="确定", command=self.ok, width=10).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="取消", command=self.cancel, width=10).pack(side=tk.RIGHT, padx=5)
        
        # 如果是编辑模式，填充数据
        if task:
            self.name_entry.insert(0, task.name)
            self.name_entry.config(state='readonly')
            self.opt_entry.insert(0, str(task.optimistic))
            self.ml_entry.insert(0, str(task.most_likely))
            self.pess_entry.insert(0, str(task.pessimistic))
            self.pred_entry.insert(0, ", ".join(task.predecessors))
            
    def ok(self):
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showerror("错误", "任务名称不能为空", parent=self.dialog)
            return
            
        # 检查任务名是否重复
        if name in self.existing_names:
            is_edit = (self.name_entry.cget('state') == 'readonly')
            if not is_edit:
                messagebox.showerror("错误", f"任务名称 '{name}' 已存在，请使用不同的名称", parent=self.dialog)
                return
            
        try:
            opt = float(self.opt_entry.get())
            ml = float(self.ml_entry.get())
            pess = float(self.pess_entry.get())
            if opt < 0 or ml < 0 or pess < 0:
                messagebox.showerror("错误", "时间不能为负数", parent=self.dialog)
                return
        except ValueError:
            messagebox.showerror("错误", "时间必须是数字", parent=self.dialog)
            return
            
        preds_str = self.pred_entry.get().strip()
        predecessors = [p.strip() for p in preds_str.split(',') if p.strip()]
        
        # 验证前置任务是否存在
        all_names = self.existing_names.copy()
        current_name = self.name_entry.get().strip()
        if current_name in all_names:
            all_names.remove(current_name)
            
        for pred in predecessors:
            if pred not in all_names:
                if not all_names:
                    msg = f"前置任务 '{pred}' 不存在\n\n当前没有其他任务，请先添加其他任务再设置依赖关系。"
                else:
                    msg = f"前置任务 '{pred}' 不存在\n\n当前已有的任务: {', '.join(all_names)}"
                messagebox.showerror("错误", msg, parent=self.dialog)
                return
                
        self.result = {
            'name': name,
            'optimistic': opt,
            'most_likely': ml,
            'pessimistic': pess,
            'predecessors': predecessors
        }
        self.dialog.destroy()
        
    def cancel(self):
        self.dialog.destroy()


def main():
    root = tk.Tk()
    # 设置样式
    style = ttk.Style()
    style.configure("Primary.TButton", font=("微软雅黑", 10, "bold"))
    style.configure("Success.TButton", font=("微软雅黑", 10, "bold"))
    style.configure("Warning.TButton", font=("微软雅黑", 10, "bold"))
    
    app = PERTApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()