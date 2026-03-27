import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np

# 设置中文字体
plt.rcParams["font.sans-serif"] = ["SimHei", "Arial Unicode MS", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False

# 创建图形
fig, ax = plt.subplots(1, 1, figsize=(16, 12))
ax.set_xlim(0, 16)
ax.set_ylim(0, 12)
ax.axis("off")

# 定义颜色方案
colors = {
    "user": "#E3F2FD",  # 浅蓝
    "logic": "#FFF3E0",  # 浅橙
    "data": "#E8F5E9",  # 浅绿
    "adapter": "#F3E5F5",  # 浅紫
    "border": "#1976D2",  # 深蓝边框
    "text": "#212121",  # 深灰文字
    "arrow": "#757575",  # 灰色箭头
}


# 绘制圆角矩形函数
def draw_rounded_box(ax, x, y, width, height, label, color, fontsize=10):
    box = FancyBboxPatch(
        (x, y),
        width,
        height,
        boxstyle="round,pad=0.1",
        facecolor=color,
        edgecolor=colors["border"],
        linewidth=2,
    )
    ax.add_patch(box)
    ax.text(
        x + width / 2,
        y + height / 2,
        label,
        ha="center",
        va="center",
        fontsize=fontsize,
        color=colors["text"],
        weight="bold",
    )


# 绘制层级标题
def draw_layer_title(ax, x, y, title):
    ax.text(x, y, title, fontsize=14, weight="bold", color=colors["border"])


# 绘制连接线
def draw_connection(ax, x1, y1, x2, y2):
    arrow = FancyArrowPatch(
        (x1, y1),
        (x2, y2),
        arrowstyle="->",
        mutation_scale=20,
        linewidth=2,
        color=colors["arrow"],
    )
    ax.add_patch(arrow)


# ==================== 第一层：用户交互层 ====================
draw_layer_title(ax, 0.5, 11.3, "用户交互层")

# 三个模块
draw_rounded_box(ax, 1, 10, 3.5, 1, "Web管理界面", colors["user"])
draw_rounded_box(ax, 6, 10, 3.5, 1, "命令行工具", colors["user"])
draw_rounded_box(ax, 11, 10, 3.5, 1, "配置文件", colors["user"])

# 连接线到下一层
for x in [2.75, 7.75, 12.75]:
    draw_connection(ax, x, 10, 8, 9)

# ==================== 第二层：业务逻辑层 ====================
draw_layer_title(ax, 0.5, 8.8, "业务逻辑层")

# 第一行三个模块
draw_rounded_box(ax, 0.5, 7.2, 2.8, 1.2, "规则引擎\n模块", colors["logic"], 9)
draw_rounded_box(ax, 4, 7.2, 2.8, 1.2, "AI处理引擎\n模块", colors["logic"], 9)
draw_rounded_box(ax, 7.5, 7.2, 2.8, 1.2, "邮件操作\n模块", colors["logic"], 9)
draw_rounded_box(ax, 11, 7.2, 2.8, 1.2, "知识库管理\n模块", colors["logic"], 9)
draw_rounded_box(ax, 13.2, 7.2, 2.3, 1.2, "问答匹配\n模块", colors["logic"], 9)

# 连接线到下一层
for x in [1.9, 5.4, 8.9, 12.4, 14.35]:
    draw_connection(ax, x, 7.2, 8, 5.5)

# ==================== 第三层：数据存储层 ====================
draw_layer_title(ax, 0.5, 5.8, "数据存储层")

# 三个模块
draw_rounded_box(ax, 2, 4.5, 3.5, 1, "SQLite日志数据库", colors["data"])
draw_rounded_box(ax, 6.25, 4.5, 3.5, 1, "JSON配置文件", colors["data"])
draw_rounded_box(ax, 10.5, 4.5, 3.5, 1, "知识库文档文件", colors["data"])

# 连接线到下一层
for x in [3.75, 8, 12.25]:
    draw_connection(ax, x, 4.5, 8, 3.5)

# ==================== 第四层：接口适配层 ====================
draw_layer_title(ax, 0.5, 3.3, "接口适配层")

# 三个模块
draw_rounded_box(ax, 1.5, 2, 3.5, 1, "Outlook COM接口", colors["adapter"])
draw_rounded_box(ax, 6.25, 2, 3.5, 1, "Microsoft Graph API", colors["adapter"])
draw_rounded_box(ax, 11, 2, 3.5, 1, "LMStudio AI接口", colors["adapter"])

# 添加标题
ax.text(
    8,
    11.8,
    "Outlook邮件助手系统架构图",
    fontsize=20,
    weight="bold",
    color=colors["border"],
    ha="center",
)

# 添加副标题
ax.text(
    8,
    11.4,
    "智能化邮件自动处理系统",
    fontsize=12,
    color=colors["text"],
    ha="center",
    style="italic",
)

# 添加技术栈说明（右下角）
tech_stack_text = """技术栈：
后端：Python + Flask + SQLite
前端：HTML5 + Bootstrap 5
AI：LMStudio + 本地大模型"""

ax.text(
    15.5,
    0.5,
    tech_stack_text,
    fontsize=8,
    color=colors["text"],
    ha="right",
    va="bottom",
    bbox=dict(
        boxstyle="round,pad=0.5",
        facecolor="white",
        edgecolor=colors["border"],
        alpha=0.8,
    ),
)

plt.tight_layout()
plt.savefig("architecture_diagram.png", dpi=300, bbox_inches="tight", facecolor="white")
print("架构图已生成：architecture_diagram.png")
plt.close()

# 生成第二个版本：详细模块图
fig, ax = plt.subplots(1, 1, figsize=(18, 14))
ax.set_xlim(0, 18)
ax.set_ylim(0, 14)
ax.axis("off")

# 标题
ax.text(
    9,
    13.5,
    "Outlook邮件助手 - 详细架构图",
    fontsize=22,
    weight="bold",
    color=colors["border"],
    ha="center",
)

# ==================== 核心模块详细图 ====================

# 中央大模块：核心控制器
draw_rounded_box(ax, 6, 10, 6, 1.5, "核心控制器\n(Core Controller)", "#FFECB3", 11)

# 左侧：输入模块
draw_rounded_box(
    ax, 0.5, 10, 4, 1.2, "邮件接收模块\n(Email Receiver)", colors["user"], 9
)
draw_connection(ax, 4.5, 10.6, 6, 10.75)

# 右侧：输出模块
draw_rounded_box(
    ax, 13.5, 10, 4, 1.2, "邮件发送模块\n(Email Sender)", colors["user"], 9
)
draw_connection(ax, 12, 10.75, 13.5, 10.6)

# 下方：处理模块
modules_y = 7.5
module_width = 2.5
module_height = 1.3
spacing = 0.3

# 第一行处理模块
start_x = 1
for i, (name, color) in enumerate(
    [
        ("规则引擎", colors["logic"]),
        ("AI引擎", colors["logic"]),
        ("模板引擎", colors["logic"]),
        ("问答引擎", colors["logic"]),
        ("知识库", colors["data"]),
    ]
):
    x = start_x + i * (module_width + spacing)
    draw_rounded_box(ax, x, modules_y, module_width, module_height, name, color, 9)
    # 连接到核心
    draw_connection(ax, x + module_width / 2, modules_y + module_height, 9, 10)

# 第二行：数据存储
storage_y = 4.5
storage_modules = [
    ("配置管理\nconfig.json", colors["data"]),
    ("日志数据库\nexecution_logs.db", colors["data"]),
    ("知识库文件\nknowledge_base/", colors["data"]),
    ("问答库\nqa_database.json", colors["data"]),
]

start_x = 1.5
for i, (name, color) in enumerate(storage_modules):
    x = start_x + i * (3.5 + 0.3)
    draw_rounded_box(ax, x, storage_y, 3.5, 1.2, name, color, 8)
    # 连接到上层
    if i < 2:
        draw_connection(ax, x + 1.75, storage_y + 1.2, 4 + i * 3, modules_y)
    else:
        draw_connection(ax, x + 1.75, storage_y + 1.2, 10 + (i - 2) * 3, modules_y)

# 最下层：外部接口
interface_y = 1.5
interfaces = [
    ("Outlook COM\n(Windows)", colors["adapter"]),
    ("Microsoft 365\nGraph API", colors["adapter"]),
    ("LMStudio\nLocal AI", colors["adapter"]),
]

start_x = 2
for i, (name, color) in enumerate(interfaces):
    x = start_x + i * (4.5 + 0.5)
    draw_rounded_box(ax, x, interface_y, 4.5, 1.2, name, color, 9)
    # 连接到数据层
    draw_connection(ax, x + 2.25, interface_y + 1.2, 5 + i * 4, storage_y)

# 添加流程箭头说明
ax.annotate(
    "",
    xy=(17, 10.5),
    xytext=(17, 2),
    arrowprops=dict(arrowstyle="->", lw=3, color=colors["border"]),
)
ax.text(
    17.3, 6, "数据流", fontsize=10, rotation=90, va="center", color=colors["border"]
)

plt.tight_layout()
plt.savefig(
    "architecture_detailed.png", dpi=300, bbox_inches="tight", facecolor="white"
)
print("详细架构图已生成：architecture_detailed.png")

# 生成第三个版本：流程图
fig, ax = plt.subplots(1, 1, figsize=(16, 10))
ax.set_xlim(0, 16)
ax.set_ylim(0, 10)
ax.axis("off")

# 标题
ax.text(
    8,
    9.5,
    "邮件处理流程图",
    fontsize=20,
    weight="bold",
    color=colors["border"],
    ha="center",
)

# 流程步骤
steps = [
    ("1. 接收邮件", 8, 8.5, colors["user"]),
    ("2. 规则匹配", 8, 7, colors["logic"]),
    ("3. 知识库检索", 5, 5.5, colors["data"]),
    ("4. AI生成回复", 11, 5.5, colors["logic"]),
    ("5. 执行动作", 8, 4, colors["adapter"]),
    ("6. 发送回复", 8, 2.5, colors["user"]),
]

for i, (label, x, y, color) in enumerate(steps):
    draw_rounded_box(ax, x - 1.5, y - 0.4, 3, 0.8, label, color, 10)
    if i < len(steps) - 1:
        next_y = steps[i + 1][2]
        next_x = steps[i + 1][1]
        if next_x == x:  # 垂直连接
            draw_connection(ax, x, y - 0.4, next_x, next_y + 0.4)
        else:  # 水平连接
            draw_connection(ax, x, y, next_x, y)
            draw_connection(ax, next_x, y, next_x, next_y + 0.4)

plt.tight_layout()
plt.savefig("process_flow.png", dpi=300, bbox_inches="tight", facecolor="white")
print("流程图已生成：process_flow.png")

print("\n所有图片生成完成！")
print("1. architecture_diagram.png - 分层架构图")
print("2. architecture_detailed.png - 详细架构图")
print("3. process_flow.png - 处理流程图")
