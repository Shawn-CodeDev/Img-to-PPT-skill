# Img-to-PPT Skill

将科研示意图重绘成可编辑的 PowerPoint，并从同一份母版导出论文用 PDF。

原生文本、Office Math 公式、形状、箭头和可拆分插画让后续修改有对象可选；逐图校样和编辑验证帮助检查成稿与参考图是否一致。

> 这是一个面向支持 Agent Skills 的助手的工作流与工具包，不是将任意图片一键转换成完美 PPT 的应用。重绘需要理解原图、编写或调整绘图代码，并进行人工验收。

## 功能概览

| 能力 | 内容 |
| --- | --- |
| 逐图重绘 | 将参考图中的模块、文字、分区和连接关系重建为 PPT 对象 |
| 可编辑公式 | 使用 Office Math（OMML）表达公式，并通过导出样张校准字形比例 |
| 可编辑图形 | 创建形状、分组、折线与圆角箭头；支持将局部插画转为可编辑区域 |
| 论文导出 | 从 PPTX 母版导出 PDF，按页输出并按指定坐标裁切 |
| 质量检查 | 检查对象结构、公式区域、PDF 导出和真实编辑前后的可见变化 |
| 案例复跑 | 附带八幅扩散方法图的源素材、绘图代码及历史检查记录 |

实际可编辑性、字体、视觉一致性和矢量效果取决于图形内容、字体和渲染环境；自动检查不能替代逐图目检。

## 快速开始

### 1. 获取项目

```bash
git clone https://github.com/Shawn-CodeDev/Img-to-PPT-skill.git
cd Img-to-PPT-skill
```

项目根目录的 [`SKILL.md`](SKILL.md) 是技能入口。可将整个仓库目录复制到支持该技能格式的助手的技能目录，例如 Codex 默认的 `~/.codex/skills/img-to-ppt-skill/`；复制后检查 `SKILL.md` 是否位于该目录下，并在新会话中调用。具体发现与安装方式取决于使用的助手。

### 2. 准备环境

项目示例使用 Python 3.12、LibreOffice Impress 和 Math。需要安装可用的数学字体及普通文本字体；案例文档说明了当时的字体替代方案。

在独立工作目录中运行：

```bash
# 将路径替换为本仓库或已安装技能目录的绝对路径
S="/path/to/img-to-ppt-skill"
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -r "$S/scripts/requirements.txt"
```

若 LibreOffice 不在 `PATH` 中，可设置 `SOFFICE` 为其可执行文件路径。集群环境请在已分配的计算节点运行生成、导出和检查任务。

### 3. 给助手提供任务

```text
使用 $img-to-ppt-skill 重绘这张科研示意图：/path/to/reference.png。
请保留文字、公式和箭头关系，生成可编辑 PPTX 和论文 PDF。
目标图宽 160 mm；请提供原图对照与逐项验收结果。
```

建议一并提供原始高清图片或 PDF、目标印刷尺寸、字体要求，以及是否允许更改原图内容。无法辨认的标签和关系应列为待确认项。

## 最小工作示例

以下命令生成一张示例图，导出 PDF 并运行结构审计。它们需要上面安装的依赖、LibreOffice 和可用的数学字体。

```bash
printf '%s\n' '["y_w"]' > terms.json
python "$S/scripts/calibrate_math.py" terms.json metrics.json
python "$S/scripts/example_figure.py" example --metrics metrics.json
python "$S/scripts/render.py" example/example.pptx example/example.pdf
python "$S/scripts/audit.py" example/example.pptx example/example.pdf \
  --expect-math > example/audit.json
```

输出包括可编辑的 `example/example.pptx`、导出的 `example/example.pdf` 和结构审计报告。公式区域检查、真实编辑探针及原图对照的完整命令见 [命令与最小工作例](references/commands.md)。

## 工作流程

1. **读图与登记**：确定页面顺序、文字、公式、模块、连接方向和目标印刷尺寸。
2. **重建布局**：以统一比例创建画布，安排面板、形状、分组、边距和层级。
3. **重建公式与连线**：用 Office Math 表达公式，检查上下标、符号、箭头端点和弯折。
4. **处理局部插画**：优先使用有权使用的供稿素材，必要时重绘或转为可编辑区域。
5. **导出和验收**：从 PPTX 导出 PDF，对照原图检查内容、字体、拓扑及目标尺寸下的可读性。
6. **验证可编辑性**：在副本中修改文本、公式和组件，比较编辑前后的实际渲染结果。

详细规范见 [重绘手册](references/redraw.md) 和 [逐图验收](references/acceptance.md)。

## 常用工具

| 脚本 | 用途 |
| --- | --- |
| `scripts/calibrate_math.py` | 测量项目公式的字体尺寸，辅助布局 |
| `scripts/example_figure.py` | 构建含文字、箭头、公式和分组的示例 |
| `scripts/render.py` | 通过 LibreOffice 将 PPTX 导出为 PDF |
| `scripts/audit.py` | 审计文稿结构及 PDF 产物 |
| `scripts/formula_check.py` | 按给定区域检查导出公式字形比例 |
| `scripts/probe.py` | 修改副本并对比导出结果，验证实际编辑效果 |
| `scripts/compare.py` | 生成参考图与导出 PDF 的并排校样 |
| `scripts/export_pages.py` | 将 PDF 按页导出，可选择裁切 |
| `scripts/replay_case.py` | 在独立目录复跑扩散方法图案例 |

脚本参数、检查条件和适用边界以 [命令参考](references/commands.md) 为准。

## 八图案例

[案例说明](references/diffusion-case.md) 描述了八幅扩散方法图的源材料、逐页代码、字体映射和历史验收结果。案例包位于 [`assets/diffusion-redraw-source.zip`](assets/diffusion-redraw-source.zip)，包含原始供稿图、生成素材和检查记录。

安装案例附加依赖后，可在空目录或尚不存在的目录中复跑：

```bash
python -m pip install -r "$S/scripts/requirements-case.txt"
python "$S/scripts/replay_case.py" ./case-07 --figure 7
python "$S/scripts/replay_case.py" ./case-all --all
```

案例历史记录包含 239 个可编辑文本框和 67 个原生公式；不同字体或导出环境可能产生不同结果。案例中部分图有明确标注的内容校正，自己的项目应以供稿和作者要求为准。原始供稿和第三方素材的再使用须确认相应权利。

## 仓库结构

```text
.
├── SKILL.md                  # Agent 技能入口
├── agents/openai.yaml        # 技能元数据
├── references/               # 命令、重绘、验收与案例说明
├── scripts/                  # 绘图、导出、审计与案例复跑工具
└── assets/                   # 案例源包
```

## 使用边界

- 忠实复刻时，疑似科学错误应先记录并与作者确认，再交付有差异说明的修正版。
- 结果图、显微图与定量热图应保留证据属性；没有原始数据时，不应生成替代实验结果。
- 公式结构计数、文字提取和像素差异不能单独证明语义或视觉正确；请结合原图逐页复核。
- 可编辑与全矢量要求需要按实际输出逐页验证；局部插画的转换可能丢失细节。
