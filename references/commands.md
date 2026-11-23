# 命令与工具说明

在独立工作目录中运行以下命令。先将 `S` 设为本技能目录的绝对路径。用 Python 3.12、LibreOffice Impress 和 Math 测试。为公式准备 Cambria Math 或匹配的数学字体，为普通文本准备明确的、可嵌入的字体。若系统不能自动找到 LibreOffice，设置 `SOFFICE` 为可执行文件的完整路径。已有合适环境时可跳过安装；集群内的生成、导出和校验均在分配的计算节点执行。

```bash
S="${CODEX_HOME:-$HOME/.codex}/skills/img-to-ppt-skill"
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -r "$S/scripts/requirements.txt"
```

## 绘制与检查自己的图

在工作目录创建绘图定义，运行时将 `$S/scripts` 加入 `PYTHONPATH`，并使用 `skill_portable_native.new_deck` 按参考图宽高和稿件物理宽度创建画布。所有框坐标记录在**未裁切**的 PDF 中，单位为点，原点位于左上。按项目选定并记录印刷宽度和字形尺度偏差阈值。

- `calibrate_math.py` 把项目里实际用到的 LaTeX 列表转换为实测宽高。重复公式只需校准一次。换导出器、数学字体或字体映射后重新测量。核对生成的数学结构与原稿，尤其是多级上下标、粗体和可伸缩符号。
- `vector_art.place_vector(slide, image_path, x_emu, y_emu, width_emu, height_emu, name=...)` 导入不透明局部插画，输出可逐区域编辑的彩色形状组。用 `round(坐标 * c.scale)` 将参考像素换算至 EMU。转换会限制采样尺寸和路径数，用印刷尺寸比较细节；必要时改为手绘关键轮廓。
- `python "$S/scripts/compare.py" figure.pdf comparison.pdf source-1.png source-2.jpg` 生成按给定顺序并排的校样。图片数须等于 PDF 页数。发布的论文单图取自母版 PDF；运行 `python "$S/scripts/export_pages.py" figure.pdf paper/` 可直接按页导出；可选 `--crops crops.json` 按明确的 PDF 点坐标裁边，如 `{"1":[0,12,450,220]}`。复查成品页尺寸、字体、透明度及位图数。
- `audit.py` 审计整个文稿的页序、结构和 PDF；`--expect-math` 要求至少有一个可见 OMML。按每页清单另行核对实际公式数。纯图形页可指定 `--allow-empty-text 2 5`，其余页继续要求正文或数学文字及字体。
- `formula_check.py` 接受形如 `[{"page":1,"name":"Eq 1","box_pt":[40,30,90,50]}]` 的清单，覆盖最终页面里每个公式的独立区域。它测量普通字形的横纵尺度，排除单独的可伸缩括号等符号。对只有定界符、矩阵、轮廓化算子或与标签混排的公式逐项检查，必要时另制孤立校样；确认结果中的字符确实属于该公式。
- `probe.py` 在文稿副本中改一个具名文本框、一个原生下标和一个图形或图形组。测试文件没有下标，或需要修改指定子组、其他公式结构时，按实际语义编辑副本后沿用 `probe.py diff`。允许区应分别围住旧、新位置，取互不重叠的小框。`must_change: true` 表示对应区必须产生可见像素变化。

## 从分发包安装

把 `Img-to-PPT-skill.zip` 解压到用户技能目录（默认 `~/.codex/skills/`），确认入口为 `img-to-ppt-skill/SKILL.md`。在新会话中使用 `$img-to-ppt-skill` 并附上参考图或本地路径，以及已知的论文栏宽。
