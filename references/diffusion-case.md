# 八图复刻案例

[案例源包](../assets/diffusion-redraw-source.zip) 按当时的相对路径保存原始供稿图集、逐页绘图源代码、插画与提示词、文字配准标注、公式度量以及历史验收记录。提取到新的工作目录即可离线访问。使用代码时可对照下面的文件定位技术细节。

| 文件或目录 | 内容 |
|---|---|
| `注意力增强：扩散生成语义引导.zip` | 八幅原始供稿图。压缩包内顺序由归档保留 |
| `scripts/diffusion_native/pages_a.py` | AnimateDiff、ControlNet、DiffEdit |
| `scripts/diffusion_native/pages_b.py` | DreamBooth、PhotoMaker、StoryDiffusion |
| `scripts/diffusion_native/pages_c.py` | Diffusion-DPO、Attend-and-Excite |
| `canvas.py`、`equations.py`、`calibrate.py`、`formula_metrics.json` | 原项目可编辑图元、Office Math、字形校准及当时的度量值 |
| `illustrations.py`、`vector_art.py`、`generated_sources/` | 原生插画、局部彩色曲线、实际使用的素材母图和提示词 |
| `register_text.py`、`text_registration.json`、`scripts/diffusion_facsimile/labels.json`、`docs/public_results/.ppt_reference_work/` | 参考文字位置、原始标注、八张配准参考图 |
| `build.py`、`proofcheck.py`、`formula_proof.py`、`edit_probe.py`、`package.py` | 当时的构建、逐页校样、公式比例测量、实际修改验证和打包 |
| `previous-checks/` | 上次交付的结构检查 CSV、67 项公式测量 CSV 及真实编辑检验 JSON，供比对 |
| `docs/public_results/diffusion_native_redraw/Reference_Layout/`、`使用说明.md` | 保留供稿布局的 ControlNet 补充单页及历史版本说明 |

## 独立复跑

准备通用运行环境，并安装案例附加依赖。命令假定 `S` 是技能根目录。使用含 Liberation Sans、Liberation Serif 的 Regular、Italic、Bold、BoldItalic 字体文件的目录，必要时传 `--fonts /path/to/fonts`。配置渲染器能够发现这些字体及数学字体。原项目用 Liberation Serif/Sans 分别替代 Times New Roman/Arial，给 Cambria Math 配置了 TeX Gyre Termes Math 字体映射。实际使用及嵌入的字体见导出审计。不同环境可能需要重新校准。素材已经归档，复跑不会调用图像生成服务。

```bash
python -m pip install -r "$S/scripts/requirements-case.txt"
python "$S/scripts/replay_case.py" ./case-07 --figure 7
python "$S/scripts/replay_case.py" ./case-all --all
```

两个输出目录须为空或尚不存在。单图命令生成母版 PPTX、对应 PDF、结构审计和参考并排校样。全套命令还生成八个单独 PPTX、裁切版论文 PDF、逐公式检查和实际编辑报告，最终得到 `case-all/docs/public_results/diffusion_native_redraw/Diffusion_Methods_Native_Package.zip`。复跑新报告与 `previous-checks/` 中的历史记录分开保存。

原案例合计 239 个可编辑文本框、67 个原生公式，论文目标图宽为 183 mm。ControlNet 等页存在明示的科学内容校正；ControlNet 另附供稿拓扑版本。新任务应从自己的供稿与作者指示确定重绘范围、目标栏宽、字体和检查阈值。
