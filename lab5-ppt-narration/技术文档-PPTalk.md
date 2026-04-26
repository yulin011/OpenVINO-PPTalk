# PPTalk：让 PPT 自己讲起来（OpenVINO + Qwen3-VL + Qwen3-TTS）

> 项目定位：把“做 PPT 讲解”从**写稿 + 录音 + 剪辑**，压缩为“上传一次、自动生成”的端侧流程。  
> 核心能力：**PPT→逐页图片→逐页讲解稿→全局一致性梳理→逐页 TTS→打包输出（可选回写旁白到 PPTX）**。

---

## 1. 项目名称与一句话介绍

- **项目名**：**PPTalk**
- **一句话**：上传一份 PPT，自动生成“可直接口播的逐页讲解稿”与“逐页旁白音频”，并打包下载；在 Windows + PowerPoint 环境下还可将旁白**自动嵌入回 PPTX**，得到可播放旁白的演示稿版本。

---

## 2. 应用场景与创新点

### 2.1 面向用户

- **学生/教师/答辩者**：需要将 PPT 快速转为结构清晰、可口播的讲解稿，并生成旁白音频用于练习或成品交付。
- **产品/研发/售前**：希望把 PPT 快速打包为“可听可看”的交付物（图 + 稿 + 音频 + 提纲），降低重复录音与脚本打磨成本。
- **内容创作者**：把静态 slide 扩展为多媒体素材，便于二次分发与改编。

### 2.2 解决的实际问题

- **时间成本**：从“逐页写稿 + 统一术语 + 录音”降到“一键生成 + 少量人工微调”。
- **跨页一致性**：逐页生成常见问题是“跨页重复定义、术语不统一、节奏断裂”。本项目将“逐页草稿”与“全局梳理校验”拆分为两阶段，专门处理多页 PPT 的一致性。
- **端侧可用性**：VLM 与 TTS 使用 OpenVINO 推理，可在本地 CPU/GPU 上运行；全局梳理阶段对接 OpenAI 兼容接口，便于按网络条件与合规要求替换部署形态。

### 2.3 核心方法（从局部到全局）

- **逐页结构化草稿**：VLM 输出严格 JSON，保证信息结构可控（标题/要点/讲解稿/不该念的内容）。
- **全局一致性梳理**：对多页稿件做术语统一、跨页去重、过渡自然化，并强调“不得编造 PPT 中不存在的事实”。
- **可交付产物标准化**：统一的图片、脚本、音频命名与目录结构，便于复现、回溯与二次加工。

---

## 3. 系统总体流程（Pipeline）

整体流程为 A→F 六步（对应 `gradio_app.py` 的 `pipeline()`）：

1. **A）PPT → PNG**：`ppt_utils.export_pptx_to_pngs()`  
2. **B）逐页讲解稿初稿（VLM）**：`vlm_script.generate_slide_drafts()`  
3. **C）全局梳理校验（OpenAI 兼容接口）**：`llm_polish.polish_slide_scripts()`  
4. **D）逐页语音合成（TTS）**：`tts_narration.synthesize_slides_to_wavs()`  
5. **E）可选：旁白嵌入 PPTX（PowerPoint COM）**：`ppt_utils.embed_slide_wavs_to_pptx()`  
6. **F）打包输出**：`tts_narration.package_outputs_zip()`

默认输出目录：`lab5-ppt-narration/outputs/<run_id>/`

- `slides/slide_0001.png ...`（逐页图片）
- `scripts/slides_drafts.json`（逐页初稿）
- `scripts/slides_polished.json`（逐页润色后讲解稿）
- `scripts/global_outline.md`（全局提纲）
- `audio/slide_0001.wav ...`（逐页音频）
- `package_<run_id>.zip`（打包下载）
- `narrated_<name>_<run_id>.pptx`（可选：嵌入旁白后的 PPTX）

---

## 4. 模块设计与关键实现点

### 4.1 `ppt_utils.py`：PPT 渲染与旁白回写

**目标**：将 `.ppt/.pptx` 稳定导出为按页排序的 PNG；在 Windows + PowerPoint 环境下，将逐页 wav 旁白嵌入回 PPTX。

**关键设计**：

- **双后端兜底**：
  - **PowerPoint COM（Windows-only）**：渲染一致性最好；同时支持“嵌入音频回写 PPTX”。
  - **LibreOffice + PDF + 图片**：当缺少 PowerPoint 时仍可导出图片（需要 `soffice`、`pdf2image`、Poppler）。
- **线程内 COM 初始化**：Gradio 的 pipeline 通常在 worker 线程执行；COM 需要当前线程 `CoInitialize`。通过 `_com_initialized()` 保障在 UI 场景也能稳定使用 PowerPoint 渲染/回写。
- **Windows 大小写不敏感的文件收集坑**：PowerPoint 可能导出 `*.PNG`。在 Windows 上同时 `glob("*.png")` 与 `glob("*.PNG")` 可能收集到同一批文件，导致重命名阶段出现 `FileNotFoundError`。通过“合并候选 + realpath 去重”避免重复处理同一文件。

### 4.2 `vlm_script.py`：逐页讲解稿（结构化草稿）

**目标**：让 VLM 输出“可控结构”的讲解稿，避免自由散文式输出造成不可控长度与格式漂移。

**关键设计**：

- **严格 JSON 输出**：字段包含：
  - `slide_title`：页标题（可空）
  - `key_points`：3–6 个要点（避免逐条复读）
  - `speaker_notes`：最终口播稿（可直接读）
  - `do_not_say`：不该念的内容（如 logo、页脚、版权等）
- **抗噪策略（Prompt 约束）**：明确忽略页角 logo、装饰背景、页眉页脚等；对图表/流程图使用“先结论、后解释、再 takeaway”的讲解模板。
- **时长预算**：`short/medium/long` 映射到不同口播长度区间，便于按场景控制节奏。

### 4.3 `llm_polish.py`：全局一致性梳理（OpenAI 兼容接口）

**目标**：解决逐页生成的“跨页不一致”问题：术语统一、跨页去重、自然过渡、避免编造事实。

**关键设计**：

- **强约束提示**：强调“不得新增 PPT 中不存在的关键事实；不确定时用更保守表述”。
- **JSON Mode 优先**：优先使用 OpenAI SDK 的 JSON 输出模式；失败时自动降级并做 JSON 提取，提高鲁棒性。
- **工程化 IO**：输入逐页结构化内容，输出“逐页 speaker_notes + 全局提纲”。

### 4.4 `tts_narration.py`：逐页语音合成与打包

**目标**：将每页讲解稿稳定生成 wav，并对长文本做分段，提升端侧稳定性与可重复性。

**关键设计**：

- **分段策略**：按段落/句子切分，超长再硬切（`split_text_for_tts`）。
- **音频拼接**：分段 wav 之间插入 100ms 静音，提升听感与可懂度。
- **标准化命名**：输出 `slide_0001.wav`，与“嵌入 PPTX”步骤对齐。

### 4.5 `gradio_app.py`：一键交互与产物导出

**UI 输入**：

- 上传 PPT / 选择渲染后端（`auto/powerpoint/libreoffice`）
- 每页讲解长度、是否加入过渡、听众类型
- OpenAI 兼容接口配置（Base URL / Key / Model）
- TTS speaker / language / 风格指令（instruct）

**UI 输出**：逐页图片预览、讲解稿 JSON、全局提纲、音频试听、打包 zip（可选 narrated PPTX）。

---

## 5. 环境配置与依赖安装（本机实测）

### 5.1 测试机器信息

- **CPU**：Intel(R) Core(TM) i9-14900HX @ 2.20GHz
- **内存**：64 GB
- **GPU**：NVIDIA GeForce RTX 4060 Laptop GPU（8 GB）与 Intel UHD Graphics
- **操作系统**：Windows 11

### 5.2 软件与版本

- **OpenVINO**：`openvino==2026.0`（见仓库根目录 `requirements.txt`）
- **Gradio**：`gradio==6.9.0`
- **Transformers**：`transformers==4.57.3`（与 Qwen3-TTS 依赖对齐）

### 5.3 安装步骤（PowerShell）

1) 创建并激活虚拟环境：

```bash
python -m venv ov_workshop
.\ov_workshop\Scripts\activate
```

2) 安装依赖：

```bash
pip install -r requirements.txt
```

3) （建议）启用 PowerPoint 渲染与 PPTX 回写能力：

```bash
pip install pywin32
```

4) （可选）无 PowerPoint 时启用 LibreOffice 渲染后端：

- 安装 LibreOffice（确保 `soffice` 在 PATH）
- 安装 Poppler（供 `pdf2image` 使用）并配置到 PATH
- 安装 Python 包：

```bash
pip install pdf2image
```

> 注：TTS 环节可能出现 “SoX could not be found!” 的探测提示，通常不影响主流程；如遇到确实依赖 SoX 的处理环节，再按提示安装并加入 PATH。

---

## 6. 运行方式与结果展示

### 6.1 Notebook 方式（便于复现与调参）

- 打开 `lab5-ppt-narration/lab5-ppt-narration.ipynb`
- 按顺序运行：设备选择 → 模型加载 → 跑通导出/生成/润色/TTS/打包
- 将要处理的 PPT 放到 `lab5-ppt-narration/assets/`，并设置：
  - `ppt_path = Path('assets/demo.pptx')`

### 6.2 Gradio 一键演示

从 Notebook 的 Gradio 章节启动，或在入口脚本中启动 `make_demo()`。

### 6.3 Baseline 可行性验证：豆包 AI 生成 OpenVINO 示例 PPT

为验证端到端流程在“非手工制作、具有真实排版与内容密度”的输入上同样可用，我使用**豆包 AI**生成了一份“关于 OpenVINO 的 PPT”，作为 baseline 示例输入，重点验证：

- PPT 渲染导出是否稳定（含中文环境下的文件命名、版式与图片）
- VLM 是否能按页提炼关键信息并生成可口播讲解稿
- 全局梳理是否能减少跨页重复、统一术语并形成连贯提纲
- TTS 是否能按页稳定合成，且在长页场景下通过分段策略保持可懂度

豆包生成 PPT 的过程截图如下（示例输入来源说明）：

![豆包AI生成PPT过程截图](assets/doubao2ppt.png)

### 6.4 运行结果截图（必须）

> 这里保留占位，便于你后续替换为真实截图。

- **截图 1：Gradio 主界面（上传 PPT/选择参数）**  
  `![Gradio 主界面](待补充：screenshots/gradio_home.png)`

- **截图 2：逐页图片预览 + 讲解稿 JSON 输出**  
  `![输出预览](待补充：screenshots/gradio_outputs.png)`

- **截图 3：音频试听 + 打包 zip**  
  `![音频与打包](待补充：screenshots/gradio_zip_audio.png)`

### 6.5 完整 Demo 视频（可选）

- `（待补充：demo_video.mp4 / 链接）`

---

## 7. Skill 运行展示（Copaw 创空间）

### 7.1 Skill 形态建议

- **Skill 名称**：PPTalk：PPT 一键旁白包
- **输入**：PPT/PPTX、听众类型、每页时长、是否过渡、（可选）OpenAI 接口配置、TTS 风格
- **输出**：`package_<run_id>.zip`（slides + scripts + audio + outline），可选 `narrated_*.pptx`

### 7.2 截图/视频（占位）

- **截图 1：Skill 参数填写页**  
  `![Skill 参数](待补充：screenshots/skill_form.png)`
- **截图 2：Skill 运行中进度/日志**  
  `![Skill 进度](待补充：screenshots/skill_progress.png)`
- **截图 3：Skill 输出物（zip + pptx）**  
  `![Skill 输出](待补充：screenshots/skill_outputs.png)`

---

## 8. 跑通过程与踩坑记录（工程鲁棒性）

### 8.1 PPT 导出图片：Windows 大小写不敏感导致重复收集

**现象**：`FileNotFoundError: '...\\幻灯片1.PNG' -> '...\\slide_0002.png'`  
**根因**：Windows 文件系统大小写不敏感，`glob("*.png") + glob("*.PNG")` 可能收集同一批文件，导致后续重命名重复处理同一源文件。  
**修复**：合并候选后按 `resolve().lower()` 去重，再排序重命名。

### 8.2 Gradio worker 线程下 PowerPoint COM 失败

**现象**：Notebook 单步能用 PowerPoint 渲染，但在 Gradio pipeline 中出现 “检测不到 powerpoint 后端/Dispatch 失败”。  
**根因**：COM 必须在当前线程 `CoInitialize`；而 Gradio 通常在线程池里执行用户函数。  
**修复**：为 COM 调用加入 `_com_initialized()` 上下文，保证线程内正确初始化/反初始化。

### 8.3 多后端渲染策略：一致性与可用性的平衡

- PowerPoint：一致性与回写能力最好，但依赖 Office 环境。  
- LibreOffice：可作为无 Office 时的渲染兜底，但需要额外组件（`soffice`、Poppler、`pdf2image`），且渲染细节可能与 Office 存在差异。  

### 8.4 TTS 稳定性：长文本分段是必要工程

逐页讲解稿在“长页/密集页”容易超出一次性生成的稳定区间。项目采用：

- 段落/句子切分 + 超长硬切（`split_text_for_tts`）
- 分段 wav 拼接时插入短静音（100ms）

以提升端侧输出的稳定性与可重复性。

---

## 9. 总结与展望

### 9.1 端侧部署心得

- 端侧落地的主要挑战通常不在“模型能不能跑”，而在“输入多样性”带来的工程边界：PPT 渲染环境差异、文件命名/编码、线程模型与系统组件（COM）、以及跨页一致性与语音节奏。
- OpenVINO 的价值在于让 VLM/TTS 在本地 CPU/GPU 上更可控地运行，便于在有限环境下稳定复现。

### 9.2 当前局限

- **标题页/过渡页**：尚未做专门识别，可能导致标题页讲解偏长。
- **Logo/装饰干扰**：已用提示词抑制，但不同模板噪声不同，仍可能偶发提及不重要元素。
- **全局梳理依赖接口**：全局梳理阶段依赖 OpenAI 兼容接口；离线场景需本地替代方案。
- **旁白回写依赖 PowerPoint**：嵌入旁白到 PPTX 依赖 Windows + PowerPoint（COM）。

### 9.3 优化方向

- 标题页/章节页智能识别与时长自适应
- 对图表/流程图的更强版面理解与讲解模板
- 本地离线一致性校验（轻量 LLM 或规则/模板）
- 支持逐页编辑与“单页重合成音频”的人机协同工作流

