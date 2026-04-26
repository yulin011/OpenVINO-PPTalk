# OpenVINO AI 应用实战 Workshop

本项目基于 [OpenVINO™](https://github.com/openvinotoolkit/openvino) 工具套件，聚焦 **Lab 5：PPT 旁白生成（PPTalk）**：将 PPT/PPTX 转成逐页讲解脚本，并可进一步生成语音旁白等交付物，用于演示 OpenVINO 在多模态与语音方向的端到端应用能力。

## 💻 系统要求

- **操作系统**: Windows 11
- **Python**: 3.10+
- **GPU（可选）**: Intel® Arc™ 系列独立显卡 或 Intel® Core™ Ultra 集成显卡

## 📖 安装步骤

### 1. 创建虚拟环境并安装依赖

双击运行 `setup_lab.bat`，脚本会自动：
1. 使用 Python `venv` 创建虚拟环境 `ov_workshop`
2. 安装 `requirements.txt` 中的所有依赖
3. 克隆并安装 Qwen3-ASR / Qwen3-TTS 推理所需的代码库

### 2. 启动 JupyterLab

双击运行 `run_lab.bat`，脚本会：
1. 激活虚拟环境
2. 启动 JupyterLab，浏览器会自动打开

## 🧪 实验列表

| 实验 | 主题 | 简介 | 链接 |
|------|------|------|------|
| Lab 5 | PPT 旁白生成（PPTalk） | 将 PPT/PPTX 转成逐页讲解脚本，并可生成语音旁白等交付物 | [进入](lab5-ppt-narration/) |

## 📁 项目结构

```
openvino-workshop/
├── README.md                     # 本文件
├── .gitignore                    # Git 忽略规则
├── requirements.txt              # 统一依赖
├── setup_lab.bat                 # 环境安装脚本
├── run_lab.bat                   # 启动 JupyterLab
└── lab5-ppt-narration/           # Lab 5: PPT 旁白生成（PPTalk）
```

## 🔗 参考资源

- [OpenVINO 官方文档](https://docs.openvino.ai/)
- [OpenVINO Notebooks](https://github.com/openvinotoolkit/openvino_notebooks)
