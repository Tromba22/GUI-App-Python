# 🖥️ Auto PPT — AI-Powered PowerPoint Generator

A desktop GUI application built with Python that lets you create styled PowerPoint presentations through an intuitive interface — with an optional **OpenAI integration** for AI-generated content and images.

---

## ✨ Features

- 📝 **Slide builder** — Add titles, descriptions, company names, and job IDs per slide
- 🎨 **Full styling control** — Custom text colors, background colors, font sizes, bold/italic/underline
- 🖼️ **Image support** — Insert images with configurable positioning (corners)
- 📐 **Shape overlays** — Add geometric shapes to slides
- 🤖 **AI-powered version** — Generate slide content and images via OpenAI API
- 🌗 **Light & dark themes** — Custom themed UI using CustomTkinter
- 💾 **Export to .pptx** — Save presentations directly to disk

---

## 📁 Project Structure

```
Auto_ppt/
├── auto_ppt_basic.py       # Basic GUI version (Tkinter)
├── fileppt.py              # Enhanced version with scrollable UI & shape support
└── openaippt/
    ├── openaiapp.py         # OpenAI-integrated version (CustomTkinter)
    ├── pptaiauto.py         # AI content generation + DALL·E image integration
    └── theme/               # Light & dark UI theme assets
```

---

## 🛠️ Tech Stack

<p>
  <img src="https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white" />
  <img src="https://img.shields.io/badge/Tkinter-GUI-blue?style=flat-square" />
  <img src="https://img.shields.io/badge/CustomTkinter-Modern_UI-green?style=flat-square" />
  <img src="https://img.shields.io/badge/python--pptx-PPTX-orange?style=flat-square" />
  <img src="https://img.shields.io/badge/OpenAI-412991?style=flat-square&logo=openai&logoColor=white" />
</p>

---

## 🚀 Getting Started

```bash
git clone https://github.com/YOUR_USERNAME/GUI-App-Python.git
cd GUI-App-Python/Auto_ppt

# Install dependencies
pip install python-pptx customtkinter openai pillow requests

# Run the basic version
python auto_ppt_basic.py

# Run the AI-powered version
python openaippt/openaiapp.py
```

> **Note:** The OpenAI version requires a valid API key. Set it as an environment variable or configure it in the script.

---

## 👤 Author

**Ali Trabelsi Karoui** — [LinkedIn](https://www.linkedin.com/in/ali-trabelsi-karoui-226990151/) · [Email](mailto:alitrabelsikaroui2293@gmail.com)
