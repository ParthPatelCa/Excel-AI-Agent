# 📊 ExcelAI Assistant

*A GPT-powered Excel Add-in that brings natural language data analysis, chart generation, and smart insights directly into your spreadsheet.*

---

## 🚀 Overview

**ExcelAI Assistant** is a sidebar Add-in for Microsoft Excel that lets users:
- Ask data-related questions in plain English
- Generate charts, summaries, or formulas automatically
- Interact with selected cells without needing deep Excel knowledge
- Receive AI-powered insights, all inside Excel

Built with:
- `Office.js` (for Excel integration)
- `Node.js` (hosted on Replit for backend logic)
- `OpenAI API` (for natural language processing)

---

## 🎯 Features

- 🔍 Natural language chat interface in Excel
- 📊 Auto-generate charts and summaries from selected data
- 🧮 Turn text prompts into valid Excel formulas
- 📥 Insert results, formulas, or charts into the sheet
- 🔐 Built-in **guardrails** to ensure safe and responsible use

---

## 🛡️ Misuse Prevention (Guardrails)

To ensure ethical and secure AI use, ExcelAI Assistant includes:

### Input Guardrails
- OpenAI Moderation API for toxic/unsafe prompts
- Regex filtering for sensitive terms (e.g., “password”, “delete”)
- Rate limits and prompt length caps

### Output Guardrails
- Formula validator (blocks unsafe or malicious output)
- No VBA/macros allowed
- Output truncation for safety

### Data Privacy
- Only selected cell range is processed
- Optional anonymization of data
- No storage or logging of user data
- Clear disclaimer before each request

---

## 🧱 Tech Stack

| Layer | Tools |
|-------|-------|
| Excel Integration | Office.js, Manifest XML |
| Frontend (Task Pane) | HTML, CSS, JavaScript (or React) |
| Backend | Node.js (Replit-hosted) |
| AI Engine | OpenAI GPT API |
| Hosting | Replit (HTTPS for dev; move to production later) |

---

## 🧪 How to Use (Developer Preview)

### 1. Clone the Repo
```bash
git clone https://github.com/your-org/excelai-assistant.git
