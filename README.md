# 🤖 PPT Generator — Full Stack App
### Document → PowerPoint using Your Own Template
**myrobo.in — Future Skills for Future Schools**

---

## 📁 Project Structure

```
ppt-app/
│
├── backend/
│   ├── main.py              ← FastAPI server (entry point)
│   ├── parser.py            ← Document content extractor
│   ├── template_reader.py   ← PPT template analyzer
│   ├── ppt_generator.py     ← PPT builder (core logic)
│   ├── requirements.txt     ← Python dependencies
│   ├── uploads/             ← Temp upload folder (auto-created)
│   └── outputs/             ← Generated PPT folder (auto-created)
│
├── frontend/
│   ├── public/
│   │   └── index.html
│   ├── src/
│   │   ├── App.jsx          ← Main React UI
│   │   ├── index.js
│   │   └── index.css
│   └── package.json
│
└── README.md
```

---

## ⚙️ Setup Instructions

### Prerequisites
- **Python 3.9+** — https://www.python.org/downloads/
- **Node.js 18+** — https://nodejs.org/

---

### 🐍 Backend Setup

**Step 1** — Open terminal in the `backend/` folder:
```
cd ppt-app\backend
```

**Step 2** — Create a virtual environment (recommended):
```
python -m venv venv
venv\Scripts\activate
```
*(On Mac/Linux: `source venv/bin/activate`)*

**Step 3** — Install dependencies:
```
pip install -r requirements.txt
```

**Step 4** — Start the backend server:
```
python main.py
```

You should see:
```
INFO:     Uvicorn running on http://0.0.0.0:8000
```

✅ Backend is running at **http://localhost:8000**

---

### ⚛️ Frontend Setup

Open a **new terminal window**, then:

**Step 1**
```
cd ppt-app\frontend
```

**Step 2**
```
npm install
```

**Step 3**
```
npm start
```

You should see:
```
Compiled successfully!
Local: http://localhost:3000
```

✅ Frontend is running at **http://localhost:3000**

---

## 🚀 Using the App

1. Open **http://localhost:3000** in your browser
2. **Upload Document** — drop your `.pdf`, `.docx`, `.txt`, or `.rtf` file
3. **Upload Template** — drop your `.pptx` template file
4. Click **"Generate PowerPoint"**
5. Wait 10–30 seconds
6. Click **"Download"** — your PPT is ready! 🎉

---

## 🧠 How It Works

### Step 1 — Document Parsing (`parser.py`)
- Reads PDF (PyMuPDF), DOCX (python-docx), TXT/RTF
- Detects numbered headings: `1. Introduction`, `2. Concept` etc.
- Extracts bullet points under each heading
- Returns structured JSON:
  ```json
  {
    "title": "Hall Sensor Bot",
    "sections": [
      { "heading": "1. Introduction", "points": ["point 1", "point 2"] }
    ]
  }
  ```

### Step 2 — Template Analysis (`template_reader.py`)
- Opens your .pptx template
- Extracts most-used colors (primary, secondary, accent)
- Extracts fonts (heading font, body font, sizes)
- Detects title slide vs content slide patterns
- Finds placeholder positions

### Step 3 — PPT Generation (`ppt_generator.py`)
- Loads your template as base
- Creates title slide using template's color scheme
- Creates one content slide per section
- Applies extracted fonts, colors, header/footer
- Max 5 bullet points per slide (auto-splits if more)
- Saves as .pptx

---

## 📦 Dependencies

### Backend
| Package | Purpose |
|---------|---------|
| fastapi | Web API framework |
| uvicorn | ASGI server |
| python-pptx | PPT read/write |
| PyMuPDF | PDF parsing |
| python-docx | DOCX parsing |
| python-multipart | File upload handling |

### Frontend
| Package | Purpose |
|---------|---------|
| react | UI framework |
| axios | HTTP requests |
| react-scripts | Build tooling |

---

## ❓ Troubleshooting

| Problem | Solution |
|---------|----------|
| `pip install` fails | Make sure Python 3.9+ is installed |
| `ModuleNotFoundError` | Run `pip install -r requirements.txt` again |
| `npm install` fails | Make sure Node.js 18+ is installed |
| CORS error in browser | Make sure backend is running on port 8000 |
| Port 8000 in use | Change port in `main.py`: `uvicorn.run(..., port=8001)` |
| Port 3000 in use | React will ask to use another port — say Yes |
| PDF not parsing | Make sure `PyMuPDF` is installed: `pip install PyMuPDF` |
| DOCX not parsing | Make sure `python-docx` is installed: `pip install python-docx` |

---

## 🔮 Extending the App

- **Add AI structuring**: In `parser.py`, call Gemini/OpenAI API to better structure content before generating slides
- **Add slide preview**: Use `python-pptx` to export slide thumbnails via LibreOffice
- **Multiple templates**: Let users choose from a library of templates
- **More formats**: Add `.ppt` support via LibreOffice conversion

---

*Built for myrobo.in — Future Skills for Future Schools* 🚀
