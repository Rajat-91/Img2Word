# Img2Word 📝

**Convert images into editable Word documents using OCR.**

Img2Word is a web application that extracts text from images and converts it into downloadable Microsoft Word (`.docx`) documents. It uses Optical Character Recognition (OCR) and document-structure analysis to help preserve elements such as paragraphs, headings, and tables.

## ✨ Features

- 🖼️ **Image Upload** — Upload an image or drag and drop it into the application.
- 🔍 **OCR Text Extraction** — Extract text from supported images using PaddleOCR.
- 📄 **Word Document Generation** — Convert extracted content into editable `.docx` files.
- 📊 **Table Recognition** — Recognize document tables and convert supported table structures into Word tables.
- 🧱 **Document Structure** — Preserve supported headings, paragraphs, and lists.
- ⚡ **Simple Interface** — Convert images through a clean, user-friendly web interface.
- 📥 **Download Output** — Download the generated Word document.
- 🛡️ **File Validation** — Validate uploaded file types and file sizes.

## 🛠️ Tech Stack

### Frontend
- React
- TypeScript
- Vite

### Backend
- Python
- FastAPI
- PaddleOCR
- PP-StructureV3
- python-docx
- OpenCV
- Pillow
- BeautifulSoup4

## 🏗️ Project Structure

```text
Img2Word/
├── frontend/
│   ├── src/
│   ├── public/
│   ├── package.json
│   └── vite.config.ts
├── ocr-backend/
│   ├── main.py
│   ├── requirements.txt
│   └── .venv/
└── README.md
```

The `.venv/` directory is generated locally and should not be committed to Git.

## ⚙️ Getting Started

Follow these steps to run Img2Word locally.

### Prerequisites

Install the following:

- [Node.js](https://nodejs.org/)
- [Python 3.11](https://www.python.org/downloads/)
- [Git](https://git-scm.com/)

### 1. Clone the Repository

```bash
git clone https://github.com/Rajat-91/Img2Word.git
cd Img2Word
```

### 2. Set Up the Backend

Open a terminal in the project root.

Create and activate a virtual environment:

**Windows (Command Prompt)**

```cmd
cd ocr-backend
python -m venv .venv
.venv\Scripts\activate
```

Install the dependencies:

```cmd
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install "paddleocr[doc-parser]"
```

If `beautifulsoup4` is not already listed in `requirements.txt`, install it as well:

```cmd
python -m pip install beautifulsoup4
```

Start the backend:

```cmd
python -m uvicorn main:app --reload
```

The backend should be available at:

- API: `http://127.0.0.1:8000`
- Interactive API documentation: `http://127.0.0.1:8000/docs`
- Health check: `http://127.0.0.1:8000/health`

The first OCR run may take longer because the required models may need to be downloaded.

### 3. Set Up the Frontend

Open a **second terminal** from the project root.

```bash
cd frontend
npm install
npm run dev
```

Open the local URL printed by Vite, typically:

`http://localhost:5173`

Keep the backend running while using the application.

## 🔄 How It Works

1. The user selects or drops an image into the frontend.
2. The frontend validates the file and sends it to the backend.
3. FastAPI receives the image and passes it to the OCR and document-structure pipeline.
4. PaddleOCR extracts the text and identifies supported document elements.
5. The extracted structure is converted into a Word document using `python-docx`.
6. The generated `.docx` file is returned to the frontend for download.

## 📁 Supported Image Formats

The frontend is configured to accept:

- PNG
- JPG / JPEG
- BMP
- TIFF

**Maximum upload size:** 15 MB, subject to backend validation and hosting limits.

## 🔌 API

### `GET /health`

Checks whether the backend is running.

Example response:

```json
{
  "status": "ok"
}
```

### `POST /convert`

Accepts an image and returns a generated Word document.

Use the interactive API documentation at `http://127.0.0.1:8000/docs` to inspect and test the endpoint.

## ⚠️ Limitations

- OCR accuracy depends on image quality, resolution, language, and text clarity.
- Complex layouts, merged cells, handwritten text, and unusual formatting may not be reproduced perfectly.
- The output is editable, but it may not exactly match the original image's layout.
- Large images and document-structure analysis can require substantial memory and processing time.
- Public use requires a deployed backend; running the frontend alone does not enable OCR conversion.

## 🚀 Deployment

The frontend can be deployed separately to a hosting platform such as [Vercel](https://vercel.com/).

The OCR backend requires a compatible Python hosting environment with sufficient memory and processing resources. The frontend must be configured to call the public backend URL instead of a local address.

Public backend hosting and free-tier compatibility are still to be configured.

## 🔒 Privacy

Images uploaded for conversion are processed by the backend. Avoid uploading sensitive or confidential documents unless you understand the deployment's data-handling and retention policies.

## 🤝 Contributing

Contributions, suggestions, and bug reports are welcome.

1. Fork the repository.
2. Create a feature branch.
3. Commit your changes.
4. Open a pull request.

---

**Built by [Rajat Sharma](https://github.com/Rajat-91)**
