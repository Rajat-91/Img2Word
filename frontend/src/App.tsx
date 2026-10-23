import { useRef, useState } from "react";
import type { ChangeEvent, CSSProperties, DragEvent } from "react";
import "./App.css";

const API_URL = "http://127.0.0.1:8000/convert";
const MAX_FILE_SIZE = 15 * 1024 * 1024;
const ALLOWED_TYPES = ["image/png", "image/jpeg", "image/bmp", "image/tiff"];
const OUTPUT_NAME = "Img2Word_Converted.docx";

// Placeholder page layout for the Word panel: widths in %, 0 = paragraph gap.
const SHEET_LINES = [52, 100, 96, 100, 84, 0, 100, 92, 97, 60, 0, 100, 94, 78, 0, 98, 88];

function formatSize(bytes: number) {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function saveBlob(blob: Blob) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = OUTPUT_NAME;
  document.body.appendChild(link);
  link.click();
  link.remove();

  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/* ---------- Icons ---------- */

function DocIcon() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M7 3.75h7l5 5v11.5H7a2 2 0 0 1-2-2v-12a2 2 0 0 1 2-2Z" stroke="currentColor" strokeWidth="1.7" strokeLinejoin="round" />
      <path d="M14 3.75v5h5M8.5 13h7M8.5 16.5h7" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
    </svg>
  );
}

function UploadIcon() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M12 16V4m0 0L7.5 8.5M12 4l4.5 4.5" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" />
      <path d="M5 14.5v3.75A1.75 1.75 0 0 0 6.75 20h10.5A1.75 1.75 0 0 0 19 18.25V14.5" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
    </svg>
  );
}

function ChevronIcon() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="m9 5 7 7-7 7" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function CheckIcon() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="m5 12.5 4.5 4.5L19 7.5" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function AlertIcon() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M12 6v7.5M12 17.5v.5" stroke="currentColor" strokeWidth="2.6" strokeLinecap="round" />
    </svg>
  );
}

function LockIcon() {
  return (
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <rect x="5" y="10" width="14" height="10" rx="2" stroke="currentColor" strokeWidth="1.6" />
      <path d="M8 10V7a4 4 0 1 1 8 0v3m-4 4v2" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
    </svg>
  );
}

/* ---------- App ---------- */

function App() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [previewUrl, setPreviewUrl] = useState("");
  const [resultBlob, setResultBlob] = useState<Blob | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");

  const phase = isConverting ? "working" : success ? "done" : "idle";

  function selectFile(file: File) {
    setError("");
    setSuccess("");

    if (!ALLOWED_TYPES.includes(file.type)) {
      setError("Please select a PNG, JPG, JPEG, BMP, or TIFF image.");
      return;
    }

    if (file.size > MAX_FILE_SIZE) {
      setError("Your image must be 15 MB or smaller.");
      return;
    }

    setPreviewUrl((previousUrl) => {
      if (previousUrl) URL.revokeObjectURL(previousUrl);
      return URL.createObjectURL(file);
    });

    setResultBlob(null);
    setSelectedFile(file);
  }

  function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];

    if (file) selectFile(file);

    event.target.value = "";
  }

  function handleDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    setIsDragging(false);

    const file = event.dataTransfer.files[0];

    if (file) selectFile(file);
  }

  function removeFile() {
    if (previewUrl) URL.revokeObjectURL(previewUrl);

    setPreviewUrl("");
    setSelectedFile(null);
    setResultBlob(null);
    setError("");
    setSuccess("");
  }

  async function convertImage() {
    if (!selectedFile || isConverting) return;

    setIsConverting(true);
    setError("");
    setSuccess("");
    setResultBlob(null);

    try {
      const formData = new FormData();
      formData.append("file", selectedFile);

      const response = await fetch(API_URL, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        let message = "Conversion failed. Please try again.";

        try {
          const result = await response.json();
          if (result.detail) message = result.detail;
        } catch {
          // Keep the default error message.
        }

        throw new Error(message);
      }

      const contentType = response.headers.get("content-type") || "";

      if (!contentType.includes("application/vnd.openxmlformats-officedocument.wordprocessingml.document")) {
        throw new Error("The server did not return a Word document.");
      }

      const documentBlob = await response.blob();

      saveBlob(documentBlob);
      setResultBlob(documentBlob);
      setSuccess("Conversion complete. Your Word document has been downloaded.");
    } catch (conversionError) {
      if (conversionError instanceof TypeError) {
        setError("Could not connect to the OCR server. Make sure the backend is running.");
      } else {
        setError(
          conversionError instanceof Error
            ? conversionError.message
            : "Something went wrong. Please try again.",
        );
      }
    } finally {
      setIsConverting(false);
    }
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <a className="brand" href="/" aria-label="Img2Word home">
          <span className="brand-icon">
            <DocIcon />
          </span>
          <span>Img2Word</span>
        </a>

        <div className="topbar-right">
          <span className="topbar-note">Free image to Word converter</span>
          <a className="topbar-link" href="#how-it-works">
            How it works
          </a>
        </div>
      </header>

      <main>
        <section className="hero">
          <h1>Turn images into Word documents you can edit.</h1>
          <p className="hero-description">
            Upload a photo or scan. OCR reads the text and hands you a .docx file, ready to edit in Word.
          </p>
        </section>

        <section className="workbench" data-phase={phase} aria-label="Image to Word converter">
          <div className="panels">
            <div className="panel">
              <div className="panel-head">
                <h2>Your image</h2>
                <span className="panel-meta">PNG, JPG, BMP, TIFF</span>
              </div>

              {!selectedFile ? (
                <div
                  className={`stage source-stage brackets upload-zone ${isDragging ? "dragging" : ""}`}
                  onDragOver={(event) => {
                    event.preventDefault();
                    setIsDragging(true);
                  }}
                  onDragLeave={(event) => {
                    if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
                      setIsDragging(false);
                    }
                  }}
                  onDrop={handleDrop}
                  onClick={() => fileInputRef.current?.click()}
                  onKeyDown={(event) => {
                    if (event.key === "Enter" || event.key === " ") {
                      event.preventDefault();
                      fileInputRef.current?.click();
                    }
                  }}
                  role="button"
                  tabIndex={0}
                  aria-label="Choose an image to convert"
                >
                  <div className="upload-icon">
                    <UploadIcon />
                  </div>

                  <h3>Drop your image here</h3>
                  <p>or click to browse files</p>

                  <span className="browse-button">Choose image</span>

                  <span className="size-hint">Up to 15 MB</span>
                </div>
              ) : (
                <div className="stage source-stage brackets preview-stage">
                  {previewUrl && <img src={previewUrl} alt="Selected image preview" />}
                  {isConverting && <div className="beam" aria-hidden="true" />}
                </div>
              )}
            </div>

            <div className="connector" aria-hidden="true">
              <ChevronIcon />
            </div>

            <div className="panel">
              <div className="panel-head">
                <h2>Word document</h2>
                <span className="ext-badge">.DOCX</span>
              </div>

              <div className="stage doc-stage" aria-hidden="true">
                <div className="sheet">
                  {SHEET_LINES.map((width, index) =>
                    width === 0 ? (
                      <span className="sheet-gap" key={index} />
                    ) : (
                      <span
                        className={`sheet-line ${index === 0 ? "is-title" : ""}`}
                        key={index}
                        style={{ width: `${width}%`, "--i": index } as CSSProperties}
                      />
                    ),
                  )}
                </div>

                {phase === "done" && (
                  <span className="ready-pill">
                    <CheckIcon />
                    Ready
                  </span>
                )}
              </div>
            </div>
          </div>

          <input
            ref={fileInputRef}
            className="hidden-input"
            type="file"
            accept=".png,.jpg,.jpeg,.bmp,.tif,.tiff,image/png,image/jpeg,image/bmp,image/tiff"
            onChange={handleFileChange}
          />

          {error && (
            <div className="message error-message" role="alert">
              <span className="message-icon">
                <AlertIcon />
              </span>
              <span className="message-text">{error}</span>
            </div>
          )}

          {success && (
            <div className="message success-message" role="status">
              <span className="message-icon">
                <CheckIcon />
              </span>
              <span className="message-text">{success}</span>
              {resultBlob && (
                <button className="message-action" type="button" onClick={() => saveBlob(resultBlob)}>
                  Download again
                </button>
              )}
            </div>
          )}

          <div className="action-bar">
            <div className="file-info">
              {selectedFile ? (
                <>
                  <div className="file-text">
                    <strong title={selectedFile.name}>{selectedFile.name}</strong>
                    <span>{formatSize(selectedFile.size)}</span>
                  </div>

                  <button
                    className="remove-button"
                    type="button"
                    onClick={removeFile}
                    disabled={isConverting}
                    aria-label="Remove selected image"
                  >
                    Remove
                  </button>
                </>
              ) : (
                <span className="file-empty">No image selected yet.</span>
              )}
            </div>

            <button
              className={`convert-button ${isConverting ? "is-busy" : ""}`}
              type="button"
              onClick={convertImage}
              disabled={!selectedFile || isConverting}
            >
              {isConverting ? (
                <>
                  <span className="spinner" />
                  Extracting text...
                </>
              ) : (
                <>
                  <DocIcon />
                  Convert to Word
                </>
              )}
            </button>
          </div>
        </section>

        <p className="privacy-note">
          <LockIcon />
          Your image is processed by the local OCR server.
        </p>

        <section className="steps-section" id="how-it-works">
          <h2>Three steps. One editable document.</h2>

          <ol className="steps-list">
            <li className="step-item">
              <span className="step-number">1</span>
              <div>
                <h3>Upload an image</h3>
                <p>Choose a photo or scanned image containing text.</p>
              </div>
            </li>

            <li className="step-item">
              <span className="step-number">2</span>
              <div>
                <h3>Extract the text</h3>
                <p>OCR recognises the text in your image.</p>
              </div>
            </li>

            <li className="step-item">
              <span className="step-number">3</span>
              <div>
                <h3>Download your Word file</h3>
                <p>Open and edit the extracted text in Microsoft Word.</p>
              </div>
            </li>
          </ol>
        </section>
      </main>

      <footer className="footer">
        <a className="footer-brand" href="/">
          Img2Word
        </a>
        <span>Make your text editable.</span>
      </footer>
    </div>
  );
}

export default App;
