# Connect GitHub Webpage to This Machine

This setup lets a webpage upload a DOCX file, process it on this machine, and download XLSX.

## 1. Install dependencies

```powershell
& "C:/Program Files/Python314/python.exe" -m pip install -r requirements.txt
```

## 2. Run local API

Optional: restrict CORS to your GitHub Pages origin.

```powershell
$env:CORS_ORIGINS = "https://<your-user>.github.io"
& "C:/Program Files/Python314/python.exe" app.py
```

API endpoints:
- `GET /health`
- `POST /process` with multipart field `inputFile` (`.docx`)

## 3. Expose local API with ngrok

In another terminal:

```powershell
ngrok http 5000
```

Copy the forwarding URL (for example `https://xxxx.ngrok-free.app`).

## 4. Host frontend on GitHub Pages

Use [web/index.html](web/index.html) in your GitHub Pages site.

In the page:
- Enter your ngrok URL.
- Upload DOCX.
- Click convert and download XLSX.

## Notes
- Keep `app.py` running while using the webpage.
- If ngrok URL changes, update it in the page input.
- Large DOCX files may take longer depending on your machine.
