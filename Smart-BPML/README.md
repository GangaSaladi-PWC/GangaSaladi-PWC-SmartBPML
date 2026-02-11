# XLSX upload demo (React + Express)

This repo shows a minimal setup where the frontend uploads an `.xlsx` file to an Express backend. The backend also exposes a static workbook you can download.

## Getting started

1) Install dependencies

```bash
cd server && npm install
cd client && npm install
```

2) Run the servers (two terminals)

```bash
# Backend on http://localhost:4000
cd server && npm run start

# Frontend on http://localhost:5173 (proxy to backend /api)
cd client && npm run dev
```

3) Open http://localhost:5173 and upload an `.xlsx` file.
   - "Upload" sends file to `/api/upload` and returns a sheet summary.
   - "Merge & download" sends file to `/api/merge`, finds rows where col1–col3 match between the uploaded file and the `KMD Repository - 2023` sheet in the static workbook, and downloads a merged `.xlsx` containing:
     - col1–col5 from the uploaded row
     - columns I, J, K from the matching static row
   - "Download static workbook" fetches the backend knowledgebase file (`server/static/Knowledge Base_KMD Repository (S4HANA 2023).xlsx`). Place your file there before running.

## Notes

- The backend accepts `.xlsx` only, parses it with `xlsx`, and returns sheet counts.
- A static workbook is created at startup in `server/static/sample.xlsx` if it does not exist.
- Vite dev server proxies `/api` to the backend, so no extra config is needed during development.

