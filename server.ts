/* eslint-disable @typescript-eslint/no-explicit-any */
import express from "express";
import path from "path";
import cors from "cors";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI } from "@google/genai";
import multer from "multer";

const upload = multer({ storage: multer.memoryStorage() });

async function fetchWithRetry(url: string, retries = 3, delay = 1000): Promise<Response> {
  const response = await fetch(url);
  
  if (response.status === 429 && retries > 0) {
    await new Promise(resolve => setTimeout(resolve, delay));
    return fetchWithRetry(url, retries - 1, delay * 2);
  }
  return response;
}

function cleanPrivateKey(key?: string): string | undefined {
  if (!key) return undefined;
  let cleaned = key.trim();
  if ((cleaned.startsWith('"') && cleaned.endsWith('"')) || (cleaned.startsWith("'") && cleaned.endsWith("'"))) {
    cleaned = cleaned.slice(1, -1).trim();
  }
  let formatted = cleaned.replace(/\\n/g, '\n');
  if (!formatted.includes('\n')) {
    const pemMatch = formatted.match(/(-----BEGIN [A-Z ]+-----)(.*?)(-----END [A-Z ]+-----)/);
    if (pemMatch) {
      const header = pemMatch[1];
      const body = pemMatch[2].trim().replace(/\s+/g, '\n');
      const footer = pemMatch[3];
      formatted = `${header}\n${body}\n${footer}`;
    }
  }
  return formatted;
}

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json({ limit: "50mb" }));

// API routes FIRST
app.get("/api/health", (req, res) => {
  res.json({ status: "ok" });
});

app.post("/api/gs-export", async (req, res) => {
  const { url } = req.body;
  console.log(`[API] gs-export called for: ${url}`);
  try {
    if (!url || typeof url !== "string") {
      console.warn("[API] No URL provided in body");
      return res.status(400).json({ error: "No URL provided" });
    }

    let spreadsheetId = "";
    
    // Handle Published to Web URLs (2PACX-...)
    if (url.includes("docs.google.com/spreadsheets/d/e/")) {
      const pubMatch = url.match(/\/d\/e\/([a-zA-Z0-9-_]{20,})/);
      if (pubMatch) {
        const pubId = pubMatch[1];
        const pubExportUrl = `https://docs.google.com/spreadsheets/d/e/${pubId}/pub?output=csv`;
        try {
          const resPub = await fetchWithRetry(pubExportUrl);
          if (resPub.ok) {
            const txt = await resPub.text();
            res.header("Content-Type", "text/csv; charset=utf-8");
            return res.send(txt);
          }
        } catch (e) {
          console.warn("Published sheet fetch failed:", e);
        }
      }
    }

    const dMatch = url.match(/\/d\/([a-zA-Z0-9-_]{15,})/);
    const idMatch = url.match(/[?&]id=([a-zA-Z0-9-_]{15,})/);
    
    if (dMatch) {
      spreadsheetId = dMatch[1];
    } else if (idMatch) {
      spreadsheetId = idMatch[1];
    } else if (url.match(/^[a-zA-Z0-9-_]{15,}$/)) {
      spreadsheetId = url;
    }

    if (!spreadsheetId) return res.status(400).json({ error: "Invalid Google Sheet URL" });
    
    let gid = "0";
    const gidMatch = url.match(/[#&?]gid=([0-9]+)/);
    if (gidMatch) {
      gid = gidMatch[1];
    }

    const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=${gid}`;
    
    let responseOk = false;
    let csvText = "";

    try {
      console.log(`[API] Attempting public fetch for: ${exportUrl}`);
      const response = await fetchWithRetry(exportUrl);
      if (response.ok) {
        csvText = await response.text();
        responseOk = true;
        // Check if it's an HTML login page instead of CSV
        if (csvText.trim().toLowerCase().startsWith("<!doctype html>")) {
          console.log("[API] Public fetch returned HTML (login page), will fallback to credentials.");
          responseOk = false;
        } else {
          console.log("[API] Public fetch successful.");
        }
      } else {
        console.log(`[API] Public fetch returned status ${response.status}, proceeding with credentials fallback...`);
      }
    } catch (err) {
      console.warn("[API] Public fetch exception:", err);
    }

    // Fallback to Service Account if public fetch failed or returned HTML
    if (!responseOk) {
      const { google } = await import("googleapis");

      const auth = new google.auth.GoogleAuth({
        credentials: {
          client_email: process.env.GOOGLE_CLIENT_EMAIL,
          private_key: cleanPrivateKey(process.env.GOOGLE_PRIVATE_KEY),
        },
        scopes: ['https://www.googleapis.com/auth/drive.readonly', 'https://www.googleapis.com/auth/spreadsheets.readonly'],
      });
      const drive = google.drive({ version: 'v3', auth });

      try {
        // Check file metadata first to handle Shared Drives and determine mimeType
        let fileMetadata: any = null;
        try {
          const meta = await drive.files.get({
            fileId: spreadsheetId,
            fields: 'id, name, mimeType, size, shortcutDetails',
            supportsAllDrives: true
          });
          fileMetadata = meta.data;
          
          // Resolve shortcut
          if (fileMetadata.mimeType === 'application/vnd.google-apps.shortcut' && fileMetadata.shortcutDetails?.targetId) {
            console.log(`[API] Resolving shortcut for ${fileMetadata.name} -> ${fileMetadata.shortcutDetails.targetId}`);
            spreadsheetId = fileMetadata.shortcutDetails.targetId;
            // Fetch target metadata
            const targetMeta = await drive.files.get({
              fileId: spreadsheetId,
              fields: 'id, name, mimeType, size',
              supportsAllDrives: true
            });
            fileMetadata = targetMeta.data;
          }
          
          console.log(`[API] File metadata found: ${fileMetadata.name} (${fileMetadata.mimeType})`);
        } catch (metaErr: any) {
          console.warn(`[API] Could not fetch file metadata for ID ${spreadsheetId}: ${metaErr.message}`);
        }

        const isGoogleDoc = fileMetadata?.mimeType?.startsWith('application/vnd.google-apps.');
        const isSpreadsheet = fileMetadata?.mimeType === 'application/vnd.google-apps.spreadsheet';

        if (isSpreadsheet) {
          try {
            const sheets = google.sheets({ version: 'v4', auth: auth });
            const ss = await sheets.spreadsheets.get({ spreadsheetId });
            const sheetsList = ss.data.sheets || [];
            
            console.log(`[API] Sheets found in ${fileMetadata.name || spreadsheetId}:`, sheetsList.map(s => s.properties?.title));
            
            // Tìm tất cả các sheet hiển thị trong file có sheet nào Roster hay Q_Roster không
            let targetSheet = sheetsList.find(s => {
              const title = s.properties?.title || "";
              return title === "Roster" || title === "Q_Roster";
            });

            if (!targetSheet) {
              targetSheet = sheetsList.find(s => {
                const title = (s.properties?.title || "").toLowerCase();
                return title === "roster" || title === "q_roster" || title.includes("roster") || title.includes("q_roster");
              });
            }

            const chosenSheet = targetSheet || sheetsList[0];
            const sheetTitle = chosenSheet?.properties?.title || "Sheet1";
            console.log(`[API] Chosen sheet to export: "${sheetTitle}"`);

            // Fetch values using Sheets API and convert to CSV to ensure perfect export of correct sheet
            const valuesRes = await sheets.spreadsheets.values.get({
              spreadsheetId,
              range: sheetTitle,
              valueRenderOption: 'FORMATTED_VALUE',
            });
            const rows = valuesRes.data.values || [];
            
            // Convert 2D array to CSV
            csvText = rows.map(row => 
              row.map(val => {
                const stringVal = val === null || val === undefined ? '' : String(val);
                if (stringVal.includes(',') || stringVal.includes('\n') || stringVal.includes('\r') || stringVal.includes('"')) {
                  return `"${stringVal.replace(/"/g, '""')}"`;
                }
                return stringVal;
              }).join(',')
            ).join('\n');
            
            responseOk = true;
          } catch (sheetErr: any) {
            console.warn("[API] Failed fetching via Sheets API, falling back to Drive export:", sheetErr.message);
            // Fallback to Drive export
            const resDrive = await drive.files.export({
              fileId: spreadsheetId,
              mimeType: 'text/csv'
            }, { responseType: 'text' });
            
            if (resDrive.data) {
              csvText = resDrive.data as unknown as string;
              responseOk = true;
            } else {
              throw new Error("Empty response from Drive API export fallback");
            }
          }
        } else if (isGoogleDoc) {
          try {
            const resDrive = await drive.files.export({
              fileId: spreadsheetId,
              mimeType: 'text/csv'
            }, { responseType: 'text' });
            
            if (resDrive.data) {
              csvText = resDrive.data as unknown as string;
              responseOk = true;
            } else {
              throw new Error("Empty response from Drive API export");
            }
          } catch (exportErr: any) {
            console.error("Drive API export error:", exportErr.response?.data || exportErr.message);
            throw exportErr;
          }
        } else {
          // Non-Google Doc (e.g. .xlsx, .csv file uploaded to drive)
          try {
            const resMedia = await drive.files.get({
              fileId: spreadsheetId,
              alt: 'media',
              supportsAllDrives: true
            }, { responseType: 'arraybuffer' });
            
            if (resMedia.data) {
              // If it was already a CSV file, convert to string
              if (fileMetadata?.mimeType === 'text/csv') {
                const decoder = new TextDecoder('utf-8');
                csvText = decoder.decode(resMedia.data as any);
                responseOk = true;
              } else {
                // It's likely an Excel file, return as binary
                res.header("Content-Type", fileMetadata?.mimeType || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                return res.send(Buffer.from(resMedia.data as any));
              }
            }
          } catch (mediaErr: any) {
            console.error("Direct media download error:", mediaErr.response?.data || mediaErr.message);
            throw mediaErr;
          }
        }
      } catch (driveErr: any) {
         console.error("[API] Drive API error details:", JSON.stringify(driveErr.response?.data || driveErr, null, 2));
         
         let errorMsg = driveErr.message;
         const googleError = driveErr.response?.data?.error;
         if (googleError && googleError.message) {
           errorMsg = googleError.message;
         }

         if (errorMsg && (errorMsg.includes("invalid_grant") || errorMsg.toLowerCase().includes("jwt signature"))) {
           const serviceAccountEmail = process.env.GOOGLE_CLIENT_EMAIL || "email của bạn";
           
           throw new Error(`LỖI KHÓA LIÊN KẾT GOOGLE (Invalid JWT Signature / invalid_grant)\n\nKhóa bảo mật trong biến môi trường hiện tại đã BỊ HẾT HẠN, BỊ XÓA hoặc BỊ THU HỒI trên Google Cloud Console.\n\nCÁCH KHẮC PHỤC NHANH:\n1. Truy cập vào Google Cloud Console (https://console.cloud.google.com).\n2. Vào mục IAM & Admin -> Service Accounts (Tài khoản dịch vụ).\n3. Chọn tài khoản dịch vụ của bạn (Ví dụ: ${serviceAccountEmail}).\n4. Nhấp vào tab "Keys" (Khóa), bấm "Add Key" -> "Create new key" -> Chọn định dạng JSON rồi tải về.\n5. Mở file JSON vừa tải về, copy giá trị của "client_email" và "private_key".\n6. Cập nhật các biến môi trường GOOGLE_CLIENT_EMAIL và GOOGLE_PRIVATE_KEY trong phần Settings của ứng dụng.\n7. Thử bấm "Đồng bộ" lại.`);
         }

         if (errorMsg && (errorMsg.toLowerCase().includes("file not found") || errorMsg.toLowerCase().includes("forbidden") || driveErr.status === 404 || driveErr.status === 403)) {
           const serviceAccountEmail = process.env.GOOGLE_CLIENT_EMAIL || "email của bạn";
           
           throw new Error(`BẠN CHƯA CẤP QUYỀN TRUY CẬP cho file/sheet này.\n\nHÃY LÀM THEO CÁC BƯỚC SAU:\n1. Mở file/thư mục trên Google Drive.\n2. Bấm nút "Share" (Chia sẻ).\n3. Copy và dán email sau vào ô người nhận:\n👉 ${serviceAccountEmail}\n4. Chọn quyền "Viewer" (Người xem) và bấm "Send" (Gửi).`);
         }
         
         throw new Error(`[Google Drive Error] ${errorMsg}`);
      }
    }
    
    if (!responseOk) {
      throw new Error("Không thể tải Google Sheet.");
    }
    
    res.header("Content-Type", "text/csv; charset=utf-8");
    res.send(csvText);
  } catch (error: any) {
    console.error("[API] gs-export error:", error);
    res.status(500).json({ error: `[Server Error] ${error.message}` });
  }
});

app.get("/api/drive-folder-files", async (req, res) => {
  const serviceAccountEmail = process.env.GOOGLE_CLIENT_EMAIL || "email của bạn";
  try {
    const { folderId } = req.query;
    if (!folderId || typeof folderId !== "string") {
      return res.status(400).json({ error: "No folderId provided" });
    }

    const { google } = await import("googleapis");
    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: process.env.GOOGLE_CLIENT_EMAIL,
        private_key: cleanPrivateKey(process.env.GOOGLE_PRIVATE_KEY),
      },
      scopes: [
        'https://www.googleapis.com/auth/spreadsheets.readonly',
        'https://www.googleapis.com/auth/drive.readonly'
      ],
    });

    const drive = google.drive({ version: 'v3', auth });
    
    const response = await drive.files.list({
      q: `'${folderId}' in parents and trashed = false and mimeType != 'application/vnd.google-apps.folder'`,
      fields: 'files(id, name, mimeType, shortcutDetails)',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });
    
    const files = (response.data.files || []).map((f: any) => {
      if (f.mimeType === 'application/vnd.google-apps.shortcut' && f.shortcutDetails?.targetId) {
        return {
          ...f,
          id: f.shortcutDetails.targetId,
          _isShortcut: true
        };
      }
      return f;
    }).filter((f: any) => 
      !f.name?.toLowerCase().includes("copy")
    );

    res.json({ 
      success: true,
      totalFiles: files.length,
      files: files 
    });

  } catch (error: any) {
    console.error('Lỗi khi gọi Google API:', error);
    let errorMsg = 'Lỗi khi lấy danh sách file từ Drive: ' + error.message;
    
    if (error.message && (error.message.includes('invalid_grant') || error.message.toLowerCase().includes('jwt signature'))) {
      errorMsg = `LỖI KHÓA LIÊN KẾT GOOGLE (Invalid JWT Signature / invalid_grant)\n\nKhóa bảo mật trong biến môi trường hiện tại đã BỊ HẾT HẠN, BỊ XÓA hoặc BỊ THU HỒI trên Google Cloud Console.\n\nCÁCH KHẮC PHỤC NHANH:\n1. Truy cập vào Google Cloud Console (https://console.cloud.google.com).\n2. Vào mục IAM & Admin -> Service Accounts (Tài khoản dịch vụ).\n3. Chọn tài khoản dịch vụ của bạn (Ví dụ: ${serviceAccountEmail}).\n4. Nhấp vào tab "Keys" (Khóa), bấm "Add Key" -> "Create new key" -> Chọn định dạng JSON rồi tải về.\n5. Mở file JSON vừa tải về, copy giá trị của "client_email" và "private_key".\n6. Cập nhật các biến môi trường GOOGLE_CLIENT_EMAIL và GOOGLE_PRIVATE_KEY trong phần Settings của ứng dụng.\n7. Thử bấm "Đồng bộ" lại.`;
    } else if (error.message && error.message.toLowerCase().includes('file not found')) {
      errorMsg = `Không tìm thấy thư mục trên Google Drive.\n\nLÝ DO PHỔ BIẾN:\n1. Link/ID thư mục không chính xác.\n2. BẠN CHƯA CẤP QUYỀN TRUY CẬP: Bạn phải vào Google Drive, bấm Share (Chia sẻ) thư mục đó cho email sau với quyền Viewer (Người xem):\n👉 ${serviceAccountEmail}`;
    }
    
    res.status(500).json({ error: errorMsg });
  }
});

app.post("/api/analyze-image", upload.single("image"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No image file provided" });
    }

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return res.status(500).json({ error: "GEMINI_API_KEY is not configured" });
    }

    const ai = new GoogleGenAI({ apiKey });
    
    const filePart = {
      inlineData: {
        data: req.file.buffer.toString("base64"),
        mimeType: req.file.mimetype,
      },
    };

    const prompt = req.body.prompt || "Phân tích chi tiết hình ảnh này.";

    const response = await ai.models.generateContent({
      model: "gemini-3.1-pro-preview",
      contents: [
        prompt,
        filePart,
      ],
    });

    res.json({ result: response.text() });
  } catch (error: any) {
    console.error("Gemini Error:", error);
    res.status(500).json({ error: error.message || "Failed to analyze image" });
  }
});

app.post("/api/chat", async (req, res) => {
  try {
    const { messages, model = "gemini-3.1-flash-lite", systemInstruction } = req.body;
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) return res.status(500).json({ error: "GEMINI_API_KEY is not configured" });

    const ai = new GoogleGenAI({ apiKey });
    const response = await ai.models.generateContent({
      model,
      contents: messages,
      config: { systemInstruction },
    });
    res.json({ text: response.text() });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/generate-image", async (req, res) => {
  try {
    const { prompt, aspectRatio = "1:1", model = "gemini-3.1-flash-image-preview" } = req.body;
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) return res.status(500).json({ error: "GEMINI_API_KEY is not configured" });

    const ai = new GoogleGenAI({ apiKey });
    const response = await ai.models.generateImages({
      model,
      prompt,
      config: { aspectRatio, numberOfImages: 1, outputMimeType: "image/jpeg" },
    });
    res.json({ imageBase64: response.generatedImages[0].image.imageBytes });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

async function startServer() {
  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

if (process.env.NODE_ENV !== "production" || !process.env.VERCEL) {
  startServer();
}

export default app;

