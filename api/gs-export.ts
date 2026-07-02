/* eslint-disable @typescript-eslint/no-explicit-any */
import path from "path";
import fs from "fs";
import { google } from "googleapis";

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

async function fetchWithRetry(url: string, retries = 3, delay = 1000): Promise<any> {
  const response = await fetch(url);
  
  if (response.status === 429 && retries > 0) {
    await new Promise(resolve => setTimeout(resolve, delay));
    return fetchWithRetry(url, retries - 1, delay * 2);
  }
  return response;
}

export default async function handler(req: any, res: any) {
  // Allow CORS
  res.setHeader("Access-Control-Allow-Credentials", "true");
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,OPTIONS,PATCH,DELETE,POST,PUT");
  res.setHeader(
    "Access-Control-Allow-Headers",
    "X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version"
  );

  if (req.method === "OPTIONS") {
    res.status(200).end();
    return;
  }

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
            res.setHeader("Content-Type", "text/csv; charset=utf-8");
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

    // Fallback to credentials if public fetch failed or returned HTML
    if (!responseOk) {
      const credsPath = path.join(process.cwd(), "credentials.json");
      let auth;
      const hasEnvCreds = process.env.GOOGLE_CLIENT_EMAIL && process.env.GOOGLE_PRIVATE_KEY;

      if (hasEnvCreds) {
        auth = new google.auth.GoogleAuth({
          credentials: {
            client_email: process.env.GOOGLE_CLIENT_EMAIL,
            private_key: cleanPrivateKey(process.env.GOOGLE_PRIVATE_KEY),
          },
          scopes: ['https://www.googleapis.com/auth/drive.readonly', 'https://www.googleapis.com/auth/spreadsheets.readonly'],
        });
      } else {
        if (!fs.existsSync(credsPath)) {
          throw new Error("Không thể tải Google Sheet. Vui lòng thiết lập biến môi trường GOOGLE_CLIENT_EMAIL và GOOGLE_PRIVATE_KEY trên Vercel, hoặc thêm file credentials.json.");
        }

        auth = new google.auth.GoogleAuth({
          keyFile: credsPath,
          scopes: ['https://www.googleapis.com/auth/drive.readonly', 'https://www.googleapis.com/auth/spreadsheets.readonly'],
        });
      }

      const client = await auth.getClient();
      const drive = google.drive({ version: 'v3', auth: client as any });

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
            const sheets = google.sheets({ version: 'v4', auth: client as any });
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
                res.setHeader("Content-Type", fileMetadata?.mimeType || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
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
           let serviceAccountEmail = "email trong file credentials.json";
           try {
             const credsContent = fs.readFileSync(credsPath, 'utf8');
             const credsJson = JSON.parse(credsContent);
             if (credsJson.client_email) {
               serviceAccountEmail = credsJson.client_email;
             }
           } catch { /* ignore */ }
           
           throw new Error(`LỖI KHÓA LIÊN KẾT GOOGLE (Invalid JWT Signature / invalid_grant)\n\nKhóa bảo mật trong file credentials.json hiện tại đã BỊ HẾT HẠN, BỊ XÓA hoặc BỊ THU HỒI trên Google Cloud Console.\n\nCÁCH KHẮC PHỤC NHANH:\n1. Truy cập vào Google Cloud Console (https://console.cloud.google.com).\n2. Vào mục IAM & Admin -> Service Accounts (Tài khoản dịch vụ).\n3. Chọn tài khoản dịch vụ của bạn (Ví dụ: ${serviceAccountEmail}).\n4. Nhấp vào tab "Keys" (Khóa), bấm "Add Key" -> "Create new key" -> Chọn định dạng JSON rồi tải về.\n5. Đổi tên file vừa tải về thành "credentials.json" (phải viết thường chính xác).\n6. Kéo thả hoặc upload đè file "credentials.json" mới này vào cột thư mục bên trái phần mềm.\n7. Thử bấm "Đồng bộ" lại.`);
         }

         if (errorMsg && (errorMsg.toLowerCase().includes("file not found") || errorMsg.toLowerCase().includes("forbidden") || driveErr.status === 404 || driveErr.status === 403)) {
           let serviceAccountEmail = "email trong file credentials.json";
           try {
             const credsContent = fs.readFileSync(credsPath, 'utf8');
             const credsJson = JSON.parse(credsContent);
             if (credsJson.client_email) {
               serviceAccountEmail = credsJson.client_email;
             }
           } catch { /* ignore */ }
           
           throw new Error(`BẠN CHƯA CẤP QUYỀN TRUY CẬP cho file/sheet này.\n\nHÃY LÀM THEO CÁC BƯỚC SAU:\n1. Mở file/thư mục trên Google Drive.\n2. Bấm nút "Share" (Chia sẻ).\n3. Copy và dán email sau vào ô người nhận:\n👉 ${serviceAccountEmail}\n4. Chọn quyền "Viewer" (Người xem) và bấm "Send" (Gửi).`);
         }
         
         throw new Error(`[Google Drive Error] ${errorMsg}`);
      }
    }
    
    if (!responseOk) {
      throw new Error("Không thể tải Google Sheet.");
    }
    
    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    return res.send(csvText);
  } catch (error: any) {
    console.error("[API] gs-export error:", error);
    return res.status(500).json({ error: `[Server Error] ${error.message}` });
  }
}
