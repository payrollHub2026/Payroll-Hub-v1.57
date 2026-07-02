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

  let serviceAccountEmail = "email trong file credentials.json";
  try {
    const { folderId } = req.query;
    if (!folderId || typeof folderId !== "string") {
      return res.status(400).json({ error: "No folderId provided" });
    }

    let auth;
    const hasEnvCreds = process.env.GOOGLE_CLIENT_EMAIL && process.env.GOOGLE_PRIVATE_KEY;
    
    if (hasEnvCreds) {
      serviceAccountEmail = process.env.GOOGLE_CLIENT_EMAIL!;
      auth = new google.auth.GoogleAuth({
        credentials: {
          client_email: process.env.GOOGLE_CLIENT_EMAIL,
          private_key: cleanPrivateKey(process.env.GOOGLE_PRIVATE_KEY),
        },
        scopes: [
          'https://www.googleapis.com/auth/spreadsheets.readonly',
          'https://www.googleapis.com/auth/drive.readonly'
        ],
      });
    } else {
      const credsPath = path.join(process.cwd(), "credentials.json");
      
      if (!fs.existsSync(credsPath)) {
        return res.status(500).json({ 
          error: "Không tìm thấy cấu hình kết nối Google. Vui lòng thêm biến môi trường GOOGLE_CLIENT_EMAIL và GOOGLE_PRIVATE_KEY trong phần cài đặt của Vercel (hoặc upload file credentials.json vào thư mục gốc)."
        });
      }

      let credsJson: any = {};
      try {
        const credsContent = fs.readFileSync(credsPath, 'utf8');
        credsJson = JSON.parse(credsContent);
        if (credsJson.client_email) {
          serviceAccountEmail = credsJson.client_email;
        }
      } catch (err: any) {
        console.error("Lỗi khi đọc file credentials.json:", err);
      }

      auth = new google.auth.GoogleAuth({
        keyFile: credsPath,
        scopes: [
          'https://www.googleapis.com/auth/spreadsheets.readonly',
          'https://www.googleapis.com/auth/drive.readonly'
        ],
      });
    }

    const client = await auth.getClient();
    const drive = google.drive({ version: 'v3', auth: client as any });
    
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
    
    return res.status(200).json({ 
      success: true,
      totalFiles: files.length,
      files: files 
    });

  } catch (error: any) {
    console.error('Lỗi khi gọi Google API:', error);
    let errorMsg = 'Lỗi khi lấy danh sách file từ Drive: ' + error.message;
    
    if (error.message && (error.message.includes('invalid_grant') || error.message.toLowerCase().includes('jwt signature'))) {
      errorMsg = `LỖI KHÓA LIÊN KẾT GOOGLE (Invalid JWT Signature / invalid_grant)\n\nKhóa bảo mật trong file credentials.json hiện tại đã BỊ HẾT HẠN, BỊ XÓA hoặc BỊ THU HỒI trên Google Cloud Console.\n\nCÁCH KHẮC PHỤC NHANH:\n1. Truy cập vào Google Cloud Console (https://console.cloud.google.com).\n2. Vào mục IAM & Admin -> Service Accounts (Tài khoản dịch vụ).\n3. Chọn tài khoản dịch vụ của bạn (Ví dụ: ${serviceAccountEmail}).\n4. Nhấp vào tab "Keys" (Khóa), bấm "Add Key" -> "Create new key" -> Chọn định dạng JSON rồi tải về.\n5. Đổi tên file vừa tải về thành "credentials.json" (phải viết thường chính xác).\n6. Kéo thả hoặc upload đè file "credentials.json" mới này vào cột thư mục bên trái phần mềm.\n7. Thử bấm "Đồng bộ" lại.`;
    } else if (error.message && error.message.toLowerCase().includes('file not found')) {
      errorMsg = `Không tìm thấy thư mục trên Google Drive.\n\nLÝ DO PHỔ BIẾN:\n1. Link/ID thư mục không chính xác.\n2. BẠN CHƯA CẤP QUYỀN TRUY CẬP: Bạn phải vào Google Drive, bấm Share (Chia sẻ) thư mục đó cho email sau với quyền Viewer (Người xem):\n👉 ${serviceAccountEmail}`;
    }
    
    return res.status(500).json({ error: errorMsg });
  }
}
