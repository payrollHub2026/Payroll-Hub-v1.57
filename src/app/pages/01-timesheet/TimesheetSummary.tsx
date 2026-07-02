import { PuppyLogo } from "../../components/shared/PuppyLogo";
import React, { useState, useMemo, useEffect, useRef } from "react";
import {
  FileSpreadsheet,
  Download,
  Settings,
  RefreshCw,
  Trash2,
  ArrowLeft,
  Save,
  Copy,
  Plus,
} from "lucide-react";
import * as XLSX from "xlsx";
import { toast } from "sonner";
import { useAppData } from "../../lib/contexts/AppDataContext";
import { isSupabaseConfigured } from "../../lib/supabase";
import { syncRosterToSupabase, SQL_SETUP_SCRIPT } from "../../lib/supabase-sync-utils";
import { useTimesheetCalculations } from "../../hooks/useTimesheetCalculations";
import { getDynamicEmployeeColumns, CENTER_COLUMNS } from "../../constants/timesheet-columns";
import { TimesheetInputTable } from "./components/TimesheetInputTable";
import type { TimesheetInputRow } from "./components/TimesheetInputTable";
import {
  getL07FromFileName,
  getCenterInfoByL07,
  getCenterInfoByAECode,
  mapL07,
  getBusinessFromL07,
} from "../../lib/utils/center-utils";
import { 
  generateUUID, 
  prepareDataForExport,
  getVal,
  getExcelFileBuffer,
  fetchGoogleSheetAsFile,
} from "../../lib/utils/data-utils";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "../../components/ui/dropdown-menu";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
  DialogFooter,
} from "../../components/ui/dialog";
import { Button } from "../../components/ui/button";
import {
  Tooltip,
  TooltipTrigger,
  TooltipContent,
} from "../../components/ui/tooltip";

import ExcelWorker from "../../workers/excelParser.worker?worker&inline";

const parseExcelInWorker = async (file: File): Promise<Record<string, unknown>[]> => {
  const { buffer, name } = await getExcelFileBuffer(file);

  return new Promise((resolve, reject) => {
    try {
      const worker = new ExcelWorker();
      worker.onmessage = (e: MessageEvent) => {
        worker.terminate();
        if (e.data && e.data.success) {
          resolve(e.data.allRows as Record<string, unknown>[]);
        } else {
          reject(new Error(e.data?.error || "Unknown error parsing Excel file"));
        }
      };
      worker.onerror = (err) => {
        worker.terminate();
        reject(err);
      };
      worker.postMessage({ fileBuffer: buffer, fileName: name });
    } catch (err) {
      reject(err);
    }
  });
};

const mapExcelRosterRow = (row: Record<string, unknown>, fileName?: string, fileId?: string) => {
  const rawCenter = String(getVal(row, ["center", "mã ae", "ae", "ae code"]) || "").trim();
  const info = getCenterInfoByAECode(rawCenter);
  const l07 = info?.l07 || rawCenter || "UNKNOWN";
  const business = info?.bus || "";
  const ma_nv = String(getVal(row, ["id number", "id", "teacher id", "emp id", "mã nv", "manv", "code"]) || "").trim();
  const full_name = String(getVal(row, ["full name", "name", "teacher name", "tên", "họ và tên", "họ tên"]) || "").trim();
  const ngayRaw = getVal(row, ["date", "ngay", "ngày", "tk_date", "session date", "sessiondate", "ngày học", "scheduledate", "ngày làm việc", "ngày tháng"]);
  const ngay = ngayRaw !== undefined && ngayRaw !== null ? String(ngayRaw).trim() : "";
  const type = String(getVal(row, ["type", "type code", "type_code", "typecode", "task type", "task", "loại", "loại hoạt động", "event type", "activity", "category", "task type name", "taskType"]) || "").trim();
  const className = String(getVal(row, ["class", "class code", "class_code", "classcode", "lớp", "class name", "mã lớp", "tên lớp", "code", "mã lớp học", "classCode"]) || "").trim();
  const gio_vao = String(getVal(row, ["from", "start", "start time", "từ", "giờ bắt đầu"]) || "").trim();
  const gio_ra = String(getVal(row, ["to", "end", "end time", "đến", "giờ kết thúc"]) || "").trim();
  
  const rawDuration = getVal(row, ["duration", "quy ra số giờ làm", "total", "actual hours", "working hours", "giờ làm", "số giờ", "hours", "tk_duration", "total hours", "tổng giờ", "time", "thời lượng"]);
  let duration = 0;
  if (typeof rawDuration === "number") {
    duration = rawDuration;
  } else if (rawDuration) {
    const sv = String(rawDuration).trim().replace(",", ".");
    if (sv.includes(":")) {
      const p = sv.split(":");
      duration = (parseInt(p[0]) || 0) + (parseInt(p[1]) || 0) / 60;
    } else {
      duration = parseFloat(sv) || 0;
    }
  }
  
  const notes = String(getVal(row, ["notes", "note", "ghi chú", "ghi chu", "remarks"]) || "").trim().replace(/^["']|["']$/g, "");
  const chargeToCenterMkt = String(getVal(row, ["charge to center mkt", "charge to center", "chargetocenter"]) || "").trim();

  // DỮ LIỆU CỘT TYPE VÀ CLASS BỊ TRÁO CHO NHAU TẠI L07N = MKT LOCAL NORTH
  let finalType = type;
  let finalClass = className;
  const isMktN = l07 === "MKT LOCAL NORTH" || rawCenter.toUpperCase().includes("NORTH") || chargeToCenterMkt.toUpperCase().includes("NORTH");
  
  if (isMktN) {
    const tLow = type.toLowerCase();
    const cLow = className.toLowerCase();
    const isCost = (s: string) => s.startsWith("lpar") || s.startsWith("lret") || s.startsWith("ldem") || s.startsWith("ldec") || s.startsWith("moth");
    if (isCost(cLow) && !isCost(tLow)) {
      finalType = className;
      finalClass = type; 
    }
  }

  return {
    center: rawCenter,
    l07,
    business,
    ma_nv,
    full_name,
    ngay,
    type: finalType,
    class: finalClass,
    gio_vao,
    gio_ra,
    chargeToCenterMkt,
    duration,
    notes,
    
    employeeId: ma_nv,
    fullName: full_name,
    maAE: rawCenter,
    date: ngay,
    taskType: finalType,
    classCode: finalClass,
    from: gio_vao,
    to: gio_ra,
    _sourceFile: fileName || row._sourceFile || "",
    _rowId: fileId || row._rowId || "",
    _uuid: row._uuid || generateUUID()
  };
};

const DEFAULT_FOLDER_URL = "https://drive.google.com/drive/folders/1gU6Hcrv94Bx_yv1qNTqH0vQNy7ElKzXJ";

export default function TimesheetSummaryPage({
  onBack,
}: {
  onBack?: () => void;
}) {
  const { appData, updateAppData } = useAppData();

  const [activeTab] = useState<"files">("files");
  const [fromDate] = useState("");
  const [toDate] = useState("");
  const [debouncedFromDate, setDebouncedFromDate] = useState("");
  const [debouncedToDate, setDebouncedToDate] = useState("");

  const [isSyncing, setIsSyncing] = useState(false);
  const [isCalculating, setIsCalculating] = useState(false);
  const [syncProgress, setSyncProgress] = useState(0);
  const [totalSyncRows, setTotalSyncRows] = useState(0);
  const [syncedRowsCount, setSyncedRowsCount] = useState(0);
  const [showSqlDialog, setShowSqlDialog] = useState(false);

  const [isFetchingGgSheet, setIsFetchingGgSheet] = useState(false);
  const [, setRefreshKey] = useState(0);

  const handleUrlInput = async (id: string, url: string) => {
    if (!url.trim()) return;
    const isFolder = url.includes("folders/") || url.includes("drive/folders/") || url.includes("?id=");

    setIsFetchingGgSheet(true);
    try {
      if (isFolder) {
        let folderId = url.trim();
        const match = url.match(/folders\/([a-zA-Z0-9-_]+)/);
        if (match) {
          folderId = match[1];
        } else {
          try {
            const urlObj = new URL(url);
            if (urlObj.searchParams.has("id")) {
              folderId = urlObj.searchParams.get("id") || folderId;
            }
          } catch { /* ignore */ }
        }

        const response = await fetch(`/api/drive-folder-files?folderId=${encodeURIComponent(folderId)}`);
        if (!response.ok) {
          const errData = await response.json().catch(() => ({}));
          throw new Error(errData.error || "Không thể lấy danh sách file từ thư mục. Vui lòng kiểm tra lại link hoặc quyền chia sẻ.");
        }

        const data = await response.json();
        if (!data.success || !data.files || data.files.length === 0) {
          throw new Error("Không tìm thấy file nào trong thư mục này.");
        }

        const driveFiles = (data.files || []).filter((f: Record<string, unknown>) => {
          const name = String(f.name || "").toLowerCase();
          return !name.includes("copy");
        });

        if (driveFiles.length === 0 && data.files.length > 0) {
          throw new Error("Tất cả các file trong thư mục đều là file 'copy' nên hệ thống tự động bỏ qua.");
        }

        toast.info(`Tìm thấy ${driveFiles.length} file hợp lệ. Đang tự động đối chiếu và nạp dữ liệu...`);

        const currentInputs = [...(appData.Timesheet_InputList || [])];
        const toProcess: { id: string; file: File }[] = [];
        let successCount = 0;
        let skipCount = 0;

        for (const f of driveFiles) {
          const fileName = String(f.name || "");
          const l07 = getL07FromFileName(fileName) || "";
          if (!l07) {
            skipCount++;
            continue;
          }
          const centerInfo = getCenterInfoByL07(l07);
          const aeCode = centerInfo?.aeCode || "";
          const bu = getBusinessFromL07(l07);

          // Matching logic similar to bulk Excel upload
          let matchIndex = currentInputs.findIndex((r) => {
            const rowL07 = r.l07 ? mapL07(r.l07).toLowerCase() : "";
            const rowAE = r.aeCode ? r.aeCode.toLowerCase() : "";
            const matchL07 = l07 && rowL07 === l07.toLowerCase();
            const matchAE = aeCode && rowAE === aeCode.toLowerCase();
            return matchL07 || matchAE;
          });

          if (matchIndex === -1) {
            matchIndex = currentInputs.findIndex(r => !r.l07 && !r.fileName && (r.status === "pending" || r.status === "ready"));
          }

          let rowId: string;
          if (matchIndex !== -1) {
            rowId = currentInputs[matchIndex].id;
            currentInputs[matchIndex] = {
              ...currentInputs[matchIndex],
              l07: l07,
              aeCode: aeCode,
              bus: bu,
              status: "processing",
            };
          } else {
            rowId = crypto.randomUUID();
            currentInputs.push({
              id: rowId,
              l07: l07,
              aeCode: aeCode,
              bus: bu,
              status: "processing",
              url: ""
            });
          }

          const sheetUrl = `https://docs.google.com/spreadsheets/d/${f.id}`;
          const fileContent = JSON.stringify({ url: sheetUrl });
          const blob = new Blob([fileContent], { type: 'application/json' });
          let name = fileName;
          if (!name.toLowerCase().endsWith(".gsheet")) {
            name = name.replace(/\.(xlsx|xls|csv)$/i, "") + ".gsheet";
          }
          const fileObj = new File([blob], name, { type: 'application/json' });
          toProcess.push({ id: rowId, file: fileObj });
          successCount++;
        }

        if (successCount > 0) {
          // Set matched rows to a "ready" status first, but don't start processing yet
          const readyInputs = currentInputs.map(r => {
            const match = toProcess.find(tp => tp.id === r.id);
            if (match && r.status !== "success") {
              return { ...r, status: "ready" as const };
            }
            return r;
          });
          
          updateAppData(prev => ({ ...prev, Timesheet_InputList: readyInputs }), false);
          
          // Sequential processing with delay
          for (let i = 0; i < toProcess.length; i++) {
            const item = toProcess[i];
            
            // 1. Set individual row to processing for UI feedback
            handleUpdateRow(item.id, "status", "processing");
            
            // 2. Process the file (this includes the fetch)
            try {
              await handleUploadFile(item.id, item.file);
            } catch (err: unknown) {
              const msg = err instanceof Error ? err.message : String(err);
              console.error(`Failed to process ${item.file.name}:`, err);
              handleUpdateRow(item.id, "status", "error");
              toast.error(`Lỗi xử lý ${item.file.name}: ${msg}`);
            }
            
            // 3. Wait 1500ms before next file to avoid rate limits (except for the last one)
            if (i < toProcess.length - 1) {
              await new Promise(resolve => setTimeout(resolve, 1500));
            }
          }
          
          toast.success(`Đã nạp xong từ thư mục! Thành công: ${successCount} trung tâm${skipCount > 0 ? `, Bỏ qua: ${skipCount}` : ""}.`);
        } else {
          toast.warn(`Không tìm thấy trung tâm nào khớp với các file trong thư mục.`);
        }
      } else {
        const selectedRow = inputRows.find(r => r.id === id);
        const l07 = selectedRow?.l07 || "GoogleSheet";
        
        const file = await fetchGoogleSheetAsFile(url, `${l07}_GoogleSheet.gsheet`);
        await handleUploadFile(id, file);
        toast.success(`Đã nạp dữ liệu từ link cho trung tâm ${l07}!`);
      }
    } catch (error: unknown) {
      console.error(error);
      const msg = error instanceof Error ? error.message : "Lỗi xử lý link";
      toast.error(msg);
    } finally {
      setIsFetchingGgSheet(false);
    }
  };

  const lastSummaryRef = useRef("");

  useEffect(() => {
    const timer = setTimeout(() => {
      setDebouncedFromDate(fromDate);
      setDebouncedToDate(toDate);
    }, 500);
    return () => clearTimeout(timer);
  }, [fromDate, toDate]);

  useEffect(() => {
    if (!debouncedFromDate || !debouncedToDate) return;

    updateAppData((prev) => {
      if (
        prev.Timesheet_Dates?.from === debouncedFromDate &&
        prev.Timesheet_Dates?.to === debouncedToDate
      ) {
        return prev;
      }

      return {
        ...prev,
        Timesheet_Dates: { from: debouncedFromDate, to: debouncedToDate },
      };
    }, false);
  }, [debouncedFromDate, debouncedToDate, updateAppData]);



  const rosterData = useMemo(() => appData.Q_Roster || [], [appData.Q_Roster]);
  const salaryScaleData = useMemo(() => appData.Q_Salary_Scale || [], [appData.Q_Salary_Scale]);
  const staffData = useMemo(() => appData.Q_Staff || [], [appData.Q_Staff]);
  const cacheData = useMemo(() => appData.Q_Cache || [], [appData.Q_Cache]);
  const inputRows = useMemo(() => appData.Timesheet_InputList || [
    { id: "1", l07: "", aeCode: "", bus: "", url: "", status: "pending" },
  ], [appData.Timesheet_InputList]);

  const handleAddRow = () => {
    updateAppData((prev) => ({
      ...prev,
      Timesheet_InputList: [
        ...inputRows,
        {
          id: generateUUID(),
          l07: "",
          aeCode: "",
          bus: "",
          url: "",
          status: "pending",
        },
      ],
    }));
  };
  const handleUpdateRow = (
    id: string,
    field: keyof TimesheetInputRow,
    val: string | number | boolean | Record<string, unknown> | undefined,
  ) => {
    updateAppData(
      (prev) => ({
        ...prev,
        Timesheet_InputList: (prev.Timesheet_InputList || []).map((r) =>
          r.id === id ? { ...r, [field]: val } : r,
        ),
      }),
      false,
    );
  };
  const handleClearRow = (id: string) => {
    updateAppData((prev) => ({
      ...prev,
      Timesheet_InputList: (prev.Timesheet_InputList || []).map((r) =>
        r.id === id
          ? {
              ...r,
              url: "",
              fileName: undefined,
              sheetName: undefined,
              status: "pending",
              count: undefined,
              date: undefined,
              columnMapping: undefined,
            }
          : r,
      ),
      Q_Roster: (prev.Q_Roster || []).filter((r) => r._rowId !== id),
      Q_Salary_Scale: (prev.Q_Salary_Scale || []).filter((r) => r._rowId !== id),
      Q_Staff: (prev.Q_Staff || []).filter((r) => r._rowId !== id),
      Q_Cache: (prev.Q_Cache || []).filter((r) => r._rowId !== id),
    }));
  };
  const handleClearAll = () => {
    updateAppData((prev) => ({
      ...prev,
      Timesheet_InputList: (prev.Timesheet_InputList || []).map((r) => ({
        ...r,
        url: "",
        fileName: undefined,
        sheetName: undefined,
        status: "pending",
        count: undefined,
        date: undefined,
        columnMapping: undefined,
      })),
      Q_Roster: [],
      Q_Salary_Scale: [],
      Q_Staff: [],
      Q_Cache: [],
    }));
    toast?.success("Đã xóa toàn bộ dữ liệu (đã giữ lại thông tin center).");
  };

  const handleClearEmptyL07 = () => {
    updateAppData((prev) => ({
      ...prev,
      Timesheet_InputList: (prev.Timesheet_InputList || []).filter(
        (r) => r.l07 && r.l07.trim() !== "",
      ),
    }));
    toast?.success("Đã xóa các dòng chưa có mã L07.");
  };

  useEffect(() => {
    if (rosterData.length === 0) return;

    const centerSet = new Map<
      string,
      { l07: string; aeCode: string; bus: string }
    >();
    rosterData.forEach((t) => {
      const rawCenterCol = String(
        getVal(t, ["center", "location", "cơ sở"]) || "",
      ).trim();
      const rawAECol = String(getVal(t, ["mã ae", "ae"]) || "").trim();
      const info =
        getCenterInfoByAECode(rawAECol) ||
        getCenterInfoByL07(rawCenterCol) ||
        getCenterInfoByL07(mapL07(rawCenterCol));

      const l07 = info?.l07 || rawCenterCol || rawAECol || "UNKNOWN";
      const aeCode = info?.aeCode || rawAECol || "";
      const bus = info?.bus || "";
      const key = `${l07}|${aeCode}|${bus}`;

      if (!centerSet.has(key)) {
        centerSet.set(key, { l07, aeCode, bus });
      }
    });

    updateAppData((prev) => {
      const currentInputs = prev.Timesheet_InputList || [];
      const existingKeys = new Set(
        currentInputs.map((r) => `${r.l07}|${r.aeCode}|${r.bus}`),
      );
      
      let hasChanges = false;
      let newInputs = [...currentInputs];

      if (
        centerSet.size > 0 &&
        newInputs.length === 1 &&
        !newInputs[0].l07 &&
        !newInputs[0].url
      ) {
        newInputs = [];
        hasChanges = true;
      }

      centerSet.forEach((val, key) => {
        if (!existingKeys.has(key)) {
          newInputs.push({
            id: generateUUID(),
            l07: val.l07,
            aeCode: val.aeCode,
            bus: val.bus,
            url: "",
            status: "pending",
          });
          hasChanges = true;
        }
      });

      if (hasChanges) {
        return {
          ...prev,
          Timesheet_InputList: newInputs,
        };
      }
      return prev;
    }, false);
  }, [rosterData, updateAppData]);

  const handleRecalculate = () => {
    setRefreshKey((prev) => prev + 1);
    toast?.success("Đã tổng hợp lại dữ liệu.");
  };

  const handleSaveData = async () => {
    updateAppData(prev => ({
      ...prev,
      updatedAt: new Date().toISOString()
    }), true);
    
    if (isSupabaseConfigured()) {
      toast.info("Đang tự động đồng bộ dữ liệu hiện tại lên Supabase...");
      await handleSyncToSupabase();
    } else {
      toast.success("Đã lưu dữ liệu hiện tại offline thành công!");
    }
  };

  const handleSyncToSupabase = async () => {
    if (!isSupabaseConfigured()) {
      toast.error("Supabase chưa được cấu hình! Vui lòng cài đặt URL và Anon Key trong phần cấu hình.");
      return;
    }

    if (!rosterData || rosterData.length === 0) {
      toast.warning("Không có dữ liệu Roster để đồng bộ.");
      return;
    }

    setIsSyncing(true);
    setTotalSyncRows(rosterData.length);
    setSyncedRowsCount(0);
    setSyncProgress(0);

    try {
      const dataToSync = (computedData.processedRosterData && computedData.processedRosterData.length > 0) 
        ? computedData.processedRosterData 
        : rosterData;

      const { successCount, totalRows } = await syncRosterToSupabase(
        dataToSync as Record<string, unknown>[],
        (current, total) => {
          setSyncedRowsCount(current);
          setTotalSyncRows(total);
          setSyncProgress(Math.round((current / total) * 100));
        }
      );

      toast.success(`Đồng bộ thành công ${successCount.toLocaleString()}/${totalRows.toLocaleString()} dòng lên Supabase.`);
      
      updateAppData((prev: Record<string, unknown>) => ({
        ...prev,
        updatedAt: new Date().toISOString(),
        lastSupabaseSyncAt: new Date().toISOString()
      }), true);
      toast.success("Đã tự động lưu cứng dữ liệu trên web.");
    } catch (err: unknown) {
      console.error("Supabase Sync Error:", err);
      let errMsg = err instanceof Error ? err.message : String(err);
      
      if (errMsg.includes("Failed to fetch") || errMsg.includes("fetch")) {
        errMsg = "Không thể kết nối tới Supabase (Failed to fetch). Vui lòng kiểm tra lại URL Supabase trong phần Settings và đảm bảo Project của bạn đang hoạt động (không bị tạm dừng).";
      }

      // Detailed alert as requested for debugging RLS and column issues
      alert('Lỗi Supabase: ' + errMsg);
      toast.error(`Đồng bộ thất bại: ${errMsg}`);
      if (errMsg.includes("Bảng 'roster_cham_cong' chưa tồn tại") || errMsg.includes("Thiếu cột 'charge_to_center_mkt'")) {
        setShowSqlDialog(true);
      }
    } finally {
      setIsSyncing(false);
    }
  };

  const handleSyncRow = async (id: string) => {
    const row = (appData.Timesheet_InputList || []).find(r => r.id === id);
    if (!row || !row.url) {
      toast.error("Vui lòng nhập URL/ID Google Sheet trước.");
      return;
    }

    handleUpdateRow(id, "status", "processing");
    try {
      const file = await fetchGoogleSheetAsFile(row.url, row.sheetName || "Sheet1");
      if (file) {
         const allRows = await parseExcelInWorker(file);
         
         updateAppData((prev) => {
            const next = { ...prev };
            next.Q_Roster = (next.Q_Roster || []).filter(r => r._rowId !== id);
            const mapped = allRows.map(r => mapExcelRosterRow(r, file.name, id));
            next.Q_Roster = next.Q_Roster.concat(mapped);
            
            const newList = (prev.Timesheet_InputList || []).map(r => 
              r.id === id ? { ...r, status: "success", count: mapped.length, date: new Date().toLocaleString() } : r
            );
            next.Timesheet_InputList = newList;
            return next;
         }, true);
         
         toast.success(`Đã đồng bộ ${row.l07}: ${allRows.length} dòng.`);
      } else {
        throw new Error("Không lấy được nội dung file.");
      }
    } catch (err: unknown) {
      console.error(err);
      handleUpdateRow(id, "status", "error");
      const msg = err instanceof Error ? err.message : String(err);
      toast.error(`Lỗi: ${msg}`);
      if (msg.includes("BẠN CHƯA CẤP QUYỀN")) {
        alert(msg);
      }
    }
  };

  const handleSyncAll = async () => {
    const list = appData.Timesheet_InputList || [];
    const toSync = list.filter(r => r.url && r.status !== "success");
    if (toSync.length === 0) {
      toast.info("Không có dòng nào cần đồng bộ (hoặc chưa nhập URL).");
      return;
    }

    setIsCalculating(true);
    let successCount = 0;
    for (const row of toSync) {
      try {
        await handleSyncRow(row.id);
        successCount++;
        await new Promise(r => setTimeout(r, 1000)); // Rate limit safety
      } catch (e) {
        console.error(e);
      }
    }
    setIsCalculating(false);
    toast.success(`Đã hoàn thành đồng bộ ${successCount}/${toSync.length} trung tâm.`);
  };

  const handleUploadFiles = async (files: File[]) => {
    const currentInputs = appData.Timesheet_InputList || [];
    const updatedInputs = [...currentInputs];
    const toProcess: { id: string; file: File }[] = [];
    let hasChanges = false;

    const filteredFiles = files.filter(f => !f.name.toLowerCase().includes("copy"));
    if (filteredFiles.length === 0 && files.length > 0) {
      toast.info("Tất cả các file đã chọn đều là file copy nên hệ thống tự động bỏ qua.");
      return;
    }

    for (const file of filteredFiles) {
      const l07 = getL07FromFileName(file.name) || "";
      const centerInfo = l07 ? getCenterInfoByL07(l07) : null;
      const aeCode = centerInfo?.aeCode || "";

      const matchIndex = updatedInputs.findIndex((r) => {
        const matchL07 =
          l07 && r.l07 && r.l07.toLowerCase() === l07.toLowerCase();
        const matchAE =
          aeCode && r.aeCode && r.aeCode.toLowerCase() === aeCode.toLowerCase();
        return matchL07 || matchAE;
      });

      if (matchIndex !== -1) {
        updatedInputs[matchIndex] = {
          ...updatedInputs[matchIndex],
          status: "processing",
        };
        toProcess.push({ id: updatedInputs[matchIndex].id, file });
        hasChanges = true;
      } else {
        const newId = crypto.randomUUID();
        updatedInputs.push({
          id: newId,
          l07: l07,
          aeCode: aeCode,
          bus: centerInfo?.bus || "",
          status: "processing",
          url: ""
        });
        toProcess.push({ id: newId, file });
        hasChanges = true;
      }
    }

    if (hasChanges) {
      updateAppData(
        (prev) => ({
          ...prev,
          Timesheet_InputList: updatedInputs,
        }),
        false
      );
    }

    if (toProcess.length > 0) {
      for (const p of toProcess) {
        try {
          await handleUploadFile(p.id, p.file);
        } catch (err: unknown) {
          console.error(`Error parsing ${p.file.name}:`, err);
          handleUpdateRow(p.id, "status", "error");
        }
      }
    }
  };

  const handleUploadFile = async (rowId: string, file: File) => {
    if (file.name.toLowerCase().includes("copy")) {
      toast?.info(`Hệ thống tự động bỏ qua file có tên 'copy': ${file.name}`);
      return;
    }

    handleUpdateRow(rowId, "status", "processing");
    try {
      let isSalary = false,
        isStaff = false,
        isCache = false;
      const fn = file.name.toLowerCase();
      if (fn.includes("salary")) isSalary = true;
      else if (fn.includes("staff")) isStaff = true;
      else if (fn.includes("cache")) isCache = true;

      const allRows = await parseExcelInWorker(file);

      allRows.forEach((r: Record<string, unknown>) => {
        r._sourceFile = file.name;
        r._rowId = rowId;
      });

      if (allRows.length > 0) {
        const headers = Object.keys(allRows[0] as Record<string, unknown>).map((k) =>
          k.toLowerCase().trim(),
        );

        updateAppData((prev) => {
          const next = { ...prev };
          
          next.Q_Roster = (next.Q_Roster || []).filter((r: Record<string, unknown>) => r._rowId !== rowId);
          next.Q_Salary_Scale = (next.Q_Salary_Scale || []).filter((r: Record<string, unknown>) => r._rowId !== rowId);
          next.Q_Staff = (next.Q_Staff || []).filter((r: Record<string, unknown>) => r._rowId !== rowId);
          next.Q_Cache = (next.Q_Cache || []).filter((r: Record<string, unknown>) => r._rowId !== rowId);

          if (
            headers.includes("academic price") ||
            isSalary ||
            headers.includes("s code")
          )
            next.Q_Salary_Scale = next.Q_Salary_Scale.concat(allRows);
          else if (headers.includes("bank account number") || isStaff)
            next.Q_Staff = next.Q_Staff.concat(allRows);
          else if (headers.includes("today") || isCache)
            next.Q_Cache = next.Q_Cache.concat(allRows);
          else {
            const mappedRosters = allRows.map((r: Record<string, unknown>) => mapExcelRosterRow(r, file.name, rowId));
            next.Q_Roster = next.Q_Roster.concat(mappedRosters);
          }

          const d = new Date();
          const detectedL07 = getL07FromFileName(file.name);
          const centerInfo = detectedL07 ? getCenterInfoByL07(detectedL07) : null;
          const bu = detectedL07 ? getBusinessFromL07(detectedL07) : "";

          next.Timesheet_InputList = (next.Timesheet_InputList || []).map((input) =>
            input.id === rowId
              ? {
                  ...input,
                  l07: input.l07 || detectedL07 || "",
                  aeCode: input.aeCode || centerInfo?.aeCode || "",
                  bus: input.bus || bu || "",
                  status: "success",
                  fileName: file.name,
                  date: `${d.getHours()}:${String(d.getMinutes()).padStart(2, "0")} ${d.getDate()}/${d.getMonth() + 1}`,
                }
              : input
          );

          return next;
        }, false);

        toast?.success(`Đọc thành công ${file.name}`);
      }
    } catch (err: unknown) {
      const errMsg = err instanceof Error ? err.message : "Unknown error";
      const errName = file.name;
      console.error(`[TimesheetSummary] Error reading ${errName}:`, err);
      handleUpdateRow(rowId, "status", "error");
      toast?.error(
        `Lỗi đọc ${errName}: ${errMsg}`,
      );
    }
  };

  const computedData = useTimesheetCalculations(
    rosterData,
    salaryScaleData,
    staffData,
    cacheData,
    debouncedFromDate,
    debouncedToDate
  );

  useEffect(() => {
    const signature = JSON.stringify({
      emp: computedData.employeeSummary?.length || 0,
      center: computedData.centerSummary?.length || 0,
    });

    if (lastSummaryRef.current === signature) return;
    lastSummaryRef.current = signature;

    updateAppData(
      (prev: Record<string, unknown>) => ({
        ...prev,
        TA_Employee_Summary: {
          headers: getDynamicEmployeeColumns(rosterData).map((c) => c.label),
          data: computedData.employeeSummary,
        },
        TA_Center_Summary: {
          headers: CENTER_COLUMNS.map((c) => c.label),
          data: computedData.centerSummary,
        },
      }),
      false,
    );
  }, [computedData.employeeSummary, computedData.centerSummary, rosterData, updateAppData]);

  const activeData = inputRows;

  const handleUploadFileA = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const allRows = await parseExcelInWorker(file);

      console.log("Parsed File A:", allRows.slice(0, 5));
      updateAppData((prev) => ({ ...prev, Q_TeacherHours: allRows }));
      toast?.success(`Tải lên File A thành công (${allRows.length} dòng)`);
      if (e.target) e.target.value = "";
    } catch (err: unknown) {
      console.error(err);
      const msg = err instanceof Error ? err.message : "Lỗi khi đọc File A";
      toast?.error(msg);
      if (e.target) e.target.value = "";
    }
  };

  const handleExport = () => {
    if (activeData.length === 0) {
      toast?.error("Không có dữ liệu");
      return;
    }
    const ws = XLSX.utils.json_to_sheet(prepareDataForExport(activeData));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, activeTab);
    XLSX.writeFile(wb, `Timesheet_Export_${activeTab}.xlsx`);
  };

  return (
    <div className="page-timesheet-summary flex-1 flex flex-col min-h-0 bg-transparent p-2 sm:p-3 md:p-4 gap-4 w-full h-full overflow-hidden">
      <button data-action="save-data" className="hidden" onClick={handleSaveData} />
      
      <input
        type="file"
        id="fileA"
        accept=".xlsx,.xls,.csv"
        className="hidden"
        onChange={handleUploadFileA}
      />

      <div className="bg-white soft-card force-light flex-1 flex flex-col min-h-0 w-full relative overflow-hidden rounded-2xl border border-accent/10 shadow-sm">
        <div className="absolute inset-0 bg-accent/5 opacity-[0.05] pointer-events-none hidden" />

        <div className="p-4 sm:p-5 flex flex-col md:flex-row items-center justify-between gap-4 bg-accent/5 shrink-0 border-none relative z-10 overflow-hidden rounded-t-2xl rounded-b-none">
          <div className="absolute inset-0 bg-accent/5 opacity-[0.03] pointer-events-none rounded-t-2xl rounded-b-none" />
          {computedData?.error && (
            <div className="absolute top-0 left-0 right-0 bg-red-100 text-red-600 p-2 text-center text-xs font-bold z-50">
              WORKER ERROR: {computedData.error}
            </div>
          )}
          {isSyncing && (
            <div className="absolute top-0 left-0 right-0 bg-sky-50 border-b border-sky-200 px-8 py-3 flex flex-col sm:flex-row items-center justify-between gap-3 z-50 animate-in fade-in slide-in-from-top duration-300">
              <div className="flex items-center gap-3">
                <RefreshCw className="w-5 h-5 text-sky-600 animate-spin" />
                <div>
                  <p className="text-xs font-black text-sky-950 uppercase tracking-wider">
                    Đang đồng bộ dữ liệu lên Supabase...
                  </p>
                  <p className="text-[10px] font-bold text-sky-700 uppercase mt-0.5">
                    Đã lưu thành công: {syncedRowsCount.toLocaleString()} / {totalSyncRows.toLocaleString()} dòng ({syncProgress}%)
                  </p>
                </div>
              </div>
              <div className="w-full sm:w-64 bg-sky-200/50 rounded-full h-2.5 overflow-hidden relative">
                <div 
                  className="bg-sky-600 h-full rounded-full transition-all duration-300"
                  style={{ width: `${syncProgress}%` }}
                />
              </div>
            </div>
          )}
          <div className="flex items-center gap-5 relative z-10">
            <PuppyLogo size={56} className="shrink-0" />

            <div>
              <div className="flex flex-col sm:flex-row sm:items-center gap-3">
                <h2 className="text-3xl font-normal font-serif text-[#3D3935] tracking-tight flex items-end gap-1" style={{ lineHeight: "33px" }}>
                  Data{" "}
                  <span className="not-italic font-script text-accent text-4xl lowercase inline-block transform -translate-y-0.5" style={{ lineHeight: "33px" }}>
                    Summary
                  </span>
                  <span 
                    className="text-3xl tracking-tight" 
                    style={{ 
                      lineHeight: "33px",
                      fontFamily: "Corinthia, cursive",
                      fontWeight: "bold",
                      color: "#5d4021",
                      marginLeft: "4px",
                      paddingTop: "0px",
                      marginTop: "0px"
                    }}
                  >
                    & Source
                  </span>
                </h2>
              </div>
              <div className="flex items-center gap-3 mt-1.5 text-xs text-slate-500 font-semibold tracking-wider uppercase">
                <span className="flex items-center gap-1">
                  <span className="font-extrabold text-slate-800 text-[13px]">{inputRows.length || 0}</span>{" "}
                  <span className="text-[10px] text-slate-500 lowercase">centers</span>
                </span>
                <span className="text-accent/30 font-normal">•</span>
                <span className="flex items-center gap-1">
                  <span className="font-extrabold text-slate-800 text-[13px]">{computedData?.employeeSummary?.length || 0}</span>{" "}
                  <span className="text-[10px] text-slate-500 lowercase">employees</span>
                </span>
                <span className="text-accent/30 font-normal">•</span>
                <span className="flex items-center gap-1">
                  <span className="font-extrabold text-slate-800 text-[13px]">{(computedData?.processedRosterData?.length || 0).toLocaleString()}</span>{" "}
                  <span className="text-[10px] text-slate-500 lowercase">records</span>
                </span>
              </div>
            </div>
          </div>

          <div className="flex items-center gap-3">
            {onBack && (
              <button
                onClick={onBack}
                className="group relative h-11 px-6 bg-white overflow-hidden rounded-full shadow-[0_2px_10px_rgba(0,0,0,0.06)] hover:shadow-[0_4px_15px_rgba(0,0,0,0.1)] transition-all duration-300 flex items-center justify-center gap-2 border border-border mr-2"
              >
                <div className="absolute inset-0 bg-primary/5 opacity-0 group-hover:opacity-100 transition-opacity duration-300" />
                <ArrowLeft className="w-4 h-4 text-primary group-hover:-translate-x-1 transition-transform" />
                <span className="text-xs font-bold uppercase tracking-wider text-primary relative z-10">
                  Về bảng Roster gốc
                </span>
              </button>
            )}
            <div className="flex items-center gap-3">
              <DropdownMenu>
                <Tooltip>
                  <TooltipTrigger asChild>
                    <DropdownMenuTrigger asChild>
                      <button className="flex w-11 h-11 items-center justify-center rounded-full border border-border bg-[#3D3935] text-white hover:bg-[#3D3935]/90 transition-all group shadow-sm z-10 relative">
                        <Settings className="w-5 h-5 group-hover:rotate-45 transition-transform duration-500" />
                      </button>
                    </DropdownMenuTrigger>
                  </TooltipTrigger>
                  <TooltipContent>Cài đặt & Tiện ích</TooltipContent>
                </Tooltip>
                <DropdownMenuContent
                  align="end"
                  className="w-64 border border-border/50 shadow-2xl p-2 bg-white rounded-2xl"
                >
                  <DropdownMenuLabel className="text-[0.625rem] font-bold uppercase tracking-widest text-primary/60 px-3 py-2">
                    Tiện ích
                  </DropdownMenuLabel>
                  <DropdownMenuSeparator className="bg-border/50 mx-1" />
                  
                  <DropdownMenuItem
                    onSelect={handleSyncAll}
                    disabled={isCalculating || isFetchingGgSheet}
                    className="flex items-center gap-3 px-3 py-2.5 rounded-xl cursor-pointer hover:bg-slate-50 transition-colors font-bold text-[0.6875rem] uppercase tracking-wider"
                  >
                    <RefreshCw className={`w-4 h-4 text-[#E5A8A0] ${(isCalculating || isFetchingGgSheet) ? 'animate-spin' : ''}`} />
                    <span>Sync All Data</span>
                  </DropdownMenuItem>

                  <DropdownMenuItem
                    onSelect={handleAddRow}
                    className="flex items-center gap-3 px-3 py-2.5 rounded-xl cursor-pointer hover:bg-accent/10 transition-colors font-bold text-[0.6875rem] uppercase tracking-wider"
                  >
                    <Plus className="w-4 h-4 text-accent" />
                    <span>Thêm dòng trung tâm</span>
                  </DropdownMenuItem>

                  <DropdownMenuSeparator className="bg-border/50 mx-1" />

                  <DropdownMenuItem
                    onSelect={handleClearAll}
                    className="flex items-center gap-3 px-3 py-2.5 rounded-xl cursor-pointer hover:bg-accent/10 transition-colors text-accent"
                  >
                    <Trash2 className="w-4 h-4 text-accent" />
                    <span className="text-[0.6875rem] font-bold uppercase tracking-wider text-accent">
                      Xóa toàn bộ
                    </span>
                  </DropdownMenuItem>
                  <DropdownMenuSeparator className="bg-border/50 mx-1" />
                  <DropdownMenuItem
                    onSelect={() => handleUrlInput(inputRows[0].id, DEFAULT_FOLDER_URL)}
                    disabled={isFetchingGgSheet}
                    className="flex items-center gap-3 px-3 py-2.5 rounded-xl cursor-pointer hover:bg-amber-50 transition-colors text-amber-600 disabled:opacity-50"
                  >
                    <FileSpreadsheet className={`w-4 h-4 text-amber-600 ${isFetchingGgSheet ? "animate-spin" : ""}`} />
                    <span className="text-[0.6875rem] font-bold uppercase tracking-wider text-amber-600">
                      Đồng bộ google sheet
                    </span>
                  </DropdownMenuItem>
                  <DropdownMenuSeparator className="bg-border/50 mx-1" />
                  <DropdownMenuItem
                    onSelect={handleSyncToSupabase}
                    disabled={isSyncing}
                    className="flex items-center gap-3 px-3 py-2.5 rounded-xl cursor-pointer hover:bg-sky-50 transition-colors text-sky-600 disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    <RefreshCw className={`w-4 h-4 text-sky-600 ${isSyncing ? "animate-spin" : ""}`} />
                    <span className="text-[0.6875rem] font-bold uppercase tracking-wider text-sky-600">
                      ĐỒNG BỘ LÊN SUPABASE
                    </span>
                  </DropdownMenuItem>
                  <DropdownMenuSeparator className="bg-border/50 mx-1" />

                  <DropdownMenuItem
                    onSelect={handleSaveData}
                    className="flex items-center gap-3 px-3 py-2.5 rounded-xl cursor-pointer hover:bg-slate-50 transition-colors"
                  >
                    <Save className="w-4 h-4 text-accent" />
                    <span className="text-[0.6875rem] font-bold uppercase tracking-wider text-accent">
                      Lưu dữ liệu
                    </span>
                  </DropdownMenuItem>
                  <DropdownMenuItem
                    onSelect={handleExport}
                    className="flex items-center gap-3 px-3 py-2.5 rounded-xl cursor-pointer hover:bg-slate-50 transition-colors"
                  >
                    <Download className="w-4 h-4 text-primary" />
                    <span className="text-[0.6875rem] font-bold uppercase tracking-wider">
                      Xuất Excel
                    </span>
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            </div>
          </div>
        </div>
        
        {/* Service Account Info Card removed as requested */}


        <div className="flex-1 flex flex-col min-h-0 relative rounded-b-2xl overflow-hidden">
          <TimesheetInputTable
            rows={inputRows}
            onAddRow={handleAddRow}
            onUpdateRow={handleUpdateRow}
            onClearRow={handleClearRow}
            onClearAll={handleClearAll}
            onClearEmptyL07={handleClearEmptyL07}
            onUploadFile={handleUploadFile}
            onUploadFiles={handleUploadFiles}
            onUrlInput={handleUrlInput}
            onRefresh={handleRecalculate}
            onSyncRow={handleSyncRow}
            isProcessing={isCalculating || isFetchingGgSheet}
          />
        </div>
      </div>

      <Dialog open={showSqlDialog} onOpenChange={setShowSqlDialog}>
        <DialogContent className="max-w-2xl bg-white rounded-3xl border-none shadow-2xl p-0 overflow-hidden">
          <div className="bg-sky-600 p-8 text-white">
            <DialogHeader>
              <DialogTitle className="text-2xl font-black uppercase tracking-wider">Thiết lập Bảng Supabase</DialogTitle>
              <DialogDescription className="text-sky-100 font-medium">
                Bảng 'roster_cham_cong' chưa tồn tại hoặc thiếu cột dữ liệu. Vui lòng copy script bên dưới và chạy trong SQL Editor của Supabase để cập nhật cấu trúc bảng.
              </DialogDescription>
            </DialogHeader>
          </div>
          <div className="p-8">
            <div className="relative group">
              <pre className="bg-slate-950 text-sky-400 p-6 rounded-2xl text-[10px] font-mono leading-relaxed overflow-x-auto max-h-[300px] border border-slate-800 shadow-inner custom-scrollbar">
                {SQL_SETUP_SCRIPT}
              </pre>
              <Button
                variant="outline"
                size="sm"
                className="absolute top-4 right-4 bg-white/10 hover:bg-white/20 border-white/20 text-white gap-2 backdrop-blur-md opacity-0 group-hover:opacity-100 transition-all"
                onClick={() => {
                  navigator.clipboard.writeText(SQL_SETUP_SCRIPT);
                  toast.success("Đã copy script SQL!");
                }}
              >
                <Copy className="w-3.5 h-3.5" />
                SAO CHÉP
              </Button>
            </div>
            <div className="mt-6 space-y-4">
              <h4 className="text-xs font-black uppercase tracking-widest text-slate-400">Các bước thực hiện:</h4>
              <ol className="text-[11px] font-bold text-slate-600 space-y-2 list-decimal pl-4">
                <li>Truy cập vào Dashboard Supabase của bạn.</li>
                <li>Chọn dự án và vào phần <span className="text-sky-600">SQL Editor</span>.</li>
                <li>Bấm <span className="text-sky-600">New Query</span> và dán nội dung script trên vào.</li>
                <li>Bấm <span className="text-sky-600">Run</span> để tạo bảng và cấu hình quyền truy cập (RLS).</li>
                <li>Quay lại đây và thử Đồng bộ lại.</li>
              </ol>
            </div>
          </div>
          <DialogFooter className="p-6 bg-slate-50 border-t border-slate-100">
            <Button 
              onClick={() => setShowSqlDialog(false)}
              className="bg-sky-600 hover:bg-sky-700 text-white rounded-xl px-8 font-black uppercase tracking-widest text-[10px]"
            >
              Tôi đã hiểu
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
