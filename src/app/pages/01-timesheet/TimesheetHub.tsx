import { PuppyLogo } from "../../components/shared/PuppyLogo";
/* eslint-disable @typescript-eslint/no-explicit-any, react-hooks/set-state-in-effect, @typescript-eslint/no-unused-vars */
import React, { useMemo, useRef, useState, useEffect, useTransition, useCallback } from "react";
import { useLocation } from "react-router";
import { useAppData } from "../../lib/contexts/AppDataContext";
import { useTimesheetCalculations } from "../../hooks/useTimesheetCalculations";
import { prepareDataForExport } from "../../lib/utils/data-utils";
import { INITIAL_APP_DATA } from "../../constants/initial-data";
import {
  FileText,
  Users,
  Building2,
  Search,
  ChevronDown,
  XCircle,
  RefreshCw,
  SlidersHorizontal,
  Save,
  Plus,
  Check,
  Settings,
  Download,
} from "lucide-react";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription, DialogFooter } from "../../components/ui/dialog";
import { Button } from "../../components/ui/button";
import { Copy } from "lucide-react";
import { RosterRawTable } from "./tables/RosterRawTable";
import { EmployeeTable } from "./tables/EmployeeTable";
import { CenterTable } from "./tables/CenterTable";
import { MktLocalNorthPivotTable } from "./tables/MktLocalNorthPivotTable";
import TimesheetSummaryPage from "./TimesheetSummary";
import { useNavigate } from "react-router";
import { supabase, isSupabaseConfigured } from "@/lib/supabaseClient";
import { 
  syncRosterToSupabase, 
  syncEmployeesToSupabase, 
  syncSalaryScalesToSupabase, 
  clearSupabaseData, 
  SQL_SETUP_SCRIPT 
} from "../../lib/supabase-sync-utils";
import { toast } from "sonner";
import {
  Popover,
  PopoverContent,
  PopoverTrigger,
} from "../../components/ui/popover";
import { Calendar } from "../../components/ui/calendar";
import { format } from "date-fns";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "../../components/ui/dropdown-menu";
import { motion, AnimatePresence } from "motion/react";
import * as XLSX from "xlsx";

const containerVariants = {
  hidden: { opacity: 0 },
  visible: {
    opacity: 1,
    transition: {
      staggerChildren: 0.1,
    },
  },
} as const;

const timesheetSearchCache = new WeakMap<any, string>();

let hasFetchedSupabase = false;

export function TimesheetHub() {
  const { appData, updateAppData } = useAppData();
  const location = useLocation();
  const navigate = useNavigate();
  const [isPending, startTransition] = useTransition();

  const [activeTab, setActiveTab] = useState<
    "roster_raw" | "employee" | "center" | "mkt_local_north"
  >("roster_raw");
  const [searchTerm, setSearchTerm] = useState("");
  const [debouncedSearchTerm, setDebouncedSearchTerm] = useState("");

  const [view, setView] = useState<"final" | "upload">("final");
  const [fromDate, setFromDate] = useState(appData.Timesheet_Dates?.from || "");
  const [toDate, setToDate] = useState(appData.Timesheet_Dates?.to || "");
  const [debouncedFromDate, setDebouncedFromDate] = useState(appData.Timesheet_Dates?.from || "");
  const [debouncedToDate, setDebouncedToDate] = useState(appData.Timesheet_Dates?.to || "");
  const [showSidebar, setShowSidebar] = useState(true);
  const [showControlBar, setShowControlBar] = useState(true);

  const [isSyncing, setIsSyncing] = useState(false);
  const [syncProgress, setSyncProgress] = useState(0);
  const [totalSyncRows, setTotalSyncRows] = useState(0);
  const [syncedRowsCount, setSyncedRowsCount] = useState(0);
  const [showSqlDialog, setShowSqlDialog] = useState(false);
  const [_tableFilteredCount, setTableFilteredCount] = useState<number | null>(null);

  useEffect(() => {
    const fetchRealtimeData = async () => {
      if (!isSupabaseConfigured()) {
        console.log("Supabase is not configured yet. Using local state.");
        return;
      }
      if (hasFetchedSupabase) {
        console.log("Supabase data already loaded in this session. Skipping auto-fetch on tab switch.");
        return;
      }
      try {
        // Fetch roster_cham_cong
        const { data: dbRoster, error: rosterErr } = await supabase
          .from("roster_cham_cong")
          .select("*");
          
        // Fetch nhan_vien
        const { data: dbStaff, error: staffErr } = await supabase
          .from("nhan_vien")
          .select("*");

        // Fetch thang_luong
        const { data: dbSalary, error: salaryErr } = await supabase
          .from("thang_luong")
          .select("*");

        if (rosterErr || staffErr || salaryErr) {
          console.warn("Supabase tables might not exist yet. Please run the SQL setup script.", { rosterErr, staffErr, salaryErr });
          return;
        }

        if ((dbRoster || []).length === 0 && (dbStaff || []).length === 0 && (dbSalary || []).length === 0) {
          console.log("Supabase tables are empty. Keeping initial local data so user can sync.");
          hasFetchedSupabase = true;
          return;
        }

        // Map Roster rows
        const mappedRoster = (dbRoster || []).map((row: any) => ({
          ...(row.raw_data || {}),
          _rowId: row.unique_id || `supa-r-${row.id}`,
          _sourceFile: row.raw_data?._sourceFile || "Supabase_Live",
          center: row.l07 || "",
          l07: row.l07 || "",
          business: row.business || "",
          ma_nv: row.ma_nv || "",
          full_name: row.full_name || "",
          ngay: row.ngay || "",
          type: row.type || "",
          class: row.class || "",
          gio_vao: row.gio_vao || "",
          gio_ra: row.gio_ra || "",
          duration: Number(row.duration) || 0,
          notes: row.notes || "",
          employeeId: row.ma_nv || "",
          fullName: row.full_name || "",
          maAE: row.l07 || "",
          date: row.ngay || "",
          taskType: row.type || "",
          classCode: row.class || "",
          from: row.gio_vao || "",
          to: row.gio_ra || "",
          chargeToCenterMkt: row.charge_to_center_mkt || ""
        }));

        // Map Staff rows
        const mappedStaff = (dbStaff || []).map((row: any) => ({
          ...(row.raw_data || {}),
          _rowId: row.unique_id,
          employeeId: row.employee_id,
          fullName: row.full_name,
          bankAccountNumber: row.bank_account_number,
          salaryScale: row.salary_scale,
          business: row.business,
          center: row.center,
          from: row.from,
          to: row.to,
          className: row.class_name,
          noteDays: row.note_days
        }));

        // Map Salary scale rows
        const mappedSalary = (dbSalary || []).map((row: any) => ({
          ...(row.raw_data || {}),
          _rowId: row.unique_id,
          sCode: row.s_code,
          academicPrice: Number(row.academic_price) || 0,
          baseSalary: Number(row.base_salary) || 0,
          totalSalary: Number(row.total_salary) || 0,
          deductionHours: Number(row.deduction_hours) || 0
        }));

        hasFetchedSupabase = true;

        updateAppData((prev) => ({
          ...prev,
          Q_Roster: mappedRoster,
          Q_Staff: mappedStaff,
          Q_Salary_Scale: mappedSalary
        }), false);

        console.log("Successfully loaded real-time data from Supabase:", {
          roster: mappedRoster.length,
          staff: mappedStaff.length,
          salary: mappedSalary.length
        });
      } catch (err) {
        console.error("Error fetching realtime Supabase data:", err);
      }
    };

    fetchRealtimeData();
  }, [updateAppData]);



  useEffect(() => {
    const timer = setTimeout(() => {
      setDebouncedSearchTerm(searchTerm);
    }, 300);
    return () => clearTimeout(timer);
  }, [searchTerm]);

  const [targetDate, setTargetDate] = useState("");
  const [targetCenter, setTargetCenter] = useState("");

  const handleClearFilters = useCallback(() => {
    setSearchTerm("");
    setTargetDate("");
    setTargetCenter("");
    navigate(location.pathname, {
      replace: true,
      state: { from: "cleared" },
    });
    if (tableRef.current) {
      tableRef.current.clearAllFilters();
    }
  }, [navigate, location.pathname]);

  const containerRef = useRef<HTMLDivElement>(null);

  // Handle deep linking and navigation resets
  useEffect(() => {
    const state = location.state as any;
    if (state && state.from === "audit") {
      // Apply filters
      if (state.activeTab) setActiveTab(state.activeTab);
      if (state.searchTerm) {
        setSearchTerm(state.searchTerm);
        setDebouncedSearchTerm(state.searchTerm);
      }
      if (state.filterDate) setTargetDate(state.filterDate);
      if (state.filterCenter) setTargetCenter(state.filterCenter);

      // Scroll to the table after a brief delay to ensure rendering
      setTimeout(() => {
        if (containerRef.current) {
          containerRef.current.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
      }, 100);

      // Clear location state but DO NOT trigger cleanup
      navigate(location.pathname, {
        replace: true,
        state: { ...state, from: "audit_applied" },
      });
    }
  }, [location.state, navigate, location.pathname]);

  // Separate effect for clearing filters when navigating NOT from audit
  useEffect(() => {
    const state = location.state as any;
    // Only clear if the user manually changed the URL, not because we cleared the state internally
    if (
      !state ||
      (state.from !== "audit" &&
        state.from !== "audit_applied" &&
        state.from !== "cleared" &&
        !state.activeTab)
    ) {
      handleClearFilters();
      setActiveTab("roster_raw");
      setView("final");
    }
  }, [location.state, handleClearFilters]);



  // Remove effect syncing globalMonth down to local selectedMonth
  
  useEffect(() => {
    const timer = setTimeout(() => {
      setDebouncedFromDate(fromDate);
      setDebouncedToDate(toDate);
    }, 500);
    setTableFilteredCount(null);
    return () => clearTimeout(timer);
  }, [fromDate, toDate]);

  useEffect(() => {
    setTableFilteredCount(null);
  }, [activeTab, searchTerm, targetDate, targetCenter]);

  useEffect(() => {
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

  const calculatedRosterData = useMemo(() => appData.Q_Roster || [], [appData.Q_Roster]);
  const calculatedSalaryScaleData = useMemo(() => appData.Q_Salary_Scale || [], [appData.Q_Salary_Scale]);
  const calculatedStaffData = useMemo(() => appData.Q_Staff || [], [appData.Q_Staff]);
  const calculatedCacheData = useMemo(() => appData.Q_Cache || [], [appData.Q_Cache]);

  const { processedRosterData, employeeSummary, centerSummary, isCalculating } =
    useTimesheetCalculations(
      calculatedRosterData,
      calculatedSalaryScaleData,
      calculatedStaffData,
      calculatedCacheData,
      appData.Timesheet_Dates?.from || debouncedFromDate,
      appData.Timesheet_Dates?.to || debouncedToDate,
    );

  const tabs = useMemo(
    () =>
      [
        { id: "roster_raw", label: "Roster Gốc", icon: FileText },
        { id: "employee", label: "Số Giờ Làm Việc", icon: Users },
        { id: "center", label: "Roster Center", icon: Building2 },
        { id: "mkt_local_north", label: "Phí MKT Local North", icon: FileText },
      ] as const,
    [],
  );

  const mktLocalNorthData = useMemo(() => {
    return processedRosterData.filter((r: any) => {
      const cUpper = String(r.center || "").toUpperCase();
      const isMkt = cUpper === "MKT LOCAL NORTH" || cUpper.startsWith("MKT LOCAL NORTH_");
      // Phải loại bỏ các ca trùng lịch (overlap) khỏi bảng Pivot
      return isMkt && r.overlap_check !== "Trùng lịch";
    });
  }, [processedRosterData]);

  const currentData = useMemo(() => {
    if (activeTab === "roster_raw") return processedRosterData;
    if (activeTab === "employee") return employeeSummary;
    if (activeTab === "center") return centerSummary;
    if (activeTab === "mkt_local_north") return mktLocalNorthData;
    return [];
  }, [activeTab, processedRosterData, employeeSummary, centerSummary, mktLocalNorthData]);

  const searchData = useMemo(() => {
    let data = currentData;

    // 1. If we have a target date (from audit or manually set), filter by date first
    if (targetDate) {
      data = data.filter((row: any) => {
        const rowDate = String(row.date || "").trim();
        const tDate = String(targetDate).trim();
        return rowDate === tDate || rowDate.includes(tDate);
      });
    }

    // 2. If we have a target center (from audit), filter by center
    if (targetCenter) {
      data = data.filter((row: any) => {
        const rowCenter = String(row.center || "")
          .trim()
          .toUpperCase();
        const tCenter = String(targetCenter).trim().toUpperCase();
        return rowCenter === tCenter || rowCenter.includes(tCenter);
      });
    }

    // 3. If we have a search term (class name)
    if (debouncedSearchTerm) {
      const normalizeStr = (s: string) => {
        if (!s) return "";
        let normalized = s.toLowerCase();
        normalized = normalized.replace(/đ/g, "d");
        return normalized
          .normalize("NFD")
          .replace(/[\u0300-\u036f]/g, "")
          .replace(/\s+/g, "");
      };

      const lowerSearch = normalizeStr(debouncedSearchTerm);
      const lowerSearchTrimmedZero = lowerSearch.replace(/^0+/, "");

      const searchCache = timesheetSearchCache;

      data = data.filter((row: any) => {
        // Use precomputed _searchStr if available
        let rowSearchStr = searchCache.get(row);
        
        if (rowSearchStr !== undefined) {
          if (rowSearchStr.includes(lowerSearch)) return true;
          if (lowerSearchTrimmedZero && rowSearchStr.includes(lowerSearchTrimmedZero)) return true;
          return false;
        }

        rowSearchStr = "";
        
        // Optimize search to only search in keys that might be displayed
        for (const [key, value] of Object.entries(row)) {
          if (value == null) continue;
          
          if (key === "employeeId" || key === "ma_nv" || key.toLowerCase().includes("business") || key.toLowerCase().includes("center") || key.toLowerCase().includes("class") || key.toLowerCase().includes("name")) {
              rowSearchStr += `|${normalizeStr(String(value))}`;
          }
        }
        
        // Cache it for future filtering
        searchCache.set(row, rowSearchStr);

        return rowSearchStr.includes(lowerSearch) || rowSearchStr.includes(lowerSearchTrimmedZero);
      });
    }

    return data;
  }, [currentData, debouncedSearchTerm, targetDate, targetCenter]);

  // 1. Get unique non-empty type values for Pivot Table columns (excluding empty key values as requested)
  const mktPivotUniqueTypes = useMemo(() => {
    if (activeTab !== "mkt_local_north") return [];
    const typesSet = new Set<string>();
    searchData.forEach((r: any) => {
      const type = String(r.taskType || "").trim().toUpperCase();
      if (type) {
        typesSet.add(type);
      }
    });
    return Array.from(typesSet).sort();
  }, [activeTab, searchData]);

  // 2. Aggregate row data by business -> center -> chargeToCenterMkt
  const mktPivotRows = useMemo(() => {
    if (activeTab !== "mkt_local_north") return [];
    
    const map = new Map<string, {
      business: string;
      center: string;
      chargeToCenterMkt: string;
      values: Record<string, number>;
      total: number;
    }>();

    searchData.forEach((r: any) => {
      const type = String(r.taskType || "").trim().toUpperCase();
      if (!type) return; // skip empty data as requested

      const bus = String(r.business || "").trim();
      const charge = String(r.chargeToCenterMkt || "").trim();
      const key = `${bus}||${charge}`;

      if (!map.has(key)) {
        map.set(key, {
          business: bus,
          center: "", // No longer grouping by center as requested
          chargeToCenterMkt: charge,
          values: {},
          total: 0,
        });
      }

      const item = map.get(key)!;
      const hours = Number(r.duration ?? r.workingHours) || 0;
      // Value: working hours * 20,000 as requested
      const value = hours * 20000;

      item.values[type] = (item.values[type] || 0) + value;
      item.total += value;
    });

    return Array.from(map.values()).sort((a, b) => {
      const comp1 = a.business.localeCompare(b.business);
      if (comp1 !== 0) return comp1;
      return a.chargeToCenterMkt.localeCompare(b.chargeToCenterMkt);
    });
  }, [activeTab, searchData]);

  // 3. Compute column and grand totals for the Pivot Grid
  const mktPivotGrandTotals = useMemo(() => {
    const totals: Record<string, number> = {};
    let grandTotal = 0;
    
    mktPivotRows.forEach((row) => {
      mktPivotUniqueTypes.forEach((type) => {
        totals[type] = (totals[type] || 0) + (row.values[type] || 0);
      });
      grandTotal += row.total;
    });

    return { totals, grandTotal };
  }, [mktPivotRows, mktPivotUniqueTypes]);

  const handleExportExcel = () => {
    if (currentData.length === 0) return;

    if (activeTab === "mkt_local_north") {
      const rows = mktPivotRows.map((row) => {
        const item: any = {
          "Business": row.business,
          "Charge To Center MKT": row.chargeToCenterMkt,
        };
        mktPivotUniqueTypes.forEach((type) => {
          item[type] = row.values[type] || 0;
        });
        item["Grand Total"] = row.total;
        return item;
      });

      // Add Grand Totals Row
      const totalsRow: any = {
        "Business": "TỔNG CỘNG",
        "L07 (Region)": "",
        "Charge To Center MKT": "",
      };
      mktPivotUniqueTypes.forEach((type) => {
        totalsRow[type] = mktPivotGrandTotals.totals[type] || 0;
      });
      totalsRow["Grand Total"] = mktPivotGrandTotals.grandTotal;
      rows.push(totalsRow);

      const ws = XLSX.utils.json_to_sheet(prepareDataForExport(rows));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Phí MKT Local North (Pivot)");
      XLSX.writeFile(wb, `Pivot_Phi_MKT_Local_North.xlsx`);
      return;
    }

    const ws = XLSX.utils.json_to_sheet(prepareDataForExport(currentData));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, activeTab);
    XLSX.writeFile(wb, `Timesheet_Hub_${activeTab}.xlsx`);
  };

  const handleSyncToSupabase = useCallback(async () => {
    if (!isSupabaseConfigured()) {
      toast.error("Supabase chưa được cấu hình! Vui lòng cài đặt URL và Anon Key trong phần cấu hình.");
      return;
    }

    const rosterData = appData.Q_Roster || [];
    const staffData = appData.Q_Staff || [];
    const salaryData = appData.Q_Salary_Scale || [];

    if (rosterData.length === 0 && staffData.length === 0 && salaryData.length === 0) {
      toast.warning("Không có dữ liệu để đồng bộ.");
      return;
    }

    setIsSyncing(true);
    setTotalSyncRows(rosterData.length + staffData.length + salaryData.length);
    setSyncedRowsCount(0);
    setSyncProgress(0);

    try {
      let overallSuccessCount = 0;
      const totalToSync = rosterData.length + staffData.length + salaryData.length;

      // 1. Sync Staff
      if (staffData.length > 0) {
        const { successCount } = await syncEmployeesToSupabase(
          staffData,
          (current) => {
            setSyncedRowsCount(current);
            setSyncProgress(Math.round((current / totalToSync) * 100));
          }
        );
        overallSuccessCount += successCount;
      }

      // 2. Sync Salary Scale
      if (salaryData.length > 0) {
        const { successCount } = await syncSalaryScalesToSupabase(
          salaryData,
          (current) => {
            const currentTotal = staffData.length + current;
            setSyncedRowsCount(currentTotal);
            setSyncProgress(Math.round((currentTotal / totalToSync) * 100));
          }
        );
        overallSuccessCount += successCount;
      }

      // 3. Sync Roster
      if (rosterData.length > 0) {
        const { successCount } = await syncRosterToSupabase(
          rosterData,
          (current) => {
            const currentTotal = staffData.length + salaryData.length + current;
            setSyncedRowsCount(currentTotal);
            setSyncProgress(Math.round((currentTotal / totalToSync) * 100));
          }
        );
        overallSuccessCount += successCount;
      }

      toast.success(`Đồng bộ thành công ${overallSuccessCount.toLocaleString()}/${totalToSync.toLocaleString()} dòng lên Supabase.`);
      
      updateAppData((prev: any) => ({
        ...prev,
        updatedAt: new Date().toISOString()
      }), true);
      toast.success("Đã tự động lưu cứng dữ liệu trên web.");
    } catch (err: unknown) {
      console.error("Supabase Sync Error:", err);
      const errMsg = err instanceof Error ? err.message : String(err);
      toast.error(`Đồng bộ thất bại: ${errMsg}`);
      if (
        errMsg.includes("chưa tồn tại") || 
        errMsg.includes("relation") || 
        errMsg.includes("does not exist") ||
        errMsg.includes("Thiếu cột") ||
        errMsg.includes("unique_nv_ngay") ||
        errMsg.includes("ràng buộc") ||
        errMsg.includes("trùng lặp")
      ) {
        setShowSqlDialog(true);
      }
    } finally {
      setIsSyncing(false);
    }
  }, [appData.Q_Roster, appData.Q_Staff, appData.Q_Salary_Scale, updateAppData]);

  const tableRef = useRef<any>(null);

  const handleSaveData = useCallback(async () => {
    updateAppData((prev: any) => ({
      ...prev,
      updatedAt: new Date().toISOString()
    }), true);
    
    if (isSupabaseConfigured()) {
      toast.info("Đang tự động đồng bộ dữ liệu thay đổi lên Supabase...");
      await handleSyncToSupabase();
    } else {
      toast.success("Đã lưu dữ liệu thay đổi offline thành công!");
    }
  }, [updateAppData, handleSyncToSupabase]);

  const handleRestoreOriginal = useCallback(async () => {
    const confirmReset = window.confirm(
      "Bạn có chắc chắn muốn khôi phục dữ liệu ban đầu không? Toàn bộ thay đổi của bạn trên bảng Roster sẽ bị xóa và dữ liệu sẽ được đồng bộ lại với Supabase.",
    );
    if (!confirmReset) return;

    if (!isSupabaseConfigured()) {
      updateAppData((prev) => ({
        ...prev,
        Q_Roster: [...INITIAL_APP_DATA.Q_Roster],
        Q_Staff: INITIAL_APP_DATA.Q_Staff ? [...INITIAL_APP_DATA.Q_Staff] : [],
        Q_Salary_Scale: INITIAL_APP_DATA.Q_Salary_Scale ? [...INITIAL_APP_DATA.Q_Salary_Scale] : [],
        Q_Cache: INITIAL_APP_DATA.Q_Cache ? [...INITIAL_APP_DATA.Q_Cache] : [],
      }), true);
      toast.success("Đã khôi phục dữ liệu ban đầu offline thành công!");
      return;
    }

    const loadToastId = toast.loading("Đang khôi phục và đồng bộ dữ liệu với Supabase...");

    try {
      // 1. Clear old data on Supabase
      await clearSupabaseData();

      // 2. Sync Employees
      const staffData = INITIAL_APP_DATA.Q_Staff || [];
      if (staffData.length > 0) {
        await syncEmployeesToSupabase(staffData);
      }

      // 3. Sync Salary Scales
      const salaryData = INITIAL_APP_DATA.Q_Salary_Scale || [];
      if (salaryData.length > 0) {
        await syncSalaryScalesToSupabase(salaryData);
      }

      // 4. Sync Rosters
      const rosterData = INITIAL_APP_DATA.Q_Roster || [];
      if (rosterData.length > 0) {
        await syncRosterToSupabase(rosterData, () => {});
      }

      // 5. Update Local App Data to match
      updateAppData((prev) => ({
        ...prev,
        Q_Roster: [...rosterData],
        Q_Staff: [...staffData],
        Q_Salary_Scale: [...salaryData],
        Q_Cache: INITIAL_APP_DATA.Q_Cache ? [...INITIAL_APP_DATA.Q_Cache] : [],
      }), true);

      toast.dismiss(loadToastId);
      toast.success("Khôi phục và đồng bộ dữ liệu mẫu lên Supabase thành công!");
    } catch (error: any) {
      console.error("Lỗi khôi phục Supabase:", error);
      toast.dismiss(loadToastId);
      toast.error(`Khôi phục thất bại: ${error.message}`);
      if (
        error.message.includes("chưa tồn tại") || 
        error.message.includes("relation") || 
        error.message.includes("does not exist") ||
        error.message.includes("unique_nv_ngay") ||
        error.message.includes("ràng buộc") ||
        error.message.includes("trùng lặp")
      ) {
        setShowSqlDialog(true);
      }
    }
  }, [updateAppData]);

  const handleRosterCellChange = useCallback((row: any, colKey: string, value: any) => {
    updateAppData((prev) => {
      const qRoster = prev.Q_Roster || [];
      const updatedRoster = qRoster.map((r) => {
        if (r._rowId === row._rowId) {
          return {
            ...r,
            [colKey]: value,
            ...(colKey === "ngay" ? { date: value } : {}),
            ...(colKey === "date" ? { ngay: value } : {}),
            ...(colKey === "class" ? { classCode: value } : {}),
            ...(colKey === "classCode" ? { class: value } : {}),
            ...(colKey === "gio_vao" ? { from: value } : {}),
            ...(colKey === "from" ? { gio_vao: value } : {}),
            ...(colKey === "gio_ra" ? { to: value } : {}),
            ...(colKey === "to" ? { gio_ra: value } : {}),
            ...(colKey === "notes" ? { notes: value } : {}),
          };
        }
        return r;
      });
      return {
        ...prev,
        Q_Roster: updatedRoster,
      };
    });
    toast.success("Đã cập nhật dữ liệu!");
  }, [updateAppData]);

  return (
    <>
      <AnimatePresence initial={false}>
        {view === "final" && (
          <motion.div
            key="final"
            variants={containerVariants}
            initial="hidden"
            animate="visible"
            exit={{ y: "100%", opacity: 0 }}
            className="flex-1 flex flex-col min-h-0 gap-3.5 relative overflow-hidden bg-transparent"
            style={{
              marginLeft: 0,
              paddingLeft: 12,
              marginTop: 0,
              marginRight: 0,
              marginBottom: 0
            }}
          >
            {/* Header Overview - Stat Banner in Editorial Softness Style */}
            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 pb-3.5 border-b border-[rgba(61,57,53,0.08)] shrink-0 select-none">
              <div className="flex flex-col sm:flex-row sm:items-center gap-4">
                <div className="flex items-center gap-4">
                  <div className="flex flex-col">
                    <span 
                      className="font-mono text-[8.5px] tracking-[0.2em] uppercase text-[#3D3935]/50 leading-none"
                      style={{ fontWeight: 'bold', fontSize: '9.5px', lineHeight: '9.5px' }}
                    >
                      DATASET SELECTOR
                    </span>
                    <div className="flex items-center gap-3 mt-1">
                      <button
                        onClick={() => setShowSidebar(!showSidebar)}
                        className={`p-2 rounded-xl transition-all flex items-center justify-center cursor-pointer ${
                          showSidebar ? "text-white" : "bg-white border border-[rgba(61,57,53,0.08)] text-[#3D3935]/40 hover:text-[#E5A8A0] shadow-sm"
                        }`}
                        style={{ backgroundColor: showSidebar ? "#e5c1f2" : undefined }}
                        title={showSidebar ? "Ẩn bộ lọc" : "Hiện bộ lọc"}
                      >
                        <SlidersHorizontal className="w-4 h-4" />
                      </button>
                    <DropdownMenu>
                      <DropdownMenuTrigger asChild>
                        <button className="group flex items-center gap-2.5 cursor-pointer focus:outline-none">
                          <h2 className="text-3xl font-playpen font-normal text-[#630203] tracking-tight select-text group-hover:text-[#E5A8A0] transition-colors leading-tight">
                            {tabs.find((t) => t.id === activeTab)?.label || "General Roster"}
                          </h2>
                          <ChevronDown className="w-5 h-5 text-[#3D3935]/30 group-hover:text-[#E5A8A0] transition-all" />
                        </button>
                      </DropdownMenuTrigger>
                      <DropdownMenuContent align="start" className="w-64 bg-white border border-slate-200/80 shadow-2xl rounded-2xl p-1 z-[100] animate-in fade-in zoom-in-95 duration-200">
                        <DropdownMenuLabel className="font-sans font-bold uppercase text-[9px] tracking-widest text-[#3D3935]/50 px-3 py-2 select-none">
                          CHỌN BẢNG DỮ LIỆU
                        </DropdownMenuLabel>
                        <DropdownMenuSeparator className="bg-[#3D3935]/5 mx-1" />
                        {tabs.map((tab) => (
                          <DropdownMenuItem
                            key={tab.id}
                            onSelect={() => {
                              startTransition(() => {
                                setActiveTab(tab.id as any);
                                setTargetDate("");
                                setTargetCenter("");
                                setSearchTerm("");
                                navigate(location.pathname, {
                                  replace: true,
                                  state: { from: "cleared", activeTab: tab.id },
                                });
                              });
                            }}
                            className={`flex items-center justify-between gap-2 px-3 py-2.5 rounded-xl cursor-pointer text-xs font-bold transition-all ${
                              activeTab === tab.id 
                                ? "bg-[#E1F1F8] text-[#3D3935]" 
                                : "text-slate-600 hover:bg-slate-50"
                            }`}
                          >
                            <span>{tab.label}</span>
                            {activeTab === tab.id && <Check className="w-3.5 h-3.5 text-[#E5A8A0]" />}
                          </DropdownMenuItem>
                        ))}
                      </DropdownMenuContent>
                    </DropdownMenu>
                  </div>
                </div>
              </div>
            </div>

              {/* Action Buttons & Settings on the right */}
              <div className="flex items-center gap-3">
                <DropdownMenu>
                  <DropdownMenuTrigger asChild>
                    <button className="p-2 px-3 rounded-xl border border-[rgba(61,57,53,0.08)] bg-white hover:bg-slate-50 text-[#3D3935]/70 transition-all flex items-center gap-2 text-xs font-bold shadow-sm select-none cursor-pointer">
                      <Plus className="w-4 h-4 text-[#E5A8A0]" />
                      <span>Thao tác khác</span>
                      <ChevronDown className="w-3 h-3 opacity-60 text-[#3D3935]" />
                    </button>
                  </DropdownMenuTrigger>
                  <DropdownMenuContent align="end" className="w-56 bg-white border border-slate-200 shadow-xl rounded-xl p-1 z-[100]">
                    <DropdownMenuItem
                      onSelect={handleSyncToSupabase}
                      disabled={isSyncing}
                      className="flex items-center gap-2 px-3 py-2.5 rounded-lg cursor-pointer text-xs font-bold text-slate-700 hover:bg-slate-50 disabled:opacity-40"
                    >
                      <Check className="w-4 h-4 text-emerald-500" />
                      <span>{isSyncing ? "Đang đồng bộ..." : "Đồng bộ Supabase"}</span>
                    </DropdownMenuItem>

                    <DropdownMenuItem
                      onSelect={handleSaveData}
                      className="flex items-center gap-2 px-3 py-2.5 rounded-lg cursor-pointer text-xs font-bold text-slate-700 hover:bg-slate-50"
                    >
                      <Save className="w-4 h-4 text-sky-500" />
                      <span>Lưu dữ liệu thay đổi</span>
                    </DropdownMenuItem>

                    <DropdownMenuItem
                      onSelect={handleRestoreOriginal}
                      className="flex items-center gap-2 px-3 py-2.5 rounded-lg cursor-pointer text-xs font-bold text-rose-600 hover:bg-rose-50"
                    >
                      <RefreshCw className="w-4 h-4 text-rose-500" />
                      <span>Khôi phục ban đầu</span>
                    </DropdownMenuItem>

                    <DropdownMenuItem
                      onSelect={handleExportExcel}
                      disabled={currentData.length === 0}
                      className="flex items-center gap-2 px-3 py-2.5 rounded-lg cursor-pointer text-xs font-bold text-slate-700 hover:bg-slate-50 disabled:opacity-40"
                    >
                      <Download className="w-4 h-4 text-slate-500" />
                      <span>Xuất Excel</span>
                    </DropdownMenuItem>
                  </DropdownMenuContent>
                </DropdownMenu>

                <button
                  onClick={() => window.dispatchEvent(new CustomEvent("open-ui-settings"))}
                  className="p-2 rounded-xl border border-[rgba(61,57,53,0.08)] bg-white hover:bg-slate-50 text-slate-400 hover:text-[#E5A8A0] transition-all shadow-sm flex items-center justify-center cursor-pointer"
                  title="Cài đặt Giao diện"
                >
                  <Settings className="w-4 h-4" />
                </button>
              </div>
            </div>

            {/* Inner Content Area holding Sidebar and Table */}
            <div className="flex-1 flex flex-col lg:flex-row min-h-0 gap-4 relative overflow-hidden">
              {/* Left Panel: Sidebar Controls */}
              {showSidebar && (
                <div 
                  className="w-full lg:w-[220px] shrink-0 flex flex-col pr-0 lg:pr-1 h-full select-none animate-in fade-in slide-in-from-left duration-500"
                  style={{ borderRadius: '48px', borderWidth: '0px' }}
                >
                <div className="bg-[#FAF5EE]/80 backdrop-blur-sm p-4 rounded-[48px] border border-[#3D3935]/5 flex flex-col h-full justify-between shadow-sm">
                  <div className="flex flex-col gap-4">
                    <div className="flex items-center justify-between mb-2">
                      <span className="font-sans font-black text-[10px] tracking-widest text-[#3D3935]/40 uppercase">Filters</span>
                      <button 
                        onClick={() => setShowSidebar(false)}
                        className="text-[#3D3935]/30 hover:text-rose-500 transition-colors cursor-pointer p-1 rounded-md"
                        style={{ backgroundColor: '#efc9ea' }}
                      >
                        <XCircle className="w-4 h-4" />
                      </button>
                    </div>
                    {/* Start Date Selection */}
                    <div className="flex flex-col gap-1 relative">
                      <span 
                        className="font-mono text-[8px] tracking-[0.2em] uppercase text-[#3D3935]/50 leading-none"
                        style={{ fontWeight: 'bold', fontSize: '10px', lineHeight: '10px' }}
                      >
                        START DATE
                      </span>
                      <Popover>
                        <PopoverTrigger asChild>
                          <button className="bg-white rounded-lg px-3 py-2 border border-[rgba(61,57,53,0.08)] hover:border-[#E5A8A0] focus:outline-none transition-all w-full flex items-center justify-between cursor-pointer select-none text-[11px] font-bold text-[#3D3935]">
                            <span>
                              {fromDate
                                ? format(new Date(`${fromDate}T00:00:00`), "dd/MM/yyyy")
                                : "Chọn ngày"}
                            </span>
                          </button>
                        </PopoverTrigger>
                        <PopoverContent className="w-auto p-0 z-[100] bg-white border border-slate-200 shadow-xl rounded-xl">
                          <Calendar
                            mode="single"
                            selected={fromDate ? new Date(`${fromDate}T00:00:00`) : undefined}
                            onSelect={(d) => {
                              startTransition(() => {
                                const newDate = d ? format(d, "yyyy-MM-dd") : "";
                                setFromDate(newDate);
                                setTargetDate("");
                                setTargetCenter("");
                              });
                            }}
                            className="p-3 pointer-events-auto bg-white"
                          />
                        </PopoverContent>
                      </Popover>
                    </div>

                    {/* End Date Selection */}
                    <div className="flex flex-col gap-1 relative">
                      <span 
                        className="font-mono text-[8px] tracking-[0.2em] uppercase text-[#3D3935]/50 leading-none"
                        style={{ fontWeight: 'bold', fontSize: '10px', lineHeight: '10px' }}
                      >
                        END DATE
                      </span>
                      <Popover>
                        <PopoverTrigger asChild>
                          <button className="bg-white rounded-lg px-3 py-2 border border-[rgba(61,57,53,0.08)] hover:border-[#E5A8A0] focus:outline-none transition-all w-full flex items-center justify-between cursor-pointer select-none text-[11px] font-bold text-[#3D3935]">
                            <span>
                              {toDate
                                ? format(new Date(`${toDate}T00:00:00`), "dd/MM/yyyy")
                                : "Chọn ngày"}
                            </span>
                          </button>
                        </PopoverTrigger>
                        <PopoverContent className="w-auto p-0 z-[100] bg-white border border-slate-200 shadow-xl rounded-xl">
                          <Calendar
                            mode="single"
                            selected={toDate ? new Date(`${toDate}T00:00:00`) : undefined}
                            onSelect={(d) => {
                              startTransition(() => {
                                const newDate = d ? format(d, "yyyy-MM-dd") : "";
                                setToDate(newDate);
                                setTargetDate("");
                                setTargetCenter("");
                              });
                            }}
                            className="p-3 pointer-events-auto bg-white"
                          />
                        </PopoverContent>
                      </Popover>
                    </div>

                    {/* Search Term input */}
                    <div className="flex flex-col gap-1 relative">
                      <span 
                        className="font-mono text-[8px] tracking-[0.2em] uppercase text-[#3D3935]/50 leading-none"
                        style={{ fontWeight: 'bold', fontSize: '10px', lineHeight: '10px' }}
                      >
                        KEYWORD
                      </span>
                      <div className="relative">
                        <input
                          type="text"
                          placeholder="Tìm kiếm..."
                          value={searchTerm}
                          onChange={(e) => setSearchTerm(e.target.value)}
                          className="bg-white rounded-lg pl-8 pr-2.5 py-2 border border-[rgba(61,57,53,0.08)] hover:border-[#E5A8A0] focus:border-[#E5A8A0] focus:outline-none transition-all w-full text-[11px] font-bold text-[#3D3935]"
                        />
                        <Search className="w-3.5 h-3.5 text-[#3D3935]/30 absolute left-2.5 top-1/2 -translate-y-1/2" />
                      </div>
                    </div>

                    <button
                      onClick={() => {
                        startTransition(() => {
                          setDebouncedFromDate(fromDate);
                          setDebouncedToDate(toDate);
                        });
                        toast.success("Đã cập nhật bảng dữ liệu!");
                      }}
                      className="w-full bg-[#3D3935] hover:bg-[#3D3935]/90 text-white font-black uppercase tracking-wider text-[10px] py-2 rounded-lg transition-all cursor-pointer font-sans shadow-sm mt-1"
                    >
                      UPDATE GRID
                    </button>
                  </div>

                  <div className="flex flex-col gap-2 pt-3 border-t border-[rgba(61,57,53,0.08)] mt-auto w-full">
                    <button
                      onClick={handleClearFilters}
                      className="text-[9px] font-mono tracking-[0.2em] text-[#3D3935]/50 hover:text-[#E5A8A0] text-center transition-colors cursor-pointer select-none block uppercase"
                    >
                      RESET DEFAULTS
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* Right Panel: Content Grid */}
            <div 
              className="flex-1 flex flex-col min-h-0 h-full overflow-hidden relative animate-in fade-in slide-in-from-right duration-500 bg-white"
              style={{ borderRadius: '48px', borderWidth: '0px' }}
            >
              <div className="flex-1 flex flex-col min-h-0 relative overflow-hidden">
                {isSyncing && (
                  <div className="absolute top-0 right-0 p-4 z-[100]">
                    <div className="flex items-center gap-2 px-3 py-1.5 bg-sky-50 rounded-full border border-sky-100 shadow-sm animate-pulse">
                      <RefreshCw className="w-3 h-3 text-sky-600 animate-spin" />
                      <span className="text-[9px] font-black text-sky-700 uppercase tracking-wider">{syncProgress}% Synced</span>
                    </div>
                  </div>
                )}
                {currentData.length === 0 ? (
                  <div className="flex-1 flex flex-col items-center justify-center text-primary/10 p-12 select-none">
                    <div className="w-20 h-20 bg-[#FAF5EE] rounded-full flex items-center justify-center mb-6 border border-[#3D3935]/5">
                      <PuppyLogo size={44} className="opacity-20 grayscale" />
                    </div>
                    <p className="font-bold uppercase text-base tracking-tight text-[#3D3935]/40">
                      Chưa có dữ liệu
                    </p>
                    <p className="text-[10px] font-bold uppercase opacity-30 tracking-widest mt-2 text-center max-w-xs font-sans leading-relaxed">
                      Dữ liệu trống hoặc không khớp với ngày đang chọn.<br/>Vui lòng vào phần Summary để tải lên dữ liệu.
                    </p>
                  </div>
                ) : (
                  <div className="flex-1 flex flex-col min-h-0 overflow-hidden relative">
                      {/* Search Result Feedback when empty */}
                      {searchTerm && searchData.length === 0 && (
                        <div className="absolute inset-0 z-[100] flex items-center justify-center bg-white/85 backdrop-blur-sm animate-in fade-in duration-300 rounded-[32px] overflow-hidden">
                          <div className="bg-white p-8 rounded-2xl border border-slate-200 shadow-xl flex flex-col items-center text-center max-w-sm">
                            <div className="w-16 h-16 bg-rose-50 rounded-full flex items-center justify-center mb-4 border border-rose-100 text-rose-500 shadow-inner">
                              <XCircle className="w-8 h-8" />
                            </div>
                            <h3 
                              className="text-xl font-bold text-slate-800 tracking-tight mb-2"
                              style={{ fontSize: '14px' }}
                            >
                              Không tìm thấy kết quả
                            </h3>
                            <p className="text-[11px] font-medium text-slate-500 leading-relaxed mb-6 font-sans">
                              Không tìm thấy bản ghi nào khớp với từ khóa "{searchTerm}" trong khoảng thời gian này.
                            </p>
                            <button
                              onClick={handleClearFilters}
                              className="py-2.5 px-6 bg-slate-900 text-white text-[10px] font-black uppercase tracking-wider rounded-full hover:bg-slate-800 transition-all cursor-pointer font-sans"
                            >
                              Xóa lọc
                            </button>
                          </div>
                        </div>
                      )}

                      {/* Sync/Load Indicator */}
                      {isCalculating || isPending ? (
                        <div className="flex-1 flex flex-col items-center justify-center bg-white/60 relative z-10 p-12">
                          <div className="relative">
                            <div className="w-12 h-12 border-3 border-rose-200 border-t-rose-500 rounded-full animate-spin" />
                          </div>
                          <p className="mt-6 text-[10px] font-black uppercase tracking-[0.25em] text-rose-500/80 animate-pulse font-sans">
                            {isPending
                              ? "Đang chuyển bảng..."
                              : `Đang xử lý ${appData.Q_Roster?.length || 0} dòng dữ liệu...`}
                          </p>
                        </div>
                      ) : activeTab === "mkt_local_north" ? (
                        <MktLocalNorthPivotTable
                          rows={mktPivotRows}
                          types={mktPivotUniqueTypes}
                          grandTotals={mktPivotGrandTotals}
                        />
                      ) : activeTab === "roster_raw" ? (
                        <RosterRawTable
                          data={searchData}
                          onFilteredDataChange={(d) => setTableFilteredCount(d.length)}
                          onCellChange={handleRosterCellChange}
                        />
                      ) : activeTab === "employee" ? (
                        <EmployeeTable
                          data={searchData}
                          calculatedRosterData={calculatedRosterData}
                          onFilteredDataChange={(d) => setTableFilteredCount(d.length)}
                        />
                      ) : activeTab === "center" ? (
                        <CenterTable
                          data={searchData}
                          onFilteredDataChange={(d) => setTableFilteredCount(d.length)}
                        />
                      ) : null}
                    </div>
                  )}
                </div>
              </div>
            </div>
          </motion.div>
        )}
        {view === "upload" && (
          <motion.div
            key="upload"
            initial={{ y: "-100%", opacity: 0 }}
            animate={{ y: 0, opacity: 1 }}
            exit={{ y: "-100%", opacity: 0 }}
            transition={{ type: "spring", damping: 25, stiffness: 200 }}
            className="absolute inset-3 flex flex-col p-0 m-3"
          >
            <TimesheetSummaryPage onBack={() => setView("final")} />
          </motion.div>
        )}
      </AnimatePresence>

      <Dialog open={showSqlDialog} onOpenChange={setShowSqlDialog}>
        <DialogContent className="max-w-2xl bg-white rounded-3xl border-none shadow-2xl p-0 overflow-hidden">
          <div className="bg-sky-600 p-8 text-white">
            <DialogHeader>
              <DialogTitle className="text-2xl font-black uppercase tracking-wider">Thiết lập & Cập nhật Supabase</DialogTitle>
              <DialogDescription className="text-sky-100 font-medium text-[11px] leading-relaxed">
                Bảng 'roster_cham_cong' chưa tồn tại, thiếu cột (như charge_to_center_mkt) hoặc đang bị ràng buộc cũ (như unique_nv_ngay - giới hạn mỗi người 1 ca/ngày). Vui lòng copy toàn bộ script bên dưới và chạy trong SQL Editor của Supabase để cập nhật cấu trúc bảng chính xác nhất.
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
    </>
  );
}
