/* eslint-disable react-hooks/set-state-in-effect */
import { useState, useEffect, useRef } from "react";
import { Outlet, useLocation, Link } from "react-router";
import { Navbar } from "../components/layouts/Navbar";
import { UiSettingsModal } from "../components/UiSettingsModal";
import { ErrorBoundary } from "../components/shared/ErrorBoundary";
import { BarChart3, Database, ShieldCheck, Wallet, Menu } from "lucide-react";

// ── Root không dùng framer-motion để tránh layout thrashing trên shell layout ──
export function Root() {
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const location = useLocation();

  const [isShortcutMenuOpen, setIsShortcutMenuOpen] = useState(false);
  const menuRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    setIsMobileMenuOpen(false);
  }, [location.pathname]);

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (menuRef.current && !menuRef.current.contains(event.target as Node)) {
        setIsShortcutMenuOpen(false);
      }
    }
    if (isShortcutMenuOpen) {
      document.addEventListener("mousedown", handleClickOutside);
    }
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [isShortcutMenuOpen]);

  useEffect(() => {
    const handleOpenSettings = () => setIsSettingsOpen(true);
    window.addEventListener("open-ui-settings", handleOpenSettings);
    return () => window.removeEventListener("open-ui-settings", handleOpenSettings);
  }, []);

  return (
    <div className="flex h-screen overflow-hidden font-sans text-foreground bg-transparent">
      {/* Mobile Sidebar Overlay — CSS transition thay vì framer-motion */}
      <div
        onClick={() => setIsMobileMenuOpen(false)}
        className={`fixed inset-0 bg-black/40 backdrop-blur-md z-[60] lg:hidden transition-opacity duration-300
          ${isMobileMenuOpen ? "opacity-100 pointer-events-auto" : "opacity-0 pointer-events-none"}`}
      />

      {/* Main Content Area */}
      <div className="flex-1 flex flex-col overflow-hidden min-w-0 relative bg-transparent pt-[15px] pb-[15px] px-3">
        <div className="bg-transparent mx-4 mt-[-13px] relative z-40">
          <Navbar
            onToggleMobileMenu={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
            onOpenSettings={() => setIsSettingsOpen(true)}
          />
        </div>

        <main className="flex-1 flex flex-col min-h-0 relative">
          <ErrorBoundary>
            <div className="flex-1 flex flex-col min-h-0">
              <Outlet />
            </div>
          </ErrorBoundary>

          {/* Quick Shortcuts Exclamation Floating Dock */}
          <div ref={menuRef} className="fixed bottom-3 right-6 z-[99] flex flex-col items-end gap-2">
            {isShortcutMenuOpen && (
              <div
                className="bg-white/95 backdrop-blur-md border border-primary/20 shadow-2xl rounded-2xl p-2 w-60 mb-2 transition-all duration-300 transform scale-100 origin-bottom-right"
                style={{ transformOrigin: 'bottom right' }}
              >
                <div className="text-[10px] font-black uppercase tracking-widest text-primary/70 px-3 py-2 border-b border-[#e6dfd3] mb-1.5 whitespace-nowrap">
                  Chọn Bảng Đích
                </div>
                <div className="flex flex-col gap-1">
                  {[
                    { to: "/centers", icon: BarChart3, label: "Bảng Timesheet" },
                    { to: "/audit", icon: ShieldCheck, label: "Bảng Audit" },
                    { to: "/master-ae", icon: Database, label: "Bảng Master" },
                    { to: "/hold-dashboard", icon: Wallet, label: "Bảng Balance" },
                  ].map((item) => {
                    const isActive = location.pathname === item.to;
                    return (
                      <Link
                        key={item.to}
                        to={item.to}
                        onClick={() => setIsShortcutMenuOpen(false)}
                        className={`flex items-center gap-2.5 px-3 py-2.5 rounded-xl transition-all duration-200 text-xs font-bold ${
                          isActive
                            ? "bg-primary text-white shadow-md relative z-10"
                            : "hover:bg-primary/5 text-slate-700 hover:text-primary relative z-10"
                        }`}
                      >
                        <item.icon className={`w-4 h-4 shrink-0 ${isActive ? "text-white" : "text-primary"}`} />
                        <span className="truncate">{item.label}</span>
                      </Link>
                    );
                  })}
                </div>
              </div>
            )}

            <button
              onClick={() => setIsShortcutMenuOpen(!isShortcutMenuOpen)}
              className={`w-10 h-10 rounded-full bg-primary text-white font-black text-lg flex items-center justify-center hover:scale-110 active:scale-95 transition-all duration-300 shadow-xl border border-primary/20 cursor-pointer ${
                isShortcutMenuOpen ? "bg-primary/90 ring-4 ring-primary/20" : ""
              }`}
              title="Menu"
            >
              <Menu className="w-5 h-5" />
            </button>
          </div>
        </main>

        <UiSettingsModal
          isOpen={isSettingsOpen}
          onClose={() => setIsSettingsOpen(false)}
        />
      </div>
    </div>
  );
}
