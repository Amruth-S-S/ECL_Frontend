"use client";

import React, { useState, useEffect, useCallback, useRef, useMemo } from 'react';
import { useTranslation } from 'react-i18next';
import { useRouter, usePathname } from 'next/navigation';
import { useLanguage } from '../context/LanguageContext';
import { translateBatch } from '../utils/translateService';
import Link from 'next/link';
import {
  LayoutDashboard,
  ChevronDown,
  ChevronRight,
  Plus,
  Edit2,
  Trash2,
  ChevronLeft,
  X,
  User,
  NotebookText,
  Menu,
  Settings,
  LogOut,
  ChartColumnDecreasing,
  Search,
  Eye,
  EyeOff
} from 'lucide-react';
import './Toast.css';
import { toast, ToastContainer } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';

// ─── Interfaces ────────────────────────────────────────────────────────────────
interface Board {
  name: string;
  is_active: boolean;
  path?: string;
}

interface MainBoard {
  id?: string | number;
  main_board_id: string;
  name: string;
  boards: { [key: string]: Board };
}

type SelectedBoard = {
  mainBoardId: string;
  boardId?: string;
  boardName?: string;
} | null;

interface DemoMainBoard {
  id: number;
  name: string;
  main_board_type?: string;
}

interface DemoBoard {
  id: number;
  name: string;
  main_board_id: number;
  is_active?: boolean;
  customer_db_key?: string;
}

interface SidebarProps {
  clientUserId?: string | number;
}

interface UserData {
  email: string;
  userId: string;
  userRole: string;
  userName: string;
}

// Reads the logged-in user out of sessionStorage/localStorage synchronously.
// Used as a lazy useState initializer so the value is available on the very
// first client render — not two effect-cycles later (isMounted -> read ->
// setState), which was leaving the sidebar's nav/logo fetches with an empty
// user id whenever this component mounted fresh (e.g. navigating straight to
// /Container with no board already selected, instead of reloading a page
// that already had ?main_board_id=&board_id= in the URL from a prior visit).
function readStoredUser(): UserData {
  if (typeof window === 'undefined') {
    return { email: '', userId: '', userRole: '', userName: '' };
  }
  try {
    const sessionData = localStorage.getItem('currentUserData');
    if (sessionData) {
      const p = JSON.parse(sessionData);
      if (p.userId) {
        return { email: p.email || '', userId: String(p.userId), userRole: p.userRole || '', userName: p.userName || '' };
      }
    }
  } catch { /* ignore corrupt session data */ }
  try {
    const userId = localStorage.getItem('loggedInUserId') || '';
    if (userId) {
      return {
        email: localStorage.getItem('loggedInUserEmail') || '',
        userId,
        userRole: localStorage.getItem('loggedInUserRole') || '',
        userName: localStorage.getItem('loggedInUserName') || '',
      };
    }
  } catch { /* ignore */ }
  return { email: '', userId: '', userRole: '', userName: '' };
}

// ─── Global Loader Component ────────────────────────────────────────────────────
const GlobalLoader = ({ message = "Loading..." }: { message?: string }) => (
  <div className="fixed inset-0 z-[999] flex items-center justify-center bg-black/60 backdrop-blur-sm">
    <div className="flex flex-col items-center justify-center">
      <div className="relative w-20 h-20">
        {Array.from({ length: 12 }).map((_, i) => {
          const angle = (i / 12) * 360;
          const rad = (angle * Math.PI) / 180;
          const x = 50 + 40 * Math.sin(rad);
          const y = 50 - 40 * Math.cos(rad);
          return (
            <div
              key={i}
              className="absolute w-2 h-2 rounded-full bg-white"
              style={{
                left: `${x}%`,
                top: `${y}%`,
                transform: 'translate(-50%, -50%)',
                opacity: 0.2 + (i / 12) * 0.8,
                animation: `spin-dot 1.2s linear infinite`,
                animationDelay: `${-(12 - i) * (1.2 / 12)}s`,
              }}
            />
          );
        })}
      </div>
      <p className="mt-3 text-white text-sm font-medium tracking-wide animate-pulse">
        {message}
      </p>
    </div>
    <style>{`
      @keyframes spin-dot {
        0% { opacity: 0.2; }
        50% { opacity: 1; }
        100% { opacity: 0.2; }
      }
    `}</style>
  </div>
);

// ─── Main Sidebar Component ─────────────────────────────────────────────────────
const Sidebar: React.FC<SidebarProps> = ({ }) => {
  const { t } = useTranslation();
  const { language } = useLanguage();
  // Map of original name → translated name for current language
  const [boardNameMap, setBoardNameMap] = useState<Record<string, string>>({});
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  const [sidebarWidth, setSidebarWidth] = useState(260);
  const [isResizing, setIsResizing] = useState(false);
  const sidebarRef = useRef(null);
  const [userRole, setUserRole] = useState<string | null>(null);
  const [showModal, setShowModal] = useState<boolean>(false);
  const [showSubMenu, setShowSubMenu] = useState(false);
  const [showAddDialog, setShowAddDialog] = useState(false);
  const [newCollectionName, setNewCollectionName] = useState('');
  const [isDropdownOpen, setIsDropdownOpen] = useState(false);

  // Mobile
  const [isMobile, setIsMobile] = useState(false);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

  // Logo — display only; uploading now lives in Master Data (see ExcelTableComponent.tsx)
  const [currentLogo, setCurrentLogo] = useState<string | null>(null);

  // Settings / Password
  const [isSettingsModalOpen, setIsSettingsModalOpen] = useState(false);
const [passwordData, setPasswordData] = useState({ currentPassword: '', newPassword: '', confirmPassword: '' });
  const [showCurrentPassword, setShowCurrentPassword] = useState(false);
  const [showNewPassword, setShowNewPassword] = useState(false);
  const [showConfirmPassword, setShowConfirmPassword] = useState(false);
  const [isUpdatingPassword, setIsUpdatingPassword] = useState(false);

  const [clientUserId, setClientUserId] = useState<string | null>(() => readStoredUser().userId || null);
  const [loadingMainBoard, setLoadingMainBoard] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  // Board states
  const [newBoardName, setNewBoardName] = useState('');
  const [customerDbKey, setCustomerDbKey] = useState('');
  const [isEditMode, setIsEditMode] = useState(false);
  const [editingBoardId, setEditingBoardId] = useState<string | null>(null);
  // Whether the "+ CSV-Object" button should show up later on this board's
  // Manage Tables screen — no backend field for this yet, so it's persisted
  // per-board in localStorage and read back by Container.tsx.
  const [csvUploadEnabled, setCsvUploadEnabled] = useState(true);

  const [selectedBoard, setSelectedBoard] = useState<SelectedBoard>(null);
  const [navItems, setNavItems] = useState<MainBoard[]>([]);
  const [activeMainBoard, setActiveMainBoard] = useState<string | null>(null);
  const router = useRouter();
  const pathname = usePathname();
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [mainBoardName, setMainBoardName] = useState("");
  const [refreshTrigger, setRefreshTrigger] = useState(0);
  const [activeBoardId, setActiveBoardId] = useState<string | null>(null);
  const [, setMainBoardId] = useState(null);

  // ── Global Loading State ──────────────────────────────────────────────────────
  const [globalLoading, setGlobalLoading] = useState(false);
  const [globalLoadingMessage, setGlobalLoadingMessage] = useState("Loading...");

  const showGlobalLoader = (msg: string) => {
    setGlobalLoadingMessage(msg);
    setGlobalLoading(true);
  };
  const hideGlobalLoader = () => setGlobalLoading(false);

  const [deletingBoards, setDeletingBoards] = useState<{ [key: string]: boolean }>({});
  const [hoveredItem, setHoveredItem] = useState<string | null>(null);
  const [showUserDropdown, setShowUserDropdown] = useState(false);

  // Demo Reference
  const [isDemoRefOpen, setIsDemoRefOpen] = useState(false);
  const [demoMainBoards, setDemoMainBoards] = useState<DemoMainBoard[]>([]);
  const [demoBoards, setDemoBoards] = useState<DemoBoard[]>([]);
  const [activeDemoMainBoard, setActiveDemoMainBoard] = useState<string | null>(null);
  const [isDemoLoading, setIsDemoLoading] = useState(false);

  // Demo Create/Edit Board modal
  const [showDemoModal, setShowDemoModal] = useState(false);
  const [demoBoardName, setDemoBoardName] = useState('');
  const [demoCustomerDbKey, setDemoCustomerDbKey] = useState('');
  const [selectedDemoMainBoardId, setSelectedDemoMainBoardId] = useState<number | null>(null);
  const [isCreatingDemoBoard, setIsCreatingDemoBoard] = useState(false);
  const [isDemoEditMode, setIsDemoEditMode] = useState(false);
  const [editingDemoBoardId, setEditingDemoBoardId] = useState<number | null>(null);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const [userData, setUserData] = useState<UserData>(() => readStoredUser());
  const [isMounted, setIsMounted] = useState(false);

  const [searchQuery, setSearchQuery] = useState('');
  const [isSearchOpen, setIsSearchOpen] = useState(false);
  const searchInputRef = useRef<HTMLInputElement>(null);
  const [customerDbOptions, setCustomerDbOptions] = useState<string[]>([]);
  // Display-only — the real value (with the "customer_db_" prefix) is what
  // actually gets sent to the API; this just hides the prefix in the UI.
  const displayDbName = (db: string) => db.replace(/^customer_db_/i, '');

  const [confirmation, setConfirmation] = useState({ isOpen: false, message: '', boardId: '', mainBoardId: '' });

  const API_BASE_URL = process.env.NEXT_PUBLIC_API_BASE_URL;
  const EXCEL_API_KEY = process.env.NEXT_PUBLIC_API_KEY || '';

  const toggleDropdown = () => setIsDropdownOpen(!isDropdownOpen);

  const fetchCustomerDbKeys = async () => {
  try {
    let userId = '';

    if (typeof window !== 'undefined') {
      const s = localStorage.getItem('currentUserData');
      if (s) {
        const d = JSON.parse(s);
        userId = d.userId || d.user_id || d.id;
      }
    }

    if (!userId) return;

    const res = await fetch(
      `${API_BASE_URL}/main-boards/boards/available-customer-dbs?user_id=${userId}`,
      {
        method: "GET",
        headers: {
          "accept": "application/json",
          "X-API-Key": EXCEL_API_KEY
        }
      }
    );

    if (res.ok) {
      const data = await res.json();
      setCustomerDbOptions(data); // <-- important
    } else {
      console.error("Failed to fetch DB keys");
    }
  } catch (err) {
    console.error("Error:", err);
  }
};

  // ─── Demo Reference API ───────────────────────────────────────────────────────
  const fetchDemoMainBoards = async () => {
    try {
      const res = await fetch(`${API_BASE_URL}/demo/main-boards`, {
        headers: { accept: "application/json", "X-API-Key": EXCEL_API_KEY },
      });
      if (res.ok) {
        const json = await res.json();
        const list = Array.isArray(json) ? json : (json.data ?? json.items ?? []);
        setDemoMainBoards(list);
      }
    } catch (err) {
      console.error("Error fetching demo main boards:", err);
    }
  };

  const fetchDemoBoards = async () => {
    try {
      const res = await fetch(`${API_BASE_URL}/demo/boards`, {
        headers: { accept: "application/json", "X-API-Key": EXCEL_API_KEY },
      });
      if (res.ok) {
        const json = await res.json();
        const list = Array.isArray(json) ? json : (json.data ?? json.items ?? []);
        setDemoBoards(list);
      }
    } catch (err) {
      console.error("Error fetching demo boards:", err);
    }
  };

  const toggleDemoRef = () => {
    const opening = !isDemoRefOpen;
    setIsDemoRefOpen(opening);
    setActiveDemoMainBoard(null);
    if (opening) {
      setIsDemoLoading(true);
      Promise.all([fetchDemoMainBoards(), fetchDemoBoards()]).finally(() =>
        setIsDemoLoading(false)
      );
    }
  };

  const openDemoCreateModal = (e: React.MouseEvent, mbId: number) => {
    e.stopPropagation();
    setSelectedDemoMainBoardId(mbId);
    setDemoBoardName('');
    setDemoCustomerDbKey('');
    setShowDemoModal(true);
    fetchCustomerDbKeys();
  };

  const closeDemoModal = () => {
    setShowDemoModal(false);
    setDemoBoardName('');
    setDemoCustomerDbKey('');
    setSelectedDemoMainBoardId(null);
    setIsDemoEditMode(false);
    setEditingDemoBoardId(null);
  };

  const openDemoEditModal = (e: React.MouseEvent, board: DemoBoard) => {
    e.stopPropagation();
    setIsDemoEditMode(true);
    setEditingDemoBoardId(board.id);
    setSelectedDemoMainBoardId(board.main_board_id);
    setDemoBoardName(board.name);
    setDemoCustomerDbKey(board.customer_db_key ?? '');
    setShowDemoModal(true);
    fetchCustomerDbKeys();
  };

  const handleSaveDemoBoard = async () => {
    if (!demoBoardName.trim() || !demoCustomerDbKey || !selectedDemoMainBoardId) return;
    setIsCreatingDemoBoard(true);
    try {
      if (isDemoEditMode && editingDemoBoardId) {
        const params = new URLSearchParams({
          demo_user_id: String(clientUserId),
          name: demoBoardName.trim(),
          customer_db_key: demoCustomerDbKey,
        });
        const res = await fetch(`${API_BASE_URL}/demo/boards/${editingDemoBoardId}?${params}`, {
          method: 'PUT',
          headers: { accept: 'application/json', 'X-API-Key': EXCEL_API_KEY },
        });
        if (res.ok) {
          toast.success('Demo board updated!');
          closeDemoModal();
          fetchDemoBoards();
        } else {
          const err = await res.json();
          toast.error(err.detail || 'Failed to update demo board');
        }
      } else {
        const params = new URLSearchParams({
          demo_user_id: String(clientUserId),
          main_board_id: String(selectedDemoMainBoardId),
          name: demoBoardName.trim(),
          customer_db_key: demoCustomerDbKey,
        });
        const res = await fetch(`${API_BASE_URL}/demo/boards?${params}`, {
          method: 'POST',
          headers: { accept: 'application/json', 'X-API-Key': EXCEL_API_KEY },
        });
        if (res.ok) {
          toast.success('Demo board created!');
          closeDemoModal();
          fetchDemoBoards();
        } else {
          const err = await res.json();
          toast.error(err.detail || 'Failed to create demo board');
        }
      }
    } catch {
      toast.error(isDemoEditMode ? 'Error updating demo board' : 'Error creating demo board');
    } finally {
      setIsCreatingDemoBoard(false);
    }
  };

  const handleDeleteDemoBoard = (e: React.MouseEvent, board: DemoBoard) => {
    e.stopPropagation();
    const ConfirmToast = ({ closeToast }: { closeToast: () => void }) => (
      <div className="p-3 bg-white rounded-lg shadow-lg">
        <p className="text-gray-800 text-sm mb-3">Delete <strong>{board.name}</strong>? This cannot be undone.</p>
        <div className="flex justify-end space-x-2">
          <button onClick={() => closeToast()} className="px-3 py-1.5 text-sm bg-gray-200 text-gray-800 hover:bg-gray-300 rounded">Cancel</button>
          <button
            onClick={async () => {
              closeToast();
              try {
                const res = await fetch(
                  `${API_BASE_URL}/demo/boards/${board.id}?demo_user_id=${clientUserId}`,
                  { method: 'DELETE', headers: { accept: 'application/json', 'X-API-Key': EXCEL_API_KEY } }
                );
                if (res.ok) { toast.success('Demo board deleted!'); fetchDemoBoards(); }
                else toast.error('Failed to delete demo board');
              } catch { toast.error('Error deleting demo board'); }
            }}
            className="px-3 py-1.5 text-sm bg-red-500 text-white hover:bg-red-600 rounded"
          >Delete</button>
        </div>
      </div>
    );
    toast(<ConfirmToast closeToast={() => {}} />, {
      position: 'bottom-center', autoClose: false, closeButton: false, closeOnClick: false, draggable: false,
      className: '!bg-transparent !shadow-none',
    });
  };

  // ─── Password Update ──────────────────────────────────────────────────────────
  const handlePasswordUpdate = async () => {
    if (!passwordData.currentPassword || !passwordData.newPassword || !passwordData.confirmPassword) {
      toast.error('Please fill in all password fields'); return;
    }
    if (passwordData.newPassword !== passwordData.confirmPassword) {
      toast.error('New password and confirm password do not match'); return;
    }
    if (passwordData.newPassword.length < 8) {
      toast.error('Password must be at least 8 characters long'); return;
    }
    setIsUpdatingPassword(true);
    try {
      const response = await fetch(`${API_BASE_URL}/auth/change-password?user_id=${userData.userId}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-API-Key': EXCEL_API_KEY },
        body: JSON.stringify({
          current_password: passwordData.currentPassword,
          new_password: passwordData.newPassword,
          confirm_password: passwordData.confirmPassword,
        }),
      });
      if (response.ok) {
        const result = await response.json();
        toast.success(result.message || 'Password updated successfully!');
        setIsSettingsModalOpen(false);
        setPasswordData({ currentPassword: '', newPassword: '', confirmPassword: '' });
      } else {
        const errorData = await response.json();
        toast.error(errorData.message || 'Failed to update password');
      }
    } catch {
      toast.error('An error occurred while updating password');
    } finally {
      setIsUpdatingPassword(false);
    }
  };

  const handleSettingsClick = () => { setIsSettingsModalOpen(true); setShowUserDropdown(false); closeMobileMenu(); };

  // ─── Mobile ──────────────────────────────────────────────────────────────────
  useEffect(() => { setIsMounted(true); }, []);

  useEffect(() => {
    const checkMobile = () => {
      const mobile = window.innerWidth <= 768;
      setIsMobile(mobile);
      if (mobile) { setIsSidebarOpen(false); setSidebarWidth(0); }
      else { setIsSidebarOpen(true); setSidebarWidth(260); setIsMobileMenuOpen(false); }
    };
    checkMobile();
    window.addEventListener('resize', checkMobile);
    return () => window.removeEventListener('resize', checkMobile);
  }, []);

  useEffect(() => {
    document.body.style.overflow = (isMobile && isMobileMenuOpen) ? 'hidden' : 'unset';
    return () => { document.body.style.overflow = 'unset'; };
  }, [isMobile, isMobileMenuOpen]);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (isMobile && isMobileMenuOpen && sidebarRef.current &&
        !(sidebarRef.current as HTMLElement).contains(event.target as Node)) {
        setIsMobileMenuOpen(false);
      }
    };
    if (isMobileMenuOpen) document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [isMobile, isMobileMenuOpen]);

  const toggleMobileMenu = () => setIsMobileMenuOpen(!isMobileMenuOpen);
  const closeMobileMenu = () => { if (isMobile) setIsMobileMenuOpen(false); };

  // ─── Admin nav items ──────────────────────────────────────────────────────────
  const adminNavigationItems = [
    { id: 'users', label: 'User', href: '/UserList' },
    { id: 'organization', label: 'Organization', href: '/create-org-list' },
    { id: 'members', label: 'Members', href: '/member-list' },
     { id: 'groups', label: 'Groups', href: '/user-groups' },
    // { id: 'board-assignment', label: 'Assign Boards to Roles', href: '/BoardRoleAssignment' },
    // { id: 'user-assignment', label: 'Assign User to Roles', href: '/UserRoleAssignment' }
  ];

  // ─── Search filtering ─────────────────────────────────────────────────────────
  const filteredNavItems = useMemo(() => {
    const sorted = [...navItems].sort((a, b) => a.name.localeCompare(b.name));
    if (!searchQuery.trim()) return sorted;
    const query = searchQuery.toLowerCase();
    const matchingMainBoards = new Set<string>();
    return sorted.map(item => {
      const mainBoardMatches = item.name.toLowerCase().includes(query);
      const matchingBoards: { [key: string]: Board } = {};
      let hasBoardMatches = false;
      Object.entries(item.boards).forEach(([boardId, board]) => {
        if (board.is_active && board.name.toLowerCase().includes(query)) {
          matchingBoards[boardId] = board; hasBoardMatches = true;
        }
      });
      if (mainBoardMatches) {
        Object.entries(item.boards).forEach(([boardId, board]) => { if (board.is_active) matchingBoards[boardId] = board; });
      }
      if (mainBoardMatches || hasBoardMatches) {
        matchingMainBoards.add(item.main_board_id);
        return { ...item, boards: mainBoardMatches ? Object.fromEntries(Object.entries(item.boards).filter(([, b]) => b.is_active)) : matchingBoards };
      }
      return null;
    }).filter(Boolean) as MainBoard[];
  }, [navItems, searchQuery]);

  const filteredAdminItems = useMemo(() => {
    if (!searchQuery.trim()) return adminNavigationItems;
    return adminNavigationItems.filter(item => item.label.toLowerCase().includes(searchQuery.toLowerCase()));
  }, [searchQuery]);

  useEffect(() => {
    if (searchQuery.trim() && filteredNavItems.length > 0) setActiveMainBoard(filteredNavItems[0].main_board_id);
  }, [filteredNavItems, searchQuery]);

  useEffect(() => { if (isSearchOpen && searchInputRef.current) searchInputRef.current.focus(); }, [isSearchOpen]);

  const clearSearch = () => { setSearchQuery(''); if (searchInputRef.current) searchInputRef.current.focus(); };
  const handleSearchChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => setSearchQuery(e.target.value), []);

  // ─── Load user data ───────────────────────────────────────────────────────────
  useEffect(() => {
    if (!isMounted || typeof window === 'undefined') return;
    try {
      const sessionData = localStorage.getItem('currentUserData');
      if (sessionData) {
        const p = JSON.parse(sessionData);
        setUserData({ email: p.email || "", userId: p.userId || "", userRole: p.userRole || "", userName: p.userName || "" });
        return;
      }
      const ld = {
        email: localStorage.getItem('loggedInUserEmail') || "",
        userId: localStorage.getItem('loggedInUserId') || "",
        userRole: localStorage.getItem('loggedInUserRole') || "",
        userName: localStorage.getItem('loggedInUserName') || "",
      };
      if (ld.userId) setUserData(ld);
    } catch { /* ignore */ }
  }, [isMounted]);

  useEffect(() => {
    const h = (event: { target: any }) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) setShowUserDropdown(false);
    };
    document.addEventListener('mousedown', h);
    return () => document.removeEventListener('mousedown', h);
  }, []);

  // ─── Logo (display only) ──────────────────────────────────────────────────────
  // ── Read userId reliably from storage (never from React state) ───────────────
  const getStoredUserId = (): string => {
    if (typeof window === 'undefined') return '';
    try {
      const s = localStorage.getItem('currentUserData');
      if (s) {
        const d = JSON.parse(s);
        const id = d.userId || d.user_id || d.id || '';
        if (id) return String(id);
      }
      return sessionStorage.getItem('loggedInUserId') || localStorage.getItem('loggedInUserId') || '';
    } catch { return ''; }
  };

  // ── Fetch blob from /api/logo/{userId}/view ───────────────────────────────────
  const fetchLogoBlob = async (userId: string): Promise<string | null> => {
    try {
      const response = await fetch(`${API_BASE_URL}/api/logo/${userId}/view`, {
        method: 'GET',
        headers: { 'X-API-Key': EXCEL_API_KEY },
      });
      if (response.ok) {
        const blob = await response.blob();
        if (blob.size > 0 && blob.type.startsWith('image/')) {
          return URL.createObjectURL(blob);
        }
      }
      return null;
    } catch { return null; }
  };

  // ── Fetch logo metadata then image from server ────────────────────────────────
  // Rule: Never call setCurrentLogo(null) unless the server CONFIRMS no logo exists.
  const fetchCurrentLogo = async () => {
    const userId = getStoredUserId();
    if (!userId) return; // userId not ready — keep showing cached preview

    try {
      // First, call the metadata endpoint to verify existence
      const metadataRes = await fetch(`${API_BASE_URL}/api/logo/${userId}`, {
        method: 'GET',
        headers: { 'Accept': 'application/json', 'X-API-Key': EXCEL_API_KEY },
      });

      // If the endpoint returns 404 or any error, the user has no logo
      if (!metadataRes.ok) {
        setCurrentLogo(null);
        return;
      }

      const metadata = await metadataRes.json();
      // Check if the response indicates success and contains a logo object
      const logoExists = metadata?.success === true && metadata?.logo != null;

      if (!logoExists) {
        setCurrentLogo(null);
        return;
      }

      // Logo exists — fetch the actual image blob
      const blobUrl = await fetchLogoBlob(userId);
      if (blobUrl) {
        setCurrentLogo(blobUrl); // replace preview with real server image
      }
      // If blob fetch failed (network hiccup), keep cached preview — don't blank
    } catch {
      // Any error → keep showing cached preview, never blank
    }
  };

  // ── On mount: show cached preview for this user instantly ────────────────────
  const loadStoredLogo = () => {
    const userId = getStoredUserId();
    if (!userId) return; // no userId — server fetch will handle it via useEffect

    try {
      const cached = localStorage.getItem(`logo_cache_${userId}`);
      if (cached) {
        const d = JSON.parse(cached);
        if (d.userId === userId && d.localUrl) {
          setCurrentLogo(d.localUrl); // instant display, no network needed
        }
      }
      // Note: DO NOT call fetchCurrentLogo() here.
      // useEffect([userData.userId]) will do the server refresh once state loads.
    } catch { /* ignore corrupt cache */ }
  };

  // ─── Sidebar resize ───────────────────────────────────────────────────────────
  const toggleSidebar = () => {
    const newWidth = sidebarWidth > 100 ? 64 : 260;
    setSidebarWidth(newWidth);
    setIsSidebarOpen(newWidth > 100);
  };

  const startResizing = (e: { preventDefault: () => void }) => {
    e.preventDefault(); setIsResizing(true);
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', stopResizing);
  };
  const handleMouseMove = (e: { clientX: number }) => {
    if (isResizing && e.clientX >= 64 && e.clientX <= 360) {
      setSidebarWidth(e.clientX); setIsSidebarOpen(e.clientX > 100);
    }
  };
  const stopResizing = () => {
    setIsResizing(false);
    document.removeEventListener('mousemove', handleMouseMove);
    document.removeEventListener('mouseup', stopResizing);
  };
  useEffect(() => () => {
    document.removeEventListener('mousemove', handleMouseMove);
    document.removeEventListener('mouseup', stopResizing);
  }, [isResizing]);

  // ─── Init ─────────────────────────────────────────────────────────────────────
  useEffect(() => {
    if (typeof window === 'undefined') return;
    try {
      const s = localStorage.getItem('currentUserData');
      if (s) {
        const d = JSON.parse(s);
        setUserRole(d.userRole);
        setClientUserId(d.userId);
      } else {
        const id = localStorage.getItem('loggedInUserId');
        const role = localStorage.getItem('loggedInUserRole');
        if (id) setClientUserId(id);
        if (role) setUserRole(role);
      }
    } catch { /* ignore */ }
    loadStoredLogo();
  }, []);

  // ── Once userId is confirmed in state, fetch real blob from server ────────────
  useEffect(() => {
    if (userData.userId) fetchCurrentLogo();
  }, [userData.userId]);

  // ── Logo is uploaded from Master Data now — refetch when it tells us it changed ──
  useEffect(() => {
    const onLogoUpdated = () => fetchCurrentLogo();
    window.addEventListener('logo-updated', onLogoUpdated);
    return () => window.removeEventListener('logo-updated', onLogoUpdated);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ─── Create Main Board ────────────────────────────────────────────────────────
  const handleSave = async () => {
    if (!mainBoardName.trim()) { toast.error("Please enter a name for the main board."); return; }
    let currentUserData: { userId?: string } = {};
    if (typeof window !== 'undefined') {
      const s = localStorage.getItem('currentUserData');
      currentUserData = s ? JSON.parse(s) : {};
    }
    const userId = currentUserData.userId;
    if (!userId) { toast.error("User not found. Please log in again."); return; }

    showGlobalLoader("Creating Main Board...");
    try {
      const response = await fetch(`${API_BASE_URL}/main-boards/?user_id=${userId}`, {
        method: "POST",
        headers: { accept: "application/json", "Content-Type": "application/json", "X-API-Key": EXCEL_API_KEY },
        body: JSON.stringify({ user_id: parseInt(userId), main_board_type: "ANALYSIS", name: mainBoardName }),
      });
      const data = await response.json();
      if (!response.ok) { toast.error(`Failed to save: ${JSON.stringify(data)}`); return; }
      toast.success("Main board saved successfully!");
      setMainBoardName(""); setMainBoardId(data.id); setIsModalOpen(false);
      router.push("/Container");
      mutateNavItems();
    } catch { toast.error("An error occurred. Please try again."); }
    finally { hideGlobalLoader(); }
  };

  // ─── Create / Update Board ────────────────────────────────────────────────────
  const handleCreateBoard = async () => {
    if (!newBoardName.trim()) { toast.error('Please enter a board name'); return; }
    if (!customerDbKey.trim()) { toast.error('Please enter a customer database key'); return; }
    if (!isEditMode && !selectedBoard?.mainBoardId) { toast.error('Main board ID is missing'); return; }
    if (isEditMode && !editingBoardId) { toast.error('Board ID is missing for editing'); return; }

    let userId: string | null = null;
    if (typeof window !== 'undefined') {
      const s = localStorage.getItem('currentUserData');
      if (s) { const d = JSON.parse(s); userId = d.userId || d.user_id || d.id; }
      if (!userId) userId = sessionStorage.getItem("loggedInUserId") || localStorage.getItem('loggedInUserId');
    }
    if (!userId) { toast.error("User ID not found. Please log in again."); return; }

    showGlobalLoader(isEditMode ? "Updating Board..." : "Creating Board...");
    try {
      if (isEditMode) {
        const response = await fetch(`${API_BASE_URL}/main-boards/boards/${editingBoardId}?user_id=${userId}`, {
          method: "PUT",
          headers: { "Content-Type": "application/json", "X-API-Key": EXCEL_API_KEY },
          body: JSON.stringify({ main_board_id: parseInt(selectedBoard!.mainBoardId), name: newBoardName.trim(), customer_db_key: customerDbKey.trim() }),
        });
        if (!response.ok) {
          const errorBody = await response.text();
          let msg = `Failed to update board: ${response.status}`;
          try { const e = JSON.parse(errorBody); msg = e.detail || e.message || msg; } catch { /* ignore */ }
          toast.error(msg); return;
        }
        if (typeof window !== 'undefined') localStorage.setItem(`csvEnabled_${editingBoardId}`, String(csvUploadEnabled));
        toast.success("Board updated successfully!");
        closeModal();
        router.push(`/Container?main_board_id=${selectedBoard?.mainBoardId}&board_id=${editingBoardId}`);
        mutateNavItems();
      } else {
        const response = await fetch(`${API_BASE_URL}/main-boards/boards/?user_id=${userId}`, {
          method: "POST",
          headers: { "Content-Type": "application/json", "X-API-Key": EXCEL_API_KEY },
          body: JSON.stringify({ main_board_id: parseInt(selectedBoard!.mainBoardId), name: newBoardName.trim(), customer_db_key: customerDbKey.trim() }),
        });
        if (!response.ok) {
          const errorBody = await response.text();
          let msg = `Failed to create board: ${response.status}`;
          try { const e = JSON.parse(errorBody); msg = e.detail || e.message || msg; } catch { /* ignore */ }
          toast.error(msg); return;
        }
        const newBoard = await response.json();
        if (typeof window !== 'undefined') localStorage.setItem(`csvEnabled_${newBoard.id}`, String(csvUploadEnabled));
        toast.success("Board created successfully!");
        closeModal();
        router.push(`/Container?main_board_id=${selectedBoard!.mainBoardId}&board_id=${newBoard.id}`);
        mutateNavItems();
      }
    } catch { toast.error("An unexpected error occurred"); }
    finally { hideGlobalLoader(); }
  };

  // ─── Delete Main Board ────────────────────────────────────────────────────────
  const handleDeleteMainBoard = async (e: React.MouseEvent, mainBoardId: string, mainBoardName: string) => {
    e.stopPropagation();
    toast.info(
      <div>
        <p>Are you sure you want to delete <strong>{mainBoardName}</strong>?<br /> This action cannot be undone.</p>
        <div className="toast-actions">
          <button onClick={() => { deleteMainBoard(mainBoardId); toast.dismiss(); }} className="confirm-btn">Confirm</button>
          <button onClick={() => toast.dismiss()} className="cancel-btn">Cancel</button>
        </div>
      </div>,
      { autoClose: false, closeButton: true, closeOnClick: false, draggable: false }
    );
  };

  const deleteMainBoard = async (mainBoardId: string) => {
    showGlobalLoader("Deleting Main Board...");
    try {
      let userId = '';
      if (typeof window !== 'undefined') {
        const s = localStorage.getItem('currentUserData');
        if (s) { const d = JSON.parse(s); userId = d.userId || d.user_id || d.id || ''; }
        if (!userId) userId = sessionStorage.getItem("loggedInUserId") || localStorage.getItem('loggedInUserId') || '';
      }
      if (!userId) { toast.error("User not found. Please log in again."); return; }

      const response = await fetch(`${API_BASE_URL}/main-boards/${mainBoardId}?user_id=${userId}`, {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json', 'Accept': 'application/json', 'X-API-Key': EXCEL_API_KEY },
      });
      if (response.ok) {
        toast.success("Main board deleted successfully");
        mutateNavItems();
        if (activeMainBoard === mainBoardId) setActiveMainBoard('');
      } else {
        const errorData = await response.json();
        toast.error(errorData.message || "Failed to delete main board");
      }
    } catch { toast.error("An error occurred while deleting the main board"); }
    finally { hideGlobalLoader(); }
  };

  // ─── Delete Board ─────────────────────────────────────────────────────────────
  const handleDelete = async (boardId: string, mainBoardId: string, boardName: string) => {
    const ConfirmToast = ({ closeToast }: { closeToast: () => void }) => (
      <div className="p-3 bg-white rounded-lg shadow-lg">
        <p className="text-gray-800 text-sm mb-3">Are you sure you want to delete <strong>{boardName}</strong>?</p>
        <div className="flex justify-end space-x-2">
          <button onClick={() => closeToast()} className="px-3 py-1.5 text-sm bg-gray-200 text-gray-800 hover:bg-gray-300 rounded">Cancel</button>
          <button
            onClick={async () => {
              closeToast();
              showGlobalLoader("Deleting Board...");
              try {
                let currentUserData: { userId?: string } = {};
                if (typeof window !== 'undefined') {
                  const s = localStorage.getItem('currentUserData');
                  currentUserData = s ? JSON.parse(s) : {};
                }
                const userId = currentUserData.userId;
                if (!userId) { toast.error("User not found. Please log in again."); return; }
                const response = await fetch(`${API_BASE_URL}/main-boards/boards/${boardId}?user_id=${userId}`, {
                  method: 'DELETE',
                  headers: { 'Content-Type': 'application/json', 'Accept': 'application/json', 'X-API-Key': EXCEL_API_KEY },
                });
                if (!response.ok) {
                  const errorData = await response.json();
                  throw new Error(errorData.message || "Failed to delete board");
                }
                setNavItems(prev => prev.map(item =>
                  item.main_board_id === mainBoardId
                    ? { ...item, boards: Object.fromEntries(Object.entries(item.boards).filter(([key]) => key !== boardId)) }
                    : item
                ));
                toast.success("Board deleted successfully!");
              } catch (error) {
                toast.error(error instanceof Error ? error.message : "An error occurred while deleting the board.");
              } finally { hideGlobalLoader(); }
            }}
            className="px-3 py-1.5 text-sm bg-red-500 text-white hover:bg-red-600 rounded"
          >Delete</button>
        </div>
      </div>
    );
    toast(<ConfirmToast closeToast={() => { }} />, {
      position: 'bottom-center', autoClose: false, closeButton: false, closeOnClick: false, draggable: false, className: '!bg-transparent !shadow-none',
    });
  };

  // ─── Plus click (create board) ────────────────────────────────────────────────
  const handlePlusClick = (event: React.MouseEvent<SVGSVGElement, MouseEvent>, mainBoardId: string) => {
    event.stopPropagation();
    setSelectedBoard({ mainBoardId });
    setNewBoardName(''); setCustomerDbKey('');
    setIsEditMode(false); setEditingBoardId(null);
    setCsvUploadEnabled(true);
    setShowModal(true);
     fetchCustomerDbKeys();
  };

  // ─── Edit board click ────────────────────────────────────────────────────────
  const handleEditClick = async (boardId: string, mainBoardId: string) => {
    if (!Array.isArray(navItems)) return;
    const mainBoard = navItems.find(item => item.main_board_id === mainBoardId);
    if (!mainBoard) return;
    const boardData = mainBoard.boards[boardId];
    if (!boardData) return;

    showGlobalLoader("Loading Board Details...");
    try {
      let userId = '';
      if (typeof window !== 'undefined') {
        const s = localStorage.getItem('currentUserData');
        if (s) { const d = JSON.parse(s); userId = d.userId || d.user_id || d.id; }
        if (!userId) userId = sessionStorage.getItem("loggedInUserId") || localStorage.getItem('loggedInUserId') || '';
      }
      if (!userId) { toast.error('User ID not found. Please log in again.'); return; }

      const response = await fetch(`${API_BASE_URL}/main-boards/boards/${boardId}?user_id=${userId}`, {
        method: 'GET', headers: { 'Accept': 'application/json', 'X-API-Key': EXCEL_API_KEY },
      });

      if (response.ok) {
        const boardDetails = await response.json();
        setIsEditMode(true);
        setEditingBoardId(boardId);
        setSelectedBoard({ mainBoardId, boardId, boardName: boardData.name });
        setNewBoardName(boardData.name);
        setCustomerDbKey(boardDetails.customer_db_key || '');
        const storedCsvPref = typeof window !== 'undefined' ? localStorage.getItem(`csvEnabled_${boardId}`) : null;
        setCsvUploadEnabled(storedCsvPref === null ? true : storedCsvPref === 'true');
      } else {
        toast.error('Failed to load board details'); return;
      }
    } catch {
      toast.error('Error loading board details'); return;
    } finally { hideGlobalLoader(); }

    setShowModal(true);
  };

  // ─── Close modal ──────────────────────────────────────────────────────────────
  const closeModal = () => {
    setShowModal(false); setSelectedBoard(null);
    setNewBoardName(''); setCustomerDbKey('');
    setIsEditMode(false); setEditingBoardId(null);
    setCsvUploadEnabled(true);
  };

  // ─── Nav fetching ─────────────────────────────────────────────────────────────
  // Plain fetch-in-effect (same proven pattern the CXO page uses for its own
  // mainboard list) rather than SWR — SWR's key here was derived from
  // clientUserId, and on a client-side route change (e.g. clicking "Consultant
  // Panel" from CXO, which mounts this sidebar fresh without a full page
  // reload) the fetch wasn't reliably firing, leaving the list empty until a
  // manual refresh. A direct effect tied straight to clientUserId fetches
  // every time this component has a real user id, mount or not.
  const mutateNavItems = useCallback(async () => {
    if (!clientUserId) return;
    try {
      const res = await fetch(`${API_BASE_URL}/main-boards/get_all_info_tree?user_id=${clientUserId}`, {
        headers: { Accept: "application/json", "X-API-Key": EXCEL_API_KEY },
      });
      if (!res.ok) throw new Error("Failed to fetch");
      const data = await res.json();
      setNavItems(data);
    } catch {
      toast.error("Error loading navigation data");
    }
  }, [clientUserId, API_BASE_URL, EXCEL_API_KEY]);

  useEffect(() => { mutateNavItems(); }, [mutateNavItems]);
  useEffect(() => { if (refreshTrigger) mutateNavItems(); }, [refreshTrigger]);
  const forceRefresh = () => mutateNavItems();

  // Auto-translate all board names when language or navItems changes
  useEffect(() => {
    if (language === 'en' || navItems.length === 0) {
      setBoardNameMap({});
      return;
    }
    // Collect all unique names: main boards + sub-boards
    const allNames: string[] = [];
    navItems.forEach(item => {
      allNames.push(item.name);
      Object.values(item.boards).forEach(board => {
        if (board.is_active) allNames.push(board.name);
      });
    });
    const unique = [...new Set(allNames)];
    translateBatch(unique, language).then(translated => {
      const map: Record<string, string> = {};
      unique.forEach((name, i) => { map[name] = translated[i] || name; });
      setBoardNameMap(map);
    });
  }, [language, navItems]);

  const toggleMainBoard = (mainBoardId: string) => {
    setActiveMainBoard(prev => prev === mainBoardId ? null : mainBoardId);
    setShowSubMenu(false);
  };

 const handleLogout = () => {
  closeMobileMenu();
  localStorage.removeItem('currentUserData'); // clear auth data
  router.replace('/');                          // replace history, not push
};
  const handleBoardClick = (boardId: string) => { setActiveBoardId(boardId); closeMobileMenu(); };

  // ─── Highlight search match ───────────────────────────────────────────────────
  const highlight = (text: string, query: string) =>
    query ? text.replace(new RegExp(`(${query})`, 'gi'), '<mark class="bg-yellow-200 px-1 rounded">$1</mark>') : text;

  // ═══════════════════════════════════════════════════════════════════════════════
  // RENDER
  // ═══════════════════════════════════════════════════════════════════════════════
  return (
    <>
      {globalLoading && <GlobalLoader message={globalLoadingMessage} />}

      {isMobile && (
        <button onClick={toggleMobileMenu} className="fixed top-3 left-3 z-50 p-2 bg-blue-600 text-white rounded-lg shadow-lg hover:bg-blue-700 transition-colors" aria-label="Toggle menu">
          {isMobileMenuOpen ? <X className="w-5 h-5" /> : <Menu className="w-5 h-5" />}
        </button>
      )}

      {isMobile && isMobileMenuOpen && <div className="fixed inset-0 bg-black/50 z-40" onClick={closeMobileMenu} />}

      <div
        ref={sidebarRef}
        className="h-screen bg-white text-black flex flex-col shadow-2xl"
        style={{
          position: isMobile ? 'fixed' : 'relative',
          left: isMobile ? (isMobileMenuOpen ? 0 : '-100%') : 0,
          top: isMobile ? 0 : 'auto',
          width: isMobile ? '80%' : `${sidebarWidth}px`,
          maxWidth: isMobile ? '280px' : 'none',
          zIndex: isMobile ? 45 : 'auto',
          transition: isMobile ? 'left 0.3s ease-in-out' : 'width 0.3s',
        }}
      >
        <ToastContainer position="bottom-center" autoClose={3000} hideProgressBar={false} newestOnTop closeOnClick rtl={false} pauseOnFocusLoss draggable={false} pauseOnHover className="z-50" />

        {/* ── Header ────────────────────────────────────────────────────────── */}
        <div className="relative bg-white border-b border-blue-500/50 shadow-md">
          <div className="px-3 py-3 flex justify-between items-center">
            {(isSidebarOpen || isMobile) && currentLogo && (
              <div className="relative group flex-1 min-w-0">
                <img
                  src={currentLogo}
                  alt="Logo"
                  width={200}
                  height={68}
                  className="object-contain max-h-[68px] rounded-xl"
                  onError={() => setCurrentLogo(null)}
                />
              </div>
            )}
            {!isMobile && (
              <button onClick={toggleSidebar} className="p-2 text-black bg-white shadow-md border border-gray-200 hover:bg-blue-700/50 hover:text-white focus:outline-none rounded-lg transition-all duration-200 hover:scale-105 flex-shrink-0">
                {isSidebarOpen ? <ChevronLeft className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
              </button>
            )}
          </div>
        </div>

        {/* ── Create Main Board button ──────────────────────────────────────── */}
        {(isSidebarOpen || isMobile) && (
          <div className="px-3 py-2">
            <button
              onClick={() => { setIsModalOpen(true); closeMobileMenu(); }}
              className="w-full bg-gradient-to-r from-blue-600 to-blue-700 hover:from-blue-700 hover:to-blue-800 text-white py-2.5 px-3 rounded-2xl font-bold text-sm transition-all duration-200 shadow-md hover:shadow-lg flex items-center justify-center gap-2 group"
            >
              <Plus className="w-3.5 h-3.5 group-hover:rotate-90 transition-transform duration-200" />
              {t('sidebar.createMainBoard')}
            </button>
          </div>
        )}

        {/* ── Search bar ────────────────────────────────────────────────────── */}
        {(isSidebarOpen || isMobile) ? (
          <div className="px-3 pb-2">
            <div className="relative">
              <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 text-gray-400 w-3.5 h-3.5" />
              <input
                ref={searchInputRef} type="text" placeholder={t('sidebar.search')}
                value={searchQuery} onChange={handleSearchChange}
                suppressHydrationWarning
                className="w-full py-2 pl-8 pr-8 text-sm border border-gray-200 rounded-2xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-black bg-white shadow-sm transition-all duration-200"
              />
              {searchQuery && (
                <button onClick={clearSearch} className="absolute right-2.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 hover:bg-gray-100 rounded-full p-0.5 transition-all duration-200">
                  <X className="w-3.5 h-3.5" />
                </button>
              )}
            </div>
          </div>
        ) : (
          <div className="flex justify-center pb-2">
            <button
              onClick={toggleSidebar}
              title="Search (expand sidebar)"
              className="p-2 text-gray-500 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-all duration-200"
            >
              <Search className="w-4 h-4" />
            </button>
          </div>
        )}

        {/* ── Navigation ────────────────────────────────────────────────────── */}
        <div className="flex-1 overflow-y-auto scrollbar-hide">
          <div className="px-3 pb-3">
            {searchQuery && (
              <div className="mb-3 p-2 bg-blue-50 rounded-lg border border-blue-200">
                <div className="flex justify-between items-center">
                  <p className="text-xs text-blue-800">
                    <span className="font-medium">{filteredNavItems.length + filteredAdminItems.length}</span> {t('sidebar.resultsFor')} "{searchQuery}"
                  </p>
                  <button onClick={clearSearch} className="text-blue-600 hover:text-blue-800 text-xs font-medium hover:bg-blue-100 px-1.5 py-0.5 rounded">{t('actions.clearSearch')}</button>
                </div>
              </div>
            )}

            {/* Dashboard link */}
            {/* <div className="space-y-0.5 mb-4">
              {(!searchQuery.trim() || "dashboard".includes(searchQuery.toLowerCase())) && (
                <Link href="/Dashboard" onClick={closeMobileMenu}
                  className="flex items-center p-2 rounded-lg cursor-pointer transition-all duration-200 group hover:bg-blue-700/40 hover:shadow-sm"
                  onMouseEnter={() => setHoveredItem('dashboard')} onMouseLeave={() => setHoveredItem(null)}
                >
                  <LayoutDashboard className="w-4 h-4 flex-shrink-0 group-hover:scale-110 transition-transform duration-200" />
                  {(isSidebarOpen || isMobile) && (
                    <span className="ml-2 font-medium text-xs">
                      {searchQuery ? <span dangerouslySetInnerHTML={{ __html: highlight("Dashboard", searchQuery) }} /> : "Dashboard"}
                    </span>
                  )}
                </Link>
              )}
            </div> */}

            {/* Demo Reference dropdown */}
            {(!searchQuery.trim() || "demo reference".includes(searchQuery.toLowerCase())) && (
              <div className="space-y-0.5 mb-4">
                <div
                  className={`flex items-center justify-between p-2 rounded-lg cursor-pointer transition-all duration-200 group ${isDemoRefOpen ? 'bg-blue-700/60 shadow-md border border-blue-500/30' : 'hover:bg-blue-700/30'}`}
                  onClick={toggleDemoRef}
                  onMouseEnter={() => setHoveredItem('demo-ref')}
                  onMouseLeave={() => setHoveredItem(null)}
                >
                  <div className="flex items-center min-w-0 flex-1">
                    {isDemoRefOpen
                      ? <ChevronDown className="w-3.5 h-3.5 mr-1.5 text-gray-700 flex-shrink-0" />
                      : <ChevronRight className="w-3.5 h-3.5 mr-1.5 text-gray-700 flex-shrink-0" />
                    }
                    {(isSidebarOpen || isMobile) && (
                      <span className="font-medium text-xs truncate">{t('sidebar.demoReference')}</span>
                    )}
                  </div>
                </div>

                {isDemoRefOpen && (isSidebarOpen || isMobile) && (
                  <div className="ml-5 space-y-0.5 pb-1">
                    {isDemoLoading ? (
                      <div className="text-xs text-gray-400 px-2 py-1">{t('sidebar.loading')}</div>
                    ) : !Array.isArray(demoMainBoards) || demoMainBoards.length === 0 ? (
                      <div className="text-xs text-gray-400 px-2 py-1">{t('sidebar.noDemoBoards')}</div>
                    ) : (
                      demoMainBoards.map(mb => {
                        const mbId = String(mb.id);
                        const isActiveDemoMb = activeDemoMainBoard === mbId;
                        const boardsForMb = Array.isArray(demoBoards)
                          ? demoBoards.filter(b => b.main_board_id === mb.id)
                          : [];
                        return (
                          <div key={mbId} className="space-y-0.5">
                            <div
                              className={`flex items-center justify-between p-2 rounded-md cursor-pointer transition-all duration-200 group ${isActiveDemoMb ? 'bg-blue-600/50 shadow-sm border border-blue-400/30' : 'hover:bg-blue-700/25'}`}
                              onClick={() => setActiveDemoMainBoard(prev => prev === mbId ? null : mbId)}
                            >
                              <div className="flex items-center min-w-0 flex-1">
                                {isActiveDemoMb
                                  ? <ChevronDown className="w-3 h-3 mr-1.5 text-gray-600 flex-shrink-0" />
                                  : <ChevronRight className="w-3 h-3 mr-1.5 text-gray-600 flex-shrink-0" />
                                }
                                <span className="text-xs font-medium truncate">{mb.name}</span>
                              </div>
                              <button
                                onClick={e => openDemoCreateModal(e, mb.id)}
                                className="p-1 hover:bg-blue-600 rounded transition-colors duration-200 opacity-0 group-hover:opacity-100 flex-shrink-0"
                                title={`Add board to ${mb.name}`}
                              >
                                <Plus className="w-3.5 h-3.5" />
                              </button>
                            </div>

                            {isActiveDemoMb && (
                              <div className="ml-4 space-y-0.5">
                                {boardsForMb.length === 0 ? (
                                  <div className="text-xs text-gray-400 px-2 py-1">No boards</div>
                                ) : (
                                  boardsForMb.map(board => (
                                    <div
                                      key={board.id}
                                      className="flex items-center justify-between p-2 rounded-md hover:bg-blue-700/20 cursor-pointer group transition-all duration-200"
                                    >
                                      <Link
                                        href={{ pathname: '/DemoContainer', query: { board_id: board.id, main_board_id: mb.id } }}
                                        onClick={closeMobileMenu}
                                        className="flex-1 text-black text-xs font-medium truncate"
                                      >
                                        {board.name}
                                      </Link>
                                      <div className="flex space-x-0.5 opacity-0 group-hover:opacity-100 transition-opacity duration-200">
                                        <button
                                          onClick={e => openDemoEditModal(e, board)}
                                          className="p-1 hover:bg-blue-600 rounded transition-colors duration-200"
                                          title="Edit board"
                                        >
                                          <Edit2 className="w-3 h-3" />
                                        </button>
                                        <button
                                          onClick={e => handleDeleteDemoBoard(e, board)}
                                          className="p-1 hover:bg-red-600 rounded transition-colors duration-200"
                                          title="Delete board"
                                        >
                                          <Trash2 className="w-3 h-3" />
                                        </button>
                                      </div>
                                    </div>
                                  ))
                                )}
                              </div>
                            )}
                          </div>
                        );
                      })
                    )}
                  </div>
                )}
              </div>
            )}

            {/* Main boards */}
            <div className="space-y-1 mb-4">
              {filteredNavItems.map(item => {
                const mbId = String(item.main_board_id);
                const isExpanded = searchQuery.trim() ? true : activeMainBoard === mbId;
                return (
                  <div key={mbId} className="space-y-0.5">
                    <div
                      className={`flex items-center justify-between p-2 rounded-lg cursor-pointer transition-all duration-200 group ${isExpanded ? 'bg-blue-600 shadow-md' : 'hover:bg-blue-100'}`}
                      onClick={() => toggleMainBoard(mbId)}
                      onMouseEnter={() => setHoveredItem(item.main_board_id)}
                      onMouseLeave={() => setHoveredItem(null)}
                    >
                      <div className="flex items-center min-w-0 flex-1 group">
                        {!(item.boards && Object.keys(item.boards).length > 0 && isExpanded) && (
                          <ChartColumnDecreasing className={`w-3.5 h-3.5 mr-1.5 flex-shrink-0 group-hover:hidden ${isExpanded ? 'text-white' : 'text-gray-700'}`} />
                        )}
                        {item.boards && Object.keys(item.boards).length > 0 && (
                          <div className={`flex-shrink-0 ${isExpanded ? "block" : "hidden group-hover:block"}`}>
                            {isExpanded ? <ChevronDown className="w-3.5 h-3.5 mr-1.5 text-white" /> : <ChevronRight className="w-3.5 h-3.5 mr-1.5 text-gray-700" />}
                          </div>
                        )}
                        {(isSidebarOpen || isMobile) && (
                          <span className={`text-xs truncate ${isExpanded ? 'font-bold text-white' : 'font-medium text-gray-700 group-hover:text-black'}`}>
                            {searchQuery
                              ? <span dangerouslySetInnerHTML={{ __html: highlight(item.name, searchQuery) }} />
                              : (boardNameMap[item.name] || item.name)}
                          </span>
                        )}
                      </div>

                      {(isSidebarOpen || isMobile) && !searchQuery.trim() && (
                        <div className="flex space-x-0.5 opacity-0 group-hover:opacity-100 transition-opacity duration-200">
                          <Plus className="p-1 hover:bg-blue-600 rounded transition-colors duration-200 w-5 h-5" onClick={e => handlePlusClick(e, mbId)} />
                          <button onClick={e => handleDeleteMainBoard(e, mbId, item.name)} className="p-1 hover:bg-red-600 rounded transition-colors duration-200" title="Delete Main Board">
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      )}
                    </div>

                    {isExpanded && (isSidebarOpen || isMobile) && (
                      <div className="ml-5 space-y-0.5 pb-1">
                        {Object.keys(item.boards).filter(bId => item.boards[bId].is_active).map(boardId => {
                          const board = item.boards[boardId];
                          return (
                            <div
                              key={boardId}
                              className={`flex items-center justify-between p-2 rounded-md cursor-pointer transition-all duration-200 group ${activeBoardId === boardId ? 'bg-blue-200 shadow-sm' : 'hover:bg-blue-100'}`}
                              onClick={() => handleBoardClick(boardId)}
                            >
                              <Link
                                href={{ pathname: '/Container', query: { main_board_id: item.main_board_id, board_id: boardId } }}
                                onClick={closeMobileMenu}
                                className={`flex-1 text-xs truncate ${activeBoardId === boardId ? 'font-bold text-gray-900' : 'font-medium text-gray-700'}`}
                              >
                                {searchQuery
                                  ? <span dangerouslySetInnerHTML={{ __html: highlight(board.name, searchQuery) }} />
                                  : (boardNameMap[board.name] || board.name)}
                              </Link>
                              {!searchQuery.trim() && (
                                <div className="flex space-x-0.5 opacity-0 group-hover:opacity-100 transition-opacity duration-200">
                                  <button
                                    onClick={e => { e.stopPropagation(); handleEditClick(boardId, item.main_board_id); }}
                                    className="p-1 hover:bg-blue-600 rounded transition-colors duration-200" title="Edit Board"
                                  >
                                    <Edit2 className="w-3 h-3" />
                                  </button>
                                  <button
                                    onClick={e => { e.stopPropagation(); handleDelete(boardId, item.main_board_id, board.name); }}
                                    className="p-1 hover:bg-red-600 rounded transition-colors duration-200" title="Delete Board"
                                  >
                                    <Trash2 className="w-3 h-3" />
                                  </button>
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>

            {/* Admin items */}
            {userRole?.toLowerCase() === "admin" && (
              <div className="space-y-0.5">
                {(searchQuery.trim() ? filteredAdminItems : adminNavigationItems).map(item => (
                  <Link key={item.id} href={item.href} onClick={closeMobileMenu}
                    className={`flex items-center p-2 rounded-lg cursor-pointer transition-all duration-200 group ${pathname.startsWith(item.href) ? 'bg-blue-600 text-white shadow-md' : 'hover:bg-blue-100 hover:shadow-sm'}`}
                    onMouseEnter={() => setHoveredItem(item.id)} onMouseLeave={() => setHoveredItem(null)}
                  >
                    {item.id === 'users' && <User className="w-4 h-4 flex-shrink-0" />}
                    {(item.id !== 'users') && <NotebookText className="w-4 h-4 flex-shrink-0" />}
                    {(isSidebarOpen || isMobile) && (
                      <span className="ml-2 font-medium text-xs">
                        {searchQuery ? <span dangerouslySetInnerHTML={{ __html: highlight(item.label, searchQuery) }} /> : item.label}
                      </span>
                    )}
                  </Link>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* ── User profile / dropdown ──────────────────────────────────────── */}
        {/* Moved into the top header bar on /Container, so it isn't shown twice there. */}
        {pathname !== '/Container' && (
        <div className="border-t border-gray-200 p-2.5 relative" ref={dropdownRef}>
          {(isSidebarOpen || isMobile) ? (
            <div>
              <button onClick={toggleDropdown} className="w-full flex items-center space-x-2 p-2 bg-gray-100 hover:bg-gray-200 rounded-lg transition-colors text-left">
                <div className="w-7 h-7 bg-gray-600 rounded-full flex items-center justify-center text-white font-medium text-xs flex-shrink-0">
                  {userData.userName ? userData.userName.charAt(0).toUpperCase() : <User className="w-3 h-3" />}
                </div>
                <div className="flex-1 min-w-0">
                  <p className="text-xs font-medium text-gray-900 truncate">{userData.userName || "N/A"}</p>
                  <p className="text-xs text-blue-600 truncate">{userData.email || "N/A"}</p>
                </div>
                <div className="flex-shrink-0">
                  {isDropdownOpen ? <ChevronLeft className="w-3.5 h-3.5 text-gray-500" /> : <ChevronRight className="w-3.5 h-3.5 text-gray-500" />}
                </div>
              </button>
              {isDropdownOpen && (
                <div className="absolute left-full bottom-0 ml-2 w-56 bg-white border border-gray-200 rounded-lg shadow-lg z-50 overflow-hidden">
                  <div className="p-2.5 border-b border-gray-100 relative">
                    <button onClick={() => setIsDropdownOpen(false)} className="absolute top-2 right-2 p-0.5 text-gray-400 hover:text-gray-600 hover:bg-gray-100 rounded-full"><X className="w-3.5 h-3.5" /></button>
                    <div className="flex items-center space-x-2 pr-5">
                      <div className="w-7 h-7 bg-gray-600 rounded-full flex items-center justify-center text-white font-medium text-xs flex-shrink-0">
                        {userData.userName ? userData.userName.charAt(0).toUpperCase() : <User className="w-3 h-3" />}
                      </div>
                      <div className="flex-1 min-w-0">
                        <p className="text-xs font-medium text-gray-900 truncate">{userData.userName || "N/A"}</p>
                        <p className="text-xs text-blue-600 truncate">{userData.email || "N/A"}</p>
                      </div>
                    </div>
                  </div>
                  <div>
                    <button onClick={handleSettingsClick} className="w-full flex items-center space-x-2 p-2.5 text-left text-xs text-gray-700 hover:bg-gray-50 transition-colors">
                      <Settings className="w-3.5 h-3.5" /><span>Settings</span>
                    </button>
                    <button onClick={handleLogout} className="w-full flex items-center space-x-2 p-2.5 text-left text-xs text-gray-700 hover:bg-gray-50 transition-colors">
                      <LogOut className="w-3.5 h-3.5" /><span>{t('header.logout')}</span>
                    </button>
                  </div>
                </div>
              )}
            </div>
          ) : (
            <div>
              <button onClick={toggleDropdown} className="w-full flex justify-center p-2 bg-gray-100 hover:bg-gray-200 rounded-lg transition-colors">
                <div className="w-7 h-7 bg-gray-600 rounded-full flex items-center justify-center text-white font-medium text-xs">
                  {userData.userName ? userData.userName.charAt(0).toUpperCase() : <User className="w-3 h-3" />}
                </div>
              </button>
              {isDropdownOpen && (
                <div className="absolute left-full bottom-0 ml-2 w-56 bg-white border border-gray-200 rounded-lg shadow-lg z-50 overflow-hidden">
                  <div className="p-2.5 border-b border-gray-100">
                    <div className="flex items-center space-x-2">
                      <div className="w-7 h-7 bg-gray-600 rounded-full flex items-center justify-center text-white font-medium text-xs">
                        {userData.userName ? userData.userName.charAt(0).toUpperCase() : <User className="w-3 h-3" />}
                      </div>
                      <div className="flex-1 min-w-0">
                        <p className="text-xs font-medium text-gray-900 truncate">{userData.userName || "N/A"}</p>
                        <p className="text-xs text-blue-600 truncate">{userData.email || "N/A"}</p>
                      </div>
                    </div>
                  </div>
                  <div>
                    <button onClick={handleSettingsClick} className="w-full flex items-center space-x-2 p-2.5 text-left text-xs text-gray-700 hover:bg-gray-50 transition-colors">
                      <Settings className="w-3.5 h-3.5" /><span>Settings</span>
                    </button>
                    <button onClick={handleLogout} className="w-full flex items-center space-x-2 p-2.5 text-left text-xs text-gray-700 hover:bg-gray-50 transition-colors">
                      <LogOut className="w-3.5 h-3.5" /><span>{t('header.logout')}</span>
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
        )}

        {/* ── Settings / Change Password Modal ─────────────────────────────── */}
        {isSettingsModalOpen && (
          <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-50">
            <div className="bg-white rounded-xl shadow-2xl max-w-sm w-full mx-4 overflow-hidden">
              <div className="flex justify-between items-center px-5 py-4 border-b border-gray-200">
                <h2 className="text-base font-bold text-gray-900">Change Password</h2>
                <button onClick={() => { setIsSettingsModalOpen(false); setPasswordData({ currentPassword: '', newPassword: '', confirmPassword: '' }); }} className="p-1.5 text-gray-400 hover:text-gray-600 hover:bg-gray-100 rounded-full transition-all duration-200">
                  <X className="w-4 h-4" />
                </button>
              </div>
              <div className="p-5 space-y-3">
                {([
                  { label: 'Current Password', key: 'currentPassword' as const, show: showCurrentPassword, toggle: () => setShowCurrentPassword(!showCurrentPassword) },
                  { label: 'New Password', key: 'newPassword' as const, show: showNewPassword, toggle: () => setShowNewPassword(!showNewPassword) },
                  { label: 'Confirm New Password', key: 'confirmPassword' as const, show: showConfirmPassword, toggle: () => setShowConfirmPassword(!showConfirmPassword) },
                ]).map(({ label, key, show, toggle }) => (
                  <div key={key}>
                    <label className="block text-xs font-semibold text-gray-700 mb-1.5">{label}</label>
                    <div className="relative">
                      <input type={show ? "text" : "password"} value={passwordData[key]} onChange={e => setPasswordData({ ...passwordData, [key]: e.target.value })}
                        className="w-full px-3 py-2 text-sm border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-gray-900 pr-9" placeholder={`Enter ${label.toLowerCase()}`} />
                      <button type="button" onClick={toggle} className="absolute right-2.5 top-1/2 -translate-y-1/2 text-gray-500 hover:text-gray-700">
                        {show ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                      </button>
                    </div>
                  </div>
                ))}
              </div>
              <div className="flex justify-end gap-2 px-5 py-3 border-t bg-gray-50">
                <button onClick={() => { setIsSettingsModalOpen(false); setPasswordData({ currentPassword: '', newPassword: '', confirmPassword: '' }); }} disabled={isUpdatingPassword} className="px-4 py-2 text-xs font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 disabled:opacity-50">Cancel</button>
                <button onClick={handlePasswordUpdate} disabled={isUpdatingPassword} className="px-4 py-2 text-xs font-medium text-white bg-blue-600 rounded-lg hover:bg-blue-700 disabled:opacity-50 flex items-center gap-1.5">
                  {isUpdatingPassword ? (<><div className="animate-spin rounded-full h-3 w-3 border-b-2 border-white" />Updating...</>) : 'Update Password'}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ── Create Main Board Modal ───────────────────────────────────────── */}
        {isModalOpen && (
          <div className="fixed inset-0 flex items-center justify-center bg-black/60 backdrop-blur-sm z-50">
            <div className="bg-white rounded-xl shadow-2xl p-5 w-80 mx-4">
              <h2 className="text-base font-bold mb-4 text-gray-900">{t('modal.createMainBoard')}</h2>
              <input
                style={{ color: "black" }} type="text" value={mainBoardName}
                onChange={e => setMainBoardName(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && handleSave()}
                placeholder={t('modal.enterMainBoardName')}
                className="w-full p-2.5 text-sm border border-gray-300 rounded-lg mb-4 focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
              <div className="flex justify-end space-x-3">
                <button onClick={() => setIsModalOpen(false)} className="bg-gray-300 hover:bg-gray-400 text-black py-2 px-4 text-sm rounded-lg font-medium transition-all duration-200">{t('modal.cancel')}</button>
                <button onClick={handleSave} className="bg-blue-600 hover:bg-blue-700 text-white py-2 px-4 text-sm rounded-lg font-medium transition-all duration-200">{t('modal.save')}</button>
              </div>
            </div>
          </div>
        )}

        {/* ── Create / Edit Board Modal ─────────────────────────────────────── */}
        {showModal && selectedBoard && (
          <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center" onClick={closeModal}>
            <div className="bg-white rounded-xl shadow-2xl p-5 w-full max-w-sm mx-4 relative" onClick={e => e.stopPropagation()}>
              <div className="mb-4">
                <button className="absolute top-3 right-3 text-gray-500 hover:text-gray-700 hover:bg-gray-100 rounded-full p-1.5 transition-all duration-200" onClick={closeModal}>
                  <X className="w-4 h-4" />
                </button>
                <h2 className="text-base font-bold text-gray-900">{isEditMode ? 'Edit Board' : 'Create New Board'}</h2>
                <p className="text-xs text-gray-500 mt-0.5">
                  {isEditMode ? `Board ID: ${editingBoardId} • Main Board ID: ${selectedBoard.mainBoardId}` : `Main Board ID: ${selectedBoard.mainBoardId}`}
                </p>
              </div>
              <div className="space-y-3">
                <div>
                  <label className="block text-xs font-semibold text-gray-700 mb-1.5">Board Name <span className="text-red-500">*</span></label>
                  <input type="text" value={newBoardName} onChange={e => setNewBoardName(e.target.value)}
                    placeholder="Enter board name"
                    className="w-full px-3 py-2 text-sm border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-gray-900"
                  />
                </div>
               <div>
  <label className="block text-xs font-semibold text-gray-700 mb-1.5">
    Customer Database Key <span className="text-red-500">*</span>
  </label>

  <select
  value={customerDbKey}
  onChange={(e) => setCustomerDbKey(e.target.value)}
  className="w-full px-3 py-2 text-sm border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-gray-900"
>
  <option value="">Select database</option>

  {customerDbOptions.map((db) => (
    <option key={db} value={db}>
      {displayDbName(db)}
    </option>
  ))}
</select>

  {/* <p className="mt-1 text-xs text-gray-500">
    Examples: customer_db_tally, customer_db_onegcp
  </p> */}
</div>
                <label className="flex items-center gap-2 cursor-pointer select-none">
                  <input
                    type="checkbox"
                    checked={csvUploadEnabled}
                    onChange={(e) => setCsvUploadEnabled(e.target.checked)}
                    className="w-4 h-4 rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                  />
                  <span className="text-xs font-semibold text-gray-700">Enable CSV file upload for this board</span>
                </label>
                {newBoardName.trim() && customerDbKey.trim() && (
                  <div className="bg-blue-50 border border-blue-200 rounded-lg p-2.5">
                    <p className="text-xs text-blue-800"><span className="font-semibold">{isEditMode ? 'Ready to update:' : 'Ready to create:'}</span> {newBoardName}</p>
                    <p className="text-xs text-blue-600 mt-0.5">Database: {displayDbName(customerDbKey)}</p>
                    {csvUploadEnabled ? (
                      <p className="text-xs text-green-600 mt-0.5">✓ CSV is enabled for this board</p>
                    ) : (
                      <p className="text-xs text-gray-500 mt-0.5">CSV upload is disabled for this board</p>
                    )}
                  </div>
                )}
              </div>
              <div className="flex justify-end gap-2 mt-5">
                <button onClick={closeModal} className="px-4 py-2 text-xs font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-all duration-200">Cancel</button>
                <button onClick={handleCreateBoard} disabled={!newBoardName.trim() || !customerDbKey.trim()}
                  className="px-4 py-2 text-xs font-medium text-white bg-blue-600 rounded-lg hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed transition-all duration-200">
                  {isEditMode ? 'Update Board' : 'Create Board'}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ── Create Demo Board Modal ──────────────────────────────────────── */}
        {showDemoModal && selectedDemoMainBoardId !== null && (
          <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center" onClick={closeDemoModal}>
            <div className="bg-white rounded-xl shadow-2xl p-5 w-full max-w-sm mx-4 relative" onClick={e => e.stopPropagation()}>
              <div className="mb-4">
                <button className="absolute top-3 right-3 text-gray-500 hover:text-gray-700 hover:bg-gray-100 rounded-full p-1.5 transition-all duration-200" onClick={closeDemoModal}>
                  <X className="w-4 h-4" />
                </button>
                <h2 className="text-base font-bold text-gray-900">{isDemoEditMode ? 'Edit Demo Board' : 'Create a Demo Board'}</h2>
                <p className="text-xs text-gray-500 mt-0.5">
                  {isDemoEditMode ? `Board ID: ${editingDemoBoardId} • Main Board ID: ${selectedDemoMainBoardId}` : `Main Board ID: ${selectedDemoMainBoardId}`}
                </p>
              </div>

              <div className="space-y-3">
                <div>
                  <label className="block text-xs font-semibold text-gray-700 mb-1.5">
                    Board Name <span className="text-red-500">*</span>
                  </label>
                  <input
                    type="text"
                    value={demoBoardName}
                    onChange={e => setDemoBoardName(e.target.value)}
                    placeholder="Enter board name"
                    className="w-full px-3 py-2 text-sm border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-gray-900"
                  />
                </div>

                <div>
                  <label className="block text-xs font-semibold text-gray-700 mb-1.5">
                    Customer Database Key <span className="text-red-500">*</span>
                  </label>
                  <select
                    value={demoCustomerDbKey}
                    onChange={e => setDemoCustomerDbKey(e.target.value)}
                    className="w-full px-3 py-2 text-sm border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-gray-900"
                  >
                    <option value="">Select database</option>
                    {customerDbOptions.map(db => (
                      <option key={db} value={db}>{displayDbName(db)}</option>
                    ))}
                  </select>
                </div>

                {demoBoardName.trim() && demoCustomerDbKey && (
                  <div className="bg-blue-50 border border-blue-200 rounded-lg p-2.5">
                    <p className="text-xs text-blue-800"><span className="font-semibold">{isDemoEditMode ? 'Ready to update:' : 'Ready to create:'}</span> {demoBoardName}</p>
                    <p className="text-xs text-blue-600 mt-0.5">Database: {displayDbName(demoCustomerDbKey)}</p>
                    <p className="text-xs text-green-600 mt-0.5">✓ CSV is enabled for this board</p>
                  </div>
                )}
              </div>

              <div className="flex justify-end gap-2 mt-5">
                <button onClick={closeDemoModal} className="px-4 py-2 text-xs font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-all duration-200">
                  Cancel
                </button>
                <button
                  onClick={handleSaveDemoBoard}
                  disabled={!demoBoardName.trim() || !demoCustomerDbKey || isCreatingDemoBoard}
                  className="px-4 py-2 text-xs font-medium text-white bg-blue-600 rounded-lg hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed transition-all duration-200"
                >
                  {isCreatingDemoBoard ? (isDemoEditMode ? 'Updating...' : 'Creating...') : (isDemoEditMode ? 'Update Board' : 'Create Board')}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ── Resize handle ─────────────────────────────────────────────────── */}
        {!isMobile && (
          <div className="absolute top-0 right-0 w-1 h-full cursor-ew-resize bg-blue-600/20 opacity-0 hover:opacity-100 transition-opacity duration-200" onMouseDown={startResizing} />
        )}
      </div>
    </>
  );
};

export default Sidebar;