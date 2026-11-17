"use client";

import { useEffect, useRef, useState } from 'react';
import { Bricolage_Grotesque, DM_Sans } from 'next/font/google';
import { useTranslation } from 'react-i18next';

const bricolage = Bricolage_Grotesque({ subsets: ['latin'], weight: ['600', '700', '800'], variable: '--font-bricolage' });
const dmSans = DM_Sans({ subsets: ['latin'], weight: ['400', '500', '600', '700'], variable: '--font-dmsans' });
import axios from 'axios';
import LanguageSelector from '../components/LanguageSelector';
import TrialCountdown from '../components/TrialCountdown';
import { useLanguage } from '../context/LanguageContext';
import { translateBatch, formatNumber } from '../utils/translateService';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  LineElement,
  PointElement,
  ArcElement,
  Tooltip,
  Legend,
  Filler,
} from 'chart.js';
import { Bar, Line, Pie } from 'react-chartjs-2';
import { useRouter } from 'next/navigation';
import { Menu, X, Settings, BarChart2, FileText, PieChart, TrendingUp, Database, Users, LayoutDashboard, BookOpen, Play, ChevronRight, ChevronLeft, ChevronDown, ArrowUp, ArrowDown, ArrowUpDown, Maximize2, MessageSquare, Sparkles } from 'lucide-react';
import KPIDashboard from '../Dashboard/page';
import dynamic from 'next/dynamic';

const LiveDataModal = dynamic(
  () => import('../components/LiveData').then(mod => ({ default: mod.LiveDataModal })),
  { ssr: false }
);
import Image from 'next/image';
import loginImage from '../assets/onevega_ai.png';
import RunningPrompt from "@/components/RunningPrompt";
import { FiMic } from "react-icons/fi";
import PptxGenJS from "pptxgenjs";
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';


// Draws the value on top of each bar/point, and the share (%) inside each
// pie slice — chartjs-plugin-datalabels isn't a dependency here, so this is
// a small inline plugin instead of pulling one in just for this.
const valueLabelPlugin = {
  id: 'valueLabel',
  afterDatasetsDraw(chart: import("chart.js").Chart) {
    const { ctx } = chart;
    chart.data.datasets.forEach((dataset, datasetIndex) => {
      const meta = chart.getDatasetMeta(datasetIndex);
      if (meta.hidden || !meta.data) return;
      const chartType = (chart.config as { type?: string }).type;
      const isPie = chartType === 'pie' || chartType === 'doughnut';
      const total = isPie
        ? (dataset.data as number[]).reduce((a, b) => a + (Number(b) || 0), 0)
        : 0;
      // With many bars/points or many thin pie slices, drawing a label on
      // every single one makes them overlap into unreadable mush — skip
      // labels past a point and rely on the tooltip (hover) instead.
      const tooCrowded = !isPie && meta.data.length > 8;
      let lastLabelX = -Infinity;
      meta.data.forEach((element, index) => {
        if (tooCrowded) return;
        const value = (dataset.data as (number | null)[])[index];
        if (value === null || value === undefined) return;
        const pos = element.tooltipPosition(false);
        if (isPie) {
          const pct = total ? Math.round((Number(value) / total) * 100) : 0;
          if (pct < 4) return;
          ctx.save();
          ctx.font = 'bold 11px sans-serif';
          ctx.textAlign = 'center';
          ctx.fillStyle = '#ffffff';
          ctx.textBaseline = 'middle';
          ctx.fillText(`${pct}%`, pos.x, pos.y);
          ctx.restore();
        } else {
          if (pos.x - lastLabelX < 38) return;
          lastLabelX = pos.x;
          ctx.save();
          ctx.font = 'bold 11px sans-serif';
          ctx.textAlign = 'center';
          ctx.fillStyle = '#374151';
          ctx.textBaseline = 'bottom';
          const label = typeof value === 'number' ? value.toFixed(2).replace(/\.00$/, '') : String(value);
          ctx.fillText(label, pos.x, pos.y - 6);
          ctx.restore();
        }
      });
    });
  },
};

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  LineElement,
  PointElement,
  ArcElement,
  Tooltip,
  Legend,
  Filler,
  valueLabelPlugin
);


interface Board {
  name: string;
  is_active: boolean;
  path?: string;
}

interface MainBoard {
  main_board_id: string;
  name: string;
  boards: {
    [key: string]: Board;
  };
}

interface DemoMainBoard { id: number; name: string; }
interface DemoBoard { id: number; name: string; main_board_id: number; is_active?: boolean; }

interface Prompt {
  prompt_text: string;
  id: string;
  prompt_title: string;
  prompt_content: string;
  user_name: string;
  created_at: string;
}

interface UserData {
  email: string;
  userId: string;
  userRole: string;
  userName: string;
}

export default function CXO() {
  const { t } = useTranslation();
  const { language } = useLanguage();
  const [boardNameMap, setBoardNameMap] = useState<Record<string, string>>({});
  const [translatedCxoTexts, setTranslatedCxoTexts] = useState<Record<string, string>>({});
  const [translatedCxoTitles, setTranslatedCxoTitles] = useState<Record<string, string>>({});
  const [navItems, setNavItems] = useState<MainBoard[]>([]);
  // Prompt counts per sub-board, shown on its tile — keyed by board id
  // (string for real boards, demo board ids coerced to string too).
  const [boardPromptCounts, setBoardPromptCounts] = useState<Record<string, number>>({});
  const [selectedMainBoardId, setSelectedMainBoardId] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [showBoardModal, setShowBoardModal] = useState(false);
  const [showPromptsModal, setShowPromptsModal] = useState(false);
  const [selectedBoardId, setSelectedBoardId] = useState<string | null>(null);
  const [prompts, setPrompts] = useState<Prompt[]>([]);
  const [promptsLoading, setPromptsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [boardCheckLoading, setBoardCheckLoading] = useState<string | null>(null);
  const [, setCurrentPromptIndex] = useState(0);
  const [newPromptName, setNewPromptName] = useState('');
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // Auto-grow the prompt textarea to fit its content instead of scrolling
  // internally — runs on every value change, including ones set
  // programmatically (voice input, Re Prompt, Clear), not just typing.
  useEffect(() => {
    const el = textareaRef.current;
    if (!el) return;
    el.style.height = 'auto';
    el.style.height = `${Math.min(el.scrollHeight, 240)}px`;
    el.style.overflowY = el.scrollHeight > 240 ? 'auto' : 'hidden';
  }, [newPromptName]);
  const [activeTab, setActiveTab] = useState("prompts");
  const [cxoView, setCxoView] = useState<"home" | "dashboard" | "livedata">("home");
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);
  const [, setShowCharts] = useState(false);
  const [isRunClicked, setIsRunClicked] = useState(false);
  const router = useRouter();
  const [showDropdown, setShowDropdown] = useState(false);
  const [showBoardDropdown, setShowBoardDropdown] = useState(false);
  const [showRoleDropdown, setShowRoleDropdown] = useState(false);
  const [showModalRoleDropdown, setShowModalRoleDropdown] = useState(false);
  const [tableSort, setTableSort] = useState<{ index: number; direction: 'asc' | 'desc' } | null>(null);
  const [tableFilterTerm, setTableFilterTerm] = useState('');
  const [tablePage, setTablePage] = useState(1);
  const [tablePageSize, setTablePageSize] = useState(10);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const roleDropdownRef = useRef<HTMLDivElement>(null);
  const modalRoleDropdownRef = useRef<HTMLDivElement>(null);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [orgLogoUrl, setOrgLogoUrl] = useState<string | null>(null);
  const [showScrollTop, setShowScrollTop] = useState(false);

  const [userData, setUserData] = useState<UserData>({
    email: "",
    userId: "",
    userRole: "",
    userName: "",
  });
  const [isMounted, setIsMounted] = useState(false);

  interface ChartData {
    chart_type: string;
    data_format: ChartDataFormat;
    insight?: string[];
  }

  interface ChartDataFormat {
    labels?: string[];
    categories?: string[];
    values?: number[] | number[][];
    isStacked?: boolean;
  }

  type RunResult = {
    message: string[];
    table: { columns: string[]; data: string[][] };
    charts: ChartData[];
  };

  const [runResult, setRunResult] = useState<RunResult | null>(null);
  // Index into runResult.charts of the chart currently shown full-screen in
  // the expand lightbox, or null when the lightbox is closed.
  const [expandedChartIdx, setExpandedChartIdx] = useState<number | null>(null);

  useEffect(() => {
    setTableSort(null);
    setTableFilterTerm('');
    setTablePage(1);
  }, [runResult]);

  useEffect(() => {
    setTablePage(1);
  }, [tableFilterTerm, tablePageSize]);

  const [translatedCxoResultColumns, setTranslatedCxoResultColumns] = useState<string[]>([]);
  const [translatedCxoResultData, setTranslatedCxoResultData] = useState<string[][]>([]);
  const [translatedCxoResultMessages, setTranslatedCxoResultMessages] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  // Separate from isLoading (which drives the "Running..." Run button/skeleton)
  // so saving a prompt doesn't make the UI claim it's running the prompt.
  const [isSaving, setIsSaving] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const [filteredPrompt, setFilteredPrompt] = useState<Prompt[]>([]);
  // Prompt header (optional), set from the Consultant panel — stored as a
  // specially-tagged comment (the prompts API itself has no header field).
  // Same convention as Container/page.tsx's HEADER_COMMENT_PREFIX.
  const HEADER_COMMENT_PREFIX = '[prompt-header] ';
  const PROMPT_HEADER_MAX_LENGTH = 100;
  const [cxoCommentsMap, setCxoCommentsMap] = useState<Record<string, { id?: number; comment_text: string }[]>>({});
  const getPromptHeaderText = (promptId: string): string => {
    const c = (cxoCommentsMap[promptId] || []).find(x => x.comment_text.startsWith(HEADER_COMMENT_PREFIX));
    return c ? c.comment_text.slice(HEADER_COMMENT_PREFIX.length) : '';
  };
  // Header can now also be set from CXO itself — previously it was
  // display-only here, populated from whatever Container had already saved.
  const [promptHeader, setPromptHeader] = useState('');
  // Track which prompt IDs are AI-generated (loaded from localStorage, same key as Container page)
  const [cxoGeneratedPromptIds, setCxoGeneratedPromptIds] = useState<Set<string>>(new Set());
  const [toasts, setToasts] = useState<{ id: number; message: string; type: 'error' | 'info' }[]>([]);

  const showToast = (message: string, type: 'error' | 'info' = 'error') => {
    const id = Date.now();
    setToasts(prev => [...prev, { id, message, type }]);
    setTimeout(() => setToasts(prev => prev.filter(t => t.id !== id)), 4000);
  };

  const API_BASE_URL = process.env.NEXT_PUBLIC_API_BASE_URL;
  const EXCEL_API_KEY = process.env.NEXT_PUBLIC_API_KEY || '';

  const scrollRef = useRef<HTMLDivElement>(null);
  const [isListening, setIsListening] = useState(false);
  const [activeMainBoardInSidebar, setActiveMainBoardInSidebar] = useState<string | null>(null);
  const [noPromptsBoard, setNoPromptsBoard] = useState<string | null>(null);
  const [demoMainBoards, setDemoMainBoards] = useState<DemoMainBoard[]>([]);
  const [demoBoards, setDemoBoards] = useState<DemoBoard[]>([]);
  const [isDemoRefOpen, setIsDemoRefOpen] = useState(false);
  const [activeDemoMainBoard, setActiveDemoMainBoard] = useState<string | null>(null);

  // Fetch prompt counts for each sub-board of whichever main board's tile
  // grid is currently shown, so every tile can display "N prompts".
  useEffect(() => {
    if (!selectedMainBoardId) return;
    const mb = navItems.find(i => i.main_board_id === selectedMainBoardId);
    if (!mb) return;
    Object.entries(mb.boards).filter(([, b]) => b.is_active).forEach(async ([bid]) => {
      try {
        const res = await fetch(`${API_BASE_URL}/main-boards/boards/prompts/boards/${bid}`, { headers: { "X-API-Key": EXCEL_API_KEY } });
        if (!res.ok) return;
        const data = await res.json();
        if (Array.isArray(data)) setBoardPromptCounts(prev => ({ ...prev, [bid]: data.length }));
      } catch {}
    });
  }, [selectedMainBoardId, navItems]);

  useEffect(() => {
    if (!activeDemoMainBoard) return;
    const mbId = parseInt(activeDemoMainBoard);
    demoBoards.filter(b => b.main_board_id === mbId).forEach(async (board) => {
      try {
        const res = await fetch(`${API_BASE_URL}/demo/prompts/board/${mbId}/${board.id}`, { headers: { "X-API-Key": EXCEL_API_KEY } });
        if (!res.ok) return;
        const data = await res.json();
        const list = Array.isArray(data) ? data : (data.data ?? data.items ?? []);
        if (Array.isArray(list)) setBoardPromptCounts(prev => ({ ...prev, [String(board.id)]: list.length }));
      } catch {}
    });
  }, [activeDemoMainBoard, demoBoards]);

  const [isDemoLoading, setIsDemoLoading] = useState(false);
  const [sidebarSearch, setSidebarSearch] = useState('');
  const [selectedPrompt, setSelectedPrompt] = useState<Prompt | null>(null);
  const [isDemoBoard, setIsDemoBoard] = useState(false);
  const [selectedDemoBoardId, setSelectedDemoBoardId] = useState<number | null>(null);
  const [demoBoardName, setDemoBoardName] = useState('');
  const [showDownloadModal, setShowDownloadModal] = useState(false);
  const [showEmailModal, setShowEmailModal] = useState(false);
  const [emailData, setEmailData] = useState({ email: '', subject: '', message: '', tableOption: 'limited', reportType: '' });
  const [isSendingEmail, setIsSendingEmail] = useState(false);
  const [showWhatsappModal, setShowWhatsappModal] = useState(false);
  const [whatsappPhone, setWhatsappPhone] = useState('');
  const [whatsappReportType, setWhatsappReportType] = useState<'charts-only' | 'complete'>('charts-only');
  const [whatsappTableOption, setWhatsappTableOption] = useState('limited');
  const [isSendingWhatsapp, setIsSendingWhatsapp] = useState(false);
  const [homeSection, setHomeSection] = useState<"mainboard" | "demo" | null>(null);

  const handleVoiceInput = () => {
  const SpeechRecognition =
    (window as any).SpeechRecognition ||
    (window as any).webkitSpeechRecognition;
  
    if (!SpeechRecognition) {
      alert("Speech Recognition not supported in this browser");
      return;
    }
  
    const recognition = new SpeechRecognition();
    recognition.lang = "en-IN"; // change if needed
    recognition.interimResults = false;
  
    recognition.start();
    setIsListening(true);
  
    recognition.onresult = (event: { results: { transcript: any; }[][]; }) => {
      const transcript = event.results[0][0].transcript;
      setNewPromptName((prev) => prev + " " + transcript);
    };
  
    recognition.onerror = () => {
      setIsListening(false);
    };
  
    recognition.onend = () => {
      setIsListening(false);
    };
  };
  

  const handleResultsScroll = () => {
    const el = scrollRef.current;
    if (!el) return;
    setShowScrollTop(el.scrollTop > 200);
  };

  const handleScrollTop = () => {
    scrollRef.current?.scrollTo({ top: 0, behavior: 'smooth' });
  };

  // Fetch each prompt's comments (so the header — a specially-tagged comment —
  // is available to look up) whenever the prompt list changes.
  useEffect(() => {
    if (!prompts.length) { setCxoCommentsMap({}); return; }
    let cancelled = false;
    (async () => {
      const entries = await Promise.all(
        prompts.map(async (p) => {
          try {
            const res = await fetch(
              `${API_BASE_URL}/main-boards/boards/prompts/${p.id}/comments?order_by=created_at&order_dir=DESC`,
              { headers: { "X-API-Key": EXCEL_API_KEY } }
            );
            if (!res.ok) return [String(p.id), []] as const;
            const data = await res.json();
            const comments = Array.isArray(data) ? data : data.comments || [];
            return [String(p.id), comments] as const;
          } catch {
            return [String(p.id), []] as const;
          }
        })
      );
      if (!cancelled) setCxoCommentsMap(Object.fromEntries(entries));
    })();
    return () => { cancelled = true; };
  }, [prompts]);

  useEffect(() => {
    if (!searchTerm.trim()) {
      setFilteredPrompt(prompts);
    } else {
      const query = searchTerm.toLowerCase();
      setFilteredPrompt(prompts.filter(p =>
        getPromptHeaderText(String(p.id)).toLowerCase().includes(query) ||
        (p.prompt_text && p.prompt_text.toLowerCase().includes(query))
      ));
    }
  }, [prompts, searchTerm, cxoCommentsMap]);

  const filteredPrompts = prompts.filter(p =>
    p.prompt_text.toLowerCase().includes(searchTerm.toLowerCase()) ||
    (p.user_name && p.user_name.toLowerCase().includes(searchTerm.toLowerCase()))
  );

  useEffect(() => {
    if (typeof window !== 'undefined' && !localStorage.getItem('currentUserData')) {
      router.replace('/Login');
    }
  }, []);

  useEffect(() => { setIsMounted(true); }, []);

  useEffect(() => {
    if (!isMounted) return;
    setIsDemoLoading(true);
    Promise.all([fetchDemoMainBoards(), fetchDemoBoards()]).finally(() => setIsDemoLoading(false));
  }, [isMounted]);

  useEffect(() => {
    if (!isMounted || typeof window === 'undefined') return;
    try {
      const sessionData = localStorage.getItem('currentUserData');
      if (sessionData) {
        const p = JSON.parse(sessionData);
        setUserData({ email: p.email || "", userId: p.userId || "", userRole: p.userRole || "", userName: p.userName || "" });
        return;
      }
      const local = {
        email: localStorage.getItem('loggedInUserEmail') || "",
        userId: localStorage.getItem('loggedInUserId') || "",
        userRole: localStorage.getItem('loggedInUserRole') || "",
        userName: localStorage.getItem('loggedInUserName') || "",
      };
      if (local.userId) setUserData(local);
    } catch (e) { console.error('Error loading user data:', e); }
  }, [isMounted]);

  useEffect(() => {
    if (!isMounted || !userData.userId) return;
    const fetchNavItems = async () => {
      try {
        setLoading(true);
        const res = await fetch(`${API_BASE_URL}/main-boards/get_all_info_tree?user_id=${userData.userId}`, {
          method: 'GET',
          headers: { Accept: 'application/json', "X-API-Key": EXCEL_API_KEY },
        });
        if (res.ok) setNavItems(await res.json());
        else setError(`Failed to fetch data: ${res.statusText}`);
      } catch (e) {
        setError('Failed to load data. Please try again.');
      } finally { setLoading(false); }
    };
    fetchNavItems();
  }, [isMounted, userData.userId]);

  useEffect(() => {
    if (!isMounted) return;
    const handleClickOutside = (e: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) setShowDropdown(false);
      if (roleDropdownRef.current && !roleDropdownRef.current.contains(e.target as Node)) setShowRoleDropdown(false);
      if (modalRoleDropdownRef.current && !modalRoleDropdownRef.current.contains(e.target as Node)) setShowModalRoleDropdown(false);
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [isMounted]);

  useEffect(() => {
    if (!isMounted || !userData.userId) return;
    const fetchOrgLogo = async () => {
      try {
        const cached = localStorage.getItem(`logo_cache_${userData.userId}`);
        if (cached) {
          const d = JSON.parse(cached);
          if (d.userId === userData.userId && d.localUrl) setOrgLogoUrl(d.localUrl);
        }
        const metaRes = await fetch(`${API_BASE_URL}/api/logo/${userData.userId}`, {
          headers: { Accept: 'application/json', 'X-API-Key': EXCEL_API_KEY },
        });
        if (!metaRes.ok) return;
        const meta = await metaRes.json();
        if (meta?.success !== true || meta?.logo == null) return;
        const blobRes = await fetch(`${API_BASE_URL}/api/logo/${userData.userId}/view`, {
          headers: { 'X-API-Key': EXCEL_API_KEY },
        });
        if (blobRes.ok) {
          const blob = await blobRes.blob();
          if (blob.size > 0 && blob.type.startsWith('image/')) setOrgLogoUrl(URL.createObjectURL(blob));
        }
      } catch { /* keep cached */ }
    };
    fetchOrgLogo();
  }, [isMounted, userData.userId]);

  const handleRunPrompt = async () => {
    setIsLoading(true);
    setIsRunClicked(true);
    // Clear the previous result so the "running" card actually shows instead
    // of staying hidden behind stale results from the last run.
    setRunResult(null);
    if (!newPromptName?.trim()) { showToast("Please enter a valid prompt.", 'info'); setIsLoading(false); return; }
    try {
      if (isDemoBoard && selectedDemoBoardId !== null) {
        const url = new URL(`${API_BASE_URL}/demo/prompts/${selectedDemoBoardId}/run`);
        url.searchParams.append("input_text", newPromptName.trim());
        url.searchParams.append("use_cache", "true");
        const response = await axios.post(url.href, null, { headers: { "X-API-Key": EXCEL_API_KEY } });
        if (response?.data) {
          setRunResult(response.data);
          if (response.data.message?.length > 0) setActiveTab("message");
          else if (response.data.charts?.length > 0) setActiveTab("charts");
          else if (response.data.table?.columns?.length > 0) setActiveTab("table");
        } else { showToast("No data returned from the server."); }
      } else {
        if (!selectedBoardId) { showToast("Board ID is required.", 'info'); setIsLoading(false); return; }
        const url = new URL(`${API_BASE_URL}/main-boards/boards/prompts/run_prompt_v5?`);
        url.searchParams.append("input_text", newPromptName.trim());
        url.searchParams.append("board_id", selectedBoardId);
        url.searchParams.append("user_name", userData.userName || "Unknown User");
        url.searchParams.append("use_cache", "true");
        const response = await axios.post(url.href,
          { input_text: newPromptName.trim(), board_id: selectedBoardId, user_name: userData.userName || "Unknown User", use_cache: true },
          { headers: { "X-API-Key": EXCEL_API_KEY } }
        );
        if (response?.data) {
          setRunResult(response.data);
          if (response.data.message?.length > 0) setActiveTab("message");
          else if (response.data.charts?.length > 0) setActiveTab("charts");
          else if (response.data.table?.columns?.length > 0) setActiveTab("table");
          setShowCharts(["chart","visualization"].some(k => newPromptName.toLowerCase().includes(k)));
        } else { showToast("No data was returned from the server."); }
      }
    } catch (error: unknown) {
      if (axios.isAxiosError(error)) showToast(`Server Error (${error.response?.status || "Unknown"}): ${error.response?.data?.message || error.message}`);
      else if (error instanceof Error) showToast(`Error: ${error.message}`);
      else showToast("An unknown error occurred.");
    } finally { setIsLoading(false); }
  };

  useEffect(() => {
    const fetchPrompts = async () => {
      if (!selectedBoardId) return;
      setPromptsLoading(true);
      setError(null);
      try {
        const res = await fetch(`${API_BASE_URL}/main-boards/boards/prompts/boards/${selectedBoardId}`, { headers: { "X-API-Key": EXCEL_API_KEY } });
        if (!res.ok) throw new Error(`Failed to fetch prompts: ${res.status} ${res.statusText}`);
        const data = await res.json();
        if (!Array.isArray(data)) { setPrompts([]); throw new Error("Invalid response format"); }
        setPrompts(data);
        // Load generated prompt IDs from localStorage so G1 labels show in View Prompts modal
        try {
          const genStored = localStorage.getItem(`gen_prompt_ids_${selectedBoardId}`);
          setCxoGeneratedPromptIds(genStored ? new Set(JSON.parse(genStored).map(String)) : new Set());
        } catch { setCxoGeneratedPromptIds(new Set()); }
      } catch (e) {
        setError(e instanceof Error ? e.message : "An unknown error occurred");
      } finally { setPromptsLoading(false); }
    };
    fetchPrompts();
  }, [selectedBoardId]);

  // Auto-translate all board names when language or navItems changes
  useEffect(() => {
    if (language === 'en' || navItems.length === 0) {
      setBoardNameMap({});
      return;
    }
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

  // Auto-translate prompt text and titles when language or prompts change
  useEffect(() => {
    if (language === 'en' || prompts.length === 0) {
      setTranslatedCxoTexts({});
      setTranslatedCxoTitles({});
      return;
    }
    const texts = prompts.map(p => p.prompt_text || '');
    translateBatch(texts, language).then(translated => {
      const map: Record<string, string> = {};
      prompts.forEach((p, i) => { map[p.id] = translated[i] || texts[i]; });
      setTranslatedCxoTexts(map);
    });
    const titles = prompts.map(p => p.prompt_title || '');
    if (titles.some(t => t.trim() !== '')) {
      translateBatch(titles, language).then(translated => {
        const map: Record<string, string> = {};
        prompts.forEach((p, i) => { if (p.prompt_title) map[p.id] = translated[i] || titles[i]; });
        setTranslatedCxoTitles(map);
      });
    }
  }, [language, prompts]);

  // Translate run result when language or runResult changes
  useEffect(() => {
    if (!runResult || language === 'en') {
      setTranslatedCxoResultColumns([]);
      setTranslatedCxoResultData([]);
      setTranslatedCxoResultMessages([]);
      return;
    }
    const columns = runResult.table?.columns ?? [];
    const messages: string[] = Array.isArray(runResult.message) ? runResult.message : (runResult.message ? [runResult.message as unknown as string] : []);
    const tableData = runResult.table?.data ?? [];

    // Format numbers synchronously; collect strings for translation
    const formattedData: string[][] = tableData.map(row => [...row]);
    const stringCellCoords: { row: number; col: number }[] = [];
    const stringCellTexts: string[] = [];
    tableData.forEach((row, r) => {
      row.forEach((cell, c) => {
        const str = String(cell ?? '');
        if (!str) return;
        if (!isNaN(Number(str))) {
          formattedData[r][c] = formatNumber(str, language);
        } else {
          stringCellCoords.push({ row: r, col: c });
          stringCellTexts.push(str);
        }
      });
    });

    const allTexts = [...columns, ...messages, ...stringCellTexts];
    if (!allTexts.some(t => t.trim())) {
      setTranslatedCxoResultData(formattedData);
      return;
    }

    translateBatch(allTexts, language).then(translated => {
      const colEnd = columns.length;
      const msgEnd = colEnd + messages.length;
      setTranslatedCxoResultColumns(translated.slice(0, colEnd));
      setTranslatedCxoResultMessages(translated.slice(colEnd, msgEnd));
      stringCellCoords.forEach(({ row, col }, i) => {
        formattedData[row][col] = translated[msgEnd + i] || tableData[row][col];
      });
      setTranslatedCxoResultData(formattedData);
    });
  }, [runResult, language]);

  // Also reload cxoGeneratedPromptIds when selectedBoardId changes (handles modal open without re-fetch)
  useEffect(() => {
    if (!selectedBoardId) { setCxoGeneratedPromptIds(new Set()); return; }
    try {
      const genStored = localStorage.getItem(`gen_prompt_ids_${selectedBoardId}`);
      setCxoGeneratedPromptIds(genStored ? new Set(JSON.parse(genStored).map(String)) : new Set());
    } catch { setCxoGeneratedPromptIds(new Set()); }
  }, [selectedBoardId]);

  const handleMainBoardClick = (id: string) => setSelectedMainBoardId(id);
  const handleBackClick = () => { setActiveTab("prompts"); setSelectedMainBoardId(null); };
  const handleBoardClick = async (id: string) => {
    setBoardCheckLoading(id);
    setNoPromptsBoard(null);
    setRunResult(null);
    setNewPromptName('');
    setPromptHeader('');
    setIsRunClicked(false);
    try {
      const res = await fetch(`${API_BASE_URL}/main-boards/boards/prompts/boards/${id}`, { headers: { "X-API-Key": EXCEL_API_KEY } });
      if (!res.ok) { setNoPromptsBoard(id); showToast("No prompts in this board.", 'info'); return; }
      const data = await res.json();
      if (!Array.isArray(data) || data.length === 0) { setNoPromptsBoard(id); showToast("No prompts in this board.", 'info'); return; }
      setPrompts(data);
      setActiveTab("prompts");
      setSelectedBoardId(id);
      setShowBoardModal(true);
    } catch {
      setNoPromptsBoard(id);
      showToast("Error loading board.");
    } finally {
      setBoardCheckLoading(null);
    }
  };
  const handleDemoBoardClick = async (boardId: number, mainBoardId: number, boardName: string) => {
    setBoardCheckLoading(String(boardId));
    setRunResult(null);
    setNewPromptName('');
    setPromptHeader('');
    setIsRunClicked(false);
    setNoPromptsBoard(null);
    try {
      const res = await fetch(`${API_BASE_URL}/demo/prompts/board/${mainBoardId}/${boardId}`, { headers: { "X-API-Key": EXCEL_API_KEY } });
      if (!res.ok) { showToast("Failed to load demo board prompts."); return; }
      const data = await res.json();
      const promptsList: Prompt[] = Array.isArray(data) ? data : (data.data ?? data.items ?? []);
      if (promptsList.length === 0) { showToast("No prompts available in this board.", 'info'); return; }
      setPrompts(promptsList);
      setActiveTab("prompts");
      setIsDemoBoard(true);
      setSelectedDemoBoardId(boardId);
      setDemoBoardName(boardName);
      setSelectedBoardId(null);
      setShowBoardModal(true);
    } catch { showToast("Error loading demo board."); }
    finally { setBoardCheckLoading(null); }
  };

  const handleCloseBoardModal = () => {
    setShowBoardModal(false); setSelectedBoardId(null); setActiveTab("prompts");
    setSelectedPrompt(null); setNewPromptName(''); setPromptHeader(''); setIsRunClicked(false); setRunResult(null);
    setIsDemoBoard(false); setSelectedDemoBoardId(null); setDemoBoardName('');
    setShowBoardDropdown(false);
    // The View Prompts drawer must not stay open once we've left this board's
    // page — otherwise it keeps floating over whatever screen comes next.
    setShowPromptsModal(false); setCurrentPromptIndex(0); setSearchTerm('');
  };
  const handleViewPromptsClick = () => setShowPromptsModal(true);
  const handleClosePromptsModal = () => { setShowPromptsModal(false); setCurrentPromptIndex(0); setSearchTerm(''); };
  const handlePromptClick = (prompt: Prompt) => {
    setNewPromptName(prompt.prompt_text);
    setPromptHeader(getPromptHeaderText(String(prompt.id)));
    setShowPromptsModal(false);
    textareaRef.current?.focus();
  };

  const selectedMainBoard = navItems.find(i => i.main_board_id === selectedMainBoardId)
    ?? navItems.find(i => selectedBoardId != null && selectedBoardId in i.boards);

  const handleLogout = () => {
    if (typeof window !== 'undefined') {
      localStorage.removeItem('currentUserData');
      ['loggedInUserEmail','loggedInUserId','loggedInUserRole','loggedInUserName','client_user_id'].forEach(k => localStorage.removeItem(k));
    }
    router.replace('/Login');
  };

  const toggleMobileMenu = () => setIsMobileMenuOpen(!isMobileMenuOpen);

  const fetchDemoMainBoards = async () => {
    try {
      const res = await fetch(`${API_BASE_URL}/demo/main-boards`, { headers: { accept: "application/json", "X-API-Key": EXCEL_API_KEY } });
      if (res.ok) { const json = await res.json(); setDemoMainBoards(Array.isArray(json) ? json : (json.data ?? [])); }
    } catch {}
  };
  const fetchDemoBoards = async () => {
    try {
      const res = await fetch(`${API_BASE_URL}/demo/boards`, { headers: { accept: "application/json", "X-API-Key": EXCEL_API_KEY } });
      if (res.ok) { const json = await res.json(); setDemoBoards(Array.isArray(json) ? json : (json.data ?? [])); }
    } catch {}
  };
  const toggleDemoRef = () => {
    const opening = !isDemoRefOpen;
    setIsDemoRefOpen(opening);
    setActiveDemoMainBoard(null);
    if (opening) {
      setIsDemoLoading(true);
      Promise.all([fetchDemoMainBoards(), fetchDemoBoards()]).finally(() => setIsDemoLoading(false));
      setCxoView("home");
      setHomeSection("demo");
    }
  };

  const downloadExcel = () => {
    if (!runResult?.table || runResult.table.data.length === 0) { showToast("No data to download.", 'info'); return; }
    const ws = XLSX.utils.aoa_to_sheet([runResult.table.columns, ...runResult.table.data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Table Data");
    const excelBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const data = new Blob([excelBuffer], { type: "application/octet-stream" });
    saveAs(data, "table_data.xlsx");
  };

  const downloadPPT = async (includeTableData = true, tableRowOption = 'limited', returnAsBase64 = false): Promise<string | void> => {
    try {
      let ppt = new PptxGenJS();
      ppt.author = "Data Analysis Tool";
      ppt.subject = "Data Analysis Results";
      ppt.title = "Insight Analysis Report";

      const THEME = {
        primary: "2B579A", secondary: "4472C4", accent1: "ED7D31",
        accent2: "70AD47", accent3: "5B9BD5", background: "FFFFFF",
        text: "2F3542", headerBackground: "F2F2F2"
      };

      ppt.defineSlideMaster({
        title: "MASTER_SLIDE",
        background: { color: THEME.background },
        margin: [0.5, 0.25, 0.5, 0.25],
        slideNumber: { x: 0.5, y: "95%", fontFace: "Arial", fontSize: 8, color: "666666" },
        objects: [
          { rect: { x: 0, y: 0, w: "100%", h: 0.6, fill: { color: THEME.primary } } },
          { rect: { x: 0, y: "97%", w: "100%", h: 0.2, fill: { color: THEME.primary } } }
        ]
      });
      ppt.defineSlideMaster({
        title: "CLEAN_MASTER_SLIDE",
        background: { color: THEME.background },
        margin: [0.5, 0.25, 0.5, 0.25],
        slideNumber: { x: 0.5, y: "95%", fontFace: "Arial", fontSize: 8, color: "666666" }
      });

      // Title slide
      const titleSlide = ppt.addSlide({ masterName: "MASTER_SLIDE" });
      titleSlide.addText("Insights Analysis Report", { x: 0.5, y: 2.0, fontFace: "Arial", fontSize: 36, color: THEME.primary, bold: true, align: "center" });
      titleSlide.addText("Generated on " + new Date().toLocaleDateString(), { x: 0.5, y: 3.0, fontFace: "Arial", fontSize: 18, color: THEME.text, align: "center" });

      // Prompt slide
      if (newPromptName.trim()) {
        const promptSlide = ppt.addSlide({ masterName: "MASTER_SLIDE" });
        promptSlide.addText("Current Prompt", { x: 0.5, y: 0.12, w: 8.5, fontFace: "Arial", fontSize: 20, color: "FFFFFF", bold: true, align: "left" });
        promptSlide.addText("Query entered by the user", { x: 0.5, y: 0.78, w: 8.5, fontFace: "Arial", fontSize: 11, color: "888888", italic: true, align: "left" });
        const estimatedLines = Math.ceil(newPromptName.trim().length / 80);
        const boxHeight = Math.min(Math.max(estimatedLines * 0.35 + 0.4, 1.0), 4.5);
        promptSlide.addText(newPromptName.trim(), { x: 0.5, y: 1.15, w: 8.5, h: boxHeight, fontFace: "Arial", fontSize: 15, color: THEME.text, fill: { color: "EEF2FF" }, line: { color: "4472C4", pt: 1 }, wrap: true, valign: "middle", align: "left", lineSpacing: 22, margin: [10, 14, 10, 14] });
      }

      // Table data slides
      if (includeTableData && runResult?.table && runResult.table.data.length > 0) {
        try {
          const columns = runResult.table.columns;
          const tableHeader = columns.map(col => ({ text: col, fontFace: "Arial", bold: true, fill: THEME.headerBackground, color: THEME.primary, fontSize: 11 }));
          const dataToDisplay = tableRowOption === 'all' ? runResult.table.data : runResult.table.data.slice(0, 20);
          const COLUMNS_PER_SLIDE_THRESHOLD = 8;

          if (columns.length > COLUMNS_PER_SLIDE_THRESHOLD) {
            const columnsPerSlide = 8;
            const totalColumnSlides = Math.ceil(columns.length / columnsPerSlide);
            for (let colSlideIndex = 0; colSlideIndex < totalColumnSlides; colSlideIndex++) {
              const startCol = colSlideIndex * columnsPerSlide;
              const endCol = Math.min(startCol + columnsPerSlide, columns.length);
              const currentColumnSet = columns.slice(startCol, endCol);
              const partialTableHeader = currentColumnSet.map(col => ({ text: col, fontFace: "Arial", bold: true, fill: THEME.headerBackground, color: THEME.primary, fontSize: 11 }));
              const rowsPerSlide = 15;
              const rowSlidesNeeded = Math.ceil(dataToDisplay.length / rowsPerSlide);
              for (let rowSlideIndex = 0; rowSlideIndex < rowSlidesNeeded; rowSlideIndex++) {
                const startRow = rowSlideIndex * rowsPerSlide;
                const endRow = Math.min(startRow + rowsPerSlide, dataToDisplay.length);
                const currentRows = dataToDisplay.slice(startRow, endRow).map(row => row.slice(startCol, endCol).map(cell => ({ text: String(cell || ''), fontFace: "Arial", fontSize: 10, color: THEME.text })));
                const tableSlide = ppt.addSlide({ masterName: "CLEAN_MASTER_SLIDE" });
                tableSlide.addText(`Table Data - Columns ${startCol + 1}-${endCol}`, { x: 0.5, y: 0.5, fontSize: 18, fontFace: "Arial", color: THEME.primary, bold: true });
                tableSlide.addText(`Rows ${startRow + 1}-${endRow} of ${dataToDisplay.length}`, { x: 0.5, y: 1.0, fontSize: 14, fontFace: "Arial", color: THEME.secondary });
                const availableWidth = 8.5;
                const colWidth = availableWidth / currentColumnSet.length;
                tableSlide.addTable([partialTableHeader, ...currentRows], { x: 0.5, y: 1.4, w: availableWidth, border: { pt: 0.5, color: "CFCFCF" }, colW: currentColumnSet.map(() => colWidth), rowH: Array(currentRows.length + 1).fill(0.3), fill: { color: "FFFFFF" }, valign: "middle", align: "center", fontSize: 10, autoPage: true });
              }
            }
          } else {
            const rowsPerSlide = 10;
            const totalSlides = Math.ceil(dataToDisplay.length / rowsPerSlide);
            for (let slideIndex = 0; slideIndex < totalSlides; slideIndex++) {
              const startRow = slideIndex * rowsPerSlide;
              const endRow = Math.min(startRow + rowsPerSlide, dataToDisplay.length);
              const currentRows = dataToDisplay.slice(startRow, endRow).map(row => row.map(cell => ({ text: String(cell || ''), fontFace: "Arial", fontSize: 10, color: THEME.text })));
              const tableSlide = ppt.addSlide({ masterName: "CLEAN_MASTER_SLIDE" });
              tableSlide.addText(`Table Data (${slideIndex + 1}/${totalSlides})`, { x: 0.5, y: 0.7, fontSize: 20, fontFace: "Arial", color: THEME.primary, bold: true });
              tableSlide.addTable([tableHeader, ...currentRows], { x: 0.5, y: 1.3, w: 8.5, border: { pt: 0.5, color: "CFCFCF" }, colW: columns.map(() => 8.5 / columns.length), rowH: Array(currentRows.length + 1).fill(0.3), fill: { color: "FFFFFF" }, valign: "middle" });
              const rowInfoText = tableRowOption === 'all'
                ? `Showing rows ${startRow + 1} to ${endRow} of ${dataToDisplay.length} total rows`
                : `Showing rows ${startRow + 1} to ${endRow} of 20 ${runResult.table.data.length > 20 ? `(limited from ${runResult.table.data.length} total rows)` : ''}`;
              tableSlide.addText(rowInfoText, { x: 0.5, y: 6.5, fontSize: 10, fontFace: "Arial", italic: true, color: "666666" });
            }
          }
        } catch (error) {
          console.error("Error creating table slides:", error);
        }
      }

      // Chart slides — capture canvas images
      if (runResult?.charts && runResult.charts.length > 0) {
        const chartsGrid = document.getElementById('cxo-charts-grid');
        const canvases = chartsGrid ? chartsGrid.querySelectorAll("canvas") : document.querySelectorAll("canvas");

        runResult.charts.forEach((chart, index) => {
          const slide = ppt.addSlide({ masterName: "CLEAN_MASTER_SLIDE" });
          slide.addText(chart.chart_type.toUpperCase() + " Chart", { x: 0.5, y: 0.7, fontSize: 20, fontFace: "Arial", color: THEME.primary, bold: true, align: "left" });
          const canvas = canvases[index] as HTMLCanvasElement;
          if (canvas) {
            const imgData = canvas.toDataURL("image/png", 1.0);
            slide.addImage({ data: imgData, x: 0.5, y: 1.3, w: 4.5, h: 3.5 });
          } else {
            slide.addText("Chart not available", { x: 0.5, y: 2, fontSize: 14, color: "FF0000" });
          }
          if (chart.insight?.length) {
            slide.addText("Key Insights:", { x: 5.5, y: 1.3, fontSize: 14, fontFace: "Arial", color: THEME.primary, bold: true });
            const maxInsights = Math.min(6, chart.insight.length);
            chart.insight.slice(0, maxInsights).forEach((insight, i) => {
              const text = insight.length > 80 ? insight.substring(0, 77) + "..." : insight;
              slide.addText(text, { x: 5.5, y: 1.7 + i * 0.4, w: 3.5, fontSize: 11, bullet: true, color: THEME.text });
            });
          }
        });
      }

      let fileName = "CXO_Report";
      if (!includeTableData) fileName += "_Charts_Only";
      else if (tableRowOption === 'all') fileName += "_All_Data";
      else fileName += "_Limited_Data";
      fileName += ".pptx";

      if (returnAsBase64) {
        return await ppt.write({ outputType: 'base64' }) as string;
      } else {
        ppt.writeFile({ fileName });
      }
    } catch (e) {
      console.error(e);
      if (returnAsBase64) throw e;
    }
  };

  const sendViaEmail = async (includeTable: boolean | undefined, tableOption: string | undefined) => {
    if (!emailData.email) {
      showToast('Please enter a recipient email address');
      return;
    }
    setIsSendingEmail(true);
    try {
      // Generate PPT client-side using PptxGenJS (same as download, avoids CORS/backend issues)
      const pptBase64 = await downloadPPT(
        includeTable ?? true,
        tableOption ?? 'limited',
        true
      ) as string;

      if (!pptBase64) {
        throw new Error('Failed to generate PPT — please check that there is chart/table data available.');
      }

      // Send via Next.js API route (uses nodemailer server-side, no CORS) —
      // same route the Consultant panel's report emails already go through.
      const response = await fetch('/api/send-report-email', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          recipientEmail: emailData.email,
          subject: emailData.subject || 'Your Data Analysis Report',
          message: emailData.message || '',
          reportType: includeTable ? 'complete' : 'charts-only',
          tableOption: tableOption || 'limited',
          pptBase64,
        }),
      });

      const result = await response.json();

      if (!response.ok || !result.success) {
        const errMsg = result.error || result.message || `HTTP ${response.status}`;
        if (errMsg.includes('OverQuota') || errMsg.includes('out of storage space')) {
          throw new Error("The recipient's email inbox is full. Please use a different email address.");
        }
        throw new Error(errMsg);
      }

      showToast('Email sent successfully!', 'info');
    } catch (error) {
      console.error('Error sending email:', error);
      const errorMessage = error instanceof Error ? error.message : String(error);
      showToast(`Failed to send email: ${errorMessage}`);
    } finally {
      setIsSendingEmail(false);
      setShowEmailModal(false);
      setShowDownloadModal(false);
    }
  };

  const sendViaWhatsapp = async () => {
    const digitsOnly = whatsappPhone.replace(/[^0-9]/g, '');
    if (!digitsOnly) {
      showToast('Please enter a phone number');
      return;
    }
    if (digitsOnly.length < 10) {
      showToast('Please enter a valid phone number with country code');
      return;
    }
    setIsSendingWhatsapp(true);
    try {
      // Generate PPT client-side using PptxGenJS (same as download/email, avoids CORS/backend issues)
      const pptBase64 = await downloadPPT(
        whatsappReportType === 'complete',
        whatsappTableOption,
        true
      ) as string;

      if (!pptBase64) {
        throw new Error('Failed to generate PPT — please check that there is chart/table data available.');
      }

      const response = await fetch('/api/send-report-whatsapp', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ phoneNumber: digitsOnly, pptBase64 }),
      });

      const result = await response.json();

      if (!response.ok || !result.success) {
        throw new Error(result.message || result.error || `HTTP ${response.status}`);
      }

      showToast('Report sent via WhatsApp successfully!', 'info');
      setWhatsappPhone('');
    } catch (error) {
      console.error('Error sending WhatsApp message:', error);
      const errorMessage = error instanceof Error ? error.message : String(error);
      showToast(`Failed to send via WhatsApp: ${errorMessage}`);
    } finally {
      setIsSendingWhatsapp(false);
      setShowWhatsappModal(false);
      setShowDownloadModal(false);
    }
  };

  const handleRePrompt = async () => {
    setIsLoading(true);
    try {
      const res = await axios.post(`${API_BASE_URL}/main-boards/boards/prompts/re_prompt`, null, {
        params: { input_text: newPromptName, board_id: selectedBoardId },
        headers: { "X-API-Key": EXCEL_API_KEY },
      });
      setNewPromptName(res.data.newPromptName || res.data);
      textareaRef.current?.focus();
    } catch (error) {
      if (axios.isAxiosError(error)) showToast(`Server Error (${error.response?.status || 'Unknown'}): ${error.response?.data?.message || error.message}`);
      else if (error instanceof Error) showToast(`Error: ${error.message}`);
      else showToast('An unknown error occurred.');
    } finally { setIsLoading(false); }
  };

  const handleSavePrompt = async () => {
    if (!newPromptName.trim()) { showToast('Prompt cannot be empty!'); return; }
    if (!selectedBoardId) { showToast('Error: board is not selected.'); return; }
    setIsSaving(true);
    try {
      let loggedInUserName: string | null = null;
      try {
        const stored = localStorage.getItem('currentUserData');
        if (stored) { const d = JSON.parse(stored); loggedInUserName = d.userName; }
      } catch {}
      if (!loggedInUserName) loggedInUserName = localStorage.getItem('loggedInUserName');
      if (!loggedInUserName || loggedInUserName.trim() === '' || loggedInUserName === 'Unknown User') {
        showToast('Error: User name missing. Please log in again.'); setIsSaving(false); return;
      }
      const response = await fetch(`${API_BASE_URL}/main-boards/boards/prompts/`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-API-Key': EXCEL_API_KEY },
        body: JSON.stringify({ board_id: selectedBoardId, prompt_text: newPromptName.trim(), prompt_out: 'out_string', user_name: loggedInUserName, created_by: loggedInUserName }),
      });
      if (!response.ok) {
        const err = await response.json();
        showToast(`Failed to save prompt: ${err.message || 'Unknown error'}`); setIsSaving(false); return;
      }
      const newPromptData = await response.json();
      setPrompts(prev => [...prev, newPromptData]);

      // Save the prompt header via the comments API (CXO always creates a
      // new prompt here — there's no edit-in-place flow — so this only
      // ever needs to create a header comment, never update one).
      const trimmedHeader = promptHeader.trim();
      if (trimmedHeader && newPromptData?.id) {
        try {
          const res = await fetch(
            `${API_BASE_URL}/main-boards/boards/prompts/${newPromptData.id}/comments?comment_text=${encodeURIComponent(HEADER_COMMENT_PREFIX + trimmedHeader)}`,
            { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-API-Key': EXCEL_API_KEY } }
          );
          if (res.ok) {
            const data = await res.json();
            const created = data.comment ?? data;
            setCxoCommentsMap(prev => ({ ...prev, [String(newPromptData.id)]: [created, ...(prev[String(newPromptData.id)] || [])] }));
          }
        } catch (headerError) {
          console.error('Failed to save prompt header:', headerError);
          showToast('Prompt saved, but the header failed to save.');
        }
      }

      setPromptHeader('');
      showToast('Prompt saved successfully!', 'info');
    } catch (error) {
      showToast('Network error: Failed to save the prompt.');
    } finally { setIsSaving(false); }
  };

  if (!isMounted || loading) return <div className="flex items-center justify-center h-screen text-sm text-gray-500">Loading...</div>;

  if (!userData.userId) return (
    <div className="flex items-center justify-center min-h-screen">
      <div className="text-center">
        <h2 className="text-base font-semibold mb-3">Authentication Required</h2>
        <p className="mb-3 text-sm text-gray-600">Please log in to access this page.</p>
        <button onClick={() => router.push('/')} className="bg-blue-600 text-white px-3 py-1.5 rounded text-sm hover:bg-blue-700">Go to Login</button>
      </div>
    </div>
  );

  const CHART_COLORS = [
    'rgba(54, 162, 235, 0.8)',
    'rgba(255, 99, 132, 0.8)',
    'rgba(75, 192, 192, 0.8)',
    'rgba(255, 206, 86, 0.8)',
    'rgba(153, 102, 255, 0.8)',
    'rgba(255, 159, 64, 0.8)',
    'rgba(199, 199, 199, 0.8)',
    'rgba(83, 102, 255, 0.8)',
    'rgba(40, 159, 64, 0.8)',
    'rgba(210, 99, 132, 0.8)',
  ];

  const getPieData = (chartData: ChartData) => {
    if (!chartData?.data_format) return { labels: [], datasets: [{ data: [], backgroundColor: [] }] };
    const { labels = [], values = [] } = chartData.data_format;
    return {
      labels,
      datasets: [{
        data: values as number[],
        backgroundColor: labels.map((_, idx) => CHART_COLORS[idx % CHART_COLORS.length]),
        borderColor: '#fff',
        borderWidth: 2,
      }],
    };
  };

  // Richer, fully-opaque line colors — the shared CHART_COLORS palette is
  // intentionally muted/pastel for bar & pie fills, but that reads as washed
  // out for line strokes, so line charts get their own crisper set.
  const LINE_COLORS = ['#2563EB', '#14B8A6', '#8B5CF6', '#F59E0B', '#EF4444'];

  const hexToRgba = (hex: string, alpha: number) => {
    const h = hex.replace('#', '');
    const r = parseInt(h.substring(0, 2), 16);
    const g = parseInt(h.substring(2, 4), 16);
    const b = parseInt(h.substring(4, 6), 16);
    return `rgba(${r},${g},${b},${alpha})`;
  };

  // Scriptable Chart.js fill — a soft vertical gradient under the line,
  // fading from the series color down to almost nothing.
  const makeLineGradient = (hexColor: string) => (context: any) => {
    const { ctx, chartArea } = context.chart;
    if (!chartArea) return hexToRgba(hexColor, 0.04);
    const gradient = ctx.createLinearGradient(0, chartArea.top, 0, chartArea.bottom);
    gradient.addColorStop(0, hexToRgba(hexColor, 0.32));
    gradient.addColorStop(1, hexToRgba(hexColor, 0.04));
    return gradient;
  };

  const getChartData = (chartData: ChartData, type: "bar" | "line") => {
    if (!chartData?.data_format) return { labels: [], datasets: [] };
    const { labels = [], categories, values = [] } = chartData.data_format;
    return {
      labels,
      datasets: (categories || []).map((cat, i) => ({
        label: cat,
        data: Array.isArray(values) && Array.isArray((values as number[][])[i]) ? (values as number[][])[i] : [],
        backgroundColor: type === 'bar'
          ? labels.map((_, idx) => CHART_COLORS[idx % CHART_COLORS.length])
          : makeLineGradient(LINE_COLORS[i % LINE_COLORS.length]),
        borderColor: type === 'line' ? LINE_COLORS[i % LINE_COLORS.length] : CHART_COLORS[i % CHART_COLORS.length],
        borderWidth: 2,
        pointBackgroundColor: type === 'line' ? LINE_COLORS[i % LINE_COLORS.length] : undefined,
        pointBorderColor: type === 'line' ? '#fff' : undefined,
        pointBorderWidth: type === 'line' ? 2 : undefined,
        pointRadius: type === 'line' ? 4 : undefined,
        fill: type === 'line',
        tension: type === 'line' ? 0.35 : 0.3,
      })),
    };
  };

  const cardIconStyles = [
    { Icon: BarChart2,      iconColor: 'text-orange-400',  bg: 'bg-orange-50',  textColor: 'text-orange-500',  hoverBorder: 'hover:border-orange-400'  },
    { Icon: FileText,       iconColor: 'text-purple-400',  bg: 'bg-purple-50',  textColor: 'text-purple-500',  hoverBorder: 'hover:border-purple-400'  },
    { Icon: PieChart,       iconColor: 'text-blue-400',    bg: 'bg-blue-50',    textColor: 'text-blue-500',    hoverBorder: 'hover:border-blue-400'    },
    { Icon: TrendingUp,     iconColor: 'text-green-400',   bg: 'bg-green-50',   textColor: 'text-green-500',   hoverBorder: 'hover:border-green-400'   },
    { Icon: Database,       iconColor: 'text-rose-400',    bg: 'bg-rose-50',    textColor: 'text-rose-500',    hoverBorder: 'hover:border-rose-400'    },
    { Icon: Users,          iconColor: 'text-indigo-400',  bg: 'bg-indigo-50',  textColor: 'text-indigo-500',  hoverBorder: 'hover:border-indigo-400'  },
    { Icon: LayoutDashboard,iconColor: 'text-teal-400',    bg: 'bg-teal-50',    textColor: 'text-teal-500',    hoverBorder: 'hover:border-teal-400'    },
    { Icon: BookOpen,       iconColor: 'text-amber-400',   bg: 'bg-amber-50',   textColor: 'text-amber-500',   hoverBorder: 'hover:border-amber-400'   },
  ];


  return (
    <div className={`flex h-screen overflow-hidden bg-[#F5F8FD] ${dmSans.variable} ${bricolage.variable}`} style={{ fontFamily: 'var(--font-dmsans)' }}>
      <style>{`
        .cxo-sidebar-nav::-webkit-scrollbar { width: 4px; }
        .cxo-sidebar-nav::-webkit-scrollbar-track { background: #e5e7eb; border-radius: 4px; }
        .cxo-sidebar-nav::-webkit-scrollbar-thumb { background: #3b82f6; border-radius: 4px; }
        .cxo-sidebar-nav::-webkit-scrollbar-thumb:hover { background: #2563eb; }
      `}</style>

      {/* Sidebar */}
      <div className={`hidden md:flex flex-col bg-white border-r border-gray-200 flex-shrink-0 transition-all duration-300 overflow-hidden ${isSidebarCollapsed ? 'w-14' : 'w-60'}`}>
        {/* Logo + collapse button row */}
        <div className="flex items-center justify-between px-3 py-4 border-b border-gray-100 flex-shrink-0">
          {!isSidebarCollapsed && (
            <div className="flex-1 flex items-center justify-center">
              {orgLogoUrl ? (
                <img src={orgLogoUrl} alt="Logo" className="max-h-14 object-contain rounded-xl" />
              ) : (
                <div className="h-14 w-24 bg-gray-200 animate-pulse rounded" />
              )}
            </div>
          )}
          <button
            onClick={() => setIsSidebarCollapsed(v => !v)}
            className="flex items-center justify-center w-7 h-7 rounded border border-gray-300 bg-white hover:bg-gray-100 transition-colors flex-shrink-0 ml-auto"
          >
            <ChevronRight className={`w-4 h-4 text-gray-600 transition-transform duration-300 ${isSidebarCollapsed ? '' : 'rotate-180'}`} />
          </button>
        </div>

        {/* Nav items */}
        <nav className="cxo-sidebar-nav flex-1 py-2 px-2 flex flex-col gap-1 overflow-y-auto" style={{ scrollbarWidth: 'thin', scrollbarColor: '#3b82f6 #e5e7eb' }}>
          {/* Search bar */}
          {!isSidebarCollapsed && (
            <div className="relative mb-2">
              <input type="text" placeholder={t('sidebar.searchBoards')} value={sidebarSearch}
                onChange={e => setSidebarSearch(e.target.value)}
                className="w-full py-1.5 pl-3 pr-7 text-xs border border-gray-300 rounded-lg focus:outline-none focus:ring-1 focus:ring-blue-400 bg-white" />
              {sidebarSearch && (
                <button onClick={() => setSidebarSearch('')} className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 text-xs">×</button>
              )}
            </div>
          )}

          {/* search highlight helper */}
          {(() => {
            const q = sidebarSearch.trim().toLowerCase();
            const hl = (text: string) => {
              if (!q) return <span>{text}</span>;
              const parts = text.split(new RegExp(`(${q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'gi'));
              return (
                <span>
                  {parts.map((p, i) =>
                    p.toLowerCase() === q
                      ? <mark key={i} className="bg-yellow-200 text-yellow-900 rounded-sm px-0.5 not-italic">{p}</mark>
                      : p
                  )}
                </span>
              );
            };

            // compute visible items based on search
            const liveDataVisible = !q || 'live data'.includes(q);
            const demoMbFiltered = demoMainBoards.filter(mb => {
              if (!q) return true;
              if (mb.name.toLowerCase().includes(q)) return true;
              return demoBoards.some(b => b.main_board_id === mb.id && b.name.toLowerCase().includes(q));
            });
            const demoSectionVisible = !q || 'demo reference'.includes(q) || demoMbFiltered.length > 0;
            const navFiltered = navItems.filter(item => {
              if (!q) return true;
              if (item.name.toLowerCase().includes(q)) return true;
              return Object.values(item.boards).some(b => b.is_active && b.name.toLowerCase().includes(q));
            });

            return (
              <>
                {/* ── Live Data ── */}
                {liveDataVisible && (
                  <button
                    onClick={() => setCxoView("livedata")}
                    title="Live Data"
                    className={`w-full flex items-center gap-1.5 py-1.5 rounded-lg text-xs font-medium transition-colors ${isSidebarCollapsed ? "justify-center px-0" : "px-2"} ${cxoView === "livedata" ? "bg-blue-500 text-white" : "text-gray-700 hover:bg-blue-100"}`}
                  >
                    {!isSidebarCollapsed && <span className="w-3 h-3 flex-shrink-0" />}
                    <Database className="w-4 h-4 flex-shrink-0" />
                    {!isSidebarCollapsed && <span className="ml-0.5">{hl("Live Data")}</span>}
                  </button>
                )}

                {/* ── Demo Reference ── */}
                {demoSectionVisible && (
                  <div>
                    <button
                      onClick={toggleDemoRef}
                      title="Demo Reference"
                      className={`w-full flex items-center gap-1.5 py-1.5 rounded-lg text-xs font-medium transition-colors ${isSidebarCollapsed ? "justify-center px-0" : "px-2"} ${isDemoRefOpen || q ? "bg-blue-500 text-white" : "text-gray-700 hover:bg-blue-100"}`}
                    >
                      {!isSidebarCollapsed && (
                        <ChevronRight className={`w-3 h-3 flex-shrink-0 transition-transform duration-200 ${isDemoRefOpen || q ? "rotate-90" : ""}`} />
                      )}
                      <Database className="w-4 h-4 flex-shrink-0" />
                      {!isSidebarCollapsed && <span className="ml-0.5">{hl("Demo Reference")}</span>}
                    </button>
                    {(isDemoRefOpen || !!q) && !isSidebarCollapsed && (
                      <div className="ml-3 mt-0.5 space-y-0.5">
                        {isDemoLoading ? (
                          <div className="text-xs text-gray-400 px-3 py-1">Loading...</div>
                        ) : demoMbFiltered.map(mb => {
                          const mbId = String(mb.id);
                          const mbMatches = mb.name.toLowerCase().includes(q);
                          const boardsForMb = demoBoards.filter(b => {
                            if (b.main_board_id !== mb.id) return false;
                            if (!q) return true;
                            return mbMatches || b.name.toLowerCase().includes(q);
                          });
                          const isExpMb = activeDemoMainBoard === mbId || (!!q && boardsForMb.some(b => b.name.toLowerCase().includes(q)));
                          return (
                            <div key={mbId}>
                              <button
                                onClick={() => setActiveDemoMainBoard(prev => prev === mbId ? null : mbId)}
                                className={`w-full flex items-center gap-1.5 px-2 py-1.5 rounded-lg text-xs font-medium transition-colors ${isExpMb ? "bg-blue-400 text-white" : "text-gray-700 hover:bg-blue-50"}`}
                              >
                                <ChevronRight className={`w-3 h-3 flex-shrink-0 transition-transform duration-200 ${isExpMb ? "rotate-90" : ""}`} />
                                <span className="truncate">{hl(mb.name)}</span>
                              </button>
                              {isExpMb && (
                                <div className="ml-5 mt-0.5 space-y-0.5">
                                  {boardsForMb.length === 0
                                    ? <div className="text-xs text-gray-400 px-2 py-1">No boards</div>
                                    : boardsForMb.map(board => (
                                        <button key={board.id}
                                          onClick={() => handleDemoBoardClick(board.id, mb.id, board.name)}
                                          disabled={boardCheckLoading === String(board.id)}
                                          className={`w-full text-left px-2 py-1.5 rounded text-xs transition-colors disabled:opacity-50 ${q && board.name.toLowerCase().includes(q) ? 'bg-yellow-50 text-yellow-900 hover:bg-yellow-100' : 'text-gray-600 hover:bg-blue-50 hover:text-blue-700'}`}
                                        >
                                          {boardCheckLoading === String(board.id)
                                            ? <span className="flex items-center gap-1"><span className="animate-spin inline-block w-3 h-3 border-b border-blue-600 rounded-full" />Loading...</span>
                                            : hl(board.name)}
                                        </button>
                                      ))}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                )}

                {/* ── Mainboards (tree) ── */}
                {navFiltered.map(item => {
                  const mbId = item.main_board_id;
                  const mbMatches = item.name.toLowerCase().includes(q);
                  const activeBoards = Object.entries(item.boards).filter(([, b]) => {
                    if (!b.is_active) return false;
                    if (!q) return true;
                    return mbMatches || b.name.toLowerCase().includes(q);
                  });
                  const hasMatchingBoard = activeBoards.some(([, b]) => b.name.toLowerCase().includes(q));
                  const isExpMb = activeMainBoardInSidebar === mbId || (!!q && hasMatchingBoard);
                  return (
                    <div key={mbId}>
                      <button
                        onClick={() => setActiveMainBoardInSidebar(prev => prev === mbId ? null : mbId)}
                        title={item.name}
                        className={`w-full flex items-center gap-1.5 py-1.5 rounded-lg text-xs font-medium transition-colors ${isSidebarCollapsed ? "justify-center px-0" : "px-2"} ${isExpMb ? "bg-blue-500 text-white" : "text-gray-700 hover:bg-blue-100"}`}
                      >
                        {!isSidebarCollapsed && (
                          <ChevronRight className={`w-3 h-3 flex-shrink-0 transition-transform duration-200 ${isExpMb ? "rotate-90" : ""}`} />
                        )}
                        <BarChart2 className="w-4 h-4 flex-shrink-0" />
                        {!isSidebarCollapsed && <span className="ml-0.5 truncate">{hl(boardNameMap[item.name] || item.name)}</span>}
                      </button>
                      {isExpMb && !isSidebarCollapsed && (
                        <div className="ml-5 mt-0.5 space-y-0.5">
                          {activeBoards.length === 0
                            ? <div className="text-xs text-gray-400 px-2 py-1">No boards</div>
                            : activeBoards.map(([boardId, board]) => (
                                <button key={boardId}
                                  onClick={() => handleBoardClick(boardId)}
                                  disabled={boardCheckLoading === boardId}
                                  className={`w-full text-left px-2 py-1.5 rounded text-xs transition-colors disabled:opacity-50 ${q && board.name.toLowerCase().includes(q) ? 'bg-yellow-50 text-yellow-900 hover:bg-yellow-100' : 'text-gray-600 hover:bg-blue-50 hover:text-blue-700'}`}
                                >
                                  {boardCheckLoading === boardId
                                    ? <span className="flex items-center gap-1"><span className="animate-spin inline-block w-3 h-3 border-b border-blue-600 rounded-full" />Loading...</span>
                                    : hl(boardNameMap[board.name] || board.name)}
                                </button>
                              ))}
                        </div>
                      )}
                    </div>
                  );
                })}
              </>
            );
          })()}

        </nav>
      </div>

      {/* Mobile sidebar overlay */}
      {isMobileMenuOpen && (
        <div className="md:hidden fixed inset-0 bg-black bg-opacity-50 z-40" onClick={toggleMobileMenu} />
      )}

      {/* Mobile sidebar */}
      <div className={`md:hidden fixed top-0 left-0 h-full bg-gray-200 z-50 transition-transform duration-300 w-60 ${isMobileMenuOpen ? 'translate-x-0' : '-translate-x-full'}`}>
        <div className="p-3">
          <div className="flex justify-between items-center mb-3">
            {orgLogoUrl
              ? <img src={orgLogoUrl} alt="Logo" className="h-12 object-contain rounded-xl" />
              : <Image src={loginImage} alt="Logo" width={110} height={48} className="rounded-md object-contain" />}
            <button onClick={toggleMobileMenu} className="p-1.5"><X className="w-5 h-5" /></button>
          </div>
          <nav className="mt-3 space-y-1.5">
            <button
              onClick={() => { setCxoView("home"); setSelectedMainBoardId(null); toggleMobileMenu(); }}
              className={`w-full text-left py-2 px-3 text-sm rounded ${cxoView === "home" ? "bg-blue-100 text-blue-700 font-medium" : "text-gray-700 hover:bg-gray-300"}`}
            >
              {t('header.home')}
            </button>
            <a href="/Consultant" className="block py-2 px-3 text-blue-600 text-sm hover:bg-gray-300 rounded">{t('header.consultant')}</a>
            <button onClick={handleLogout} className="w-full py-2 px-3 bg-blue-600 hover:bg-red-500 rounded text-white text-sm text-left">{t('header.logout')}</button>
          </nav>
        </div>
      </div>

      {/* Main area */}
      <div className="flex-1 flex flex-col overflow-hidden">

        {/* Header */}
        <header className="bg-white shadow-sm px-3 sm:px-5 py-2.5 flex items-center gap-2 sm:gap-4 w-full z-30 sticky top-0">
          {/* Mobile menu */}
          <button className="md:hidden flex-shrink-0" onClick={toggleMobileMenu}><Menu className="w-5 h-5 text-gray-600" /></button>

          {/* Panel switcher — plain links, no dropdown */}
          <div className="flex-1 flex justify-center items-center gap-3 sm:gap-6 min-w-0">
            <button
              type="button"
              onClick={() => router.push('/Container')}
              className="text-xs sm:text-sm font-semibold text-blue-500 hover:text-blue-700 transition-colors whitespace-nowrap"
            >
              {t('header.consultantRole')}
            </button>
            <span className="text-xs sm:text-sm font-bold text-blue-700 whitespace-nowrap">
              {t('header.cxoRole')}
            </span>
          </div>

          {/* Language selector + User info + settings. Wraps instead of
              scrolling on a narrow phone — overflow-x-auto was tried here
              first, but it forces overflow-y to auto too, which clips the
              settings dropdown below (a position: absolute child). */}
          <div className="flex items-center flex-wrap sm:flex-nowrap justify-end gap-2 sm:gap-3 flex-shrink-0">
            {/* Trial countdown badge */}
            <TrialCountdown />

            {/* Language Selector */}
            <LanguageSelector />

            <div className="flex items-center gap-3" ref={dropdownRef}>
              <div className="text-right hidden sm:block">
                <p className="text-sm font-semibold text-gray-800 leading-tight">{userData.userName || 'User'}</p>
                <p className="text-xs text-gray-500 leading-tight">{userData.email}</p>
              </div>
              <div className="relative">
                <button
                  onClick={() => setShowDropdown(v => !v)}
                  className="w-9 h-9 rounded-full bg-blue-600 hover:bg-blue-700 flex items-center justify-center transition-colors"
                >
                  <Settings className="w-4 h-4 text-white" />
                </button>
                {showDropdown && (
                  <div className="absolute right-0 top-full mt-1.5 bg-white shadow-lg rounded-md border border-gray-100 min-w-[120px] z-50">
                    <button
                      onClick={handleLogout}
                      className="w-full text-left px-4 py-2.5 text-sm text-gray-700 hover:bg-gray-50 rounded-md"
                    >
                      {t('header.logout')}
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>
        </header>

        {/* Content */}
        <div className="flex-1 overflow-auto p-8">
          {cxoView === "dashboard" ? (
            <>
              <div className="mb-4">
                <button onClick={() => setCxoView("home")}
                  className="flex items-center gap-1 px-3 py-1.5 text-xs font-medium text-blue-600 border border-blue-400 rounded-md hover:bg-blue-50 transition-colors">
                  {t('header.back')}
                </button>
              </div>
              <KPIDashboard />
            </>
          ) : cxoView === "livedata" ? (
            <LiveDataModal onClose={() => setCxoView("home")} />
          ) : (
            <div className="p-6">

              {/* ── Sub-navigation: Mainboard boards ── */}
              {selectedMainBoardId ? (
                <>
                  <div className="flex items-center gap-3 mb-5">
                    <button onClick={() => setSelectedMainBoardId(null)}
                      className="flex items-center gap-1 px-3 py-1.5 text-xs font-medium text-blue-600 border border-blue-400 rounded-md hover:bg-blue-50 transition-colors">
                      ← Back
                    </button>
                    <h2 className="text-sm font-bold text-gray-700">
                      {(() => { const n = navItems.find(i => i.main_board_id === selectedMainBoardId)?.name ?? "Boards"; return boardNameMap[n] || n; })()}
                    </h2>
                  </div>
                  {/* grid-cols-4 was fixed regardless of screen width, so on
                      a phone the cards got squeezed to the point of
                      overlapping each other's text — 2/3/4 columns now
                      scale with the available width. */}
                  <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-4 pb-3">
                    {Object.entries(navItems.find(i => i.main_board_id === selectedMainBoardId)?.boards ?? {})
                      .filter(([, b]) => b.is_active)
                      .map(([bid, board], idx) => {
                        const style = cardIconStyles[idx % cardIconStyles.length];
                        const isChecking = boardCheckLoading === bid;
                        const promptCount = boardPromptCounts[bid];
                        return (
                          <button key={bid} onClick={() => handleBoardClick(bid)} disabled={isChecking}
                            className={`group bg-white border border-gray-200 rounded-xl p-6 flex flex-col items-center gap-3 shadow-sm hover:shadow-xl hover:scale-[1.03] ${style.hoverBorder} transition-all duration-200 disabled:opacity-50 w-full`}>
                            <div className={`w-16 h-16 rounded-full ${style.bg} flex items-center justify-center`}>
                              {isChecking ? <div className="animate-spin rounded-full h-6 w-6 border-b-2 border-blue-600" /> : <style.Icon className={`w-8 h-8 ${style.iconColor}`} />}
                            </div>
                            <span className={`text-sm font-semibold text-center w-full break-words ${style.textColor}`}>{boardNameMap[board.name] || board.name}</span>
                            {promptCount !== undefined && (
                              <span className="text-xs text-gray-400 -mt-2">{promptCount} prompt{promptCount !== 1 ? 's' : ''}</span>
                            )}
                            <span className={`text-xs font-bold ${style.textColor} opacity-0 group-hover:opacity-100 transition-opacity -mt-1`}>
                              Open board →
                            </span>
                          </button>
                        );
                      })}
                  </div>
                </>

              ) : activeDemoMainBoard ? (
                /* ── Sub-navigation: Demo boards ── */
                <>
                  <div className="flex items-center gap-3 mb-5">
                    <button onClick={() => setActiveDemoMainBoard(null)}
                      className="flex items-center gap-1 px-3 py-1.5 text-xs font-medium text-blue-600 border border-blue-400 rounded-md hover:bg-blue-50 transition-colors">
                      ← Back
                    </button>
                    <h2 className="text-sm font-bold text-gray-700">
                      {demoMainBoards.find(m => String(m.id) === activeDemoMainBoard)?.name ?? "Demo Boards"}
                    </h2>
                  </div>
                  <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-4 pb-3">
                    {demoBoards.filter(b => b.main_board_id === parseInt(activeDemoMainBoard)).map((board, idx) => {
                      const style = cardIconStyles[idx % cardIconStyles.length];
                      const isChecking = boardCheckLoading === String(board.id);
                      const promptCount = boardPromptCounts[String(board.id)];
                      return (
                        <button key={board.id}
                          onClick={() => handleDemoBoardClick(board.id, parseInt(activeDemoMainBoard), board.name)}
                          disabled={isChecking}
                          className={`group bg-white border border-gray-200 rounded-xl p-6 flex flex-col items-center gap-3 shadow-sm hover:shadow-xl hover:scale-[1.03] ${style.hoverBorder} transition-all duration-200 disabled:opacity-50 w-full`}>
                          <div className={`w-16 h-16 rounded-full ${style.bg} flex items-center justify-center`}>
                            {isChecking ? <div className="animate-spin rounded-full h-6 w-6 border-b-2 border-blue-600" /> : <style.Icon className={`w-8 h-8 ${style.iconColor}`} />}
                          </div>
                          <span className={`text-sm font-semibold text-center w-full break-words ${style.textColor}`}>{board.name}</span>
                          {promptCount !== undefined && (
                            <span className="text-xs text-gray-400 -mt-2">{promptCount} prompt{promptCount !== 1 ? 's' : ''}</span>
                          )}
                          <span className={`text-xs font-bold ${style.textColor} opacity-0 group-hover:opacity-100 transition-opacity -mt-1`}>
                            Open board →
                          </span>
                        </button>
                      );
                    })}
                  </div>
                </>

              ) : (
                /* ── Combined home: all tiles in one flat 4-column grid ── */
                <>
                  <div className="mb-6">
                    <h1 className="text-2xl font-bold text-gray-900" style={{ fontFamily: 'var(--font-bricolage)' }}>
                      Welcome back, <span className="text-blue-600">{t('header.cxoRole')}</span>
                    </h1>
                    <p className="text-sm text-gray-500 mt-1">Pick a board below to continue your work.</p>
                  </div>
                  <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-4 pb-3">

                  {/* Live Data */}
                  <button
                    onClick={() => setCxoView("livedata")}
                    className="group bg-white border border-gray-200 rounded-xl p-6 flex flex-col items-center gap-3 shadow-sm hover:shadow-xl hover:scale-[1.03] hover:border-blue-400 transition-all duration-200 w-full"
                  >
                    <div className="w-16 h-16 rounded-full bg-blue-50 flex items-center justify-center">
                      <Database className="w-8 h-8 text-blue-400" />
                    </div>
                    <span className="text-sm font-semibold text-center w-full break-words text-blue-500">Live Data</span>
                    <span className="text-xs text-gray-400 -mt-2">Real-time reports</span>
                    <span className="text-xs font-bold text-blue-600 opacity-0 group-hover:opacity-100 transition-opacity -mt-1">
                      Open board →
                    </span>
                  </button>

                  {/* Demo Reference tiles */}
                  {isDemoLoading ? (
                    <div className="text-xs text-gray-400 flex items-center">Loading...</div>
                  ) : (
                    demoMainBoards.map((mb, idx) => {
                      const style = cardIconStyles[(idx + 4) % cardIconStyles.length];
                      const subCount = demoBoards.filter(b => b.main_board_id === mb.id).length;
                      return (
                        <button key={mb.id} onClick={() => setActiveDemoMainBoard(String(mb.id))}
                          className={`group bg-white border border-gray-200 rounded-xl p-6 flex flex-col items-center gap-3 shadow-sm hover:shadow-xl hover:scale-[1.03] ${style.hoverBorder} transition-all duration-200 w-full`}>
                          <div className={`w-16 h-16 rounded-full ${style.bg} flex items-center justify-center`}>
                            <style.Icon className={`w-8 h-8 ${style.iconColor}`} />
                          </div>
                          <span className={`text-sm font-semibold text-center w-full break-words ${style.textColor}`}>{mb.name}</span>
                          <span className="text-xs text-gray-400 -mt-2">{subCount} sub board{subCount !== 1 ? 's' : ''}</span>
                          <span className={`text-xs font-bold ${style.textColor} opacity-0 group-hover:opacity-100 transition-opacity -mt-1`}>
                            Open board →
                          </span>
                        </button>
                      );
                    })
                  )}

                  {/* Mainboard tiles */}
                  {loading ? (
                    <div className="text-xs text-gray-400 flex items-center">Loading...</div>
                  ) : (
                    navItems.map((item, idx) => {
                      const style = cardIconStyles[idx % cardIconStyles.length];
                      const subCount = Object.keys(item.boards).length;
                      return (
                        <button key={item.main_board_id} onClick={() => setSelectedMainBoardId(item.main_board_id)}
                          className={`group bg-white border border-gray-200 rounded-xl p-6 flex flex-col items-center gap-3 shadow-sm hover:shadow-xl hover:scale-[1.03] ${style.hoverBorder} transition-all duration-200 w-full`}>
                          <div className={`w-16 h-16 rounded-full ${style.bg} flex items-center justify-center`}>
                            <style.Icon className={`w-8 h-8 ${style.iconColor}`} />
                          </div>
                          <span className={`text-sm font-semibold text-center w-full break-words ${style.textColor}`}>{boardNameMap[item.name] || item.name}</span>
                          <span className="text-xs text-gray-400 -mt-2">{subCount} sub board{subCount !== 1 ? 's' : ''}</span>
                          <span className={`text-xs font-bold ${style.textColor} opacity-0 group-hover:opacity-100 transition-opacity -mt-1`}>
                            Open board →
                          </span>
                        </button>
                      );
                    })
                  )}

                  </div>
                </>
              )}
            </div>
          )}
        </div>
      </div>

      {/* Board Modal — full page overlay matching main layout */}
      {showBoardModal && (
        <div className="fixed inset-0 z-50 flex">

          {/* Replicate sidebar */}
           <div className="hidden md:flex flex-col items-center w-28 bg-white border-r border-gray-200 flex-shrink-0 pt-3 pb-4 gap-1 px-2">
  <div className="w-full h-16 flex items-center justify-center">
    {orgLogoUrl ? (
      <img
        src={orgLogoUrl}
        alt="Logo"
        className="max-h-14 object-contain rounded-xl"
      />
    ) : (
      <div className="h-14 w-full bg-gray-300 animate-pulse rounded"></div>
    )}
  </div>
</div>

          {/* Main panel — the whole panel scrolls as one page; only the
              header below stays pinned (sticky), not the breadcrumb or the
              prompt composer underneath it. */}
          <div ref={scrollRef} onScroll={handleResultsScroll} className="flex-1 flex flex-col bg-gray-100 overflow-y-auto">

            {/* Replicate header */}
            <header className="bg-white shadow-md px-3 sm:px-5 py-2.5 flex items-center gap-2 sm:gap-4 w-full z-30 sticky top-0 flex-shrink-0">
              {/* Panel switcher — plain links, no dropdown */}
              <div className="flex-1 flex justify-center items-center gap-3 sm:gap-6 min-w-0">
                <button
                  type="button"
                  onClick={() => router.push('/Container')}
                  className="text-xs sm:text-sm font-semibold text-blue-500 hover:text-blue-700 transition-colors whitespace-nowrap"
                >
                  {t('header.consultantRole')}
                </button>
                <span className="text-xs sm:text-sm font-bold text-blue-700 whitespace-nowrap">
                  {t('header.cxoRole')}
                </span>
              </div>
              <div className="flex items-center gap-2 sm:gap-3 flex-shrink-0">
                <div className="text-right hidden sm:block">
                  <p className="text-sm font-semibold text-gray-800 leading-tight">{userData.userName || 'User'}</p>
                  <p className="text-xs text-gray-500 leading-tight">{userData.email}</p>
                </div>
                <div className="relative">
                  <button onClick={() => setShowBoardDropdown(v => !v)}
                    className="w-9 h-9 rounded-full bg-blue-600 hover:bg-blue-700 flex items-center justify-center">
                    <Settings className="w-4 h-4 text-white" />
                  </button>
                  {showBoardDropdown && (
                    <div className="absolute right-0 top-full mt-1.5 bg-white shadow-lg rounded-md border border-gray-100 min-w-[120px] z-[200]">
                      <button onClick={handleLogout} className="w-full text-left px-4 py-2.5 text-sm text-gray-700 hover:bg-gray-50 rounded-md">Logout</button>
                    </div>
                  )}
                </div>
              </div>
            </header>

            {/* Breadcrumb */}
            <div className="px-6 py-2.5 flex items-center gap-1 text-xs text-gray-500">
              <span onClick={handleCloseBoardModal} className="text-blue-500 hover:underline cursor-pointer font-medium">CXO Panel</span>
              {isDemoBoard ? (
                <>
                  <ChevronRight className="w-3 h-3" />
                  <span onClick={handleCloseBoardModal} className="text-blue-500 hover:underline cursor-pointer font-medium">Demo Reference</span>
                  <ChevronRight className="w-3 h-3" />
                  <span className="text-gray-700 font-semibold">{demoBoardName}</span>
                </>
              ) : (
                <>
                  {selectedMainBoard && (
                    <>
                      <ChevronRight className="w-3 h-3" />
                      {selectedBoardId && selectedMainBoard?.boards[selectedBoardId] ? (
                        <span onClick={handleCloseBoardModal} className="text-blue-500 hover:underline cursor-pointer font-medium">
                          {selectedMainBoard.name}
                        </span>
                      ) : (
                        <span className="text-gray-700 font-semibold">{selectedMainBoard.name}</span>
                      )}
                    </>
                  )}
                  {selectedBoardId && selectedMainBoard?.boards[selectedBoardId] && (
                    <>
                      <ChevronRight className="w-3 h-3" />
                      <span className="text-gray-700 font-semibold">{selectedMainBoard.boards[selectedBoardId].name}</span>
                    </>
                  )}
                </>
              )}
            </div>

            {/* Back + Clear action bar */}
            <div className="mx-5 mb-2 flex items-center justify-end gap-2">
              <button
                onClick={handleCloseBoardModal}
                className="flex items-center gap-1 px-4 py-1.5 text-xs font-bold text-gray-600 border border-gray-300 rounded-full hover:bg-gray-100 transition-colors bg-white"
              >
                ← Back
              </button>
              <button
                onClick={() => { setNewPromptName(""); setPromptHeader(""); setRunResult(null); setIsRunClicked(false); }}
                className="px-4 py-1.5 text-xs font-bold text-red-600 bg-red-50 hover:bg-red-100 rounded-full transition-colors border border-red-200"
              >
                Clear
              </button>
            </div>

            {/* Prompt header (optional) */}
            <div className="relative mx-5 mb-2">
              <input
                type="text"
                className="w-full p-2.5 pr-14 border-2 border-blue-400 rounded-2xl text-xs bg-blue-50 focus:outline-none focus:ring-2 focus:ring-blue-500"
                placeholder="Prompt header (optional)"
                value={promptHeader}
                maxLength={PROMPT_HEADER_MAX_LENGTH}
                onChange={(e) => setPromptHeader(e.target.value.slice(0, PROMPT_HEADER_MAX_LENGTH))}
              />
              <span className="absolute right-2 top-1/2 -translate-y-1/2 text-[10px] text-gray-400 pointer-events-none">
                {PROMPT_HEADER_MAX_LENGTH - promptHeader.length}
              </span>
            </div>

            {/* Prompt toolbar. Stacks on mobile instead of cramming the
                input and 4 non-shrinking buttons onto one row — that used
                to squeeze the input down to almost nothing, so the prompt
                text itself was barely visible. */}
            <div className="mx-5 mb-4 bg-[#1E3A8A] rounded-2xl flex flex-col sm:flex-row sm:items-center px-3 py-2 gap-2 sm:gap-3 shadow-lg">
              {/* textarea instead of a single-line input: a long prompt in
                  an <input> just scrolls sideways out of view, showing only
                  whatever's near the cursor. This wraps it and auto-grows
                  to fit the content (see the effect above) instead of
                  scrolling internally. */}
              <textarea
                ref={textareaRef}
                rows={1}
                className="w-full sm:flex-1 bg-white text-gray-800 placeholder-gray-400 text-sm outline-none min-w-0 rounded-xl px-3 py-2 border-0 resize-none leading-snug overflow-hidden"
                placeholder="Dynamic Prompt Entry..."
                value={newPromptName}
                onChange={e => setNewPromptName(e.target.value)}
                onKeyDown={e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleRunPrompt(); } }}
              />
              <div className="flex items-center gap-2 flex-wrap sm:flex-nowrap">
                <button
                  onClick={handleVoiceInput}
                  title="Click to speak"
                  className={`flex-shrink-0 w-9 h-9 rounded-full flex items-center justify-center text-xs font-medium whitespace-nowrap transition-colors ${isListening ? 'bg-green-500 hover:bg-green-600' : 'bg-blue-600 hover:bg-blue-700'} text-white`}
                >
                  <FiMic className="text-white text-lg" />
                </button>
                <button
                  onClick={handleRePrompt}
                  disabled={isLoading || isSaving}
                  className="flex-shrink-0 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-full text-xs font-bold disabled:opacity-[0.45] disabled:cursor-not-allowed transition-colors"
                >
                  Re Prompt
                </button>
                <button
                  onClick={handleRunPrompt}
                  disabled={!newPromptName.trim() || isLoading || isSaving}
                  title="Run"
                  className="flex-shrink-0 w-9 h-9 rounded-full flex items-center justify-center bg-white border-2 border-blue-600 text-blue-600 hover:bg-blue-50 disabled:opacity-[0.45] disabled:cursor-not-allowed transition-colors"
                >
                  {isLoading ? <span className="w-4 h-4 border-2 border-blue-600 border-t-transparent rounded-full animate-spin inline-block" /> : <Play className="w-4 h-4" />}
                </button>
                {!isDemoBoard && (
                  <button
                    onClick={handleSavePrompt}
                    disabled={!newPromptName.trim() || isLoading || isSaving}
                    className="flex-shrink-0 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-full text-xs font-bold disabled:opacity-[0.45] disabled:cursor-not-allowed transition-colors"
                  >
                    {isSaving
                      ? <><span className="w-3.5 h-3.5 border-2 border-white border-t-transparent rounded-full animate-spin inline-block" /> Saving...</>
                      : 'Save'}
                  </button>
                )}
              </div>
            </div>

            {isDemoBoard && (
              <div className="mx-5 mb-4 px-3 py-2 bg-amber-50 border border-amber-200 rounded-lg text-xs text-amber-800">
                This board is for viewing and demo purposes only — you can run prompts here, but saving is not available. If you want to save a prompt, please use an already created main board and continue there.
              </div>
            )}

            {/* Results area */}
            <div className="px-5 pb-20">
              {isLoading && !runResult && <RunningPrompt />}

              {/* Toolbar: Message/Table/Charts pill segment + downloads only
                  once a result exists, View Prompts always shown on the right. */}
              <div className="flex items-center gap-2 mb-3 flex-wrap">
                {isRunClicked && runResult && (
                  <div className="inline-flex items-center gap-1 bg-blue-50 rounded-full p-1">
                    {['message', 'table', 'charts'].map(tab => {
                      const hasData = tab === 'message'
                        ? (runResult?.message?.length ?? 0) > 0
                        : tab === 'table'
                          ? (runResult?.table?.columns?.length ?? 0) > 0
                          : (runResult?.charts?.length ?? 0) > 0;
                      return (
                        <button key={tab}
                          disabled={!hasData}
                          onClick={() => hasData && setActiveTab(tab)}
                          className={`px-4 py-1.5 rounded-full text-sm font-bold transition-colors ${!hasData ? 'text-gray-300 cursor-not-allowed' : activeTab === tab ? 'bg-blue-600 text-white' : 'text-gray-600 hover:bg-blue-100'}`}>
                          {tab.charAt(0).toUpperCase() + tab.slice(1)}
                        </button>
                      );
                    })}
                  </div>
                )}
                <div className="ml-auto flex gap-2">
                  {isRunClicked && runResult && activeTab !== 'message' && (
                    <>
                      {activeTab === 'table' && runResult?.table?.columns?.length > 0 && (
                        <button onClick={downloadExcel} className="flex items-center gap-1.5 px-4 py-2 text-xs font-bold bg-green-600 hover:bg-green-700 text-white rounded-full transition-colors">
                          <Database className="w-3.5 h-3.5" /> Download Excel
                        </button>
                      )}
                      <button onClick={() => setShowDownloadModal(true)} className="flex items-center gap-1.5 px-4 py-2 text-xs font-bold bg-blue-600 hover:bg-blue-700 text-white rounded-full transition-colors">
                        <FileText className="w-3.5 h-3.5" /> Download PPT
                      </button>
                    </>
                  )}
                  <button
                    onClick={handleViewPromptsClick}
                    className="flex items-center gap-1.5 px-4 py-2 text-xs font-bold bg-white border-2 border-blue-600 text-blue-600 hover:bg-blue-50 rounded-full transition-colors whitespace-nowrap"
                  >
                    <MessageSquare className="w-3.5 h-3.5" /> View Prompts
                  </button>
                </div>
              </div>

              {isRunClicked && runResult && (
                <div>
                  <div>
                    {activeTab === 'message' && (
                      <div className="text-sm text-gray-700 p-4 bg-white rounded-xl shadow-sm">
                        {runResult?.message?.length > 0
                          ? <p>{translatedCxoResultMessages.length > 0 ? translatedCxoResultMessages[0] : runResult.message[0]}</p>
                          : <p className="text-gray-400">No message found.</p>}
                      </div>
                    )}
                    {activeTab === 'table' && (
                      runResult.table?.columns?.length > 0 ? (() => {
                        const baseRows = translatedCxoResultData.length > 0 ? translatedCxoResultData : runResult.table.data;
                        const sortedRows = tableSort
                          ? [...baseRows].sort((a, b) => {
                              const av = a[tableSort.index] ?? '';
                              const bv = b[tableSort.index] ?? '';
                              const an = parseFloat(av);
                              const bn = parseFloat(bv);
                              const bothNumeric = av.trim() !== '' && bv.trim() !== '' && !isNaN(an) && !isNaN(bn);
                              const cmp = bothNumeric ? an - bn : String(av).localeCompare(String(bv));
                              return tableSort.direction === 'asc' ? cmp : -cmp;
                            })
                          : baseRows;
                        const filteredRows = tableFilterTerm.trim()
                          ? sortedRows.filter(row => row.some(cell => String(cell ?? '').toLowerCase().includes(tableFilterTerm.toLowerCase())))
                          : sortedRows;
                        const totalPages = Math.max(1, Math.ceil(filteredRows.length / tablePageSize));
                        const safePage = Math.min(tablePage, totalPages);
                        const pagedRows = filteredRows.slice((safePage - 1) * tablePageSize, safePage * tablePageSize);
                        const toggleSort = (i: number) => {
                          setTableSort(prev =>
                            prev?.index === i
                              ? (prev.direction === 'asc' ? { index: i, direction: 'desc' } : null)
                              : { index: i, direction: 'asc' }
                          );
                        };
                        // Right-align + bold any column whose non-empty values are all numeric.
                        const isNumericCol = (idx: number) => {
                          const vals = filteredRows.map(r => r[idx]).filter(v => v !== undefined && v !== null && String(v).trim() !== '');
                          return vals.length > 0 && vals.every(v => !isNaN(parseFloat(String(v))));
                        };
                        // A column with only a handful of repeating short text values reads as a
                        // status/category field — render those as colored pills instead of plain text.
                        const pillPalette = ['bg-blue-50 text-blue-700', 'bg-emerald-50 text-emerald-700', 'bg-amber-50 text-amber-700', 'bg-violet-50 text-violet-700'];
                        const pillColCache = new Map<number, Map<string, string> | null>();
                        const getPillColor = (idx: number, value: string): string | null => {
                          if (isNumericCol(idx)) return null;
                          if (!pillColCache.has(idx)) {
                            const vals = filteredRows.map(r => String(r[idx] ?? '').trim()).filter(Boolean);
                            const unique = Array.from(new Set(vals));
                            const looksLikeStatus = unique.length > 1 && unique.length <= 4 && vals.length >= filteredRows.length * 0.6 && unique.every(v => v.length <= 20);
                            if (!looksLikeStatus) {
                              pillColCache.set(idx, null);
                            } else {
                              const map = new Map<string, string>();
                              unique.forEach((v, i) => map.set(v, pillPalette[i % pillPalette.length]));
                              pillColCache.set(idx, map);
                            }
                          }
                          return pillColCache.get(idx)?.get(value.trim()) ?? null;
                        };
                        return (
                      <div className="bg-blue-50/40 rounded-2xl p-4">
                        {/* Title + record count + filter */}
                        <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 mb-3 px-1">
                          <div className="flex items-center gap-2 flex-wrap">
                            <h3 className="text-base font-bold text-gray-900">{promptHeader || 'Results'}</h3>
                            <span className="text-xs font-bold text-blue-600">{filteredRows.length} record{filteredRows.length !== 1 ? 's' : ''}</span>
                          </div>
                          <div className="relative w-full sm:w-64">
                            <svg className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg>
                            <input
                              type="text"
                              value={tableFilterTerm}
                              onChange={e => setTableFilterTerm(e.target.value)}
                              placeholder="Filter results..."
                              className="w-full pl-9 pr-3 py-2 text-sm bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-400"
                            />
                          </div>
                        </div>

                        <div className="border border-gray-200 rounded-xl bg-white shadow-sm overflow-auto" style={{ maxHeight: 'calc(100vh - 320px)', scrollbarWidth: 'auto', scrollbarColor: '#93c5fd #f1f1f1' }}>
                          <table className="min-w-full table-auto text-sm whitespace-nowrap border-collapse">
                            <thead style={{ position: 'sticky', top: 0, zIndex: 10 }}>
                              <tr>
                                {runResult.table.columns.map((col, i) => (
                                  <th
                                    key={i}
                                    onClick={() => toggleSort(i)}
                                    className={`px-4 py-3.5 bg-blue-600 font-bold text-white text-xs uppercase tracking-wide cursor-pointer select-none hover:bg-blue-700 transition-colors ${isNumericCol(i) ? 'text-right' : 'text-left'}`}
                                  >
                                    <span className={`inline-flex items-center gap-1 ${isNumericCol(i) ? 'flex-row-reverse' : ''}`}>
                                      {translatedCxoResultColumns[i] || col}
                                      {tableSort?.index === i
                                        ? (tableSort.direction === 'asc' ? <ArrowUp className="w-3 h-3" /> : <ArrowDown className="w-3 h-3" />)
                                        : <ArrowUpDown className="w-3 h-3 text-blue-200" />}
                                    </span>
                                  </th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              {pagedRows.length > 0
                                ? pagedRows.map((row, ri) => (
                                    <tr key={ri} className={`${ri % 2 === 0 ? 'bg-white' : 'bg-blue-50/40'} hover:bg-blue-50 transition-colors`}>
                                      {row.map((cell, ci) => {
                                        const pillColor = getPillColor(ci, String(cell ?? ''));
                                        return (
                                          <td
                                            key={ci}
                                            className={`px-4 py-3.5 border-b border-gray-100 text-sm ${isNumericCol(ci) ? 'text-right font-semibold text-gray-900' : 'font-semibold text-gray-900'}`}
                                          >
                                            {pillColor ? (
                                              <span className={`inline-block px-2.5 py-1 rounded-full text-xs font-bold ${pillColor}`}>{cell}</span>
                                            ) : cell}
                                          </td>
                                        );
                                      })}
                                    </tr>
                                  ))
                                : <tr><td colSpan={runResult.table.columns.length} className="text-center p-3 text-gray-400">No data available.</td></tr>}
                            </tbody>
                          </table>
                        </div>

                        {/* Pagination */}
                        {filteredRows.length > 0 && (
                          <div className="flex flex-col sm:flex-row items-center justify-between gap-2 mt-3 px-1 text-sm text-gray-500">
                            <div className="flex items-center gap-2">
                              <span>
                                {(safePage - 1) * tablePageSize + 1}–{Math.min(safePage * tablePageSize, filteredRows.length)} of {filteredRows.length}
                              </span>
                              <select
                                value={tablePageSize}
                                onChange={e => setTablePageSize(Number(e.target.value))}
                                className="border border-gray-200 rounded-lg px-2 py-1 text-xs bg-white focus:outline-none focus:ring-2 focus:ring-blue-400"
                              >
                                {[10, 25, 50].map(n => <option key={n} value={n}>{n} per page</option>)}
                              </select>
                            </div>
                            <div className="flex items-center gap-1">
                              <button
                                onClick={() => setTablePage(p => Math.max(1, p - 1))}
                                disabled={safePage <= 1}
                                className="w-7 h-7 rounded-full flex items-center justify-center bg-white border border-gray-200 hover:bg-gray-100 disabled:opacity-30"
                              >
                                <ChevronLeft className="w-4 h-4" />
                              </button>
                              {Array.from({ length: totalPages }).slice(0, 5).map((_, i) => (
                                <button
                                  key={i}
                                  onClick={() => setTablePage(i + 1)}
                                  className={`w-7 h-7 rounded-full text-xs font-bold ${safePage === i + 1 ? 'bg-blue-600 text-white' : 'bg-white border border-gray-200 hover:bg-gray-100 text-gray-600'}`}
                                >
                                  {i + 1}
                                </button>
                              ))}
                              <button
                                onClick={() => setTablePage(p => Math.min(totalPages, p + 1))}
                                disabled={safePage >= totalPages}
                                className="w-7 h-7 rounded-full flex items-center justify-center bg-white border border-gray-200 hover:bg-gray-100 disabled:opacity-30"
                              >
                                <ChevronRight className="w-4 h-4" />
                              </button>
                            </div>
                          </div>
                        )}
                      </div>
                        );
                      })() : (
                        <div className="text-sm text-gray-400 p-4 bg-white rounded-xl shadow-sm">No table data found.</div>
                      )
                    )}
                    {activeTab === 'charts' && (
                      runResult.charts && runResult.charts.length > 0 ? (
                        <div id="cxo-charts-grid" className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                          {runResult.charts.map((chart: ChartData, i: number) => {
                            if (chart.chart_type === 'pie') return (
                              <div key={i} className="bg-white rounded-xl shadow-sm p-3 border border-gray-200">
                                <div className="relative mb-2">
                                  <h5 className="text-sm font-semibold text-center">Pie Chart</h5>
                                  <button onClick={() => setExpandedChartIdx(i)} title="Expand chart" className="absolute top-0 right-0 text-gray-400 hover:text-blue-600 p-0.5">
                                    <Maximize2 className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                                <div style={{ height: '260px' }}><Pie data={getPieData(chart)} options={{ maintainAspectRatio: false, plugins: { legend: { display: true, position: 'top' } } }} /></div>
                                {chart.insight && chart.insight.length > 0 && <div className="mt-3 p-4 bg-blue-50 rounded-xl text-sm text-gray-700"><p className="font-bold text-gray-900 mb-1.5">Insights</p><ul className="list-disc list-inside space-y-0.5">{chart.insight.map((ins, j) => <li key={j}>{ins}</li>)}</ul></div>}
                              </div>
                            );
                            if (chart.chart_type === 'bar') return (
                              <div key={i} className="bg-white rounded-xl shadow-sm p-3 border border-gray-200">
                                <div className="relative mb-2">
                                  <h5 className="text-sm font-semibold text-center">Bar Chart</h5>
                                  <button onClick={() => setExpandedChartIdx(i)} title="Expand chart" className="absolute top-0 right-0 text-gray-400 hover:text-blue-600 p-0.5">
                                    <Maximize2 className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                                <div style={{ height: '260px' }}><Bar data={getChartData(chart, 'bar')} options={{ maintainAspectRatio: false, plugins: { legend: { display: true, position: 'top' } }, scales: { y: { beginAtZero: true } } }} /></div>
                                {chart.insight && chart.insight.length > 0 && <div className="mt-3 p-4 bg-blue-50 rounded-xl text-sm text-gray-700"><p className="font-bold text-gray-900 mb-1.5">Insights</p><ul className="list-disc list-inside space-y-0.5">{chart.insight.map((ins, j) => <li key={j}>{ins}</li>)}</ul></div>}
                              </div>
                            );
                            if (chart.chart_type === 'line') return (
                              <div key={i} className="bg-white rounded-xl shadow-sm p-3 border border-gray-200">
                                <div className="relative mb-2">
                                  <h5 className="text-sm font-semibold text-center">Line Chart</h5>
                                  <button onClick={() => setExpandedChartIdx(i)} title="Expand chart" className="absolute top-0 right-0 text-gray-400 hover:text-blue-600 p-0.5">
                                    <Maximize2 className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                                <div style={{ height: '260px' }}><Line data={getChartData(chart, 'line')} options={{ maintainAspectRatio: false, plugins: { legend: { display: true, position: 'top' } }, scales: { y: { beginAtZero: true } } }} /></div>
                                {chart.insight && chart.insight.length > 0 && <div className="mt-3 p-4 bg-blue-50 rounded-xl text-sm text-gray-700"><p className="font-bold text-gray-900 mb-1.5">Insights</p><ul className="list-disc list-inside space-y-0.5">{chart.insight.map((ins, j) => <li key={j}>{ins}</li>)}</ul></div>}
                              </div>
                            );
                            return null;
                          })}
                        </div>
                      ) : (
                        <div className="text-sm text-gray-400 p-4 bg-white rounded-xl shadow-sm">No charts found.</div>
                      )
                    )}
                  </div>
                </div>
              )}

              {/* Empty state — shown before anything has been run, and again
                  after Clear. Fills the space below the prompt composer with
                  a starting point instead of leaving it blank. */}
              {!isLoading && !(isRunClicked && runResult) && (
                <div
                  className="flex flex-col items-center justify-center text-center py-16 px-6 bg-white rounded-2xl mt-2"
                  style={{ border: '2px dashed #B7CBF5' }}
                >
                  <div
                    className="relative flex items-center justify-center mb-5 rounded-3xl"
                    style={{ width: 92, height: 92, backgroundColor: '#E8EFFD' }}
                  >
                    <MessageSquare className="w-10 h-10 text-blue-600" />
                    <Sparkles className="w-5 h-5 text-blue-600 absolute -top-1 -right-1" />
                  </div>
                  <h3
                    className="text-xl font-bold text-gray-900 mb-2"
                    style={{ fontFamily: 'var(--font-bricolage)' }}
                  >
                    Ask a question about {(selectedBoardId && selectedMainBoard?.boards[selectedBoardId]?.name) || 'this board'}
                  </h3>
                  <p className="text-sm text-gray-500 max-w-md mb-6">
                    Type a prompt above and press Run, or use View Prompts to pick a saved one.
                  </p>
                  {prompts.length > 0 && (
                    <div className="flex flex-wrap items-center justify-center gap-2">
                      {prompts.slice(0, 3).map((p, i) => (
                        <button
                          key={p.id || i}
                          onClick={() => handlePromptClick(p)}
                          className="inline-flex items-center gap-1.5 px-4 py-2 rounded-full border border-gray-200 bg-white text-xs font-semibold text-gray-700 hover:border-blue-300 hover:bg-blue-50 transition-colors max-w-[240px]"
                        >
                          <Sparkles className="w-3.5 h-3.5 text-blue-600 flex-shrink-0" />
                          <span className="truncate">{getPromptHeaderText(String(p.id)) || p.prompt_text}</span>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              )}

              {showScrollTop && (
    <button
      onClick={handleScrollTop}
      className="absolute bottom-6 right-6 bg-blue-600 hover:bg-blue-700 text-white text-xs px-4 py-2 rounded-full shadow-lg transition-all z-50"
    >
      ↑ Top
    </button>
  )}
            </div>
          </div>

        </div>
      )}
       {/* {showScrollTop && (
  <button
    onClick={handleScrollTop}
    className="fixed bottom-6 right-6 bg-blue-600 hover:bg-blue-700 text-white text-xs px-4 py-2 rounded-full shadow-lg transition-all"
  >
    ↑ Top
  </button>
)} */}

      {/* Chart expand lightbox — lets you step through every chart one at a
          time at full size. */}
      {expandedChartIdx !== null && runResult?.charts && runResult.charts[expandedChartIdx] && (() => {
        const charts = runResult.charts;
        const chart = charts[expandedChartIdx];
        const chartTitle = chart.chart_type === 'pie' ? 'Pie Chart' : chart.chart_type === 'bar' ? 'Bar Chart' : 'Line Chart';
        const bigOptions = {
          maintainAspectRatio: false,
          plugins: { legend: { display: true, position: 'top' as const, labels: { font: { size: 13 }, boxWidth: 14 } } },
          scales: chart.chart_type === 'pie' ? undefined : { y: { beginAtZero: true } }
        };
        return (
          <div
            className="fixed inset-0 bg-black bg-opacity-50 z-[90] flex items-center justify-center px-4"
            onClick={() => setExpandedChartIdx(null)}
          >
            <div
              className="bg-white rounded-2xl shadow-2xl w-full max-w-3xl max-h-[85vh] flex flex-col overflow-hidden"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="flex-shrink-0 flex items-center justify-between px-5 py-3 border-b border-gray-100">
                <h3 className="text-base font-bold text-gray-900">
                  {chartTitle} <span className="text-xs font-normal text-gray-400">({expandedChartIdx + 1} / {charts.length})</span>
                </h3>
                <button
                  onClick={() => setExpandedChartIdx(null)}
                  className="text-gray-400 hover:text-gray-700 text-xl leading-none"
                >
                  ×
                </button>
              </div>

              <div className="flex-1 overflow-y-auto">
                <div className="flex items-center gap-3 px-3 py-3">
                  <button
                    onClick={() => setExpandedChartIdx((expandedChartIdx - 1 + charts.length) % charts.length)}
                    disabled={charts.length < 2}
                    title="Previous chart"
                    className="flex-shrink-0 w-9 h-9 rounded-full bg-gray-100 hover:bg-gray-200 flex items-center justify-center disabled:opacity-30"
                  >
                    <ChevronLeft className="w-5 h-5 text-gray-600" />
                  </button>

                  <div style={{ height: '300px', position: 'relative' }} className="flex-1">
                    {chart.chart_type === 'pie' ? (
                      <Pie data={getPieData(chart)} options={bigOptions as any} />
                    ) : chart.chart_type === 'bar' ? (
                      <Bar data={getChartData(chart, 'bar')} options={bigOptions as any} />
                    ) : (
                      <Line data={getChartData(chart, 'line')} options={bigOptions as any} />
                    )}
                  </div>

                  <button
                    onClick={() => setExpandedChartIdx((expandedChartIdx + 1) % charts.length)}
                    disabled={charts.length < 2}
                    title="Next chart"
                    className="flex-shrink-0 w-9 h-9 rounded-full bg-gray-100 hover:bg-gray-200 flex items-center justify-center disabled:opacity-30"
                  >
                    <ChevronRight className="w-5 h-5 text-gray-600" />
                  </button>
                </div>

                {chart.insight && chart.insight.length > 0 && (
                  <div className="mx-5 mb-4 p-4 bg-blue-50 rounded-xl text-sm text-gray-700">
                    <p className="font-bold text-gray-900 mb-1.5">Insights</p>
                    <ul className="list-disc list-inside space-y-0.5">{chart.insight.map((ins, i) => <li key={i}>{ins}</li>)}</ul>
                  </div>
                )}
              </div>
            </div>
          </div>
        );
      })()}

      {/* Prompts panel — fixed right overlay */}
      {showPromptsModal && (
        <div className="fixed top-0 right-0 h-full w-96 bg-white shadow-2xl border-l border-gray-200 flex flex-col z-[60] transition-transform duration-300">
          <div className="flex items-center justify-between px-5 py-4 border-b border-gray-100 flex-shrink-0">
            <div>
              <h4 className="text-lg font-bold text-gray-900">Prompts</h4>
              <p className="text-xs text-gray-400 mt-0.5">{prompts.length} prompt{prompts.length !== 1 ? 's' : ''} saved</p>
            </div>
            <button onClick={handleClosePromptsModal} className="text-gray-400 hover:text-gray-700 text-xl leading-none">×</button>
          </div>
          <div className="px-4 py-3 border-b border-gray-100 flex-shrink-0">
            <div className="relative">
              <input
                type="text"
                placeholder="Search prompts..."
                value={searchTerm}
                onChange={e => setSearchTerm(e.target.value)}
                autoFocus
                suppressHydrationWarning
                className="w-full py-2 px-4 pr-8 border-2 border-blue-100 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"
              />
              {searchTerm && (
                <button onClick={() => setSearchTerm('')} className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 text-sm">×</button>
              )}
            </div>
          </div>
          <div className="flex-1 overflow-y-auto px-4 py-3">
            {promptsLoading ? (
              <div className="flex justify-center items-center h-16">
                <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-blue-600" />
              </div>
            ) : error ? (
              <p className="text-red-500 text-xs p-2">{error}</p>
            ) : filteredPrompt.length === 0 ? (
              <div className="text-center text-gray-400 text-xs p-3">
                {searchTerm ? `No prompts found for "${searchTerm}"` : 'No prompts found for this board.'}
                {searchTerm && <button onClick={() => setSearchTerm('')} className="block mx-auto mt-1 text-blue-600 hover:underline">Clear</button>}
              </div>
            ) : (
              <div className="space-y-5 pb-2">
                {filteredPrompt.map((prompt: Prompt, index: number) => {
                  const isGenerated = cxoGeneratedPromptIds.has(String(prompt.id));
                  const gCount = filteredPrompt.slice(0, index).filter(p => cxoGeneratedPromptIds.has(String(p.id))).length + 1;
                  const nCount = filteredPrompt.slice(0, index).filter(p => !cxoGeneratedPromptIds.has(String(p.id))).length + 1;
                  const label = isGenerated ? `G${gCount}` : `${nCount}`;
                  const promptText = translatedCxoTexts[prompt.id] || prompt.prompt_text;
                  const headerText = getPromptHeaderText(String(prompt.id));
                  const inUse = newPromptName.trim() === prompt.prompt_text.trim();
                  return (
                    <div key={prompt.id || index} className="relative pb-3 pl-2">
                      {/* Speech bubble */}
                      <div
                        onClick={() => handlePromptClick(prompt)}
                        className={`relative rounded-2xl p-4 pb-3 cursor-pointer transition-all duration-200 hover:scale-[1.02] hover:shadow-lg ${inUse ? 'bg-blue-600' : 'bg-[#E8EFFD] border border-blue-100 hover:border-blue-300'}`}
                      >
                        <h4 className={`text-sm font-bold leading-snug line-clamp-2 ${inUse ? 'text-white' : 'text-gray-900'}`}>
                          {headerText || promptText}
                        </h4>
                        <p className={`text-xs mt-1 line-clamp-3 leading-relaxed ${inUse ? 'text-blue-100' : 'text-gray-500'}`}>
                          {promptText}
                        </p>
                        <div className="flex items-center justify-between mt-2.5">
                          <span className={`text-[11px] ${inUse ? 'text-blue-100' : 'text-gray-400'}`}>Created by {prompt.user_name || 'Unknown'}</span>
                          {inUse ? (
                            <span className="text-[11px] font-bold text-white">In use</span>
                          ) : (
                            <span className="text-[11px] font-bold text-blue-600 hover:underline">Use prompt</span>
                          )}
                        </div>
                        {/* Bubble tail */}
                        <div className={`absolute -bottom-1.5 left-6 w-4 h-4 rotate-45 ${inUse ? 'bg-blue-600' : 'bg-[#E8EFFD] border-r border-b border-blue-100'}`} />
                      </div>
                      {/* Number avatar */}
                      <span className={`absolute -bottom-1 left-0 w-7 h-7 rounded-full flex items-center justify-center text-xs font-bold ring-2 ring-white shadow-sm ${inUse ? 'bg-blue-600 text-white' : 'bg-white text-gray-700 border border-gray-200'}`}>
                        {label}
                      </span>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      )}

      {/* Download Report Modal */}
      {showDownloadModal && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-[200] p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-sm w-full overflow-hidden">
            <div className="flex items-center justify-between px-5 py-3.5 border-b border-gray-100 bg-gradient-to-r from-blue-50 to-indigo-50">
              <h3 className="text-base font-bold text-blue-700">Download Report</h3>
              <button
                onClick={() => setShowDownloadModal(false)}
                className="p-1 text-gray-400 hover:text-gray-600 hover:bg-white rounded-full transition-colors"
              >
                <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                </svg>
              </button>
            </div>

            <div className="p-5 space-y-4">
              <div>
                <p className="text-xs font-semibold text-gray-500 uppercase tracking-wide mb-2">Charts only</p>
                <div className="grid grid-cols-3 gap-1.5">
                  <button
                    onClick={() => { setShowDownloadModal(false); downloadPPT(false, 'limited'); }}
                    className="py-1.5 px-1 text-xs bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 transition-colors"
                  >
                    Download
                  </button>
                  <button
                    onClick={() => {
                      setEmailData(prev => ({ ...prev, reportType: 'charts-only', tableOption: 'limited' }));
                      setShowEmailModal(true);
                    }}
                    className="py-1.5 px-1 text-xs bg-green-600 text-white rounded-lg font-medium hover:bg-green-700 transition-colors"
                  >
                    Email
                  </button>
                  <button
                    onClick={() => {
                      setWhatsappReportType('charts-only');
                      setWhatsappTableOption('limited');
                      setShowWhatsappModal(true);
                    }}
                    className="py-1.5 px-1 text-xs bg-emerald-500 text-white rounded-lg font-medium hover:bg-emerald-600 transition-colors flex items-center justify-center gap-1"
                  >
                    <svg className="w-3 h-3 flex-shrink-0" viewBox="0 0 24 24" fill="currentColor"><path d="M12.04 2C6.58 2 2.13 6.45 2.13 11.91c0 1.75.46 3.45 1.32 4.95L2.05 22l5.25-1.38a9.9 9.9 0 004.74 1.21h.01c5.46 0 9.9-4.45 9.9-9.91 0-2.65-1.03-5.14-2.9-7.01A9.82 9.82 0 0012.04 2m0 1.67a8.23 8.23 0 018.24 8.24c0 4.54-3.7 8.23-8.24 8.23a8.2 8.2 0 01-4.19-1.15l-.3-.17-3.12.82.83-3.04-.2-.32a8.2 8.2 0 01-1.26-4.37c0-4.54 3.7-8.24 8.24-8.24m-4.52 4.31c-.16 0-.42.06-.64.3-.22.24-.85.83-.85 2.02s.87 2.35.99 2.51c.12.16 1.7 2.7 4.2 3.68 2.08.82 2.5.66 2.95.62.45-.04 1.46-.6 1.66-1.18.2-.58.2-1.07.14-1.18-.06-.1-.22-.16-.46-.28-.24-.12-1.46-.72-1.68-.8-.23-.08-.39-.12-.56.12-.16.24-.63.8-.78.96-.14.16-.28.18-.52.06-.24-.12-1.02-.38-1.94-1.2-.72-.64-1.2-1.43-1.34-1.67-.14-.24-.02-.37.11-.49.11-.11.24-.28.36-.42.12-.14.16-.24.24-.4.08-.16.04-.3-.02-.42-.06-.12-.56-1.35-.77-1.85-.2-.48-.4-.42-.56-.43z"/></svg>
                    WhatsApp
                  </button>
                </div>
              </div>

              <div className="border-t border-gray-100 pt-4">
                <p className="text-xs font-semibold text-gray-500 uppercase tracking-wide mb-2">Include table data</p>
                <div className="flex gap-4 mb-3">
                  <label className="flex items-center gap-1.5 text-xs text-gray-700 cursor-pointer">
                    <input type="radio" id="cxoLimitedRows" name="cxoTableRows" value="limited" defaultChecked className="accent-blue-600" />
                    First 20 rows
                  </label>
                  <label className="flex items-center gap-1.5 text-xs text-gray-700 cursor-pointer">
                    <input type="radio" id="cxoAllRows" name="cxoTableRows" value="all" className="accent-blue-600" />
                    All rows
                  </label>
                </div>
                <div className="grid grid-cols-3 gap-1.5">
                  <button
                    onClick={() => {
                      const el = document.querySelector('input[name="cxoTableRows"]:checked') as HTMLInputElement | null;
                      const opt = el ? el.value : 'limited';
                      setShowDownloadModal(false);
                      downloadPPT(true, opt);
                    }}
                    className="py-1.5 px-1 text-xs bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 transition-colors"
                  >
                    Download
                  </button>
                  <button
                    onClick={() => {
                      const el = document.querySelector('input[name="cxoTableRows"]:checked') as HTMLInputElement | null;
                      const opt = el ? el.value : 'limited';
                      setEmailData(prev => ({ ...prev, reportType: 'complete', tableOption: opt }));
                      setShowEmailModal(true);
                    }}
                    className="py-1.5 px-1 text-xs bg-green-600 text-white rounded-lg font-medium hover:bg-green-700 transition-colors"
                  >
                    Email
                  </button>
                  <button
                    onClick={() => {
                      const el = document.querySelector('input[name="cxoTableRows"]:checked') as HTMLInputElement | null;
                      const opt = el ? el.value : 'limited';
                      setWhatsappReportType('complete');
                      setWhatsappTableOption(opt);
                      setShowWhatsappModal(true);
                    }}
                    className="py-1.5 px-1 text-xs bg-emerald-500 text-white rounded-lg font-medium hover:bg-emerald-600 transition-colors flex items-center justify-center gap-1"
                  >
                    <svg className="w-3 h-3 flex-shrink-0" viewBox="0 0 24 24" fill="currentColor"><path d="M12.04 2C6.58 2 2.13 6.45 2.13 11.91c0 1.75.46 3.45 1.32 4.95L2.05 22l5.25-1.38a9.9 9.9 0 004.74 1.21h.01c5.46 0 9.9-4.45 9.9-9.91 0-2.65-1.03-5.14-2.9-7.01A9.82 9.82 0 0012.04 2m0 1.67a8.23 8.23 0 018.24 8.24c0 4.54-3.7 8.23-8.24 8.23a8.2 8.2 0 01-4.19-1.15l-.3-.17-3.12.82.83-3.04-.2-.32a8.2 8.2 0 01-1.26-4.37c0-4.54 3.7-8.24 8.24-8.24m-4.52 4.31c-.16 0-.42.06-.64.3-.22.24-.85.83-.85 2.02s.87 2.35.99 2.51c.12.16 1.7 2.7 4.2 3.68 2.08.82 2.5.66 2.95.62.45-.04 1.46-.6 1.66-1.18.2-.58.2-1.07.14-1.18-.06-.1-.22-.16-.46-.28-.24-.12-1.46-.72-1.68-.8-.23-.08-.39-.12-.56.12-.16.24-.63.8-.78.96-.14.16-.28.18-.52.06-.24-.12-1.02-.38-1.94-1.2-.72-.64-1.2-1.43-1.34-1.67-.14-.24-.02-.37.11-.49.11-.11.24-.28.36-.42.12-.14.16-.24.24-.4.08-.16.04-.3-.02-.42-.06-.12-.56-1.35-.77-1.85-.2-.48-.4-.42-.56-.43z"/></svg>
                    WhatsApp
                  </button>
                </div>
              </div>
            </div>

            <button
              onClick={() => setShowDownloadModal(false)}
              className="w-full py-2.5 text-xs font-medium text-gray-600 bg-gray-50 hover:bg-gray-100 border-t border-gray-100 transition-colors"
            >
              Cancel
            </button>
          </div>
        </div>
      )}

      {/* Email Modal */}
      {showEmailModal && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-[250] p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full overflow-hidden">
            <div className="flex items-center justify-between px-5 py-3.5 border-b border-gray-100 bg-gradient-to-r from-green-50 to-emerald-50">
              <h3 className="text-base font-bold text-green-700">Send Report via Email</h3>
              <button
                onClick={() => {
                  if (isSendingEmail) return;
                  setShowEmailModal(false);
                  setEmailData({ email: '', subject: '', message: '', tableOption: 'limited', reportType: '' });
                }}
                className="p-1 text-gray-400 hover:text-gray-600 hover:bg-white rounded-full transition-colors"
              >
                <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                </svg>
              </button>
            </div>

            <div className="p-5">
              <div className="mb-3 px-3 py-2 bg-blue-50 border-l-4 border-blue-400 rounded-lg">
                <p className="text-xs text-blue-800">
                  <strong>Report:</strong> {emailData.reportType === 'charts-only' ? 'Charts Only' : 'Complete Report'}
                  {emailData.reportType === 'complete' && (
                    <> · <strong>Table:</strong> {emailData.tableOption === 'all' ? 'All rows' : 'First 20 rows'}</>
                  )}
                </p>
              </div>

              <form className="space-y-3" onSubmit={(e) => e.preventDefault()}>
                <div>
                  <label htmlFor="cxoRecipientEmail" className="block text-xs font-medium text-gray-700 mb-1">
                    Recipient Email *
                  </label>
                  <input
                    type="email"
                    id="cxoRecipientEmail"
                    value={emailData.email}
                    onChange={(e) => setEmailData(prev => ({ ...prev, email: e.target.value }))}
                    disabled={isSendingEmail}
                    className="w-full px-3 py-2 text-sm border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-green-500 focus:border-transparent disabled:bg-gray-50"
                    placeholder="recipient@example.com"
                    required
                  />
                </div>

                <div>
                  <label htmlFor="cxoEmailSubject" className="block text-xs font-medium text-gray-700 mb-1">
                    Subject
                  </label>
                  <input
                    type="text"
                    id="cxoEmailSubject"
                    value={emailData.subject}
                    onChange={(e) => setEmailData(prev => ({ ...prev, subject: e.target.value }))}
                    disabled={isSendingEmail}
                    className="w-full px-3 py-2 text-sm border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-green-500 focus:border-transparent disabled:bg-gray-50"
                    placeholder="Data Analysis Report"
                  />
                </div>

                <div>
                  <label htmlFor="cxoEmailMessage" className="block text-xs font-medium text-gray-700 mb-1">
                    Message (Optional)
                  </label>
                  <textarea
                    id="cxoEmailMessage"
                    value={emailData.message}
                    onChange={(e) => setEmailData(prev => ({ ...prev, message: e.target.value }))}
                    rows={2}
                    disabled={isSendingEmail}
                    className="w-full px-3 py-2 text-sm border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-green-500 focus:border-transparent disabled:bg-gray-50"
                    placeholder="Enter any additional message..."
                  />
                </div>

                <p className="text-[11px] text-gray-400">The report is generated and attached automatically — nothing to attach yourself.</p>

                <div className="flex gap-2 pt-1">
                  <button
                    type="button"
                    onClick={() => {
                      if (!emailData.email) {
                        showToast('Please enter a recipient email address');
                        return;
                      }
                      const includeTable = emailData.reportType === 'complete';
                      const tableOption = emailData.tableOption || 'limited';
                      sendViaEmail(includeTable, tableOption);
                    }}
                    disabled={isSendingEmail}
                    className="flex-1 py-2 text-sm bg-green-600 text-white rounded-lg font-medium hover:bg-green-700 disabled:opacity-60 disabled:cursor-not-allowed transition-colors flex items-center justify-center gap-1.5"
                  >
                    {isSendingEmail ? (<><div className="animate-spin rounded-full h-3.5 w-3.5 border-2 border-white border-t-transparent" />Sending...</>) : 'Send Email'}
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setShowEmailModal(false);
                      setEmailData({ email: '', subject: '', message: '', tableOption: 'limited', reportType: '' });
                    }}
                    disabled={isSendingEmail}
                    className="flex-1 py-2 text-sm bg-gray-100 text-gray-700 rounded-lg hover:bg-gray-200 disabled:opacity-60 transition-colors"
                  >
                    Cancel
                  </button>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}

      {/* WhatsApp Modal */}
      {showWhatsappModal && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-[250] p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-sm w-full overflow-hidden">
            <div className="flex items-center justify-between px-5 py-3.5 border-b border-gray-100 bg-gradient-to-r from-emerald-50 to-teal-50">
              <h3 className="text-base font-bold text-emerald-700 flex items-center gap-1.5">
                <svg className="w-4 h-4 flex-shrink-0" viewBox="0 0 24 24" fill="currentColor"><path d="M12.04 2C6.58 2 2.13 6.45 2.13 11.91c0 1.75.46 3.45 1.32 4.95L2.05 22l5.25-1.38a9.9 9.9 0 004.74 1.21h.01c5.46 0 9.9-4.45 9.9-9.91 0-2.65-1.03-5.14-2.9-7.01A9.82 9.82 0 0012.04 2m0 1.67a8.23 8.23 0 018.24 8.24c0 4.54-3.7 8.23-8.24 8.23a8.2 8.2 0 01-4.19-1.15l-.3-.17-3.12.82.83-3.04-.2-.32a8.2 8.2 0 01-1.26-4.37c0-4.54 3.7-8.24 8.24-8.24m-4.52 4.31c-.16 0-.42.06-.64.3-.22.24-.85.83-.85 2.02s.87 2.35.99 2.51c.12.16 1.7 2.7 4.2 3.68 2.08.82 2.5.66 2.95.62.45-.04 1.46-.6 1.66-1.18.2-.58.2-1.07.14-1.18-.06-.1-.22-.16-.46-.28-.24-.12-1.46-.72-1.68-.8-.23-.08-.39-.12-.56.12-.16.24-.63.8-.78.96-.14.16-.28.18-.52.06-.24-.12-1.02-.38-1.94-1.2-.72-.64-1.2-1.43-1.34-1.67-.14-.24-.02-.37.11-.49.11-.11.24-.28.36-.42.12-.14.16-.24.24-.4.08-.16.04-.3-.02-.42-.06-.12-.56-1.35-.77-1.85-.2-.48-.4-.42-.56-.43z"/></svg>
                Send via WhatsApp
              </h3>
              <button
                onClick={() => {
                  if (isSendingWhatsapp) return;
                  setShowWhatsappModal(false);
                  setWhatsappPhone('');
                }}
                className="p-1 text-gray-400 hover:text-gray-600 hover:bg-white rounded-full transition-colors"
              >
                <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                </svg>
              </button>
            </div>

            <div className="p-5">
              <div className="mb-3 px-3 py-2 bg-emerald-50 border-l-4 border-emerald-400 rounded-lg">
                <p className="text-xs text-emerald-800">
                  <strong>Report:</strong> {whatsappReportType === 'complete' ? 'Complete Report' : 'Charts Only'}
                  {whatsappReportType === 'complete' && (
                    <> · <strong>Table:</strong> {whatsappTableOption === 'all' ? 'All rows' : 'First 20 rows'}</>
                  )}
                </p>
              </div>

              <label htmlFor="cxoWhatsappPhone" className="block text-xs font-medium text-gray-700 mb-1">
                WhatsApp Number *
              </label>
              <input
                type="tel"
                id="cxoWhatsappPhone"
                value={whatsappPhone}
                onChange={(e) => setWhatsappPhone(e.target.value)}
                disabled={isSendingWhatsapp}
                className="w-full px-3 py-2 text-sm border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent disabled:bg-gray-50 mb-1"
                placeholder="e.g. 919876543210"
              />
              <p className="text-[11px] text-gray-400 mb-4">Include the country code, digits only.</p>

              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={sendViaWhatsapp}
                  disabled={isSendingWhatsapp}
                  className="flex-1 py-2 text-sm bg-emerald-600 text-white rounded-lg font-medium hover:bg-emerald-700 disabled:opacity-60 disabled:cursor-not-allowed transition-colors flex items-center justify-center gap-1.5"
                >
                  {isSendingWhatsapp ? (<><div className="animate-spin rounded-full h-3.5 w-3.5 border-2 border-white border-t-transparent" />Sending...</>) : 'Send'}
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setShowWhatsappModal(false);
                    setWhatsappPhone('');
                  }}
                  disabled={isSendingWhatsapp}
                  className="flex-1 py-2 text-sm bg-gray-100 text-gray-700 rounded-lg hover:bg-gray-200 disabled:opacity-60 transition-colors"
                >
                  Cancel
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Toast notifications */}
      <div className="fixed bottom-6 left-1/2 -translate-x-1/2 flex flex-col gap-2 z-[200] pointer-events-none">
        {toasts.map(t => (
          <div key={t.id} className={`pointer-events-auto flex items-center gap-3 px-4 py-3 rounded-lg shadow-lg text-sm font-medium text-white animate-fade-in ${t.type === 'error' ? 'bg-red-600' : 'bg-blue-600'}`}>
            <span>{t.type === 'error' ? '✕' : 'ℹ'}</span>
            <span>{t.message}</span>
          </div>
        ))}
      </div>

    </div>
    
  );
}