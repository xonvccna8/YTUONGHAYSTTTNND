import { useState, useRef, useEffect } from 'react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { Loader2, Sparkles, RefreshCw, GitCompare, Menu, X, Lightbulb, Target, Settings, BookOpen, Layers, History, Trash2, Eye, Clock, Calendar, Heart, Lock, Key, Download, ChevronRight } from 'lucide-react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { motion, AnimatePresence } from 'motion/react';
import { saveAs } from 'file-saver';
import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeadingLevel,
  Packer,
  PageBreak,
  PageNumber,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';

declare global {
  interface Window {
    aistudio?: {
      hasSelectedApiKey: () => Promise<boolean>;
      openSelectKey: () => Promise<void>;
    };
  }
}

// --- CẤU HÌNH PHIÊN BẢN THỬ NGHIỆM (TESTING MODE) ---
const TEST_CONFIG = {
  // 1. Công tắc chính: Đổi thành 'false' để khóa app ngay lập tức
  isActive: true, 
  
  // 2. Tự động hết hạn: Đặt ngày giờ hết hạn (VD: '2026-03-30T23:59:59'). Để null nếu không dùng.
  expireAt: null, // '2026-03-30T23:59:59',
  
  // 3. Mật khẩu truy cập: Bật 'true' để yêu cầu nhập mã.
  requirePasscode: true,
  passcode: 'ideagpt2026', // Mã truy cập gửi cho tester
};
// ---------------------------------------------------

const IDEA_MEMORY_KEY = 'ideagpt_idea_memory';
const IDEA_MEMORY_LIMIT = 320;
const IDEA_MEMORY_PROMPT_LIMIT = 120;

interface IdeaMemoryEntry {
  title: string;
  signature: string;
  timestamp: number;
  source?: string;
  field?: string;
  grade?: string;
}

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

function normalizeIdeaSignature(value: string) {
  return value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function cleanIdeaTitle(value: string) {
  return value
    .replace(/\*\*/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function extractIdeaMemoryEntries(
  content: string,
  metadata: Partial<Pick<IdeaMemoryEntry, 'source' | 'field' | 'grade'>> = {}
): IdeaMemoryEntry[] {
  const matches = [...content.matchAll(/^###\s*(?:💡\s*)?Ý TƯỞNG\s+\d+\s*:\s*(.+)$/gmi)];
  const timestamp = Date.now();

  return matches
    .map(match => cleanIdeaTitle(match[1] || ''))
    .filter(title => title.length > 0)
    .map(title => ({
      title,
      signature: normalizeIdeaSignature(title),
      timestamp,
      ...metadata,
    }))
    .filter(entry => entry.signature.length > 0);
}

function readIdeaMemory(): IdeaMemoryEntry[] {
  try {
    const saved = localStorage.getItem(IDEA_MEMORY_KEY);
    const parsed = saved ? JSON.parse(saved) : [];
    if (!Array.isArray(parsed)) return [];

    return parsed
      .filter((entry): entry is IdeaMemoryEntry =>
        Boolean(entry?.title && entry?.signature && typeof entry.title === 'string' && typeof entry.signature === 'string')
      )
      .slice(0, IDEA_MEMORY_LIMIT);
  } catch {
    return [];
  }
}

function writeIdeaMemory(entries: IdeaMemoryEntry[]) {
  const unique = new Map<string, IdeaMemoryEntry>();

  for (const entry of entries) {
    if (!entry.signature || unique.has(entry.signature)) continue;
    unique.set(entry.signature, entry);
  }

  localStorage.setItem(IDEA_MEMORY_KEY, JSON.stringify([...unique.values()].slice(0, IDEA_MEMORY_LIMIT)));
}

function clearIdeaMemory() {
  localStorage.removeItem(IDEA_MEMORY_KEY);
}

async function readApiJson(response: Response) {
  const text = await response.text();
  if (!text) return {};

  try {
    return JSON.parse(text);
  } catch {
    throw new Error(`Máy chủ trả về dữ liệu không hợp lệ (HTTP ${response.status}). Vui lòng tải lại trang.`);
  }
}

const FIELD_OPTIONS = [
  'Đồ dùng học tập',
  'Phần mềm tin học',
  'Sản phẩm thân thiện môi trường',
  'Dụng cụ sinh hoạt gia đình & đồ chơi',
  'Giải pháp kỹ thuật liên quan cuộc sống và bảo vệ môi trường',
];

const CAP_HOC_OPTIONS = ['Tiểu học', 'THCS', 'THPT'];

const GRADE_OPTIONS = [
  'Tiểu học: Lớp 1–5',
  'THCS: Lớp 6–9',
  'THPT: Lớp 10–12',
];

const TECH_LIMIT_OPTIONS = ['Cơ bản', 'Trung bình', 'Nâng cao'];

const MUC_TIEU_OPTIONS = [
  'Cấp lớp',
  'Cấp trường',
  'Cấp huyện',
  'Cấp tỉnh',
  'Cấp quốc gia',
];

interface SavedSession {
  id: string;
  timestamp: number;
  inputs: {
    field: string;
    capHoc: string;
    grade: string;
    techLimit: string;
    mucTieu: string;
    context: string;
    resources: string;
  };
  result: string;
  compareResult?: string;
  mode?: 'basic' | 'advanced';
}

interface SavedIdea {
  id: string;
  timestamp: number;
  title: string;
  content: string;
  inputs: {
    field: string;
    capHoc: string;
    grade: string;
  };
}

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState(!TEST_CONFIG.requirePasscode);
  const [passcodeInput, setPasscodeInput] = useState('');
  const [passcodeError, setPasscodeError] = useState(false);
  const [passcodeErrorMessage, setPasscodeErrorMessage] = useState('');
  const [isCheckingPasscode, setIsCheckingPasscode] = useState(false);
  const [expiresAt, setExpiresAt] = useState<number | null>(null);

  const getDeviceId = () => {
    let id = localStorage.getItem('device_id');
    if (!id) {
      id = crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).substring(2, 15);
      localStorage.setItem('device_id', id);
    }
    return id;
  };

  useEffect(() => {
    if (TEST_CONFIG.requirePasscode) {
      const savedPasscode = localStorage.getItem('savedPasscode');
      if (savedPasscode) {
        verifyPasscode(savedPasscode, true);
      }
    }
  }, []);

  useEffect(() => {
    if (expiresAt) {
      const interval = setInterval(() => {
        if (Date.now() > expiresAt) {
          handleLogout(true);
        }
      }, 1000);
      return () => clearInterval(interval);
    }
  }, [expiresAt]);

  const handleLogout = (isExpired = false) => {
    setIsAuthenticated(false);
    localStorage.removeItem('savedPasscode');
    setExpiresAt(null);
    if (isExpired) {
      setPasscodeError(true);
      setPasscodeErrorMessage('Thời gian sử dụng 30 phút đã hết. Vui lòng liên hệ Admin để nhận mã dự phòng.');
    }
  };

  const verifyPasscode = async (code: string, isAutoLogin = false) => {
    setIsCheckingPasscode(true);
    setPasscodeError(false);
    setPasscodeErrorMessage('');
    
    try {
      const response = await fetch('/api/verify-passcode', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ passcode: code, deviceId: getDeviceId() })
      });
      
      const data = await readApiJson(response);
      
      if (data.success) {
        setIsAuthenticated(true);
        setExpiresAt(data.expiresAt);
        localStorage.setItem('savedPasscode', code);
      } else {
        if (!isAutoLogin) {
          setPasscodeError(true);
          setPasscodeErrorMessage(data.message || 'Mã truy cập không hợp lệ.');
        } else {
          // Auto login failed, clear saved passcode
          handleLogout(data.isExpired);
        }
      }
    } catch (error) {
      console.error('Error verifying passcode:', error);
      if (!isAutoLogin) {
        setPasscodeError(true);
        setPasscodeErrorMessage('Lỗi kết nối máy chủ. Vui lòng thử lại.');
      }
    } finally {
      setIsCheckingPasscode(false);
    }
  };

  const [field, setField] = useState(FIELD_OPTIONS[0]);
  const [capHoc, setCapHoc] = useState(CAP_HOC_OPTIONS[0]);
  const [grade, setGrade] = useState(GRADE_OPTIONS[0]);
  const [techLimit, setTechLimit] = useState(TECH_LIMIT_OPTIONS[0]);
  const [mucTieu, setMucTieu] = useState(MUC_TIEU_OPTIONS[0]);
  const [context, setContext] = useState('');
  const [resources, setResources] = useState('');

  const [isGenerating, setIsGenerating] = useState(false);
  const [loadingMessage, setLoadingMessage] = useState('Đang khởi tạo IdeaGPT...');
  const [result, setResult] = useState('');
  const [generationMode, setGenerationMode] = useState<'basic' | 'advanced'>('basic');
  const [advancedModel, setAdvancedModel] = useState<'gpt' | 'deepseek'>('gpt');
  const [compareResult, setCompareResult] = useState('');
  const [isComparing, setIsComparing] = useState(false);
  const [inlineComparisons, setInlineComparisons] = useState<Record<string, string>>({});
  const [inlineEnhancements, setInlineEnhancements] = useState<Record<string, string>>({});
  const [inlineDetailedGuides, setInlineDetailedGuides] = useState<Record<string, string>>({});
  const [loadingInline, setLoadingInline] = useState<Record<string, 'comparing' | 'enhancing' | 'detailing'>>({});
  const [activeTab, setActiveTab] = useState<'main' | 'compare' | 'history' | 'favorites'>('main');
  const [sidebarOpen, setSidebarOpen] = useState(true);

  const [history, setHistory] = useState<SavedSession[]>(() => {
    try {
      const saved = localStorage.getItem('ideagpt_history');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });
  const [savedIdeas, setSavedIdeas] = useState<SavedIdea[]>(() => {
    try {
      const saved = localStorage.getItem('ideagpt_saved_ideas');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });
  const [currentSessionId, setCurrentSessionId] = useState<string | null>(null);
  const [showClearConfirm, setShowClearConfirm] = useState(false);
  const [showAdvancedModal, setShowAdvancedModal] = useState(false);

  const resultEndRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    localStorage.setItem('ideagpt_history', JSON.stringify(history));
  }, [history]);

  useEffect(() => {
    localStorage.setItem('ideagpt_saved_ideas', JSON.stringify(savedIdeas));
  }, [savedIdeas]);

  const extractIdea = (fullText: string, headingText: string) => {
    const lines = fullText.split('\n');
    let startIndex = -1;
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].startsWith('###') && lines[i].includes(headingText.trim())) {
        startIndex = i;
        break;
      }
    }
    if (startIndex === -1) return '';

    let endIndex = lines.length;
    for (let i = startIndex + 1; i < lines.length; i++) {
      if (lines[i].startsWith('## ') || lines[i].startsWith('### ')) {
        endIndex = i;
        break;
      }
    }
    return lines.slice(startIndex, endIndex).join('\n');
  };

  const getIdeaMemorySnapshot = () => {
    const allEntries: IdeaMemoryEntry[] = [
      ...readIdeaMemory(),
      ...history.flatMap(session =>
        extractIdeaMemoryEntries(session.result, {
          source: 'Lịch sử',
          field: session.inputs.field,
          grade: session.inputs.grade,
        })
      ),
      ...savedIdeas.flatMap(idea =>
        extractIdeaMemoryEntries(idea.content, {
          source: 'Đã lưu',
          field: idea.inputs.field,
          grade: idea.inputs.grade,
        })
      ),
    ];
    const unique = new Map<string, IdeaMemoryEntry>();

    for (const entry of allEntries) {
      if (!entry.signature || unique.has(entry.signature)) continue;
      unique.set(entry.signature, entry);
    }

    return [...unique.values()]
      .sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0))
      .slice(0, IDEA_MEMORY_LIMIT);
  };

  const buildIdeaExclusionPrompt = () => {
    const entries = getIdeaMemorySnapshot().slice(0, IDEA_MEMORY_PROMPT_LIMIT);
    if (entries.length === 0) {
      return 'Chưa có ý tưởng cũ trong bộ nhớ. Vẫn phải tự tạo các ý tưởng khác biệt nhau trong cùng một lần sinh.';
    }

    return entries
      .map((entry, index) => `${index + 1}. ${entry.title}`)
      .join('\n');
  };

  const persistIdeaMemoryFromResult = (content: string) => {
    const newEntries = extractIdeaMemoryEntries(content, {
      source: 'Đã tạo',
      field,
      grade,
    });

    if (newEntries.length === 0) return;
    writeIdeaMemory([...newEntries, ...getIdeaMemorySnapshot()]);
  };

  const generateIdeas = async (isReroll = false) => {
    setIsGenerating(true);
    setResult('');
    setInlineComparisons({});
    setInlineEnhancements({});
    setInlineDetailedGuides({});
    setLoadingInline({});
    setLoadingMessage('Đang phân tích yêu cầu và bối cảnh...');
    setActiveTab('main');
    if (window.innerWidth < 1024) {
      setSidebarOpen(false);
    }

    // Simulated deep thinking phases
    const loadingSteps = [
      'Đang phân tích bối cảnh và các vấn đề nhức nhối trong thực tế...',
      'Đang đối chiếu với kho ý tưởng đã có để tránh trùng lặp...',
      'Đang tìm kiếm các giải pháp đột phá, đảm bảo Tính mới và Tính sáng tạo...',
      'Đang đánh giá Tính khả thi và Tính bền vững cho học sinh...',
      'Đang tinh chỉnh và chọn lọc ra 20 ý tưởng xuất sắc nhất...',
      'Đang hoàn thiện báo cáo phân tích chuyên sâu...'
    ];

    let stepIndex = 0;
    const loadingInterval = setInterval(() => {
      if (stepIndex < loadingSteps.length) {
        setLoadingMessage(loadingSteps[stepIndex]);
        stepIndex++;
      }
    }, 2500);

    // Add a simulated delay to give the impression of deep research
    await new Promise(resolve => setTimeout(resolve, 8000));

    try {
      const ideaExclusionList = buildIdeaExclusionPrompt();
      const noveltySeed = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
      const prompt = `
Bạn là một Chuyên gia AI hàng đầu về Đổi mới Sáng tạo và là Giám khảo cấp quốc gia của "Cuộc thi Sáng tạo Thanh thiếu niên Nhi đồng" tại Việt Nam.
Nhiệm vụ của bạn là tạo ra ĐÚNG 20 Ý TƯỞNG xuất sắc nhất. Các ý tưởng phải có tính mới, tính sáng tạo cao, tính khả thi, tính bền vững cao và đặc biệt phải có ý nghĩa lớn trong cuộc sống học tập.

YÊU CẦU TỐI QUAN TRỌNG (THINKING PROCESS):
1. Phân tích kỹ nhưng KHÔNG trình bày suy luận nội bộ dài dòng; chỉ viết phần tóm tắt định hướng cần thiết.
2. TÌM KIẾM SỰ ĐỘT PHÁ: Tuyệt đối KHÔNG đề xuất các ý tưởng cũ rích, sáo mòn. Phải tìm và so sánh với những dự án/sản phẩm đã làm trước đây, từ đó đưa ra giải pháp sáng tạo và hữu ích hơn nhiều.
3. TÍNH MỚI & SÁNG TẠO CAO: Ý tưởng phải thực sự độc đáo, chưa từng có ai làm hoặc áp dụng một góc nhìn/công nghệ hoàn toàn mới vào một vấn đề cũ.
4. ĐƠN GIẢN NHƯNG THỰC TẾ: Giải pháp cần đơn giản nhưng có tính ứng dụng thực tế cao, giải quyết đúng "nỗi đau" trong học tập và đời sống.
5. PHÙ HỢP LỨA TUỔI: Đảm bảo tính khả thi cho học sinh ${grade} với giới hạn công nghệ là ${techLimit}. Học sinh phải vận dụng được các kiến thức đã học trên lớp (Toán, Lý, Hóa, Sinh, Tin học...) để nghiên cứu và thực hiện dự án.

THÔNG TIN ĐẦU VÀO:
- Lĩnh vực: ${field}
- Cấp học: ${capHoc} (Lớp: ${grade})
- Giới hạn công nghệ: ${techLimit}
- Mục tiêu: ${mucTieu}
- Bối cảnh: ${context || 'Không có'}
- Nguồn lực: ${resources || 'Không có'}

MÃ PHIÊN SÁNG TẠO: ${noveltySeed}
Hãy dùng mã phiên này để mở một nhánh tư duy mới. Nếu người dùng bấm tạo nhiều lần với cùng thông tin đầu vào, kết quả lần sau vẫn phải khác hẳn lần trước.

KHO Ý TƯỞNG ĐÃ CÓ TRÊN MÁY NGƯỜI DÙNG (DANH SÁCH CẦN TRÁNH):
${ideaExclusionList}

QUY TẮC CHỐNG TRÙNG LẶP BẮT BUỘC:
- Tuyệt đối KHÔNG dùng lại tên ý tưởng, vấn đề chính, cơ chế hoạt động, tính năng nổi bật hoặc cách làm cốt lõi của danh sách đã có.
- Không được chỉ đổi vật liệu, đổi địa điểm, đổi cấp học hoặc đổi tên cho một ý tưởng cũ. Phải đổi cả "nỗi đau", đối tượng phục vụ, cơ chế giải pháp và sản phẩm đầu ra.
- Trong 20 ý tưởng của lần này cũng không được trùng nhau. Mỗi ý tưởng phải đi theo một hướng công nghệ/cơ chế/đối tượng ứng dụng khác biệt rõ.
- Nếu một hướng có nguy cơ giống ý tưởng cũ, hãy tự loại bỏ và thay bằng hướng mới lạ hơn.

${isReroll ? 'YÊU CẦU ĐẶC BIỆT: Đây là lần [TÌM LẠI]. Hãy tạo 20 ý tưởng MỚI HOÀN TOÀN, không trùng lặp với lần trước.' : ''}

TIÊU CHÍ ĐÁNH GIÁ CỐT LÕI (BẮT BUỘC ĐÁP ỨNG CHO CẢ 20 Ý TƯỞNG):
1. Tính mới: Chưa từng có trên thị trường hoặc có cải tiến vượt bậc so với các giải pháp cũ.
2. Tính sáng tạo: Cách tiếp cận độc đáo, thông minh, có yếu tố "WOW".
3. Tính khả thi: Phù hợp với năng lực học sinh (${grade}), học sinh có thể vận dụng kiến thức đã học để làm mô hình thực tế, vật liệu dễ tìm, an toàn tuyệt đối. Đơn giản nhưng hiệu quả cao.
4. Tính bền vững: Thân thiện môi trường, có khả năng nhân rộng, chi phí hợp lý.
5. Ý nghĩa trong cuộc sống học tập: Giải quyết một vấn đề thực tế, bức xúc trong học tập hoặc đời sống học đường, mang lại giá trị thiết thực.

RÀNG BUỘC BẮT BUỘC VỀ KẾT QUẢ:
- Phải hoàn thành đủ từ ### 💡 Ý TƯỞNG 1 đến ### 💡 Ý TƯỞNG 20. Tuyệt đối không dừng ở Ý TƯỞNG 8, 10 hoặc 12.
- Nếu cần rút gọn để đủ 20 ý tưởng, hãy rút gọn từng gạch đầu dòng nhưng vẫn giữ đủ 8 mục phân tích cho mỗi ý tưởng.
- Chỉ viết TOP 3 sau khi đã trình bày xong Ý TƯỞNG 20.
- Mỗi ý tưởng phải có điểm đánh giá theo thang 100 để người dùng lọc nhanh ý tưởng mạnh nhất.
- KHÔNG viết phần "Cách làm chi tiết" trong kết quả chính. Ứng dụng sẽ có nút riêng để người dùng bấm khi cần AI hướng dẫn chi tiết.

THANG ĐIỂM CHẤM MỖI Ý TƯỞNG:
- Tính mới: 20 điểm
- Tính sáng tạo: 20 điểm
- Tính khả thi với học sinh ${grade}: 20 điểm
- Tác động thực tiễn/học tập: 20 điểm
- Bền vững, chi phí, khả năng nhân rộng: 20 điểm

HÃY XUẤT KẾT QUẢ THEO ĐÚNG ĐỊNH DẠNG MARKDOWN SAU:

## 🧠 1. QUÁ TRÌNH AI SUY NGHĨ (BRAINSTORMING)
(Trình bày chi tiết quá trình bạn tư duy: Phân tích vấn đề -> Liệt kê & loại bỏ các ý tưởng cũ rích, nhàm chán -> Động não các hướng đi đột phá -> Chốt 20 ý tưởng)

## 🎯 2. PHÂN TÍCH BỐI CẢNH & ĐỊNH HƯỚNG
(Phân tích ngắn gọn các vấn đề nhức nhối nhất hiện nay trong học tập/đời sống và định hướng giải quyết phù hợp với học sinh ${grade})

## 🚀 3. DANH SÁCH 20 Ý TƯỞNG ĐỘT PHÁ
(Trình bày chi tiết 20 ý tưởng. Mỗi ý tưởng phải phân tích sâu sắc, thuyết phục)

### 💡 Ý TƯỞNG 1: [Tên ý tưởng thật ấn tượng, rõ nghĩa]
- **Vấn đề & Ý nghĩa thực tiễn:** (Nỗi đau nào trong học tập/cuộc sống đang được giải quyết?)
- **So sánh với giải pháp cũ:** (Những cái đã làm là gì? Tại sao giải pháp này sáng tạo và hữu ích hơn nhiều?)
- **Tính năng nổi bật duy nhất:** (Nêu đúng 1 tính năng/cải tiến then chốt làm ý tưởng này khác biệt mạnh nhất)
- **Cơ chế hoạt động & Giải pháp:** (Mô tả rõ cách sản phẩm hoạt động, đơn giản nhưng thực tế cao)
- **Kiến thức vận dụng:** (Học sinh cần vận dụng kiến thức môn học nào đã học để nghiên cứu dự án này?)
- **Tính khả thi & Bền vững:** (Phân tích vật liệu, độ khó kỹ thuật, tính an toàn cho học sinh ${grade}, tác động môi trường, chi phí)
- **Điểm đánh giá:** [Tổng điểm]/100 — Tính mới [x]/20; Sáng tạo [x]/20; Khả thi [x]/20; Tác động [x]/20; Bền vững [x]/20. (1 câu giải thích điểm)
- **🛠 Cách làm ngắn gọn:** (3-5 bước chính, dễ hiểu, để học sinh nắm ngay lộ trình thực hiện)

### 💡 Ý TƯỞNG 2: [Tên ý tưởng thật ấn tượng, rõ nghĩa]
... (Tương tự như trên)

... (Tiếp tục trình bày đầy đủ đến Ý TƯỞNG 20)

### 💡 Ý TƯỞNG 20: [Tên ý tưởng thật ấn tượng, rõ nghĩa]
... (Tương tự như trên)

## 🏆 4. TOP 3 Ý TƯỞNG "CHAMPION" (KHUYÊN CHỌN NHẤT)
(Chọn ra 3 ý tưởng xuất sắc toàn diện nhất trong 20 ý tưởng trên. Giải thích lý do tại sao các ý tưởng này có khả năng đạt giải Nhất cao nhất dựa trên các tiêu chí của ban giám khảo)
      `;

      let fullText = '';
      const fetchGPT = async (promptText = prompt) => {
        const response = await fetch('/api/generate', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ 
            prompt: promptText,
            passcode: localStorage.getItem('savedPasscode'),
            deviceId: getDeviceId(),
            mode: generationMode === 'advanced' ? 'advanced-gpt' : 'basic'
          })
        });

        if (!response.ok) {
          const errorData = await response.json().catch(() => ({}));
          if (response.status === 403 && errorData.isExpired) {
            handleLogout(true);
            throw new Error('Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.');
          }
          throw new Error(errorData.error || `HTTP error! status: ${response.status}`);
        }

        const reader = response.body?.getReader();
        const decoder = new TextDecoder();
        let done = false;
        let buffer = '';

        while (!done && reader) {
          const { value, done: readerDone } = await reader.read();
          done = readerDone;
          if (value) {
            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split('\n');
            buffer = lines.pop() || '';
            
            for (const line of lines) {
              if (line.startsWith('data: ') && line !== 'data: [DONE]') {
                try {
                  const data = JSON.parse(line.slice(6));
                  if (data.error) {
                    throw new Error(data.error);
                  }
                  if (data.text) {
                    fullText += data.text;
                    setResult(fullText);
                  }
                } catch (e: any) {
                  if (e.message && e.message !== 'Unexpected end of JSON input') throw e;
                  console.error('Error parsing stream data:', e, line);
                }
              }
            }
          }
        }
      };

      const fetchDeepSeek = async (promptText = prompt) => {
        const response = await fetch('/api/generate-deepseek', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ 
            prompt: promptText,
            passcode: localStorage.getItem('savedPasscode'),
            deviceId: getDeviceId()
          })
        });

        if (!response.ok) {
          const errorData = await response.json().catch(() => ({}));
          if (response.status === 403 && errorData.isExpired) {
            handleLogout(true);
            throw new Error('Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.');
          }
          throw new Error(errorData.error || `HTTP error! status: ${response.status}`);
        }

        const reader = response.body?.getReader();
        const decoder = new TextDecoder();
        let done = false;
        let buffer = '';

        while (!done && reader) {
          const { value, done: readerDone } = await reader.read();
          done = readerDone;
          if (value) {
            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split('\n');
            buffer = lines.pop() || '';
            
            for (const line of lines) {
              if (line.startsWith('data: ') && line !== 'data: [DONE]') {
                try {
                  const data = JSON.parse(line.slice(6));
                  if (data.error) {
                    throw new Error(data.error);
                  }
                  if (data.text) {
                    fullText += data.text;
                    setResult(fullText);
                  }
                } catch (e: any) {
                  if (e.message && e.message !== 'Unexpected end of JSON input') throw e;
                  console.error('Error parsing stream data:', e, line);
                }
              }
            }
          }
        }
      };

      const getLastIdeaNumber = (text: string) => {
        const matches = [...text.matchAll(/^###\s*(?:💡\s*)?Ý TƯỞNG\s+(\d+)/gmi)];
        return matches.reduce((max, match) => Math.max(max, Number(match[1] || 0)), 0);
      };

      const runSelectedModel = async (promptText = prompt) => {
        if (generationMode === 'advanced') {
          if (advancedModel === 'gpt') {
            await fetchGPT(promptText);
          } else {
            await fetchDeepSeek(promptText);
          }
        } else {
          await fetchGPT(promptText);
        }
      };

      await runSelectedModel();

      let lastIdeaNumber = getLastIdeaNumber(fullText);
      let continuationAttempts = 0;
      const removePrematureChampionSection = (text: string) => {
        const championIndex = text.search(/^##\s*(?:🏆\s*)?4\.\s*TOP\s+3/im);
        if (championIndex === -1) return text;

        const ideasBeforeChampion = getLastIdeaNumber(text.slice(0, championIndex));
        return ideasBeforeChampion < 20 ? text.slice(0, championIndex).trimEnd() : text;
      };

      while (lastIdeaNumber > 0 && lastIdeaNumber < 20 && continuationAttempts < 3) {
        continuationAttempts++;
        fullText = removePrematureChampionSection(fullText);
        setResult(fullText);
        lastIdeaNumber = getLastIdeaNumber(fullText);

        const nextIdeaNumber = lastIdeaNumber + 1;
        setLoadingMessage(`Đang viết tiếp từ Ý tưởng ${nextIdeaNumber} đến Ý tưởng 20...`);
        const continuationPrompt = `
Bạn đang viết tiếp một báo cáo ý tưởng sáng tạo. Kết quả trước mới hoàn thành đến Ý TƯỞNG ${lastIdeaNumber}.

NHIỆM VỤ BẮT BUỘC:
1. KHÔNG viết lại các ý tưởng đã có.
2. Bắt đầu chính xác bằng heading: ### 💡 Ý TƯỞNG ${nextIdeaNumber}: [Tên ý tưởng]
3. Viết tiếp đầy đủ đến ### 💡 Ý TƯỞNG 20.
4. Sau Ý TƯỞNG 20, viết mục ## 🏆 4. TOP 3 Ý TƯỞNG "CHAMPION" (KHUYÊN CHỌN NHẤT).
5. Giữ đúng cấu trúc 8 gạch đầu dòng cho mỗi ý tưởng: Vấn đề & Ý nghĩa thực tiễn; So sánh với giải pháp cũ; Tính năng nổi bật duy nhất; Cơ chế hoạt động & Giải pháp; Kiến thức vận dụng; Tính khả thi & Bền vững; Điểm đánh giá; Cách làm ngắn gọn.
6. Mọi ý tưởng viết tiếp phải khác hoàn toàn các ý tưởng trong bộ nhớ cũ và khác các ý tưởng đã có ở phần trước.
7. KHÔNG viết phần "Cách làm chi tiết"; phần này chỉ sinh khi người dùng bấm nút riêng.

KHO Ý TƯỞNG ĐÃ CÓ TRÊN MÁY NGƯỜI DÙNG (DANH SÁCH CẦN TRÁNH):
${ideaExclusionList}

THÔNG TIN ĐẦU VÀO:
- Lĩnh vực: ${field}
- Cấp học: ${capHoc} (Lớp: ${grade})
- Giới hạn công nghệ: ${techLimit}
- Mục tiêu: ${mucTieu}
- Bối cảnh: ${context || 'Không có'}
- Nguồn lực: ${resources || 'Không có'}

PHẦN CUỐI KẾT QUẢ TRƯỚC ĐỂ TRÁNH TRÙNG LẶP:
${fullText.slice(-6000)}
        `;

        fullText += '\n\n';
        setResult(fullText);
        const lengthBeforeContinuation = fullText.length;
        await runSelectedModel(continuationPrompt);
        const updatedLastIdeaNumber = getLastIdeaNumber(fullText);
        if (updatedLastIdeaNumber <= lastIdeaNumber || fullText.length === lengthBeforeContinuation) {
          break;
        }
        lastIdeaNumber = updatedLastIdeaNumber;
      }

      if (getLastIdeaNumber(fullText) < 20) {
        fullText += `\n\n> ⚠️ Kết quả hiện chưa đủ 20 ý tưởng do mô hình dừng sớm. Hãy bấm "Tìm Lại" hoặc chọn chế độ nâng cao khác để tạo lại danh sách đầy đủ.`;
        setResult(fullText);
      }

      persistIdeaMemoryFromResult(fullText);

      const newId = Date.now().toString();
      setCurrentSessionId(newId);
      setHistory(prev => [{
        id: newId,
        timestamp: Date.now(),
        inputs: { field, capHoc, grade, techLimit, mucTieu, context, resources },
        result: fullText,
        mode: generationMode
      }, ...prev]);

    } catch (error: any) {
      console.error('Error generating ideas:', error);
      if (error.message?.includes('429') || error.message?.includes('insufficient_quota')) {
        setResult('🚨 **Lỗi Quá Tải Hệ Thống (Quota Exceeded)**\n\nTài khoản OpenAI của bạn đã hết hạn mức sử dụng hoặc hệ thống đang quá tải. Vui lòng kiểm tra lại Billing trên OpenAI.');
      } else {
        setResult(`Đã xảy ra lỗi khi tạo ý tưởng: ${error.message || 'Vui lòng thử lại.'}`);
      }
    } finally {
      clearInterval(loadingInterval);
      setLoadingMessage('Đang khởi tạo IdeaGPT...');
      setIsGenerating(false);
    }
  };

  const handleInlineCompare = async (title: string, sectionContent: string) => {
    setLoadingInline(prev => ({ ...prev, [title]: 'comparing' }));
    setInlineComparisons(prev => ({ ...prev, [title]: '' }));

    const prompt = `Bạn là chuyên gia đánh giá dự án khoa học kỹ thuật và khởi nghiệp. Hãy phân tích và so sánh ý tưởng sau đây với các sản phẩm/giải pháp ĐÃ CÓ TRÊN THỊ TRƯỜNG hoặc TRÊN MẠNG.
Hãy phân tích bằng DeepSeek V4 Pro. Nếu không có dữ liệu chắc chắn, hãy nói rõ mức độ tin cậy thay vì bịa nguồn.

Ý TƯỞNG CẦN ĐÁNH GIÁ:
${sectionContent}

YÊU CẦU PHÂN TÍCH (Đóng vai trò chuyên gia DeepSeek V4 Pro để phân tích sâu sắc):
1. ĐỐI CHIẾU THỰC TẾ: Chỉ ra đích danh 2-3 sản phẩm/dự án tương tự đã có trên thực tế nếu bạn biết chắc. Kèm link tham khảo khi chắc chắn, không bịa nguồn.
2. SO SÁNH ĐIỂM GIỐNG & KHÁC: Phân tích điểm giống và khác biệt cốt lõi giữa ý tưởng này và các sản phẩm đã có.
3. ĐÁNH GIÁ TÍNH MỚI: Đánh giá khách quan xem ý tưởng này có thực sự "chưa ai làm" không? Điểm nào là cải tiến ĐÁNG GIÁ NHẤT và SÁNG TẠO NHẤT so với cái cũ?
4. TÍNH ỨNG DỤNG: Đánh giá tính ứng dụng thực tế đối với học sinh lớp ${grade}.

Trình bày bằng Markdown, ngắn gọn, súc tích, chuyên nghiệp và khách quan.`;

    try {
      const response = await fetch('/api/generate-deepseek', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          prompt,
          passcode: localStorage.getItem('savedPasscode'),
          deviceId: getDeviceId()
        })
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || 'Failed to compare');
      }

      const reader = response.body?.getReader();
      const decoder = new TextDecoder();
      let done = false;
      let buffer = '';
      let fullText = '';

      while (!done && reader) {
        const { value, done: readerDone } = await reader.read();
        done = readerDone;
        if (value) {
          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split('\n');
          buffer = lines.pop() || '';
          
          for (const line of lines) {
            if (line.startsWith('data: ') && line !== 'data: [DONE]') {
              try {
                const data = JSON.parse(line.slice(6));
                if (data.error) {
                  throw new Error(data.error);
                }
                if (data.text) {
                  fullText += data.text;
                  setInlineComparisons(prev => ({ ...prev, [title]: fullText }));
                }
              } catch (e) {
                if (e instanceof Error && e.message !== "Unexpected end of JSON input" && !e.message.includes("JSON")) {
                  throw e;
                }
              }
            }
          }
        }
      }
    } catch (error: any) {
      console.error('Compare error:', error);
      setInlineComparisons(prev => ({ ...prev, [title]: `🚨 Đã xảy ra lỗi: ${error.message || 'Vui lòng thử lại.'}` }));
    } finally {
      setLoadingInline(prev => {
        const newState = { ...prev };
        delete newState[title];
        return newState;
      });
    }
  };

  const handleInlineEnhance = async (title: string, sectionContent: string) => {
    const comparisonContent = inlineComparisons[title] || '';
    setLoadingInline(prev => ({ ...prev, [title]: 'enhancing' }));
    setInlineEnhancements(prev => ({ ...prev, [title]: '' }));

    const prompt = `Bạn là một Kỹ sư Đổi mới Sáng tạo (Innovation Engineer) xuất chúng. Dựa trên ý tưởng ban đầu và bản so sánh với các sản phẩm đã có, hãy ĐỀ XUẤT THÊM CÁC TÍNH NĂNG/CẢI TIẾN ĐỘT PHÁ HƠN NỮA để ý tưởng này trở nên hoàn toàn khác biệt, cực kỳ sáng tạo, CHƯA AI LÀM và CÓ TÍNH ỨNG DỤNG CAO VÀO THỰC TẾ.
Hãy phân tích bằng DeepSeek V4 Pro, ưu tiên các hướng cải tiến mới, khả thi và có thể kiểm chứng.

Ý TƯỞNG BAN ĐẦU:
${sectionContent}

BẢN SO SÁNH THỰC TẾ:
${comparisonContent}

YÊU CẦU NÂNG CẤP (Đóng vai trò GPT-4o/GPT-5 để sáng tạo):
1. ĐỀ XUẤT ĐỘT PHÁ (Chưa ai làm): Đưa ra 3-5 tính năng/cải tiến MỚI TOANH, độc đáo, mang yếu tố "WOW" (bất ngờ, thú vị). Hãy suy nghĩ vượt ra ngoài các giải pháp thông thường.
2. TÍNH ỨNG DỤNG THỰC TẾ CAO: Các tính năng này phải giải quyết được vấn đề thực tế một cách hiệu quả, không viển vông, có thể áp dụng ngay vào đời sống.
3. TÍNH KHẢ THI: Đảm bảo các tính năng này ĐƠN GIẢN, phù hợp với trình độ học sinh lớp ${grade} (có thể làm được với công nghệ: ${techLimit}).
4. GIẢI THÍCH SỰ KHÁC BIỆT: Giải thích rõ tại sao các cải tiến này lại làm cho dự án trở nên "vô đối" và hữu ích hơn nhiều so với các sản phẩm cũ trên mạng.
5. VẬN DỤNG KIẾN THỨC: Gợi ý cách học sinh vận dụng kiến thức môn học để làm các tính năng mới này.

Trình bày bằng Markdown, văn phong hấp dẫn, truyền cảm hứng và rõ ràng.`;

    try {
      const response = await fetch('/api/generate-deepseek', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          prompt,
          passcode: localStorage.getItem('savedPasscode'),
          deviceId: getDeviceId()
        })
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || 'Failed to enhance');
      }

      const reader = response.body?.getReader();
      const decoder = new TextDecoder();
      let done = false;
      let buffer = '';
      let fullText = '';

      while (!done && reader) {
        const { value, done: readerDone } = await reader.read();
        done = readerDone;
        if (value) {
          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split('\n');
          buffer = lines.pop() || '';
          
          for (const line of lines) {
            if (line.startsWith('data: ') && line !== 'data: [DONE]') {
              try {
                const data = JSON.parse(line.slice(6));
                if (data.error) {
                  throw new Error(data.error);
                }
                if (data.text) {
                  fullText += data.text;
                  setInlineEnhancements(prev => ({ ...prev, [title]: fullText }));
                }
              } catch (e) {
                if (e instanceof Error && e.message !== "Unexpected end of JSON input" && !e.message.includes("JSON")) {
                  throw e;
                }
              }
            }
          }
        }
      }
    } catch (error: any) {
      console.error('Enhance error:', error);
      setInlineEnhancements(prev => ({ ...prev, [title]: `🚨 Đã xảy ra lỗi: ${error.message || 'Vui lòng thử lại.'}` }));
    } finally {
      setLoadingInline(prev => {
        const newState = { ...prev };
        delete newState[title];
        return newState;
      });
    }
  };

  const handleInlineDetailedGuide = async (title: string, sectionContent: string) => {
    setLoadingInline(prev => ({ ...prev, [title]: 'detailing' }));
    setInlineDetailedGuides(prev => ({ ...prev, [title]: '' }));

    const prompt = `Bạn là giáo viên hướng dẫn học sinh làm dự án sáng tạo khoa học kỹ thuật. Hãy viết HƯỚNG DẪN CÁCH LÀM CHI TIẾT cho đúng ý tưởng dưới đây.

Ý TƯỞNG CẦN HƯỚNG DẪN:
${sectionContent}

THÔNG TIN HỌC SINH:
- Cấp học: ${capHoc}
- Lớp: ${grade}
- Giới hạn công nghệ: ${techLimit}
- Nguồn lực đang có: ${resources || 'Không có'}
- Bối cảnh: ${context || 'Không có'}

YÊU CẦU TRÌNH BÀY:
1. Không viết lại toàn bộ phần phân tích ý tưởng. Chỉ tập trung vào cách làm thực tế.
2. Chia thành các mục rõ ràng:
   - Mục tiêu sản phẩm/mô hình
   - Vật liệu, công cụ, phần mềm cần chuẩn bị
   - Sơ đồ nguyên lý hoặc luồng hoạt động bằng chữ
   - Quy trình làm từng bước
   - Cách kiểm thử, đo đạc, ghi kết quả
   - Tiêu chí đánh giá sản phẩm hoàn thành
   - Lỗi thường gặp và cách khắc phục
   - Cách nâng cấp nếu còn thời gian
3. Phù hợp với học sinh ${grade}, ưu tiên vật liệu dễ kiếm, an toàn, chi phí thấp.
4. Trình bày bằng Markdown, ngắn gọn nhưng đủ chi tiết để học sinh có thể bắt tay làm ngay.`;

    try {
      const response = await fetch('/api/generate-deepseek', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          prompt,
          passcode: localStorage.getItem('savedPasscode'),
          deviceId: getDeviceId()
        })
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || 'Failed to generate detailed guide');
      }

      const reader = response.body?.getReader();
      const decoder = new TextDecoder();
      let done = false;
      let buffer = '';
      let fullText = '';

      while (!done && reader) {
        const { value, done: readerDone } = await reader.read();
        done = readerDone;
        if (value) {
          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split('\n');
          buffer = lines.pop() || '';

          for (const line of lines) {
            if (line.startsWith('data: ') && line !== 'data: [DONE]') {
              try {
                const data = JSON.parse(line.slice(6));
                if (data.error) {
                  throw new Error(data.error);
                }
                if (data.text) {
                  fullText += data.text;
                  setInlineDetailedGuides(prev => ({ ...prev, [title]: fullText }));
                }
              } catch (e) {
                if (e instanceof Error && e.message !== "Unexpected end of JSON input" && !e.message.includes("JSON")) {
                  throw e;
                }
              }
            }
          }
        }
      }
    } catch (error: any) {
      console.error('Detailed guide error:', error);
      setInlineDetailedGuides(prev => ({ ...prev, [title]: `🚨 Đã xảy ra lỗi: ${error.message || 'Vui lòng thử lại.'}` }));
    } finally {
      setLoadingInline(prev => {
        const newState = { ...prev };
        delete newState[title];
        return newState;
      });
    }
  };

  const compareIdea = async (ideaNumber: number) => {
    const sections = result.split(/(?=^###\s*(?:💡\s*)?Ý TƯỞNG\s+\d+\b)/m);
    const selectedSection = sections.find(section =>
      new RegExp(`^###\\s*(?:💡\\s*)?Ý TƯỞNG\\s+${ideaNumber}\\b`, 'm').test(section)
    );
    const ideaContent = selectedSection?.split(/^##\s+🏆/m)[0]?.trim();

    if (!ideaContent) {
      setCompareResult(`Không tìm thấy nội dung của Ý tưởng ${ideaNumber}. Vui lòng chọn ý tưởng khác.`);
      setActiveTab('compare');
      return;
    }

    setIsComparing(true);
    setCompareResult('');
    setActiveTab('compare');

    const prompt = `Bạn là chuyên gia đánh giá dự án khoa học kỹ thuật và đổi mới sáng tạo.
Hãy phân tích khách quan ý tưởng sau bằng cách so sánh với sản phẩm/dự án tương tự đã có trên thị trường hoặc trên mạng.
Hãy dùng DeepSeek V4 Pro để đối chiếu thực tế; chỉ nêu nguồn/link khi chắc chắn, không bịa nguồn.

Ý TƯỞNG CẦN SO SÁNH:
${ideaContent}

YÊU CẦU:
1. Nêu 2-3 sản phẩm/dự án tương tự đã có, kèm link tham khảo nếu tìm thấy.
2. So sánh điểm giống và khác biệt cốt lõi.
3. Đánh giá tính mới, tính sáng tạo và khả năng đạt giải.
4. Gợi ý cách cải tiến để ý tưởng khác biệt hơn nhưng vẫn phù hợp với học sinh ${grade}.

Trình bày bằng Markdown, rõ ràng, ngắn gọn và có tính thực tế.`;

    try {
      const response = await fetch('/api/generate-deepseek', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          prompt,
          passcode: localStorage.getItem('savedPasscode'),
          deviceId: getDeviceId()
        })
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || 'Failed to compare');
      }

      const reader = response.body?.getReader();
      const decoder = new TextDecoder();
      let done = false;
      let buffer = '';
      let fullText = '';

      while (!done && reader) {
        const { value, done: readerDone } = await reader.read();
        done = readerDone;
        if (value) {
          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split('\n');
          buffer = lines.pop() || '';

          for (const line of lines) {
            if (line.startsWith('data: ') && line !== 'data: [DONE]') {
              try {
                const data = JSON.parse(line.slice(6));
                if (data.error) {
                  throw new Error(data.error);
                }
                if (data.text) {
                  fullText += data.text;
                  setCompareResult(fullText);
                }
              } catch (e) {
                if (e instanceof Error && e.message !== "Unexpected end of JSON input" && !e.message.includes("JSON")) {
                  throw e;
                }
              }
            }
          }
        }
      }

      if (currentSessionId && fullText) {
        setHistory(prev => prev.map(session =>
          session.id === currentSessionId ? { ...session, compareResult: fullText } : session
        ));
      }
    } catch (error: any) {
      console.error('Compare idea error:', error);
      setCompareResult(`🚨 Đã xảy ra lỗi khi so sánh: ${error.message || 'Vui lòng thử lại.'}`);
    } finally {
      setIsComparing(false);
    }
  };

  const generateAndDownloadWordDoc = async (content: string, title: string) => {
    const colors = {
      primary: '0F766E',
      primaryDark: '134E4A',
      accent: '10B981',
      soft: 'ECFDF5',
      line: '99F6E4',
      text: '1F2937',
      muted: '475569',
      white: 'FFFFFF',
    };

    const normalRun = { font: 'Arial', size: 23, color: colors.text };

    const makeFileName = (name: string) => {
      const cleaned = name
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-zA-Z0-9]+/g, '-')
        .replace(/^-+|-+$/g, '')
        .toLowerCase();
      return cleaned || 'ideagpt-y-tuong-sang-tao';
    };

    const parseMarkdownLine = (
      line: string,
      overrides: { bold?: boolean; color?: string; size?: number; font?: string } = {}
    ): TextRun[] => {
      const runs: TextRun[] = [];
      let currentText = '';
      let isBold = false;
      let isItalic = false;

      const addRun = (text: string, options: { bold?: boolean; italics?: boolean } = {}) => {
        runs.push(new TextRun({
          ...normalRun,
          ...overrides,
          text,
          bold: overrides.bold || options.bold || false,
          italics: options.italics || false,
        }));
      };
      
      // Simple parser for **bold** and *italic*
      for (let i = 0; i < line.length; i++) {
        if (line.substring(i, i + 2) === '**') {
          if (currentText) {
            addRun(currentText, { bold: isBold, italics: isItalic });
            currentText = '';
          }
          isBold = !isBold;
          i++; // Skip the second '*'
        } else if (line[i] === '*' && line.substring(i, i + 2) !== '**') {
           if (currentText) {
            addRun(currentText, { bold: isBold, italics: isItalic });
            currentText = '';
          }
          isItalic = !isItalic;
        } else if (line[i] === '`') {
          if (currentText) {
            addRun(currentText, { bold: isBold, italics: isItalic });
            currentText = '';
          }
          const endIndex = line.indexOf('`', i + 1);
          if (endIndex !== -1) {
            runs.push(new TextRun({
              text: line.slice(i + 1, endIndex),
              font: 'Consolas',
              size: 21,
              color: colors.primaryDark,
              shading: { type: ShadingType.CLEAR, fill: colors.soft, color: 'auto' },
            }));
            i = endIndex;
          } else {
            currentText += line[i];
          }
        } else {
          currentText += line[i];
        }
      }
      
      if (currentText) {
        addRun(currentText, { bold: isBold, italics: isItalic });
      }
      
      return runs;
    };

    const border = { style: BorderStyle.SINGLE, size: 4, color: colors.line };
    const cellMargins = { top: 130, bottom: 130, left: 180, right: 180 };
    const inputRows = [
      ['Lĩnh vực', field],
      ['Cấp học', capHoc],
      ['Lớp', grade],
      ['Giới hạn công nghệ', techLimit],
      ['Mục tiêu', mucTieu],
      ['Bối cảnh', context || 'Không có'],
      ['Nguồn lực', resources || 'Không có'],
      ['Ngày xuất file', new Date().toLocaleString('vi-VN')],
    ];

    const infoTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: border,
        bottom: border,
        left: border,
        right: border,
        insideHorizontal: border,
        insideVertical: border,
      },
      rows: inputRows.map(([label, value], index) => new TableRow({
        children: [
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            margins: cellMargins,
            shading: { type: ShadingType.CLEAR, fill: index % 2 === 0 ? colors.soft : colors.white, color: 'auto' },
            children: [
              new Paragraph({
                children: [new TextRun({ text: label, bold: true, color: colors.primaryDark, font: 'Arial', size: 22 })],
              }),
            ],
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            margins: cellMargins,
            children: [
              new Paragraph({
                children: [new TextRun({ text: value, color: colors.text, font: 'Arial', size: 22 })],
              }),
            ],
          }),
        ],
      })),
    });

    const bodyBlocks: (Paragraph | Table)[] = [];

    const lines = content.split('\n');
    let inList = false;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      
      if (!line) {
        // Empty line, add some spacing if not in a list
        if (!inList) {
           bodyBlocks.push(new Paragraph({ text: "", spacing: { after: 80 } }));
        }
        continue;
      }

      if (/Cách làm chi tiết/i.test(line)) {
        continue;
      }

      // Handle Headings
      if (line.startsWith('### ')) {
        inList = false;
        const text = line.replace(/^### /, '').trim();
        const isIdeaHeading = text.includes('Ý TƯỞNG');
        bodyBlocks.push(
          new Paragraph({
            children: parseMarkdownLine(text, {
              bold: true,
              size: isIdeaHeading ? 27 : 25,
              color: colors.primaryDark,
            }),
            heading: HeadingLevel.HEADING_3,
            spacing: { before: isIdeaHeading ? 360 : 220, after: 140 },
            shading: isIdeaHeading ? { type: ShadingType.CLEAR, fill: colors.soft, color: 'auto' } : undefined,
            border: isIdeaHeading ? {
              left: { style: BorderStyle.SINGLE, size: 18, color: colors.accent, space: 8 },
              bottom: { style: BorderStyle.SINGLE, size: 4, color: colors.line, space: 2 },
            } : undefined,
          })
        );
      } else if (line.startsWith('## ')) {
        inList = false;
        const text = line.replace(/^## /, '').trim();
        bodyBlocks.push(
          new Paragraph({
            children: [new TextRun({ text, bold: true, color: colors.primary, font: 'Arial', size: 30 })],
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 420, after: 180 },
            border: {
              bottom: { style: BorderStyle.SINGLE, size: 8, color: colors.accent, space: 4 },
            },
          })
        );
      } else if (line.startsWith('# ')) {
        inList = false;
        const text = line.replace(/^# /, '').trim();
        bodyBlocks.push(
          new Paragraph({
            children: [new TextRun({ text, bold: true, color: colors.primaryDark, font: 'Arial', size: 32 })],
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 480, after: 240 },
          })
        );
      } 
      // Handle Lists
      else if (line.startsWith('- ') || line.startsWith('* ')) {
        inList = true;
        const text = line.replace(/^[-*] /, '').trim();
        bodyBlocks.push(
          new Paragraph({
            children: parseMarkdownLine(text),
            bullet: { level: 0 },
            spacing: { after: 95, line: 330 },
            indent: { left: 520, hanging: 220 },
          })
        );
      } else if (line.match(/^\d+\.\s/)) {
         inList = true;
         const text = line.replace(/^\d+\.\s/, '').trim();
         bodyBlocks.push(
          new Paragraph({
            children: parseMarkdownLine(text),
            numbering: { reference: "decimal-numbering", level: 0 },
            spacing: { after: 95, line: 330 },
          })
        );
      }
      // Handle Horizontal Rule
      else if (line === '---' || line === '***' || line === '___') {
         inList = false;
         bodyBlocks.push(
           new Paragraph({
             text: "__________________________________________________",
             alignment: AlignmentType.CENTER,
             spacing: { before: 120, after: 120 }
           })
         );
      }
      // Normal Paragraph
      else {
        inList = false;
        bodyBlocks.push(
          new Paragraph({
            children: parseMarkdownLine(line),
            spacing: { after: 130, line: 340 },
            alignment: AlignmentType.JUSTIFIED,
          })
        );
      }
    }

    const doc = new Document({
      numbering: {
        config: [
          {
            reference: "decimal-numbering",
            levels: [
              {
                level: 0,
                format: "decimal",
                text: "%1.",
                alignment: "left",
                style: {
                  paragraph: {
                    indent: { left: 720, hanging: 360 },
                  },
                },
              },
            ],
          },
        ],
      },
      sections: [
        {
          properties: {
            page: {
              margin: { top: 900, right: 820, bottom: 900, left: 900, header: 500, footer: 500 },
              borders: {
                pageBorderTop: { style: BorderStyle.SINGLE, size: 6, color: colors.line },
                pageBorderBottom: { style: BorderStyle.SINGLE, size: 6, color: colors.line },
                pageBorderLeft: { style: BorderStyle.SINGLE, size: 6, color: colors.line },
                pageBorderRight: { style: BorderStyle.SINGLE, size: 6, color: colors.line },
              },
            },
          },
          headers: {
            default: new Header({
              children: [
                new Paragraph({
                  alignment: AlignmentType.RIGHT,
                  children: [new TextRun({ text: 'IdeaGPT - Báo cáo ý tưởng sáng tạo', color: colors.primary, font: 'Arial', size: 18, bold: true })],
                }),
              ],
            }),
          },
          footers: {
            default: new Footer({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({ text: 'Trang ', color: colors.muted, font: 'Arial', size: 18 }),
                    new TextRun({ children: [PageNumber.CURRENT], color: colors.muted, font: 'Arial', size: 18 }),
                  ],
                }),
              ],
            }),
          },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 460, after: 120 },
              children: [new TextRun({ text: 'IDEAGPT', bold: true, color: colors.primary, font: 'Arial', size: 42 })],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 260 },
              children: [new TextRun({ text: 'BÁO CÁO Ý TƯỞNG SÁNG TẠO', bold: true, color: colors.primaryDark, font: 'Arial', size: 36 })],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 420 },
              children: [new TextRun({ text: title, bold: true, color: colors.text, font: 'Arial', size: 28 })],
            }),
            infoTable,
            new Paragraph({
              spacing: { before: 280, after: 80 },
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: 'Tài liệu được định dạng tự động: tiêu đề, phân mục, danh sách, bảng thông tin và số trang.',
                  italics: true,
                  color: colors.muted,
                  font: 'Arial',
                  size: 20,
                }),
              ],
            }),
            new Paragraph({ children: [new PageBreak()] }),
            ...bodyBlocks,
          ],
        },
      ],
    });

    try {
      const blob = await Packer.toBlob(doc);
      saveAs(blob, `${makeFileName(title)}.docx`);
    } catch (error) {
      console.error('Error generating or downloading Word document:', error);
      alert('Đã có lỗi xảy ra khi tạo file Word. Vui lòng thử lại.');
    }
  };

  const loadSession = (session: SavedSession) => {
    setField(session.inputs.field);
    setCapHoc(session.inputs.capHoc);
    setGrade(session.inputs.grade);
    setTechLimit(session.inputs.techLimit);
    setMucTieu(session.inputs.mucTieu);
    setContext(session.inputs.context);
    setResources(session.inputs.resources);
    setResult(session.result);
    setGenerationMode(session.mode || 'basic');
    setCompareResult(session.compareResult || '');
    setInlineComparisons({});
    setInlineEnhancements({});
    setInlineDetailedGuides({});
    setLoadingInline({});
    setCurrentSessionId(session.id);
    setActiveTab('main');
    if (window.innerWidth < 1024) {
      setSidebarOpen(false);
    }
  };

  const deleteSession = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    setHistory(prev => prev.filter(s => s.id !== id));
    if (currentSessionId === id) {
      setCurrentSessionId(null);
      setResult('');
      setCompareResult('');
      if (activeTab !== 'history') {
        setActiveTab('history');
      }
    }
  };

  useEffect(() => {
    if (isGenerating && activeTab === 'main') {
      resultEndRef.current?.scrollIntoView({ behavior: 'smooth' });
    }
  }, [result, isGenerating, activeTab]);

  const renderResultSections = (text: string) => {
    const parts = text.split(/(?=^#{2,3} )/m);
    
    return parts.map((part, index) => {
      const isIdea = part.includes('Ý TƯỞNG') && !part.includes('TOP 3');
      const titleMatch = part.match(/^#{2,3} (.*)/m);
      const title = titleMatch ? titleMatch[1].replace('💡', '').trim() : '';
      const scoreMatch = part.match(/\*\*Điểm đánh giá:\*\*\s*(\d{1,3})\s*\/\s*100/i);
      const ideaScore = scoreMatch ? Math.min(100, Math.max(0, Number(scoreMatch[1]))) : null;
      const scoreClass = ideaScore === null
        ? 'bg-teal-900/40 text-teal-300 border-teal-700/50'
        : ideaScore >= 85
          ? 'bg-emerald-900/40 text-emerald-300 border-emerald-700/50'
          : ideaScore >= 70
            ? 'bg-amber-900/40 text-amber-300 border-amber-700/50'
            : 'bg-rose-900/40 text-rose-300 border-rose-700/50';

      return (
        <div key={index} className={isIdea ? "mb-12" : "mb-8"}>
          <ReactMarkdown 
            remarkPlugins={[remarkGfm]}
            components={getMarkdownComponents(text)}
          >
            {part}
          </ReactMarkdown>
          
          {isIdea && title && (
            <div className="mt-6 flex flex-col gap-4">
              <div className="flex flex-wrap items-center gap-3">
                <div className={cn("px-4 py-2 rounded-xl font-bold text-sm border flex items-center gap-2 shadow-sm", scoreClass)}>
                  <Target className="w-4 h-4" />
                  {ideaScore !== null ? `Điểm: ${ideaScore}/100` : 'Chưa có điểm'}
                </div>

                <button
                  onClick={() => handleInlineDetailedGuide(title, part)}
                  disabled={loadingInline[title] === 'detailing'}
                  className="px-4 py-2 rounded-xl font-semibold text-sm bg-emerald-900/40 text-emerald-300 border border-emerald-700/50 hover:bg-emerald-800/50 transition-colors flex items-center gap-2 shadow-sm"
                >
                  {loadingInline[title] === 'detailing' ? <Loader2 className="w-4 h-4 animate-spin" /> : <BookOpen className="w-4 h-4" />}
                  {loadingInline[title] === 'detailing'
                    ? 'Đang viết chi tiết'
                    : inlineDetailedGuides[title]
                      ? 'Tạo lại hướng dẫn'
                      : 'Hướng dẫn chi tiết'}
                </button>

                <button
                  onClick={() => handleInlineCompare(title, part)}
                  disabled={loadingInline[title] === 'comparing'}
                  className="px-4 py-2 rounded-xl font-semibold text-sm bg-cyan-900/40 text-cyan-300 border border-cyan-700/50 hover:bg-cyan-800/50 transition-colors flex items-center gap-2 shadow-sm"
                >
                  {loadingInline[title] === 'comparing' ? <Loader2 className="w-4 h-4 animate-spin" /> : <GitCompare className="w-4 h-4" />}
                  So sánh ý tưởng này
                </button>
                
                {inlineComparisons[title] && (
                  <button
                    onClick={() => handleInlineEnhance(title, part)}
                    disabled={loadingInline[title] === 'enhancing'}
                    className="px-4 py-2 rounded-xl font-semibold text-sm bg-fuchsia-900/40 text-fuchsia-300 border border-fuchsia-700/50 hover:bg-fuchsia-800/50 transition-colors flex items-center gap-2 shadow-sm"
                  >
                    {loadingInline[title] === 'enhancing' ? <Loader2 className="w-4 h-4 animate-spin" /> : <Sparkles className="w-4 h-4" />}
                    Sáng tạo hơn nữa
                  </button>
                )}
              </div>

              {inlineDetailedGuides[title] && (
                <motion.div
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="p-6 rounded-2xl bg-emerald-950/30 border border-emerald-800/50 shadow-inner mt-2"
                >
                  <h4 className="text-lg font-bold text-emerald-400 mb-4 flex items-center gap-2 border-b border-emerald-800/50 pb-3">
                    <BookOpen className="w-5 h-5" /> Hướng dẫn cách làm chi tiết
                  </h4>
                  <div className="prose prose-invert max-w-none prose-p:text-emerald-100/90 prose-li:text-emerald-100/90 prose-strong:text-emerald-300 prose-headings:text-emerald-200">
                    <ReactMarkdown remarkPlugins={[remarkGfm]}>{inlineDetailedGuides[title]}</ReactMarkdown>
                  </div>
                </motion.div>
              )}

              {inlineComparisons[title] && (
                <motion.div 
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="p-6 rounded-2xl bg-cyan-950/30 border border-cyan-800/50 shadow-inner mt-2"
                >
                  <h4 className="text-lg font-bold text-cyan-400 mb-4 flex items-center gap-2 border-b border-cyan-800/50 pb-3">
                    <GitCompare className="w-5 h-5" /> Kết quả so sánh
                  </h4>
                  <div className="prose prose-invert max-w-none prose-p:text-cyan-100/90 prose-li:text-cyan-100/90 prose-strong:text-cyan-300 prose-headings:text-cyan-200">
                    <ReactMarkdown remarkPlugins={[remarkGfm]}>{inlineComparisons[title]}</ReactMarkdown>
                  </div>
                </motion.div>
              )}

              {inlineEnhancements[title] && (
                <motion.div 
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="p-6 rounded-2xl bg-fuchsia-950/30 border border-fuchsia-800/50 shadow-inner mt-2"
                >
                  <h4 className="text-lg font-bold text-fuchsia-400 mb-4 flex items-center gap-2 border-b border-fuchsia-800/50 pb-3">
                    <Sparkles className="w-5 h-5" /> Đề xuất sáng tạo đột phá
                  </h4>
                  <div className="prose prose-invert max-w-none prose-p:text-fuchsia-100/90 prose-li:text-fuchsia-100/90 prose-strong:text-fuchsia-300 prose-headings:text-fuchsia-200">
                    <ReactMarkdown remarkPlugins={[remarkGfm]}>{inlineEnhancements[title]}</ReactMarkdown>
                  </div>
                </motion.div>
              )}
            </div>
          )}
        </div>
      );
    });
  };

  const getMarkdownComponents = (sourceText: string) => ({
    h2: ({node, ...props}: any) => <h2 className="text-2xl font-extrabold text-transparent bg-clip-text bg-gradient-to-r from-emerald-400 to-cyan-300 border-b-2 border-teal-700/50 pb-3 mt-14 mb-8 first:mt-0 tracking-tight" {...props} />,
    h3: ({node, children, ...props}: any) => {
      const extractText = (childArray: any): string => {
        if (!childArray) return '';
        if (typeof childArray === 'string') return childArray;
        if (Array.isArray(childArray)) return childArray.map(extractText).join('');
        if (childArray.props && childArray.props.children) return extractText(childArray.props.children);
        return '';
      };
      const text = extractText(children);
      const isIdea = text.includes('Ý TƯỞNG') && !text.includes('TOP 3');
      const cleanTitle = text.replace('💡', '').trim();
      const isSaved = savedIdeas.some(idea => idea.title === cleanTitle);

      return (
        <div className="relative group flex items-start sm:items-center justify-between bg-gradient-to-r from-teal-900/60 to-teal-950/60 px-6 py-4 rounded-xl border border-teal-700/50 mt-12 mb-6 shadow-lg shadow-teal-900/20 border-l-4 border-l-emerald-400 backdrop-blur-sm">
          <h3 className="text-xl font-bold text-emerald-50 m-0 leading-snug" {...props}>{children}</h3>
          {isIdea && (
            <button
              onClick={() => {
                if (isSaved) {
                  setSavedIdeas(prev => prev.filter(i => i.title !== cleanTitle));
                } else {
                  const content = extractIdea(sourceText, text);
                  if (content) {
                    const newIdea: SavedIdea = {
                      id: Date.now().toString() + Math.random().toString(36).substr(2, 9),
                      timestamp: Date.now(),
                      title: cleanTitle,
                      content: content,
                      inputs: { field, capHoc, grade }
                    };
                    setSavedIdeas(prev => [newIdea, ...prev]);
                  }
                }
              }}
              className={cn(
                "p-2.5 rounded-lg transition-all shadow-sm shrink-0 flex items-center gap-2 text-sm font-semibold",
                isSaved 
                  ? "text-emerald-300 bg-emerald-900/50 hover:bg-emerald-800/50 opacity-100 border border-emerald-700/50" 
                  : "text-teal-300 bg-teal-800/50 hover:text-emerald-300 hover:bg-emerald-900/40 sm:opacity-0 sm:group-hover:opacity-100 border border-teal-600/50"
              )}
              title={isSaved ? "Bỏ lưu ý tưởng" : "Lưu ý tưởng này"}
            >
              <Heart className={cn("w-4 h-4", isSaved && "fill-current")} />
              {isSaved ? "Đã lưu" : "Lưu ý tưởng"}
            </button>
          )}
        </div>
      );
    },
    ul: ({node, ...props}: any) => <ul className="space-y-3 mb-6 pl-6 list-disc marker:text-emerald-400" {...props} />,
    ol: ({node, ...props}: any) => <ol className="space-y-4 mb-6 pl-6 list-decimal marker:text-emerald-400 marker:font-bold" {...props} />,
    li: ({node, children, ...props}: any) => {
      const extractText = (childArray: any): string => {
        if (!childArray) return '';
        if (typeof childArray === 'string') return childArray;
        if (Array.isArray(childArray)) return childArray.map(extractText).join('');
        if (childArray.props && childArray.props.children) return extractText(childArray.props.children);
        return '';
      };
      const text = extractText(children);

      if (/Cách làm chi tiết/i.test(text)) {
        return null;
      }

      return <li className="text-teal-50/90 leading-relaxed pl-1" {...props}>{children}</li>;
    },
    strong: ({node, ...props}: any) => <strong className="text-emerald-300 font-semibold tracking-wide" {...props} />,
    p: ({node, ...props}: any) => <p className="text-teal-50/90 leading-relaxed mb-5 text-[1.05rem]" {...props} />,
    blockquote: ({node, ...props}: any) => <blockquote className="border-l-4 border-emerald-500 bg-gradient-to-r from-emerald-900/20 to-transparent p-5 rounded-r-2xl italic text-teal-100 my-8 shadow-sm" {...props} />,
    code: ({node, inline, ...props}: any) => inline 
      ? <code className="bg-teal-900/60 text-emerald-200 px-2 py-0.5 rounded-md text-sm font-mono border border-teal-700/50" {...props} />
      : <code className="block bg-[#0f172a] text-teal-50 p-6 rounded-2xl text-sm font-mono overflow-x-auto my-8 shadow-xl border border-slate-800" {...props} />,
  });

  // Kiểm tra hết hạn hoặc bị khóa
  const isLocked = !TEST_CONFIG.isActive || (TEST_CONFIG.expireAt && new Date() > new Date(TEST_CONFIG.expireAt));

  if (isLocked) {
    return (
      <div className="min-h-screen bg-teal-950 flex flex-col items-center justify-center p-4 font-sans">
        <div className="bg-teal-900/40 backdrop-blur-md p-8 rounded-2xl shadow-xl border border-teal-700/50 max-w-md w-full text-center">
          <div className="w-16 h-16 bg-teal-800/50 rounded-full flex items-center justify-center mx-auto mb-6">
            <Lock className="w-8 h-8 text-teal-500" />
          </div>
          <h2 className="text-2xl font-bold text-teal-50 mb-3">Đã đóng thử nghiệm</h2>
          <p className="text-teal-200 leading-relaxed">
            Phiên bản thử nghiệm này đã kết thúc hoặc tạm thời bị khóa. Cảm ơn bạn đã quan tâm và trải nghiệm IdeaGPT!
          </p>
        </div>
      </div>
    );
  }

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-teal-950 flex flex-col items-center justify-center p-4 font-sans">
        <div className="bg-teal-900/40 backdrop-blur-md p-8 rounded-2xl shadow-xl border border-teal-700/50 max-w-md w-full">
          <div className="w-16 h-16 bg-emerald-900/30 rounded-full flex items-center justify-center mx-auto mb-6">
            <Key className="w-8 h-8 text-emerald-400" />
          </div>
          <h2 className="text-2xl font-bold text-teal-50 mb-2 text-center">Nhập mã truy cập</h2>
          <p className="text-teal-300 text-center mb-8 text-sm">
            Đây là phiên bản thử nghiệm giới hạn. Vui lòng nhập mã để tiếp tục.
          </p>
          <form onSubmit={(e) => {
            e.preventDefault();
            if (passcodeInput.trim()) {
              verifyPasscode(passcodeInput.trim());
            }
          }}>
            <div className="mb-6">
              <input
                type="password"
                value={passcodeInput}
                onChange={(e) => {
                  setPasscodeInput(e.target.value);
                  setPasscodeError(false);
                }}
                placeholder="Nhập mã truy cập..."
                disabled={isCheckingPasscode}
                className={cn(
                  "w-full px-4 py-3 rounded-xl border bg-teal-900/50 focus:bg-teal-800/50 text-teal-50 placeholder-teal-500 transition-colors outline-none",
                  passcodeError ? "border-red-500/50 focus:border-red-400 focus:ring-2 focus:ring-red-400/20" : "border-teal-700/50 focus:border-emerald-500 focus:ring-2 focus:ring-emerald-500/20",
                  isCheckingPasscode && "opacity-50 cursor-not-allowed"
                )}
              />
              {passcodeError && (
                <p className="text-red-400 text-sm mt-2 font-medium">{passcodeErrorMessage || 'Mã truy cập không chính xác!'}</p>
              )}
            </div>
            <button
              type="submit"
              disabled={isCheckingPasscode || !passcodeInput.trim()}
              className="w-full bg-emerald-600 hover:bg-emerald-500 disabled:bg-emerald-800 disabled:text-emerald-400/50 text-white font-bold py-3 px-4 rounded-xl transition-colors shadow-sm shadow-emerald-900/50 flex items-center justify-center gap-2"
            >
              {isCheckingPasscode ? (
                <>
                  <Loader2 className="w-5 h-5 animate-spin" />
                  Đang kiểm tra...
                </>
              ) : (
                'Truy cập hệ thống'
              )}
            </button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className="flex h-screen bg-gradient-to-br from-teal-950 via-emerald-950 to-teal-900 text-teal-50 font-sans overflow-hidden selection:bg-emerald-500/30 selection:text-emerald-50">
      {/* Mobile Sidebar Overlay */}
      <AnimatePresence>
        {!sidebarOpen && (
          <motion.button
            initial={{ opacity: 0, scale: 0.9 }}
            animate={{ opacity: 1, scale: 1 }}
            exit={{ opacity: 0, scale: 0.9 }}
            className="lg:hidden fixed top-4 left-4 z-50 p-2.5 bg-teal-900/80 backdrop-blur-md rounded-xl shadow-lg border border-teal-700/50 text-teal-200 hover:text-emerald-300 hover:bg-teal-800 transition-colors"
            onClick={() => setSidebarOpen(true)}
          >
            <Menu className="w-5 h-5" />
          </motion.button>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {sidebarOpen && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="lg:hidden fixed inset-0 bg-teal-950/60 backdrop-blur-sm z-40"
            onClick={() => setSidebarOpen(false)}
          />
        )}
      </AnimatePresence>

      {/* Sidebar Form */}
      <aside
        className={cn(
          "fixed lg:static inset-y-0 left-0 z-50 w-[340px] bg-teal-950/80 backdrop-blur-xl border-r border-teal-800/50 flex flex-col transition-transform duration-300 ease-[cubic-bezier(0.2,0.8,0.2,1)] shadow-2xl lg:shadow-none",
          sidebarOpen ? "translate-x-0" : "-translate-x-full lg:translate-x-0"
        )}
      >
        <div className="h-16 px-6 border-b border-teal-800/50 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-2.5 text-emerald-400 font-bold text-xl tracking-tight">
            <div className="w-8 h-8 rounded-lg bg-emerald-500 flex items-center justify-center text-teal-950 shadow-sm shadow-emerald-500/20">
              <Lightbulb className="w-5 h-5" />
            </div>
            IdeaGPT
          </div>
          <button 
            className="lg:hidden p-2 text-teal-400 hover:text-teal-200 hover:bg-teal-800/50 rounded-lg transition-colors"
            onClick={() => setSidebarOpen(false)}
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-6 space-y-6 custom-scrollbar">
          <div className="space-y-5">
            <div className="space-y-2">
              <label className="flex items-center gap-2 text-sm font-semibold text-teal-200">
                <Sparkles className="w-4 h-4 text-teal-400" />
                Chế độ tạo ý tưởng
              </label>
              <div className="flex flex-col gap-2">
                <div className="flex gap-2">
                  <button
                    onClick={() => setGenerationMode('basic')}
                    className={cn(
                      "flex-1 py-2 px-3 rounded-xl text-sm font-medium transition-all",
                      generationMode === 'basic'
                        ? "bg-emerald-500 text-teal-950 shadow-sm shadow-emerald-500/20"
                        : "bg-teal-900/50 text-teal-200 hover:bg-teal-800/50 border border-teal-700/50"
                    )}
                  >
                    Cơ bản
                  </button>
                  <button
                    onClick={() => setShowAdvancedModal(true)}
                    className={cn(
                      "flex-1 py-2 px-3 rounded-xl text-sm font-medium transition-all flex flex-col items-center justify-center gap-0.5",
                      generationMode === 'advanced'
                        ? "bg-emerald-500 text-teal-950 shadow-sm shadow-emerald-500/20"
                        : "bg-teal-900/50 text-teal-200 hover:bg-teal-800/50 border border-teal-700/50"
                    )}
                  >
                    <span>Nâng cao</span>
                    {generationMode === 'advanced' && (
                      <span className="text-[10px] opacity-80 font-bold leading-none">
                        ({advancedModel === 'gpt' ? 'GPT-5.4' : 'DeepSeek V4 Pro'})
                      </span>
                    )}
                  </button>
                </div>
              </div>
            </div>

            <div className="space-y-2">
              <label className="flex items-center gap-2 text-sm font-semibold text-teal-200">
                <Layers className="w-4 h-4 text-teal-400" />
                Lĩnh vực
              </label>
              <select
                value={field}
                onChange={(e) => setField(e.target.value)}
                className="w-full p-2.5 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all text-teal-50 font-medium"
              >
                {FIELD_OPTIONS.map((opt) => (
                  <option key={opt} value={opt} className="bg-teal-900 text-teal-50">{opt}</option>
                ))}
              </select>
            </div>

            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-2">
                <label className="flex items-center gap-2 text-sm font-semibold text-teal-200">
                  <BookOpen className="w-4 h-4 text-teal-400" />
                  Cấp học
                </label>
                <select
                  value={capHoc}
                  onChange={(e) => setCapHoc(e.target.value)}
                  className="w-full p-2.5 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all text-teal-50 font-medium"
                >
                  {CAP_HOC_OPTIONS.map((opt) => (
                    <option key={opt} value={opt} className="bg-teal-900 text-teal-50">{opt}</option>
                  ))}
                </select>
              </div>

              <div className="space-y-2">
                <label className="flex items-center gap-2 text-sm font-semibold text-teal-200">
                  <Target className="w-4 h-4 text-teal-400" />
                  Lớp
                </label>
                <select
                  value={grade}
                  onChange={(e) => setGrade(e.target.value)}
                  className="w-full p-2.5 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all text-teal-50 font-medium"
                >
                  {GRADE_OPTIONS.map((opt) => (
                    <option key={opt} value={opt} className="bg-teal-900 text-teal-50">{opt}</option>
                  ))}
                </select>
              </div>
            </div>

            <div className="space-y-2">
              <label className="flex items-center gap-2 text-sm font-semibold text-teal-200">
                <Settings className="w-4 h-4 text-teal-400" />
                Giới hạn công nghệ
              </label>
              <select
                value={techLimit}
                onChange={(e) => setTechLimit(e.target.value)}
                className="w-full p-2.5 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all text-teal-50 font-medium"
              >
                {TECH_LIMIT_OPTIONS.map((opt) => (
                  <option key={opt} value={opt} className="bg-teal-900 text-teal-50">{opt}</option>
                ))}
              </select>
            </div>

            <div className="space-y-2">
              <label className="flex items-center gap-2 text-sm font-semibold text-teal-200">
                <Target className="w-4 h-4 text-teal-400" />
                Mục tiêu
              </label>
              <select
                value={mucTieu}
                onChange={(e) => setMucTieu(e.target.value)}
                className="w-full p-2.5 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all text-teal-50 font-medium"
              >
                {MUC_TIEU_OPTIONS.map((opt) => (
                  <option key={opt} value={opt} className="bg-teal-900 text-teal-50">{opt}</option>
                ))}
              </select>
            </div>

            <div className="space-y-2">
              <label className="text-sm font-semibold text-teal-200">Bối cảnh <span className="text-teal-400/70 font-normal">(Tùy chọn)</span></label>
              <textarea
                value={context}
                onChange={(e) => setContext(e.target.value)}
                placeholder="Ví dụ: Trường ở vùng nông thôn, gần biển..."
                className="w-full p-3 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all min-h-[80px] resize-y text-teal-50 placeholder:text-teal-400/50"
              />
            </div>

            <div className="space-y-2">
              <label className="text-sm font-semibold text-teal-200">Nguồn lực <span className="text-teal-400/70 font-normal">(Tùy chọn)</span></label>
              <textarea
                value={resources}
                onChange={(e) => setResources(e.target.value)}
                placeholder="Ví dụ: Có sẵn bìa carton, chai nhựa, biết lập trình Scratch..."
                className="w-full p-3 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all min-h-[80px] resize-y text-teal-50 placeholder:text-teal-400/50"
              />
            </div>
          </div>
        </div>

        <div className="p-6 border-t border-teal-800/50 bg-teal-950/50 shrink-0">
          <button
            onClick={() => generateIdeas(false)}
            disabled={isGenerating}
            className="w-full py-3 px-4 bg-gradient-to-r from-emerald-500 to-teal-500 hover:from-emerald-400 hover:to-teal-400 text-teal-950 rounded-xl font-bold flex items-center justify-center gap-2.5 transition-all disabled:opacity-70 disabled:cursor-not-allowed shadow-lg shadow-emerald-500/20 hover:shadow-emerald-500/40 active:scale-[0.98]"
          >
            {isGenerating ? (
              <>
                <Loader2 className="w-5 h-5 animate-spin" />
                <span className="truncate max-w-[200px]">{loadingMessage}</span>
              </>
            ) : (
              <>
                <Sparkles className="w-5 h-5" />
                Tạo Ý Tưởng
              </>
            )}
          </button>
        </div>
      </aside>

      {/* Advanced Mode Selection Modal */}
      <AnimatePresence>
        {showAdvancedModal && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="absolute inset-0 bg-teal-950/80 backdrop-blur-sm"
              onClick={() => setShowAdvancedModal(false)}
            />
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="relative bg-teal-900 border border-teal-700/50 p-6 rounded-2xl shadow-2xl max-w-sm w-full"
            >
              <h3 className="text-xl font-bold text-teal-50 mb-2 flex items-center gap-2">
                <Sparkles className="w-5 h-5 text-emerald-400" />
                Chọn Mô Hình AI Nâng Cao
              </h3>
              <p className="text-sm text-teal-300 mb-6">
                Chế độ nâng cao sử dụng các mô hình AI mạnh mẽ nhất, suy nghĩ sâu hơn để tạo ra các ý tưởng đột phá và chưa từng có.
              </p>
              
              <div className="space-y-3">
                <button
                  onClick={() => {
                    setGenerationMode('advanced');
                    setAdvancedModel('gpt');
                    setShowAdvancedModal(false);
                  }}
                  className={cn(
                    "w-full p-4 rounded-xl border transition-all flex items-center justify-between group",
                    advancedModel === 'gpt' && generationMode === 'advanced'
                      ? "bg-teal-800 border-emerald-500 shadow-sm shadow-emerald-500/20"
                      : "bg-teal-800/30 border-teal-700/50 hover:bg-teal-700/50 hover:border-teal-500/50"
                  )}
                >
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-lg bg-teal-950 flex items-center justify-center text-emerald-400 group-hover:scale-110 transition-transform">
                      <Sparkles className="w-5 h-5" />
                    </div>
                    <div className="text-left">
                      <div className="font-bold text-teal-50">GPT-5.4</div>
                      <div className="text-xs text-teal-300">Tư duy logic, lập luận sắc bén</div>
                    </div>
                  </div>
                  <ChevronRight className="w-5 h-5 text-teal-500 group-hover:text-emerald-400 transition-colors" />
                </button>

                <button
                  onClick={() => {
                    setGenerationMode('advanced');
                    setAdvancedModel('deepseek');
                    setShowAdvancedModal(false);
                  }}
                  className={cn(
                    "w-full p-4 rounded-xl border transition-all flex items-center justify-between group",
                    advancedModel === 'deepseek' && generationMode === 'advanced'
                      ? "bg-teal-800 border-blue-500 shadow-sm shadow-blue-500/20"
                      : "bg-teal-800/30 border-teal-700/50 hover:bg-teal-700/50 hover:border-teal-500/50"
                  )}
                >
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-lg bg-teal-950 flex items-center justify-center text-blue-400 group-hover:scale-110 transition-transform">
                      <Sparkles className="w-5 h-5" />
                    </div>
                    <div className="text-left">
                      <div className="font-bold text-teal-50">DeepSeek V4 Pro</div>
                      <div className="text-xs text-teal-300">Phân tích sâu, tư duy nâng cao</div>
                    </div>
                  </div>
                  <ChevronRight className="w-5 h-5 text-teal-500 group-hover:text-blue-400 transition-colors" />
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Main Content */}
      <main className="flex-1 flex flex-col min-w-0 bg-transparent">
        {/* Header Tabs */}
        <header className="h-16 border-b border-teal-800/50 flex items-center px-6 gap-8 bg-teal-950/40 backdrop-blur-md shrink-0 shadow-sm z-10">
          <div className="w-8 lg:hidden" /> {/* Spacer for mobile menu button */}
          <button
            onClick={() => setActiveTab('main')}
            className={cn(
              "h-full px-1 text-sm font-semibold border-b-2 transition-colors flex items-center gap-2",
              activeTab === 'main' 
                ? "border-emerald-400 text-emerald-400" 
                : "border-transparent text-teal-400 hover:text-teal-200"
            )}
          >
            Kết quả Ý Tưởng
          </button>
          {compareResult && (
            <button
              onClick={() => setActiveTab('compare')}
              className={cn(
                "h-full px-1 text-sm font-semibold border-b-2 transition-colors flex items-center gap-2",
                activeTab === 'compare' 
                  ? "border-emerald-400 text-emerald-400" 
                  : "border-transparent text-teal-400 hover:text-teal-200"
              )}
            >
              Phân tích So Sánh
            </button>
          )}
          <button
            onClick={() => setActiveTab('favorites')}
            className={cn(
              "h-full px-1 text-sm font-semibold border-b-2 transition-colors flex items-center gap-2 ml-auto",
              activeTab === 'favorites' 
                ? "border-emerald-400 text-emerald-400" 
                : "border-transparent text-teal-400 hover:text-teal-200"
            )}
          >
            <Heart className={cn("w-4 h-4", activeTab === 'favorites' && "fill-current")} />
            Đã lưu
          </button>
          <button
            onClick={() => setActiveTab('history')}
            className={cn(
              "h-full px-1 text-sm font-semibold border-b-2 transition-colors flex items-center gap-2",
              activeTab === 'history' 
                ? "border-emerald-400 text-emerald-400" 
                : "border-transparent text-teal-400 hover:text-teal-200"
            )}
          >
            <History className="w-4 h-4" />
            Lịch sử
          </button>
        </header>

        {/* Content Area */}
        <div className="flex-1 overflow-y-auto p-6 lg:p-10 custom-scrollbar scroll-smooth">
          <div className={cn("mx-auto", generationMode === 'advanced' && activeTab === 'main' ? "max-w-7xl" : "max-w-4xl")}>
            {activeTab === 'favorites' ? (
              <div className="space-y-6">
                <div className="flex items-center justify-between">
                  <h2 className="text-2xl font-bold text-teal-50 flex items-center gap-2">
                    <Heart className="w-6 h-6 text-emerald-400 fill-current" />
                    Ý tưởng đã lưu
                  </h2>
                  <span className="text-sm text-teal-200 font-medium bg-teal-900/50 px-3 py-1 rounded-full border border-teal-700/50">
                    {savedIdeas.length} ý tưởng
                  </span>
                </div>

                {savedIdeas.length === 0 ? (
                  <div className="bg-teal-900/40 backdrop-blur-sm rounded-2xl shadow-sm border border-teal-700/50 p-12 text-center">
                    <Heart className="w-12 h-12 text-teal-700 mx-auto mb-4" />
                    <h3 className="text-lg font-bold text-teal-100 mb-2">Chưa có ý tưởng nào</h3>
                    <p className="text-teal-300">Hãy nhấn nút "Lưu ý tưởng" bên cạnh mỗi ý tưởng để lưu lại nhé.</p>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 gap-6">
                    {savedIdeas.map(idea => (
                      <div key={idea.id} className="bg-teal-900/40 backdrop-blur-sm rounded-2xl shadow-sm border border-teal-700/50 p-8 relative group">
                        <button 
                          onClick={() => setSavedIdeas(prev => prev.filter(i => i.id !== idea.id))}
                          className="absolute top-4 right-4 p-2 text-teal-400 hover:text-red-400 hover:bg-red-900/30 rounded-lg transition-colors opacity-0 group-hover:opacity-100"
                          title="Xóa khỏi danh sách lưu"
                        >
                          <Trash2 className="w-5 h-5" />
                        </button>
                        <div className="flex items-center gap-2 mb-6">
                           <span className="text-xs font-medium text-teal-200 bg-teal-800/50 px-2.5 py-1 rounded-lg border border-teal-700/50">{idea.inputs.field}</span>
                           <span className="text-xs font-medium text-teal-200 bg-teal-800/50 px-2.5 py-1 rounded-lg border border-teal-700/50">{idea.inputs.grade}</span>
                           <span className="text-xs font-medium text-teal-400 flex items-center gap-1 ml-auto">
                             <Clock className="w-3 h-3" />
                             {new Date(idea.timestamp).toLocaleDateString('vi-VN')}
                           </span>
                        </div>
                        <div className="prose prose-invert prose-emerald max-w-none
                          prose-headings:font-bold 
                          prose-p:text-teal-100 prose-p:leading-relaxed prose-p:text-[15px]
                          prose-li:text-teal-100 prose-li:marker:text-emerald-400
                          prose-strong:text-teal-50 prose-strong:bg-emerald-900/50 prose-strong:px-1.5 prose-strong:py-0.5 prose-strong:rounded
                        ">
                           <ReactMarkdown remarkPlugins={[remarkGfm]} components={getMarkdownComponents(idea.content)}>
                             {idea.content}
                           </ReactMarkdown>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            ) : activeTab === 'history' ? (
              <div className="space-y-6">
                <div className="flex items-center justify-between">
                  <h2 className="text-2xl font-bold text-teal-50 flex items-center gap-2">
                    <History className="w-6 h-6 text-emerald-400" />
                    Lịch sử tạo ý tưởng
                  </h2>
                  <div className="flex items-center gap-3">
                    <span className="text-sm text-teal-200 font-medium bg-teal-900/50 px-3 py-1 rounded-full border border-teal-700/50">
                      {history.length} phiên
                    </span>
                    {history.length > 0 && (
                      showClearConfirm ? (
                        <div className="flex items-center gap-2 bg-red-900/30 px-3 py-1.5 rounded-full border border-red-800/50">
                          <span className="text-xs font-semibold text-red-300">Chắc chắn xóa?</span>
                          <button
                            onClick={() => {
                              setHistory([]);
                              clearIdeaMemory();
                              setCurrentSessionId(null);
                              setResult('');
                              setCompareResult('');
                              setShowClearConfirm(false);
                            }}
                            className="text-xs bg-red-500 hover:bg-red-600 text-white font-medium px-2.5 py-1 rounded-md transition-colors"
                          >
                            Có
                          </button>
                          <button
                            onClick={() => setShowClearConfirm(false)}
                            className="text-xs bg-teal-800/50 hover:bg-teal-700/50 text-teal-100 font-medium px-2.5 py-1 rounded-md border border-teal-600/50 transition-colors"
                          >
                            Hủy
                          </button>
                        </div>
                      ) : (
                        <button
                          onClick={() => setShowClearConfirm(true)}
                          className="text-sm text-red-400 hover:text-red-300 font-medium px-3 py-1.5 rounded-full hover:bg-red-900/30 transition-colors flex items-center gap-1.5"
                        >
                          <Trash2 className="w-4 h-4" />
                          Xóa tất cả
                        </button>
                      )
                    )}
                  </div>
                </div>

                {history.length === 0 ? (
                  <div className="bg-teal-900/40 backdrop-blur-sm rounded-2xl shadow-sm border border-teal-700/50 p-12 text-center">
                    <History className="w-12 h-12 text-teal-700 mx-auto mb-4" />
                    <h3 className="text-lg font-bold text-teal-100 mb-2">Chưa có lịch sử</h3>
                    <p className="text-teal-300">Các ý tưởng bạn tạo sẽ được lưu tự động tại đây để xem lại sau.</p>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {history.map(session => (
                      <motion.div 
                        initial={{ opacity: 0, y: 10 }}
                        animate={{ opacity: 1, y: 0 }}
                        key={session.id} 
                        onClick={() => loadSession(session)}
                        className={cn(
                          "bg-teal-900/40 backdrop-blur-sm border rounded-2xl p-5 hover:shadow-md transition-all cursor-pointer group",
                          currentSessionId === session.id ? "border-emerald-400 ring-2 ring-emerald-900/50" : "border-teal-700/50 hover:border-emerald-500/50"
                        )}
                      >
                        <div className="flex justify-between items-start mb-3">
                          <div className="flex items-center gap-3 text-xs font-medium text-teal-300 bg-teal-800/50 px-2.5 py-1.5 rounded-lg border border-teal-700/50">
                            <div className="flex items-center gap-1.5">
                              <Calendar className="w-3.5 h-3.5" />
                              {new Date(session.timestamp).toLocaleDateString('vi-VN')}
                            </div>
                            <div className="w-1 h-1 bg-teal-600 rounded-full" />
                            <div className="flex items-center gap-1.5">
                              <Clock className="w-3.5 h-3.5" />
                              {new Date(session.timestamp).toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' })}
                            </div>
                          </div>
                          <button 
                            onClick={(e) => deleteSession(session.id, e)} 
                            className="text-teal-500 hover:text-red-400 p-1.5 rounded-lg hover:bg-red-900/30 transition-colors opacity-0 group-hover:opacity-100"
                            title="Xóa phiên này"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                        <h4 className="font-bold text-teal-50 mb-1.5 line-clamp-1 text-base">{session.inputs.field}</h4>
                        <p className="text-sm text-teal-200 mb-4 flex items-center gap-2">
                          <span className="bg-teal-800/50 px-2 py-0.5 rounded text-xs border border-teal-700/50">{session.inputs.capHoc}</span>
                          <span className="text-teal-600">•</span>
                          <span className="truncate">{session.inputs.grade}</span>
                        </p>
                        <div className="flex items-center justify-between pt-3 border-t border-teal-800/50">
                          <div className="flex items-center text-emerald-400 text-sm font-semibold gap-1.5 group-hover:translate-x-1 transition-transform">
                            <Eye className="w-4 h-4" /> Xem chi tiết
                          </div>
                          {session.compareResult && (
                            <span className="text-xs font-medium text-emerald-300 bg-emerald-900/30 px-2 py-1 rounded-md flex items-center gap-1 border border-emerald-800/50">
                              <GitCompare className="w-3 h-3" /> Đã so sánh
                            </span>
                          )}
                        </div>
                      </motion.div>
                    ))}
                  </div>
                )}
              </div>
            ) : activeTab === 'main' ? (
              <>
                {!result && !isGenerating ? (
                  <motion.div 
                    initial={{ opacity: 0, y: 20 }}
                    animate={{ opacity: 1, y: 0 }}
                    className="h-full flex flex-col items-center justify-center text-teal-400/50 py-32"
                  >
                    <div className="w-20 h-20 bg-teal-900/50 rounded-2xl flex items-center justify-center mb-6 shadow-sm border border-teal-700/50">
                      <Lightbulb className="w-10 h-10 text-emerald-500/50" />
                    </div>
                    <h2 className="text-xl font-bold text-teal-100 mb-2">Chưa có dữ liệu</h2>
                    <p className="text-teal-300 text-center max-w-sm">
                      Hãy điền các thông tin cần thiết ở thanh bên trái và nhấn <strong className="text-teal-50">Tạo Ý Tưởng</strong> để bắt đầu.
                    </p>
                  </motion.div>
                ) : isGenerating && !result ? (
                  <div className="bg-teal-900/40 backdrop-blur-sm rounded-2xl shadow-sm border border-teal-700/50 p-12 flex flex-col items-center justify-center py-32 text-emerald-400">
                    <div className="relative mb-8">
                      <div className="w-20 h-20 border-4 border-emerald-500/20 border-t-emerald-500 rounded-full animate-spin"></div>
                      <div className="absolute inset-0 flex items-center justify-center">
                        <Sparkles className="w-8 h-8 text-emerald-400 animate-pulse" />
                      </div>
                    </div>
                    <h3 className="text-xl font-bold mb-4 text-teal-50 text-center px-4">{loadingMessage}</h3>
                    <p className="font-medium text-center max-w-md text-teal-300">
                      Hệ thống đang phân tích bằng AI nâng cao để đảm bảo các ý tưởng đề xuất có tính mới và sáng tạo cao nhất. Quá trình này có thể mất thêm chút thời gian.
                    </p>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 gap-8">
                    <div className="bg-teal-900/40 backdrop-blur-sm rounded-2xl shadow-sm border border-teal-700/50 p-8 lg:p-12">
                      {renderResultSections(result)}
                      <div ref={resultEndRef} />
                    </div>
                  </div>
                )}

                {/* Action Buttons at the bottom of main result */}
                {result && !isGenerating && (
                  <motion.div 
                    initial={{ opacity: 0, y: 20 }}
                    animate={{ opacity: 1, y: 0 }}
                    className="mt-8 bg-teal-900/40 backdrop-blur-sm border border-teal-700/50 flex flex-col sm:flex-row gap-4 items-center p-6 rounded-2xl shadow-sm"
                  >
                    <div className="flex-1">
                      <h3 className="text-sm font-bold text-teal-50 mb-1">Hành động tiếp theo</h3>
                      <p className="text-xs text-teal-300">Tạo lại danh sách mới hoặc so sánh chuyên sâu một ý tưởng.</p>
                    </div>
                    
                    <div className="flex flex-wrap items-center gap-3 w-full sm:w-auto">
                      <button
                        onClick={() => generateIdeas(true)}
                        className="flex-1 sm:flex-none px-5 py-2.5 bg-teal-800/50 border border-teal-600/50 hover:bg-teal-700/50 hover:border-teal-500/50 text-teal-100 rounded-xl font-semibold flex items-center justify-center gap-2 transition-all shadow-sm active:scale-[0.98]"
                      >
                        <RefreshCw className="w-4 h-4" />
                        Tìm Lại
                      </button>
                      
                      <div className="flex items-center gap-2 flex-1 sm:flex-none">
                        <button
                          onClick={() => {
                            generateAndDownloadWordDoc(result, `Ý Tưởng Sáng Tạo - ${field}`);
                          }}
                          className="px-5 py-2.5 bg-teal-800/50 border border-teal-600/50 hover:bg-teal-700/50 hover:border-teal-500/50 text-teal-100 rounded-xl font-semibold flex items-center justify-center gap-2 transition-all shadow-sm active:scale-[0.98]"
                        >
                          <Download className="w-4 h-4" />
                          Tải Word
                        </button>
                      </div>
                      
                      <div className="flex items-center gap-2 flex-1 sm:flex-none">
                        <select 
                          id="compareSelect"
                          className="p-2.5 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm font-medium focus:ring-2 focus:ring-emerald-500/20 outline-none text-teal-50"
                        >
                          {Array.from({ length: 20 }, (_, index) => index + 1).map(num => (
                            <option key={num} value={num} className="bg-teal-900 text-teal-50">Ý tưởng {num}</option>
                          ))}
                        </select>
                        <button
                          onClick={() => {
                            const select = document.getElementById('compareSelect') as HTMLSelectElement;
                            compareIdea(parseInt(select.value));
                          }}
                          disabled={isComparing}
                          className="px-5 py-2.5 bg-emerald-900/30 text-emerald-300 hover:bg-emerald-800/40 border border-emerald-700/50 rounded-xl font-semibold flex items-center justify-center gap-2 transition-all active:scale-[0.98]"
                        >
                          {isComparing ? <Loader2 className="w-4 h-4 animate-spin" /> : <GitCompare className="w-4 h-4" />}
                          So Sánh
                        </button>
                      </div>
                    </div>
                  </motion.div>
                )}
              </>
            ) : (
              <div className="bg-teal-900/40 backdrop-blur-sm rounded-2xl shadow-sm border border-teal-700/50 p-8 lg:p-12">
                {isComparing ? (
                  <div className="flex flex-col items-center justify-center py-32 text-emerald-400">
                    <Loader2 className="w-12 h-12 animate-spin mb-6 opacity-80" />
                    <h3 className="text-xl font-bold mb-2 text-teal-50">Đang phân tích và so sánh chuyên sâu...</h3>
                    <p className="font-medium text-center max-w-md text-teal-300">
                      Hệ thống đang phân tích bằng DeepSeek V4 Pro để đưa ra đánh giá khách quan nhất.
                    </p>
                  </div>
                ) : (
                  <ReactMarkdown 
                    remarkPlugins={[remarkGfm]}
                    components={getMarkdownComponents(compareResult)}
                  >
                    {compareResult}
                  </ReactMarkdown>
                )}
              </div>
            )}
          </div>
        </div>
      </main>
    </div>
  );
}
