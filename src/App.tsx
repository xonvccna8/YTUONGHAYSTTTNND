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

type ContestType = 'creative' | 'khkt';
type KhktProjectType = 'auto' | 'science' | 'engineering';

type KhktFieldGuide = {
  officialScope: string;
  researchFocus: string;
  breakthroughAngles: string[];
  feasibleEvidence: string;
};

type CreativeFieldGuide = {
  officialScope: string;
  productFocus: string;
  breakthroughAngles: string[];
  feasibleEvidence: string;
};

const CONTEST_OPTIONS: {
  id: ContestType;
  label: string;
  shortLabel: string;
  title: string;
  resultTab: string;
  actionLabel: string;
  docTitle: string;
  emptyTitle: string;
}[] = [
  {
    id: 'creative',
    label: 'Sáng tạo TTNND',
    shortLabel: 'TTNND',
    title: 'Cuộc thi Sáng tạo Thanh thiếu niên Nhi đồng',
    resultTab: 'Kết quả Ý Tưởng',
    actionLabel: 'Tạo Ý Tưởng',
    docTitle: 'Báo cáo ý tưởng sáng tạo',
    emptyTitle: 'Chưa có dữ liệu',
  },
  {
    id: 'khkt',
    label: 'KHKT Quốc gia',
    shortLabel: 'KHKT',
    title: 'Cuộc thi Khoa học, kỹ thuật học sinh trung học',
    resultTab: 'Đề tài KHKT',
    actionLabel: 'Tạo Đề Tài KHKT',
    docTitle: 'Báo cáo đề tài KHKT',
    emptyTitle: 'Chưa có đề tài KHKT',
  },
];

const CREATIVE_FIELD_OPTIONS = [
  'Đồ dùng dành cho học tập',
  'Phần mềm tin học và chuyển đổi số',
  'Sản phẩm thân thiện với môi trường',
  'Các dụng cụ sinh hoạt gia đình và đồ chơi trẻ em',
  'Các giải pháp kỹ thuật nhằm ứng phó với biến đổi khí hậu, bảo vệ môi trường và phát triển kinh tế',
];

const CREATIVE_FIELD_GUIDES: Record<string, CreativeFieldGuide> = {
  'Đồ dùng dành cho học tập': {
    officialScope: 'Mô hình, thiết bị, bộ học liệu, dụng cụ thí nghiệm, đồ dùng hỗ trợ tự học, dạy học, trải nghiệm STEM, giáo dục hòa nhập và học tập an toàn.',
    productFocus: 'Sản phẩm phải giúp học sinh hoặc giáo viên học/dạy tốt hơn bằng một cơ chế cụ thể: trực quan hóa kiến thức, luyện tập cá nhân hóa, đo - phản hồi, hỗ trợ khuyết tật, giảm chi phí hoặc tăng hứng thú học.',
    breakthroughAngles: [
      'Đồ dùng học tập có phản hồi thông minh: đo lỗi, gợi ý bước tiếp theo, cá nhân hóa cho từng học sinh.',
      'Bộ học liệu vật lý số kết hợp mô hình thật với cảm biến/AR/AI nhưng vẫn rẻ, dễ làm và dùng được ở lớp học Việt Nam.',
      'Thiết bị hỗ trợ nhóm học sinh yếu thế: khiếm thị, nghe kém, rối loạn đọc, học sinh vùng khó khăn hoặc thiếu Internet.'
    ],
    feasibleEvidence: 'Mô hình hoạt động được, kịch bản dùng trong tiết học, so sánh trước - sau hoặc nhóm đối chứng nhỏ, khảo sát giáo viên/học sinh, chi phí chế tạo, độ bền và mức độ an toàn.'
  },
  'Phần mềm tin học và chuyển đổi số': {
    officialScope: 'Website, hệ thống số, ứng dụng máy tính/di động, phần mềm điều khiển, AI, dữ liệu, tự động hóa quy trình và giải pháp chuyển đổi số phục vụ học tập, gia đình, cộng đồng.',
    productFocus: 'Không chỉ làm app đẹp; phải có thuật toán, dữ liệu, luồng nghiệp vụ hoặc mô hình chuyển đổi số giải quyết một việc thật, có người dùng thật và đo được hiệu quả.',
    breakthroughAngles: [
      'AI trợ lý học tập, sức khỏe, môi trường hoặc văn hóa địa phương có dữ liệu riêng, kiểm soát sai lệch và giải thích được.',
      'Nền tảng chuyển đổi số cho trường/lớp/xã: giảm thao tác thủ công, minh bạch dữ liệu, có dashboard và quyền truy cập rõ.',
      'Phần mềm kết nối với phần cứng/IoT để điều khiển, cảnh báo, thu thập dữ liệu hoặc tự động hóa quy trình thực tế.'
    ],
    feasibleEvidence: 'Demo chạy được, tài khoản/đường dẫn truy cập, bộ cài hoặc video giả lập, dữ liệu mẫu, sơ đồ luồng, metric như độ chính xác/thời gian xử lý/tỉ lệ hoàn thành, khảo sát người dùng và kê khai thư viện bên thứ ba.'
  },
  'Sản phẩm thân thiện với môi trường': {
    officialScope: 'Sản phẩm sử dụng vật liệu xanh, tái chế, tái sử dụng, tiết kiệm tài nguyên, giảm rác thải, giảm ô nhiễm hoặc thay thế vật liệu/sản phẩm gây hại môi trường.',
    productFocus: 'Sản phẩm cần chứng minh thân thiện môi trường bằng vật liệu, vòng đời sử dụng, khả năng tái chế/phân hủy, độ bền, chi phí và tác động thực tế, không dừng ở đồ thủ công trang trí.',
    breakthroughAngles: [
      'Vật liệu mới từ phụ phẩm địa phương có tính năng đo được: chống ẩm, cách nhiệt, hấp phụ, chịu lực, phân hủy hoặc tái dùng nhiều vòng.',
      'Sản phẩm thay thế nhựa dùng một lần hoặc vật liệu độc hại trong trường học/gia đình với trải nghiệm dùng tốt hơn.',
      'Thiết kế tuần hoàn: thu gom - tái chế - sử dụng - đo tác động, biến rác thành sản phẩm có giá trị thật.'
    ],
    feasibleEvidence: 'Mẫu vật liệu/sản phẩm, nguồn nguyên liệu, thử nghiệm độ bền/khả năng thấm/cách nhiệt/khối lượng rác giảm, so sánh với sản phẩm cũ, chi phí, số lần tái sử dụng và hình ảnh quy trình.'
  },
  'Các dụng cụ sinh hoạt gia đình và đồ chơi trẻ em': {
    officialScope: 'Dụng cụ hỗ trợ sinh hoạt, an toàn gia đình, chăm sóc trẻ em/người già/người khuyết tật, đồ chơi giáo dục, đồ chơi vận động, mô hình vui học và sản phẩm nâng chất lượng sống hằng ngày.',
    productFocus: 'Sản phẩm phải giải quyết tình huống sinh hoạt cụ thể hoặc tạo trải nghiệm chơi - học có giá trị; trọng tâm là an toàn, tiện dụng, dễ chế tạo, dùng được thật và phù hợp trẻ em.',
    breakthroughAngles: [
      'Dụng cụ gia đình thông minh chi phí thấp giúp phòng tai nạn, tiết kiệm thời gian, hỗ trợ người yếu thế hoặc trẻ nhỏ.',
      'Đồ chơi giáo dục có cơ chế học qua chơi: logic, khoa học, lịch sử, văn hóa, cảm xúc xã hội hoặc kỹ năng an toàn.',
      'Thiết kế module tháo lắp, tự sửa, tái chế hoặc cá nhân hóa theo lứa tuổi để sản phẩm có vòng đời dài hơn.'
    ],
    feasibleEvidence: 'Nguyên mẫu đúng kích thước an toàn, thử tải/độ bền, thử với người dùng mục tiêu có giám sát, mô tả rủi ro và cách giảm rủi ro, chi phí, hướng dẫn sử dụng và video vận hành.'
  },
  'Các giải pháp kỹ thuật nhằm ứng phó với biến đổi khí hậu, bảo vệ môi trường và phát triển kinh tế': {
    officialScope: 'Thiết bị, máy móc, robot, hệ thống tự động hóa, cảnh báo, quan trắc, xử lý, sản xuất thông minh hoặc giải pháp kỹ thuật giúp thích ứng khí hậu, bảo vệ môi trường và tạo giá trị kinh tế.',
    productFocus: 'Khác với sản phẩm xanh đơn lẻ, lĩnh vực này cần một hệ thống/giải pháp kỹ thuật có tác động vận hành: đo, cảnh báo, xử lý, tối ưu sản xuất, giảm thiệt hại hoặc tăng năng suất cho cộng đồng.',
    breakthroughAngles: [
      'Hệ cảnh báo sớm thiên tai/ô nhiễm/nắng nóng/ngập mặn bằng cảm biến, dữ liệu địa phương và kênh thông báo dễ dùng.',
      'Thiết bị nông nghiệp/thuỷ sản/tiểu thủ công nghiệp giúp tiết kiệm nước, năng lượng, nhân công hoặc giảm thất thoát sau thu hoạch.',
      'Giải pháp kinh tế xanh cho địa phương: biến phụ phẩm thành sản phẩm, tối ưu quy trình sản xuất nhỏ, đo được lợi ích môi trường và lợi ích kinh tế.'
    ],
    feasibleEvidence: 'Mô hình hệ thống, sơ đồ nguyên lý, số liệu trước - sau, thử nghiệm ngoài hiện trường hoặc mô phỏng thực tế, chi phí/hiệu suất, độ ổn định, khả năng nhân rộng và video vận hành.'
  }
};

const CREATIVE_NATIONAL_INNOVATION_PRINCIPLES = `
NGUYÊN TẮC TÌM Ý TƯỞNG SÁNG TẠO TTNND CẤP QUỐC GIA:
- Bám thể lệ toàn quốc lần thứ 22 năm 2026: sản phẩm thuộc 5 lĩnh vực, có tính mới, tính sáng tạo, có mô hình/sản phẩm hoặc video vận hành, vật liệu an toàn, thuyết minh rõ ý tưởng - phương pháp - vật liệu - cách dùng - khả năng áp dụng.
- Ưu tiên sản phẩm "nhìn thấy, cầm được, chạy được": mô hình thông minh, dụng cụ đa năng, thiết bị máy móc, robot, tự động hóa, phần mềm điều khiển hoặc hệ thống số có demo thật.
- Đột phá phải nằm ở ít nhất một khâu: nguyên lý, kết cấu, vật liệu, tính năng, công dụng, thiết kế, phương thức triển khai, nhóm người hưởng lợi hoặc khả năng nhân rộng.
- Ý tưởng đạt giải cao thường giải quyết vấn đề rất cụ thể tại trường học, gia đình, địa phương; có người dùng thật, dữ liệu/đánh giá trước - sau, chi phí hợp lý và câu chuyện thuyết minh dễ bảo vệ.
- Tránh ý tưởng sáo mòn nếu không có cơ chế mới: thùng rác thông minh, app nhắc học, robot tưới cây, máy lọc nước mini, đồ tái chế trang trí, mô hình cảnh báo đơn giản chỉ bật còi/đèn.
`.trim();

function getCreativeFieldGuide(field: string): CreativeFieldGuide | undefined {
  return CREATIVE_FIELD_GUIDES[field];
}

function buildCreativeFieldCataloguePrompt() {
  return CREATIVE_FIELD_OPTIONS.map((field, index) => {
    const guide = getCreativeFieldGuide(field);
    return `${index + 1}. ${field}: ${guide?.officialScope || 'Thuộc danh mục 5 lĩnh vực chính thức.'} Trọng tâm sản phẩm: ${guide?.productFocus || 'Cần có mô hình, tính mới, tính sáng tạo và khả năng áp dụng.'}`;
  }).join('\n');
}

function buildCreativeSelectedFieldPrompt(field: string) {
  const guide = getCreativeFieldGuide(field);
  if (!guide) {
    return `Lĩnh vực đang chọn: ${field}. Hãy xác định đúng dạng sản phẩm, tránh nhầm sang lĩnh vực khác, và đề xuất mô hình có thể chế tạo hoặc demo được.`;
  }

  return `
LĨNH VỰC ĐANG CHỌN: ${field}
- Phạm vi sản phẩm: ${guide.officialScope}
- Trọng tâm cần hiểu: ${guide.productFocus}
- Góc đột phá nên ưu tiên:
${guide.breakthroughAngles.map((angle, index) => `  ${index + 1}. ${angle}`).join('\n')}
- Minh chứng khả thi cần có: ${guide.feasibleEvidence}
`.trim();
}

const KHKT_FIELD_OPTIONS = [
  'Khoa học động vật',
  'Khoa học xã hội và hành vi',
  'Hóa sinh',
  'Y sinh và khoa học sức khỏe',
  'Kĩ thuật y sinh',
  'Sinh học tế bào và phân tử',
  'Hóa học',
  'Sinh học trên máy tính và Sinh - Tin',
  'Khoa học Trái đất và Môi trường',
  'Hệ thống nhúng',
  'Năng lượng: Hóa học',
  'Năng lượng: Vật lí',
  'Kĩ thuật cơ khí',
  'Kĩ thuật môi trường',
  'Khoa học vật liệu',
  'Toán học',
  'Vi sinh',
  'Vật lí và Thiên văn',
  'Khoa học thực vật',
  'Rô-bốt và máy thông minh',
  'Phần mềm hệ thống',
  'Y học chuyển dịch',
];

const KHKT_FIELD_GUIDES: Record<string, KhktFieldGuide> = {
  'Khoa học động vật': {
    officialScope: 'Đời sống động vật: cấu trúc, sinh lý, phát triển, phân loại, hành vi, sinh thái, dinh dưỡng, tăng trưởng, di truyền, chăn nuôi, thủy sản và côn trùng.',
    researchFocus: 'Tập trung vào một loài/nhóm loài cụ thể, đo được thay đổi hành vi, tăng trưởng, sức khỏe, môi trường sống hoặc tương tác sinh thái bằng quan sát định lượng.',
    breakthroughAngles: [
      'Theo dõi hành vi/sức khỏe không xâm lấn bằng cảm biến rẻ tiền hoặc thị giác máy tính.',
      'Giải pháp giảm stress, tăng phúc lợi, tăng hiệu quả nuôi trồng cho vật nuôi địa phương.',
      'Mô hình cảnh báo sớm dịch hại/côn trùng có ích dựa trên dữ liệu môi trường.'
    ],
    feasibleEvidence: 'Nhật ký quan sát, ảnh/video có nhãn, chỉ số tăng trưởng, điều kiện môi trường, nhóm đối chứng và phân tích thống kê đơn giản.'
  },
  'Khoa học xã hội và hành vi': {
    officialScope: 'Hành vi, nhận thức, tâm lý, giáo dục, xã hội học, nhân học, tương tác nhóm người và các yếu tố ảnh hưởng đến quyết định của con người.',
    researchFocus: 'Thiết kế khảo sát, thí nghiệm hành vi hoặc can thiệp giáo dục có biến độc lập, biến phụ thuộc, mẫu khảo sát rõ và xử lý thiên lệch dữ liệu.',
    breakthroughAngles: [
      'Can thiệp hành vi nhỏ nhưng đo được tác động lớn trong học tập, an toàn số, môi trường hoặc sức khỏe tinh thần học đường.',
      'Kết hợp dữ liệu khảo sát với nhật ký số/AI phân tích văn bản để phát hiện mẫu hành vi mới.',
      'Mô hình ra quyết định phù hợp văn hóa địa phương thay vì sao chép nghiên cứu nước ngoài.'
    ],
    feasibleEvidence: 'Bảng hỏi chuẩn hóa, phiếu đồng thuận, trước-sau can thiệp, nhóm so sánh, thống kê mô tả, kiểm định đơn giản và phân tích hạn chế mẫu.'
  },
  'Hóa sinh': {
    officialScope: 'Hóa sinh phân tích, hóa sinh tổng hợp, hóa sinh y, hóa sinh cấu trúc; nghiên cứu phân tử, enzyme, protein, lipid, đường, chất chuyển hóa trong hệ sống.',
    researchFocus: 'Làm rõ cơ chế hóa học trong hệ sinh học bằng phép đo định lượng an toàn như pH, quang phổ màu, hoạt tính enzyme, khả năng chống oxy hóa hoặc mô phỏng phân tử.',
    breakthroughAngles: [
      'Tận dụng phụ phẩm nông nghiệp địa phương để tạo hoạt chất sinh học có dữ liệu định lượng.',
      'So sánh cơ chế chiết tách, bảo quản hoặc ổn định hoạt chất thay vì chỉ làm sản phẩm thô.',
      'Kết hợp mô phỏng tính chất phân tử với thí nghiệm đơn giản để tăng chiều sâu khoa học.'
    ],
    feasibleEvidence: 'Quy trình chiết/đo lặp lại, đường chuẩn hoặc thang màu, so sánh mẫu đối chứng, biểu đồ sai số và giải thích cơ chế hóa sinh.'
  },
  'Y sinh và khoa học sức khỏe': {
    officialScope: 'Chẩn đoán, điều trị, dược liệu, dịch tễ học, dinh dưỡng, sinh lý học, bệnh lý học và các yếu tố ảnh hưởng sức khỏe con người.',
    researchFocus: 'Ưu tiên nghiên cứu an toàn, không can thiệp y khoa rủi ro: sàng lọc, dự báo, dinh dưỡng, hành vi sức khỏe, thiết bị hỗ trợ hoặc phân tích dữ liệu công khai.',
    breakthroughAngles: [
      'Mô hình sàng lọc sớm nguy cơ sức khỏe học đường bằng dữ liệu phi xâm lấn.',
      'Cá nhân hóa khuyến nghị dinh dưỡng/vận động dựa trên dữ liệu đo được, có đối chứng.',
      'Ứng dụng AI giải thích được cho dữ liệu sức khỏe công khai hoặc dữ liệu tự thu thập hợp lệ.'
    ],
    feasibleEvidence: 'Dữ liệu ẩn danh, tiêu chí đạo đức, chỉ số sức khỏe phi xâm lấn, so sánh baseline, độ chính xác/độ nhạy/độ đặc hiệu nếu có mô hình dự báo.'
  },
  'Kĩ thuật y sinh': {
    officialScope: 'Vật liệu y sinh, cơ chế sinh học, thiết bị y sinh, kỹ thuật tế bào và mô, sinh học tổng hợp theo hướng thiết kế công cụ hoặc hệ thống hỗ trợ y tế.',
    researchFocus: 'Tạo nguyên mẫu thiết bị, cảm biến, vật liệu hoặc hệ hỗ trợ phục hồi/chăm sóc sức khỏe, chứng minh bằng thử nghiệm mô phỏng an toàn.',
    breakthroughAngles: [
      'Thiết bị trợ giúp người khuyết tật có AI/cảm biến, chi phí thấp, dùng được trong bối cảnh Việt Nam.',
      'Cảm biến y sinh phi xâm lấn kết hợp cảnh báo thông minh và kiểm thử độ tin cậy.',
      'Vật liệu/thiết kế mô phỏng chức năng y sinh nhưng không đụng thử nghiệm nguy hiểm.'
    ],
    feasibleEvidence: 'Bản vẽ kỹ thuật, nguyên mẫu, sai số cảm biến, độ trễ, độ bền, thử nghiệm với mô hình giả lập và phản hồi người dùng mục tiêu.'
  },
  'Sinh học tế bào và phân tử': {
    officialScope: 'Sinh lý tế bào, gen, miễn dịch, sinh học phân tử, sinh học thần kinh và các quá trình ở cấp tế bào/phân tử.',
    researchFocus: 'Chọn hệ mô hình an toàn như thực vật, nấm men thực phẩm, dữ liệu công khai hoặc mô phỏng để nghiên cứu biểu hiện, phát triển, stress sinh học hoặc tương tác phân tử.',
    breakthroughAngles: [
      'Dùng dữ liệu mở về gen/protein để tìm dấu hiệu sinh học có thể kiểm chứng bằng mô phỏng.',
      'Nghiên cứu phản ứng stress ở tế bào thực vật với biến môi trường gần gũi.',
      'Kết hợp kính hiển vi/ảnh tế bào với AI phân loại định lượng.'
    ],
    feasibleEvidence: 'Ảnh hiển vi, chỉ số hình thái, dữ liệu công khai có nguồn, quy trình an toàn sinh học, lặp thí nghiệm và phân tích định lượng.'
  },
  'Hóa học': {
    officialScope: 'Hóa phân tích, hóa học trên máy tính, hóa môi trường, hóa vô cơ, hóa vật liệu, hóa hữu cơ, hóa lý; nghiên cứu thành phần, cấu trúc, tính chất và phản ứng của vật chất.',
    researchFocus: 'Đặt câu hỏi hóa học có thể đo: hấp phụ, xúc tác, đổi màu, pH, độ dẫn, tốc độ phản ứng, tính bền, khả năng phân hủy hoặc mô phỏng cấu trúc.',
    breakthroughAngles: [
      'Vật liệu xanh xử lý ô nhiễm với cơ chế hấp phụ/xúc tác được chứng minh.',
      'Cảm biến hóa học màu hoặc điện hóa giá rẻ cho vấn đề trường học/địa phương.',
      'Tối ưu hóa phản ứng bằng thiết kế thí nghiệm và mô hình dự báo.'
    ],
    feasibleEvidence: 'Mẫu đối chứng, đường chuẩn, ảnh màu chuẩn hóa, dữ liệu pH/độ dẫn/khối lượng, hiệu suất, tái sử dụng vật liệu và biểu đồ sai số.'
  },
  'Sinh học trên máy tính và Sinh - Tin': {
    officialScope: 'Mô hình sinh học tính toán, dịch tễ học tính toán, tiến hóa, thần kinh tính toán, dược lý tính toán, genomics và phân tích dữ liệu sinh học.',
    researchFocus: 'Dùng thuật toán, mô hình toán, mô phỏng hoặc AI để phân tích hệ sinh học, dịch bệnh, gen, protein, thuốc hoặc hành vi sinh học từ dữ liệu đáng tin.',
    breakthroughAngles: [
      'Khai thác bộ dữ liệu mở quốc tế để tìm dấu hiệu sinh học mới, có khả năng giải thích.',
      'Mô hình dự báo dịch/bệnh/cây trồng có yếu tố địa phương và kiểm định trên dữ liệu thật.',
      'AI sinh học giải thích được, so sánh nhiều thuật toán và tránh hộp đen.'
    ],
    feasibleEvidence: 'Nguồn dữ liệu mở, tiền xử lý rõ, metric như accuracy/F1/AUC/RMSE, baseline, cross-validation và phân tích ý nghĩa sinh học.'
  },
  'Khoa học Trái đất và Môi trường': {
    officialScope: 'Khí quyển, khí hậu, tác động môi trường lên hệ sinh thái, địa chất, nước và các hệ Trái đất.',
    researchFocus: 'Nghiên cứu hiện tượng môi trường và hệ quả của nó bằng đo đạc/quan sát: vi khí hậu, nước, đất, chất lượng không khí, xói mòn, ngập, đa dạng sinh học.',
    breakthroughAngles: [
      'Bản đồ rủi ro môi trường cấp trường/xã bằng cảm biến, GIS và dữ liệu vệ tinh mở.',
      'Mô hình dự báo vi khí hậu, ngập cục bộ hoặc chất lượng nước có kiểm chứng thực địa.',
      'Liên hệ dữ liệu môi trường với tác động lên cây trồng, sức khỏe học đường hoặc hệ sinh thái.'
    ],
    feasibleEvidence: 'Chuỗi đo theo thời gian, tọa độ mẫu, ảnh hiện trường, dữ liệu vệ tinh/công khai, bản đồ, so sánh mùa/địa điểm và kiểm định tương quan.'
  },
  'Hệ thống nhúng': {
    officialScope: 'Mạch điện, vi điều khiển, IoT, truyền thông dữ liệu, quang học, cảm biến và xử lý tín hiệu trong hệ thống điện tử điều khiển/cảm nhận.',
    researchFocus: 'Thiết kế hệ phần cứng-phần mềm đo, truyền, xử lý hoặc điều khiển tín hiệu; điểm mạnh nằm ở độ ổn định, sai số, năng lượng, bảo mật và thử nghiệm thực tế.',
    breakthroughAngles: [
      'Mạng cảm biến rẻ tiền nhưng tự hiệu chuẩn, tiết kiệm năng lượng và hoạt động ngoài hiện trường.',
      'Xử lý tín hiệu tại biên để cảnh báo nhanh mà không phụ thuộc Internet.',
      'Thiết kế IoT có bảo mật, khả năng chịu lỗi và dữ liệu kiểm thử nhiều điều kiện.'
    ],
    feasibleEvidence: 'Sơ đồ mạch, mã nhúng, sai số cảm biến, độ trễ, pin, tỉ lệ mất gói, log dữ liệu, thử nghiệm nóng/ẩm/nhiễu và so sánh thiết bị chuẩn.'
  },
  'Năng lượng: Hóa học': {
    officialScope: 'Nhiên liệu thay thế, năng lượng hóa thạch, pin, tế bào nhiên liệu, vật liệu năng lượng mặt trời và quá trình chuyển hóa/lưu trữ năng lượng theo hướng hóa học.',
    researchFocus: 'Nghiên cứu vật liệu, phản ứng hoặc hệ lưu trữ/chuyển hóa năng lượng có đo hiệu suất, điện áp, dòng, độ bền, chu kỳ sạc-xả hoặc sản lượng nhiên liệu.',
    breakthroughAngles: [
      'Pin/supercapacitor từ vật liệu sinh khối, than hoạt tính, polymer hoặc composite xanh.',
      'Tế bào nhiên liệu vi sinh/sinh khối an toàn với tối ưu vật liệu điện cực.',
      'Vật liệu hấp thụ/quang xúc tác tăng hiệu suất thu năng lượng mặt trời.'
    ],
    feasibleEvidence: 'Đường cong sạc-xả, điện áp/dòng/công suất, số chu kỳ, hiệu suất, khối lượng vật liệu, điều kiện thí nghiệm và so sánh mẫu thương mại hoặc đối chứng.'
  },
  'Năng lượng: Vật lí': {
    officialScope: 'Thủy điện, năng lượng mặt trời, nhiệt, gió, hạt nhân ở mức mô phỏng an toàn và các hệ chuyển đổi năng lượng dựa trên hiện tượng vật lý.',
    researchFocus: 'Thiết kế/mô phỏng/kiểm thử hệ thu, chuyển đổi, tối ưu năng lượng bằng các đại lượng vật lý như công suất, hiệu suất, góc, tốc độ gió, nhiệt độ, bức xạ.',
    breakthroughAngles: [
      'Cơ cấu tối ưu hóa thu năng lượng trong điều kiện gió/nắng không ổn định ở địa phương.',
      'Thu hồi nhiệt thải hoặc rung động nhỏ bằng thiết kế vật lý chi phí thấp.',
      'Mô hình lai năng lượng gió-mặt trời-nhiệt có bộ điều khiển tối ưu.'
    ],
    feasibleEvidence: 'Bộ đo công suất, dữ liệu thời gian thực, so sánh cấu hình, hiệu suất, mô phỏng và kiểm thử ngoài trời nhiều điều kiện.'
  },
  'Kĩ thuật cơ khí': {
    officialScope: 'Hàng không, dân dụng, cơ học tính toán, điều khiển, phương tiện mặt đất, gia công công nghiệp, máy móc, cơ khí và hệ hàng hải.',
    researchFocus: 'Giải quyết vấn đề chuyển động, lực, kết cấu, truyền động, độ bền, tối ưu hình học hoặc điều khiển cơ khí bằng mô hình/thiết kế/nguyên mẫu.',
    breakthroughAngles: [
      'Cơ cấu mới giúp giảm lực, tiết kiệm năng lượng hoặc tăng an toàn trong lao động/học đường.',
      'Tối ưu kết cấu bằng mô phỏng và in 3D/ gia công đơn giản.',
      'Thiết bị cơ điện tử phục vụ nông nghiệp, người yếu thế hoặc ứng phó thiên tai.'
    ],
    feasibleEvidence: 'Bản CAD, mô phỏng lực, nguyên mẫu, tải trọng, độ bền, hiệu suất cơ học, thử nghiệm lặp lại và so sánh thiết kế cũ-mới.'
  },
  'Kĩ thuật môi trường': {
    officialScope: 'Xử lý môi trường bằng sinh học, cải tạo đất, kiểm soát ô nhiễm, tái chế/quản lý chất thải và quản lý nguồn nước.',
    researchFocus: 'Tạo quy trình, vật liệu, thiết bị hoặc mô hình hạ tầng để xử lý nước/khí/rác/đất; phải chứng minh hiệu quả bằng chỉ số môi trường.',
    breakthroughAngles: [
      'Hệ xử lý nước/rác quy mô trường học dùng vật liệu địa phương, có cơ chế khoa học rõ.',
      'Tái chế chất thải thành vật liệu có tính năng đo được thay vì sản phẩm trang trí.',
      'Giải pháp quan trắc và dự báo ô nhiễm kết hợp IoT, bản đồ và khuyến nghị vận hành.'
    ],
    feasibleEvidence: 'Chỉ số trước-sau như độ đục, pH, TDS, COD mô phỏng an toàn, khối lượng rác giảm, độ bền vật liệu, chi phí và thử nghiệm nhiều mẫu.'
  },
  'Khoa học vật liệu': {
    officialScope: 'Vật liệu sinh học, gốm-thủy tinh, composite, vật liệu điện tử/quang/từ, nano, polymer, cấu trúc và tính chất vật liệu.',
    researchFocus: 'Nghiên cứu mối liên hệ giữa thành phần-cấu trúc-quy trình-tính chất; điểm mạnh là đo cơ tính, hấp phụ, dẫn điện, cách nhiệt, quang học hoặc độ bền.',
    breakthroughAngles: [
      'Composite xanh từ phụ phẩm địa phương có tính năng vượt vật liệu nền.',
      'Vật liệu thông minh đổi màu, tự làm sạch, hấp phụ, cách nhiệt hoặc cảm biến.',
      'Vật liệu nano/polymer an toàn với cơ chế và khả năng tái sử dụng rõ.'
    ],
    feasibleEvidence: 'Công thức vật liệu, quy trình chế tạo, ảnh cấu trúc, độ bền kéo/nén/uốn, hấp phụ, dẫn điện/nhiệt, chu kỳ tái sử dụng và so sánh vật liệu chuẩn.'
  },
  'Toán học': {
    officialScope: 'Đại số, giải tích, toán rời rạc, lý thuyết trò chơi và đồ thị, hình học-tô pô, lý thuyết số, xác suất và thống kê.',
    researchFocus: 'Tạo định lý, mô hình, thuật toán, chứng minh hoặc phương pháp thống kê mới/hiệu quả hơn cho một bài toán rõ ràng.',
    breakthroughAngles: [
      'Mô hình tối ưu hóa lịch học, giao thông, mạng lưới hoặc phân bổ tài nguyên địa phương.',
      'Thuật toán đồ thị/xác suất có chứng minh và thử nghiệm trên dữ liệu thật.',
      'Mô hình thống kê phát hiện bất thường, dự báo rủi ro hoặc đánh giá công bằng.'
    ],
    feasibleEvidence: 'Phát biểu bài toán, chứng minh, độ phức tạp, mô phỏng, dữ liệu kiểm thử, so sánh thuật toán baseline và phân tích sai số.'
  },
  'Vi sinh': {
    officialScope: 'Vi khuẩn, vi sinh ứng dụng, vi sinh môi trường, vi-rút, kháng sinh và các tương tác vi sinh; cần tuân thủ an toàn sinh học.',
    researchFocus: 'Ưu tiên vi sinh an toàn như nấm men, men lactic, vi sinh thực phẩm hoặc dữ liệu công khai; nghiên cứu lên men, phân hủy, ức chế tự nhiên, môi trường hoặc ứng dụng sinh học.',
    breakthroughAngles: [
      'Quy trình lên men/vi sinh vật có ích tạo sản phẩm xanh với dữ liệu tối ưu hóa.',
      'Vi sinh môi trường xử lý chất hữu cơ, nước thải mô phỏng hoặc rác sinh học an toàn.',
      'Phân tích dữ liệu microbiome công khai bằng sinh tin để tìm mẫu mới.'
    ],
    feasibleEvidence: 'Quy trình vô trùng cơ bản, chủng an toàn, số khuẩn lạc hoặc chỉ số lên men, mẫu đối chứng, nhiệt độ/thời gian, ảnh đĩa cấy và quy định an toàn.'
  },
  'Vật lí và Thiên văn': {
    officialScope: 'Thiên văn, vật lý nguyên tử-phân tử-quang học, lý-sinh, vật lý tính toán, vật lý thiên văn, vật liệu đo, điện-từ-plasma, cơ học, quang học, laser, sóng điện từ, lượng tử và lý thuyết.',
    researchFocus: 'Đặt bài toán vật lý đo được hoặc mô phỏng được: quang, sóng, cơ, nhiệt, điện-từ, vật liệu đo, ảnh thiên văn hoặc mô hình tính toán.',
    breakthroughAngles: [
      'Thiết bị đo vật lý/thiên văn tự chế có hiệu chuẩn và xử lý tín hiệu tốt.',
      'Mô hình quang-cơ-điện từ giải quyết vấn đề thực tế như đo chất lượng, truyền thông, an toàn.',
      'Phân tích dữ liệu thiên văn mở bằng thuật toán để phát hiện/ phân loại hiện tượng.'
    ],
    feasibleEvidence: 'Thiết bị đo, hiệu chuẩn, sai số, chuỗi dữ liệu, mô phỏng, ảnh/ phổ/ tín hiệu, so sánh lý thuyết-thực nghiệm và phân tích độ không chắc chắn.'
  },
  'Khoa học thực vật': {
    officialScope: 'Nông nghiệp, tương tác với môi trường, gen và sinh sản, tăng trưởng-phát triển, bệnh lý, sinh lý, hệ thống và tiến hóa thực vật.',
    researchFocus: 'Nghiên cứu cây trồng/thực vật bằng biến môi trường, dinh dưỡng, stress, bệnh, sinh trưởng, sinh lý, năng suất hoặc tương tác hệ sinh thái.',
    breakthroughAngles: [
      'Canh tác thông minh thích ứng nắng nóng, mặn, thiếu nước hoặc ô nhiễm địa phương.',
      'Chế phẩm sinh học/vật liệu giữ nước/kích thích sinh trưởng có cơ chế và đối chứng.',
      'AI phân tích ảnh lá/cây để phát hiện stress sớm và đề xuất can thiệp.'
    ],
    feasibleEvidence: 'Chiều cao, số lá, diện tích lá, khối lượng, độ ẩm đất, EC/pH, ảnh theo thời gian, nhóm đối chứng, lặp mẫu và phân tích thống kê.'
  },
  'Rô-bốt và máy thông minh': {
    officialScope: 'Máy sinh học, lý thuyết điều khiển, robot động lực/động học, hệ nhận thức, máy học và hệ giảm phụ thuộc vào can thiệp con người.',
    researchFocus: 'Thiết kế robot/hệ thông minh có cảm nhận-quyết định-hành động, chứng minh bằng nhiệm vụ cụ thể, môi trường thay đổi và tiêu chí hiệu năng.',
    breakthroughAngles: [
      'Robot phục vụ cộng đồng có khả năng thích nghi môi trường Việt Nam thay vì chỉ chạy demo.',
      'AI biên cho nhận diện/điều khiển nhanh, ít dữ liệu, tiết kiệm năng lượng.',
      'Cơ cấu robot mềm, robot sinh học hoặc điều khiển lai mới có thử nghiệm định lượng.'
    ],
    feasibleEvidence: 'Nguyên mẫu, sơ đồ điều khiển, dữ liệu huấn luyện, tỉ lệ thành công, thời gian hoàn thành, độ chính xác, thử nghiệm nhiều địa hình/ánh sáng/tải trọng.'
  },
  'Phần mềm hệ thống': {
    officialScope: 'Thuật toán, an ninh máy tính, cơ sở dữ liệu, hệ điều hành, ngôn ngữ lập trình, giao diện người-máy và hệ phần mềm điều khiển/phân tích quy trình.',
    researchFocus: 'Không chỉ làm app; phải có thuật toán, kiến trúc, mô hình dữ liệu, bảo mật, hiệu năng, khả năng mở rộng hoặc phương pháp phần mềm được kiểm chứng.',
    breakthroughAngles: [
      'AI/thuật toán giải bài toán thực tế có bộ dữ liệu, baseline và khả năng giải thích.',
      'Hệ phần mềm an toàn, riêng tư, chống gian lận hoặc bảo vệ trẻ em trên môi trường số.',
      'Nền tảng học tập/giám sát/thông tin có đánh giá hiệu quả thực nghiệm với người dùng.'
    ],
    feasibleEvidence: 'Kho dữ liệu, mã nguồn, kiến trúc, metric accuracy/F1/latency, kiểm thử bảo mật, so sánh baseline, khảo sát người dùng và nhật ký lỗi.'
  },
  'Y học chuyển dịch': {
    officialScope: 'Khám bệnh và chẩn đoán, phòng bệnh, điều trị, kiểm định thuốc, nghiên cứu tiền lâm sàng; chuyển phát hiện y sinh thành công cụ lâm sàng/y tế công cộng.',
    researchFocus: 'Chuyển một phát hiện hoặc dữ liệu y sinh thành công cụ sàng lọc, dự báo, phòng bệnh hoặc đánh giá can thiệp; ưu tiên mô phỏng, dữ liệu mở và thử nghiệm phi lâm sàng an toàn.',
    breakthroughAngles: [
      'Công cụ chẩn đoán/sàng lọc sớm từ dữ liệu ảnh, tín hiệu hoặc triệu chứng phi xâm lấn.',
      'Mô hình phòng bệnh cá nhân hóa cho học sinh/cộng đồng với dữ liệu chứng minh tác động.',
      'Sàng lọc hợp chất/dược liệu bằng mô phỏng hoặc dữ liệu công khai, không tuyên bố điều trị khi chưa kiểm chứng.'
    ],
    feasibleEvidence: 'Dữ liệu ẩn danh hoặc mở, tiêu chí loại trừ rủi ro đạo đức, metric chẩn đoán, so sánh với baseline, phân tích sai lệch và tuyên bố giới hạn rõ.'
  }
};

const KHKT_NATIONAL_RESEARCH_PRINCIPLES = `
NGUYÊN TẮC TÌM ĐỀ TÀI KHKT CẤP QUỐC GIA:
- Bám Thông tư 06/2024/TT-BGDĐT và tinh thần sửa đổi 24/2025/TT-BGDĐT: liêm chính khoa học, công khai, minh bạch, an toàn, đúng quy định, có thể bảo vệ trước giám khảo.
- Ưu tiên đề tài có bài toán thời sự tại Việt Nam: chuyển đổi số, AI/IoT có dữ liệu thật, môi trường, biến đổi khí hậu, sức khỏe học đường, vật liệu xanh, nông nghiệp thông minh, hỗ trợ người yếu thế, an toàn số.
- Đột phá không có nghĩa là viển vông: phải có cơ chế mới hoặc cách kết hợp mới, thử nghiệm trong 12 tháng, chi phí/năng lực phù hợp học sinh, có sản phẩm/dữ liệu/bằng chứng đo được.
- Mỗi ý tưởng phải có baseline để đối sánh, tiêu chí thành công định lượng, kế hoạch thu thập dữ liệu, rủi ro/hạn chế và cách giải thích kết quả không thổi phồng.
- Tránh các đề tài sáo mòn nếu không có cơ chế mới: thùng rác thông minh, app nhắc học, robot tưới cây, máy lọc nước mini, phân loại rác bằng AI đơn giản, trồng cây với các loại nước tưới quen thuộc.
`.trim();

function getKhktFieldGuide(field: string): KhktFieldGuide | undefined {
  return KHKT_FIELD_GUIDES[field];
}

function buildKhktFieldCataloguePrompt() {
  return KHKT_FIELD_OPTIONS.map((field, index) => {
    const guide = getKhktFieldGuide(field);
    return `${index + 1}. ${field}: ${guide?.officialScope || 'Thuộc danh mục lĩnh vực KHKT chính thức.'} Trọng tâm: ${guide?.researchFocus || 'Cần có câu hỏi nghiên cứu, dữ liệu và kiểm chứng rõ ràng.'}`;
  }).join('\n');
}

function buildKhktSelectedFieldPrompt(field: string) {
  const guide = getKhktFieldGuide(field);
  if (!guide) {
    return `Lĩnh vực đang chọn: ${field}. Hãy xác định đúng phạm vi nghiên cứu, tránh nhầm sang lĩnh vực khác, và đề xuất câu hỏi/vấn đề có thể kiểm chứng.`;
  }

  return `
LĨNH VỰC ĐANG CHỌN: ${field}
- Phạm vi nghiên cứu: ${guide.officialScope}
- Trọng tâm cần hiểu: ${guide.researchFocus}
- Góc đột phá nên ưu tiên:
${guide.breakthroughAngles.map((angle, index) => `  ${index + 1}. ${angle}`).join('\n')}
- Minh chứng khả thi cần có: ${guide.feasibleEvidence}
`.trim();
}

const FIELD_OPTIONS = CREATIVE_FIELD_OPTIONS;

const CREATIVE_CAP_HOC_OPTIONS = ['Tiểu học', 'THCS', 'THPT'];
const KHKT_CAP_HOC_OPTIONS = ['THCS', 'THPT'];
const CAP_HOC_OPTIONS = CREATIVE_CAP_HOC_OPTIONS;

const CREATIVE_GRADE_OPTIONS = [
  'Tiểu học: Lớp 1–5',
  'THCS: Lớp 6–9',
  'THPT: Lớp 10–12',
];

const KHKT_GRADE_OPTIONS = [
  'THCS: Lớp 8',
  'THCS: Lớp 9',
  'THPT: Lớp 10',
  'THPT: Lớp 11',
  'THPT: Lớp 12',
];

const GRADE_OPTIONS = CREATIVE_GRADE_OPTIONS;

const TECH_LIMIT_OPTIONS = ['Cơ bản', 'Trung bình', 'Nâng cao'];

const MUC_TIEU_OPTIONS = [
  'Cấp lớp',
  'Cấp trường',
  'Cấp huyện',
  'Cấp tỉnh',
  'Cấp quốc gia',
];

const KHKT_PROJECT_TYPE_OPTIONS: { value: KhktProjectType; label: string }[] = [
  { value: 'auto', label: 'AI tự đề xuất loại phù hợp' },
  { value: 'science', label: 'Dự án khoa học' },
  { value: 'engineering', label: 'Dự án kỹ thuật' },
];

const KHKT_RUBRIC_ITEMS = [
  ['Câu hỏi/Vấn đề', '10'],
  ['Thiết kế & phương pháp', '15'],
  ['Thực hiện/kiểm chứng', '20'],
  ['Tính sáng tạo', '20'],
  ['Báo cáo', '10'],
  ['Nội dung khoa học', '25'],
];

function getContestMeta(type: ContestType) {
  return CONTEST_OPTIONS.find(option => option.id === type) || CONTEST_OPTIONS[0];
}

function getFieldOptions(type: ContestType) {
  return type === 'khkt' ? KHKT_FIELD_OPTIONS : CREATIVE_FIELD_OPTIONS;
}

function getCapHocOptions(type: ContestType) {
  return type === 'khkt' ? KHKT_CAP_HOC_OPTIONS : CREATIVE_CAP_HOC_OPTIONS;
}

function getGradeOptions(type: ContestType) {
  return type === 'khkt' ? KHKT_GRADE_OPTIONS : CREATIVE_GRADE_OPTIONS;
}

function inferContestTypeFromField(value?: string): ContestType {
  return value && KHKT_FIELD_OPTIONS.includes(value) ? 'khkt' : 'creative';
}

function getKhktProjectTypeLabel(value: KhktProjectType) {
  return KHKT_PROJECT_TYPE_OPTIONS.find(option => option.value === value)?.label || KHKT_PROJECT_TYPE_OPTIONS[0].label;
}

interface SavedSession {
  id: string;
  timestamp: number;
  inputs: {
    contestType?: ContestType;
    khktProjectType?: KhktProjectType;
    field: string;
    capHoc: string;
    grade: string;
    techLimit: string;
    mucTieu: string;
    context: string;
    resources: string;
    problem?: string;
    avoidIdeas?: string;
    localTraits?: string;
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
    contestType?: ContestType;
    khktProjectType?: KhktProjectType;
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

  const [contestType, setContestType] = useState<ContestType>('creative');
  const [khktProjectType, setKhktProjectType] = useState<KhktProjectType>('auto');
  const [field, setField] = useState(FIELD_OPTIONS[0]);
  const [capHoc, setCapHoc] = useState(CAP_HOC_OPTIONS[0]);
  const [grade, setGrade] = useState(GRADE_OPTIONS[0]);
  const [techLimit, setTechLimit] = useState(TECH_LIMIT_OPTIONS[0]);
  const [mucTieu, setMucTieu] = useState(MUC_TIEU_OPTIONS[0]);
  const [context, setContext] = useState('');
  const [resources, setResources] = useState('');
  const [problem, setProblem] = useState('');
  const [avoidIdeas, setAvoidIdeas] = useState('');
  const [localTraits, setLocalTraits] = useState('');

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
  const contestMeta = getContestMeta(contestType);
  const activeFieldOptions = getFieldOptions(contestType);
  const activeCapHocOptions = getCapHocOptions(contestType);
  const activeGradeOptions = getGradeOptions(contestType);
  const isKhktContest = contestType === 'khkt';

  useEffect(() => {
    const nextFields = getFieldOptions(contestType);
    const nextCaps = getCapHocOptions(contestType);
    const nextGrades = getGradeOptions(contestType);

    setField(current => nextFields.includes(current) ? current : nextFields[0]);
    setCapHoc(current => nextCaps.includes(current) ? current : nextCaps[0]);
    setGrade(current => nextGrades.includes(current) ? current : nextGrades[0]);
  }, [contestType]);

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

  const buildAdvancedContextPrompt = () => `
THÔNG TIN BỐI CẢNH NÂNG CAO DO NGƯỜI DÙNG CUNG CẤP:
- Vấn đề muốn giải quyết rõ nhất: ${problem || 'Không có'}
- Ý tưởng đã có/không muốn trùng: ${avoidIdeas || 'Không có'}
- Điểm riêng của địa phương/trường/lớp: ${localTraits || 'Không có'}

CÁCH KHAI THÁC BỐI CẢNH NÂNG CAO:
- Ưu tiên tạo ý tưởng bám sát vấn đề cụ thể người dùng nhập, không chỉ bám vào lĩnh vực chung.
- Nếu người dùng liệt kê ý tưởng đã có hoặc không muốn trùng, tuyệt đối né các hướng đó và các biến thể quá gần.
- Biến điểm riêng của địa phương/trường/lớp thành lợi thế sáng tạo để ý tưởng có dấu ấn riêng, khó trùng trên mạng.
- Nếu thông tin địa phương có văn hóa, địa hình, khí hậu, thói quen học tập hoặc nguồn lực đặc biệt, hãy dùng chúng để tạo cơ chế giải pháp mới.
  `;

  const buildContestFieldInsightPrompt = () => isKhktContest
    ? `TRỌNG TÂM LĨNH VỰC KHKT:
${buildKhktSelectedFieldPrompt(field)}
${KHKT_NATIONAL_RESEARCH_PRINCIPLES}`
    : `TRỌNG TÂM LĨNH VỰC SÁNG TẠO TTNND:
${buildCreativeSelectedFieldPrompt(field)}
${CREATIVE_NATIONAL_INNOVATION_PRINCIPLES}`;

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
    const loadingSteps = isKhktContest
      ? [
          'Đang đối chiếu 22 lĩnh vực KHKT và năng lực học sinh...',
          'Đang bám phiếu chấm 100 điểm để lọc đề tài có khả năng đạt giải...',
          'Đang xác định loại dự án khoa học/kỹ thuật phù hợp...',
          'Đang thiết kế phương pháp, dữ liệu, nguyên mẫu và cách kiểm chứng...',
          'Đang chọn lọc 20 đề tài KHKT mạnh nhất...',
          'Đang hoàn thiện ma trận rubric và Top 3 đề tài ưu tiên...'
        ]
      : [
          'Đang đối chiếu 5 lĩnh vực Sáng tạo TTNND và bối cảnh thực tế...',
          'Đang đối chiếu với kho ý tưởng đã có để tránh trùng lặp...',
          'Đang tìm điểm đột phá ở nguyên lý, kết cấu, vật liệu, tính năng hoặc cách triển khai...',
          'Đang đánh giá mô hình, tính khả thi, khả năng áp dụng và tính bền vững...',
          'Đang chọn lọc 20 ý tưởng có tiềm năng đạt giải toàn quốc...',
          'Đang hoàn thiện Top 3 ý tưởng champion và hướng phát triển...'
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
      const khktPrompt = `
Bạn là chuyên gia AI, giáo viên hướng dẫn nghiên cứu khoa học và giám khảo của "Cuộc thi nghiên cứu khoa học, kỹ thuật cấp quốc gia dành cho học sinh trung học" tại Việt Nam.
Nhiệm vụ của bạn là tạo ra ĐÚNG 20 ĐỀ TÀI KHKT xuất sắc cho học sinh, bám sát danh mục 22 lĩnh vực chính thức và phiếu chấm 100 điểm.

DANH MỤC 22 LĨNH VỰC KHKT VÀ PHẠM VI NGHIÊN CỨU CẦN HIỂU:
${buildKhktFieldCataloguePrompt()}

CHUYÊN SÂU LĨNH VỰC ĐANG CHỌN:
${buildKhktSelectedFieldPrompt(field)}

${KHKT_NATIONAL_RESEARCH_PRINCIPLES}

PHIẾU CHẤM KHKT 100 ĐIỂM CẦN BÁM SÁT:
- Câu hỏi nghiên cứu hoặc vấn đề nghiên cứu: 10 điểm.
- Thiết kế và phương pháp: 15 điểm.
- Thực hiện: thu thập, phân tích dữ liệu hoặc chế tạo, kiểm tra nguyên mẫu: 20 điểm.
- Tính sáng tạo: 20 điểm.
- Báo cáo/poster/tài liệu: 10 điểm.
- Nội dung khoa học: 25 điểm.

PHÂN BIỆT LOẠI DỰ ÁN:
- Dự án khoa học: cần câu hỏi nghiên cứu rõ, giả thuyết/biến số, phương pháp thu thập dữ liệu, phân tích dữ liệu và kết luận có thể lặp lại.
- Dự án kỹ thuật: cần vấn đề thực tế, tiêu chí giải pháp, thiết kế nguyên mẫu/mô hình, chế tạo, kiểm thử nhiều điều kiện và chứng minh mức độ hoàn chỉnh công nghệ.

THÔNG TIN ĐẦU VÀO:
- Cuộc thi: ${contestMeta.title}
- Lĩnh vực KHKT: ${field}
- Loại dự án ưu tiên: ${getKhktProjectTypeLabel(khktProjectType)}
- Cấp học: ${capHoc} (Lớp: ${grade})
- Giới hạn công nghệ: ${techLimit}
- Mục tiêu: ${mucTieu}
- Bối cảnh: ${context || 'Không có'}
- Nguồn lực: ${resources || 'Không có'}

${buildAdvancedContextPrompt()}

MÃ PHIÊN SÁNG TẠO: ${noveltySeed}
Hãy dùng mã phiên này để mở một nhánh đề tài mới. Nếu người dùng bấm tạo nhiều lần với cùng thông tin đầu vào, kết quả lần sau vẫn phải khác rõ rệt.

KHO Ý TƯỞNG/ĐỀ TÀI ĐÃ CÓ TRÊN MÁY NGƯỜI DÙNG (DANH SÁCH CẦN TRÁNH):
${ideaExclusionList}

QUY TẮC BẮT BUỘC:
1. Chỉ đề xuất đề tài phù hợp học sinh lớp ${grade}, an toàn, đúng đạo đức nghiên cứu, không dùng mầm bệnh, hóa chất độc hại hoặc nội dung có nguy cơ gây hại; thời gian thực hiện phải khả thi trong tối đa 12 tháng.
2. Đề tài phải đủ chất "nghiên cứu" hoặc "kỹ thuật": có dữ liệu/đo đạc/thử nghiệm/nguyên mẫu, không chỉ là ý tưởng tuyên truyền hoặc sản phẩm trang trí.
3. Với mỗi đề tài, phải ghi rõ nên đăng ký DỰ ÁN KHOA HỌC hay DỰ ÁN KỸ THUẬT. Nếu người dùng chọn một loại cụ thể thì ưu tiên đúng loại đó.
4. Mỗi đề tài phải có cách kiểm chứng rõ ràng để ăn điểm mục Thực hiện và Nội dung khoa học.
5. Không trùng tên, vấn đề chính, cơ chế hoạt động hoặc phương pháp nghiên cứu với danh sách đã có. Trong 20 đề tài cũng không được trùng nhau.
6. Mỗi đề tài phải có "điểm đột phá" cụ thể: cơ chế khoa học mới, kiến trúc kỹ thuật mới, dữ liệu mới, cách đo mới, hoặc ứng dụng mới trong bối cảnh Việt Nam; không dùng các cụm chung chung như "ứng dụng AI" nếu không nêu dữ liệu, mô hình, metric và baseline.
7. Ưu tiên ý tưởng có thể đạt giải cao cấp quốc gia: giải quyết vấn đề cấp thiết, có tính mới so với giải pháp phổ biến, có số liệu đối sánh, poster/phỏng vấn dễ bảo vệ, và có hướng nghiên cứu tiếp theo.
8. Không viết "cách làm chi tiết" quá dài trong kết quả chính; chỉ nêu lộ trình ngắn gọn vì ứng dụng có nút hướng dẫn riêng.

${isReroll ? 'YÊU CẦU ĐẶC BIỆT: Đây là lần [TÌM LẠI]. Hãy tạo 20 đề tài KHKT MỚI HOÀN TOÀN, không trùng lặp với lần trước.' : ''}

HÃY XUẤT KẾT QUẢ THEO ĐÚNG ĐỊNH DẠNG MARKDOWN SAU:

## 🧪 1. CHIẾN LƯỢC CHỌN ĐỀ TÀI KHKT
(Tóm tắt hướng chọn đề tài theo lĩnh vực ${field}, loại dự án phù hợp, tiêu chí giám khảo sẽ soi kỹ và cách tối đa hóa điểm theo phiếu chấm.)

## 🎯 2. MA TRẬN RUBRIC 100 ĐIỂM
(Trình bày ngắn gọn các điểm cần chuẩn bị để đạt điểm cao: vấn đề/câu hỏi 10, phương pháp 15, thực hiện 20, sáng tạo 20, báo cáo 10, nội dung khoa học 25.)

## 🚀 3. DANH SÁCH 20 ĐỀ TÀI KHKT TIỀM NĂNG

### 💡 Ý TƯỞNG 1: [Tên đề tài KHKT ngắn, rõ vấn đề, có tính nghiên cứu]
- **Loại dự án & lĩnh vực:** (Dự án khoa học hoặc Dự án kỹ thuật; lĩnh vực đăng ký trong 22 lĩnh vực)
- **Câu hỏi/Vấn đề nghiên cứu:** (Nêu đúng vấn đề cần giải quyết hoặc câu hỏi có thể kiểm chứng)
- **Mục tiêu & tiêu chí thành công:** (Nêu 2-3 tiêu chí đo được, có baseline/đối chứng để so sánh)
- **Giải pháp/Giả thuyết cốt lõi:** (Ý tưởng khoa học hoặc kỹ thuật then chốt, nêu rõ điểm đột phá)
- **Tính mới & đối sánh:** (So với giải pháp/phương pháp phổ biến, mới ở cơ chế, dữ liệu, thiết kế, thuật toán hay cách triển khai nào)
- **Thiết kế và phương pháp:** (Cách thu thập dữ liệu, biến số, mẫu thử, hoặc cách thiết kế nguyên mẫu)
- **Thực hiện và kiểm chứng:** (Cách đo đạc, thử nghiệm nhiều điều kiện, phân tích dữ liệu hoặc kiểm tra nguyên mẫu)
- **Sản phẩm/dữ liệu cần có:** (Mô hình, bộ dữ liệu, biểu đồ, bảng đo, poster, nhật ký nghiên cứu)
- **Điểm rubric KHKT:** [Tổng điểm]/100 — Câu hỏi/Vấn đề [x]/10; Thiết kế & phương pháp [x]/15; Thực hiện/kiểm chứng [x]/20; Sáng tạo [x]/20; Báo cáo [x]/10; Nội dung khoa học [x]/25. (1 câu giải thích điểm)
- **Rủi ro, hạn chế & cách khắc phục:** (Dự đoán lỗi, giới hạn kết quả, cách xử lý)
- **Lộ trình ngắn gọn:** (3-5 bước để học sinh bắt đầu)

### 💡 Ý TƯỞNG 2: [Tên đề tài KHKT ngắn, rõ vấn đề, có tính nghiên cứu]
... (Tương tự như trên)

... (Tiếp tục trình bày đầy đủ đến Ý TƯỞNG 20)

### 💡 Ý TƯỞNG 20: [Tên đề tài KHKT ngắn, rõ vấn đề, có tính nghiên cứu]
... (Tương tự như trên)

## 🏆 4. TOP 3 ĐỀ TÀI KHKT NÊN ƯU TIÊN
Chọn ra 3 đề tài mạnh nhất trong 20 đề tài trên và trình bày BẮT BUỘC bằng bảng Markdown GFM có đúng các cột sau:
| Hạng | Đề tài | Loại dự án | Điểm rubric | Vì sao mạnh | Việc cần làm ngay |
|---|---|---|---:|---|---|
| 🥇 Champion 1 | Ý tưởng số + tên đề tài | Khoa học/Kỹ thuật | xx/100 | Lý do bám rubric, sáng tạo, khả thi | Việc nên làm tiếp theo |
| 🥈 Champion 2 | Ý tưởng số + tên đề tài | Khoa học/Kỹ thuật | xx/100 | Lý do bám rubric, sáng tạo, khả thi | Việc nên làm tiếp theo |
| 🥉 Champion 3 | Ý tưởng số + tên đề tài | Khoa học/Kỹ thuật | xx/100 | Lý do bám rubric, sáng tạo, khả thi | Việc nên làm tiếp theo |

Sau bảng, viết thêm 2-3 câu "Ghi chú lựa chọn" để giải thích vì sao 3 đề tài này có khả năng phát triển thành hồ sơ dự thi tốt.
      `;

      const creativePrompt = `
Bạn là một Chuyên gia AI hàng đầu về Đổi mới Sáng tạo và là Giám khảo cấp quốc gia của "Cuộc thi Sáng tạo Thanh thiếu niên Nhi đồng" tại Việt Nam.
Nhiệm vụ của bạn là tạo ra ĐÚNG 20 Ý TƯỞNG/MÔ HÌNH/SẢN PHẨM xuất sắc nhất. Các ý tưởng phải có tính mới, tính sáng tạo cao, có khả năng chế tạo hoặc demo thật, an toàn, bền vững và đủ sức cạnh tranh ở vòng toàn quốc.

DANH MỤC 5 LĨNH VỰC DỰ THI CHÍNH THỨC CẦN HIỂU:
${buildCreativeFieldCataloguePrompt()}

CHUYÊN SÂU LĨNH VỰC ĐANG CHỌN:
${buildCreativeSelectedFieldPrompt(field)}

${CREATIVE_NATIONAL_INNOVATION_PRINCIPLES}

YÊU CẦU TỐI QUAN TRỌNG (THINKING PROCESS):
1. Phân tích kỹ nhưng KHÔNG trình bày suy luận nội bộ dài dòng; chỉ viết phần tóm tắt định hướng cần thiết.
2. TÌM KIẾM SỰ ĐỘT PHÁ: Tuyệt đối không đề xuất ý tưởng cũ rích, sáo mòn. Phải tự đối sánh với những mô hình/sản phẩm phổ biến, rồi nêu rõ khác biệt mới ở nguyên lý, kết cấu, vật liệu, tính năng, người dùng hoặc phương thức triển khai.
3. TÍNH MỚI & SÁNG TẠO CAO: Ý tưởng phải độc đáo nhưng không viển vông; nếu dùng AI/IoT/robot thì phải nêu dữ liệu, cảm biến, thuật toán, luồng vận hành và chỉ số đánh giá.
4. MÔ HÌNH THẬT & ỨNG DỤNG THẬT: Giải pháp phải có sản phẩm/mô hình/video vận hành, người dùng mục tiêu rõ, cách dùng rõ, và có thể thuyết minh trước giám khảo.
5. PHÙ HỢP LỨA TUỔI: Đảm bảo tính khả thi cho học sinh ${grade} với giới hạn công nghệ là ${techLimit}. Học sinh phải vận dụng được kiến thức đã học hoặc kỹ năng phổ thông để làm phần cốt lõi.

THÔNG TIN ĐẦU VÀO:
- Lĩnh vực: ${field}
- Cấp học: ${capHoc} (Lớp: ${grade})
- Giới hạn công nghệ: ${techLimit}
- Mục tiêu: ${mucTieu}
- Bối cảnh: ${context || 'Không có'}
- Nguồn lực: ${resources || 'Không có'}

${buildAdvancedContextPrompt()}

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
1. Tính mới: Nêu rõ điểm mới so với sản phẩm/mô hình đã có, không chỉ đổi tên hoặc đổi vật liệu.
2. Tính sáng tạo: Chỉ rõ sáng tạo nằm ở khâu nào: nguyên lý, kết cấu, vật liệu, tính năng, công dụng, thiết kế, phương thức triển khai hoặc nhóm người hưởng lợi.
3. Tính khả thi: Học sinh ${grade} có thể làm mô hình hoặc demo trong điều kiện ${techLimit}, vật liệu/linh kiện dễ kiếm, an toàn, có video vận hành hoặc thử nghiệm.
4. Khả năng áp dụng: Có người dùng thật, tình huống dùng rõ, hiệu quả đo được bằng thời gian, chi phí, độ chính xác, độ bền, mức tiết kiệm, mức hài lòng hoặc tác động học tập/đời sống.
5. Tính bền vững và nhân rộng: Chi phí hợp lý, dễ bảo trì, thân thiện môi trường, có thể mở rộng sang trường/lớp/gia đình/địa phương khác.

RÀNG BUỘC BẮT BUỘC VỀ KẾT QUẢ:
- Phải hoàn thành đủ từ ### 💡 Ý TƯỞNG 1 đến ### 💡 Ý TƯỞNG 20. Tuyệt đối không dừng ở Ý TƯỞNG 8, 10 hoặc 12.
- Nếu cần rút gọn để đủ 20 ý tưởng, hãy rút gọn từng gạch đầu dòng nhưng vẫn giữ đủ 11 mục phân tích cho mỗi ý tưởng.
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
- **Lĩnh vực & dạng sản phẩm:** (Chọn đúng 1 trong 5 lĩnh vực; nêu đây là mô hình, dụng cụ, thiết bị, robot, phần mềm, hệ thống số hay video demo)
- **Vấn đề & người hưởng lợi:** (Nỗi đau nào đang được giải quyết? Ai hưởng lợi: học sinh, giáo viên, gia đình, trẻ em, người yếu thế, cộng đồng hay địa phương?)
- **So sánh với giải pháp cũ:** (Những cái đã làm là gì? Tại sao giải pháp này sáng tạo và hữu ích hơn nhiều?)
- **Tính mới & sáng tạo ở khâu nào:** (Mới ở nguyên lý, kết cấu, vật liệu, tính năng, công dụng, thiết kế, phương thức triển khai hay nhóm người dùng?)
- **Tính năng nổi bật duy nhất:** (Nêu đúng 1 tính năng/cải tiến then chốt làm ý tưởng này khác biệt mạnh nhất)
- **Cơ chế hoạt động & mô hình:** (Mô tả rõ sản phẩm hoạt động thế nào, gồm những khối/chức năng nào, có thể vận hành hoặc demo ra sao)
- **Kiến thức vận dụng & vật liệu:** (Học sinh dùng kiến thức môn học/kỹ năng nào? Vật liệu, linh kiện, phần mềm chính là gì?)
- **Cách kiểm chứng hiệu quả:** (Đo bằng chỉ số nào, thử với ai, so sánh trước - sau hoặc so với sản phẩm cũ thế nào?)
- **Tính khả thi & Bền vững:** (Phân tích vật liệu, độ khó kỹ thuật, tính an toàn cho học sinh ${grade}, tác động môi trường, chi phí)
- **Điểm đánh giá:** [Tổng điểm]/100 — Tính mới [x]/20; Sáng tạo [x]/20; Khả thi [x]/20; Tác động [x]/20; Bền vững [x]/20. (1 câu giải thích điểm)
- **🛠 Cách làm ngắn gọn:** (3-5 bước chính, dễ hiểu, để học sinh nắm ngay lộ trình thực hiện)

### 💡 Ý TƯỞNG 2: [Tên ý tưởng thật ấn tượng, rõ nghĩa]
... (Tương tự như trên)

... (Tiếp tục trình bày đầy đủ đến Ý TƯỞNG 20)

### 💡 Ý TƯỞNG 20: [Tên ý tưởng thật ấn tượng, rõ nghĩa]
... (Tương tự như trên)

## 🏆 4. TOP 3 Ý TƯỞNG "CHAMPION" (KHUYÊN CHỌN NHẤT)
Chọn ra 3 ý tưởng xuất sắc toàn diện nhất trong 20 ý tưởng trên và trình bày BẮT BUỘC bằng bảng Markdown GFM có đúng các cột sau:
| Hạng | Ý tưởng | Điểm | Lý do chọn | Hướng phát triển |
|---|---|---:|---|---|
| 🥇 Champion 1 | Ý tưởng số + tên ý tưởng | xx/100 | Lý do ngắn gọn, thuyết phục | Việc nên làm tiếp theo |
| 🥈 Champion 2 | Ý tưởng số + tên ý tưởng | xx/100 | Lý do ngắn gọn, thuyết phục | Việc nên làm tiếp theo |
| 🥉 Champion 3 | Ý tưởng số + tên ý tưởng | xx/100 | Lý do ngắn gọn, thuyết phục | Việc nên làm tiếp theo |

Sau bảng, viết thêm 2-3 câu "Ghi chú lựa chọn" để giải thích vì sao 3 ý tưởng này đáng ưu tiên.
      `;
      const prompt = isKhktContest ? khktPrompt : creativePrompt;

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
        const khktContinuationPrompt = `
Bạn đang viết tiếp một báo cáo đề tài KHKT. Kết quả trước mới hoàn thành đến Ý TƯỞNG ${lastIdeaNumber}.

NHIỆM VỤ BẮT BUỘC:
1. KHÔNG viết lại các đề tài đã có.
2. Bắt đầu chính xác bằng heading: ### 💡 Ý TƯỞNG ${nextIdeaNumber}: [Tên đề tài KHKT]
3. Viết tiếp đầy đủ đến ### 💡 Ý TƯỞNG 20.
4. Sau Ý TƯỞNG 20, viết mục ## 🏆 4. TOP 3 ĐỀ TÀI KHKT NÊN ƯU TIÊN bằng bảng Markdown GFM có đúng 6 cột: Hạng | Đề tài | Loại dự án | Điểm rubric | Vì sao mạnh | Việc cần làm ngay.
5. Giữ đúng cấu trúc 11 gạch đầu dòng cho mỗi đề tài: Loại dự án & lĩnh vực; Câu hỏi/Vấn đề nghiên cứu; Mục tiêu & tiêu chí thành công; Giải pháp/Giả thuyết cốt lõi; Tính mới & đối sánh; Thiết kế và phương pháp; Thực hiện và kiểm chứng; Sản phẩm/dữ liệu cần có; Điểm rubric KHKT; Rủi ro, hạn chế & cách khắc phục; Lộ trình ngắn gọn.
6. Mọi đề tài viết tiếp phải khác hoàn toàn các đề tài trong bộ nhớ cũ và khác các đề tài đã có ở phần trước.
7. Không viết phần hướng dẫn chi tiết quá dài; chỉ nêu lộ trình ngắn gọn.

KHO Ý TƯỞNG/ĐỀ TÀI ĐÃ CÓ TRÊN MÁY NGƯỜI DÙNG (DANH SÁCH CẦN TRÁNH):
${ideaExclusionList}

THÔNG TIN ĐẦU VÀO:
- Cuộc thi: ${contestMeta.title}
- Lĩnh vực KHKT: ${field}
- Loại dự án ưu tiên: ${getKhktProjectTypeLabel(khktProjectType)}
- Cấp học: ${capHoc} (Lớp: ${grade})
- Giới hạn công nghệ: ${techLimit}
- Mục tiêu: ${mucTieu}
- Bối cảnh: ${context || 'Không có'}
- Nguồn lực: ${resources || 'Không có'}

CHUYÊN SÂU LĨNH VỰC ĐANG CHỌN:
${buildKhktSelectedFieldPrompt(field)}

${KHKT_NATIONAL_RESEARCH_PRINCIPLES}

${buildAdvancedContextPrompt()}

PHẦN CUỐI KẾT QUẢ TRƯỚC ĐỂ TRÁNH TRÙNG LẶP:
${fullText.slice(-6000)}
        `;

        const creativeContinuationPrompt = `
Bạn đang viết tiếp một báo cáo ý tưởng sáng tạo. Kết quả trước mới hoàn thành đến Ý TƯỞNG ${lastIdeaNumber}.

NHIỆM VỤ BẮT BUỘC:
1. KHÔNG viết lại các ý tưởng đã có.
2. Bắt đầu chính xác bằng heading: ### 💡 Ý TƯỞNG ${nextIdeaNumber}: [Tên ý tưởng]
3. Viết tiếp đầy đủ đến ### 💡 Ý TƯỞNG 20.
4. Sau Ý TƯỞNG 20, viết mục ## 🏆 4. TOP 3 Ý TƯỞNG "CHAMPION" (KHUYÊN CHỌN NHẤT) bằng bảng Markdown GFM có đúng 5 cột: Hạng | Ý tưởng | Điểm | Lý do chọn | Hướng phát triển.
5. Giữ đúng cấu trúc 11 gạch đầu dòng cho mỗi ý tưởng: Lĩnh vực & dạng sản phẩm; Vấn đề & người hưởng lợi; So sánh với giải pháp cũ; Tính mới & sáng tạo ở khâu nào; Tính năng nổi bật duy nhất; Cơ chế hoạt động & mô hình; Kiến thức vận dụng & vật liệu; Cách kiểm chứng hiệu quả; Tính khả thi & Bền vững; Điểm đánh giá; Cách làm ngắn gọn.
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

CHUYÊN SÂU LĨNH VỰC ĐANG CHỌN:
${buildCreativeSelectedFieldPrompt(field)}

${CREATIVE_NATIONAL_INNOVATION_PRINCIPLES}

${buildAdvancedContextPrompt()}

PHẦN CUỐI KẾT QUẢ TRƯỚC ĐỂ TRÁNH TRÙNG LẶP:
${fullText.slice(-6000)}
        `;
        const continuationPrompt = isKhktContest ? khktContinuationPrompt : creativeContinuationPrompt;

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
        fullText += `\n\n> ⚠️ Kết quả hiện chưa đủ 20 ${isKhktContest ? 'đề tài KHKT' : 'ý tưởng'} do mô hình dừng sớm. Hãy bấm "Tìm Lại" hoặc chọn chế độ nâng cao khác để tạo lại danh sách đầy đủ.`;
        setResult(fullText);
      }

      persistIdeaMemoryFromResult(fullText);

      const newId = Date.now().toString();
      setCurrentSessionId(newId);
      setHistory(prev => [{
        id: newId,
        timestamp: Date.now(),
        inputs: { contestType, khktProjectType, field, capHoc, grade, techLimit, mucTieu, context, resources, problem, avoidIdeas, localTraits },
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
      setLoadingMessage(isKhktContest ? 'Đang khởi tạo KHKT IdeaGPT...' : 'Đang khởi tạo IdeaGPT...');
      setIsGenerating(false);
    }
  };

  const handleInlineCompare = async (title: string, sectionContent: string) => {
    setLoadingInline(prev => ({ ...prev, [title]: 'comparing' }));
    setInlineComparisons(prev => ({ ...prev, [title]: '' }));

    const prompt = `Bạn là chuyên gia đánh giá ${isKhktContest ? 'dự án nghiên cứu khoa học, kỹ thuật học sinh trung học' : 'dự án khoa học kỹ thuật và khởi nghiệp'}. Hãy phân tích và so sánh ý tưởng sau đây với các sản phẩm/giải pháp ĐÃ CÓ TRÊN THỊ TRƯỜNG hoặc TRÊN MẠNG.
Hãy phân tích bằng DeepSeek V4 Pro. Nếu không có dữ liệu chắc chắn, hãy nói rõ mức độ tin cậy thay vì bịa nguồn.

THÔNG TIN CUỘC THI:
- Cuộc thi: ${contestMeta.title}
- Lĩnh vực: ${field}
${isKhktContest ? `- Loại dự án ưu tiên: ${getKhktProjectTypeLabel(khktProjectType)}
- Phiếu chấm cần bám: Vấn đề/Câu hỏi 10; Thiết kế & phương pháp 15; Thực hiện/kiểm chứng 20; Sáng tạo 20; Báo cáo 10; Nội dung khoa học 25.` : `- Mục tiêu: ${mucTieu}`}
${buildContestFieldInsightPrompt()}

Ý TƯỞNG CẦN ĐÁNH GIÁ:
${sectionContent}

YÊU CẦU PHÂN TÍCH (Đóng vai trò chuyên gia DeepSeek V4 Pro để phân tích sâu sắc):
1. ĐỐI CHIẾU THỰC TẾ: Chỉ ra đích danh 2-3 sản phẩm/dự án tương tự đã có trên thực tế nếu bạn biết chắc. Kèm link tham khảo khi chắc chắn, không bịa nguồn.
2. SO SÁNH ĐIỂM GIỐNG & KHÁC: Phân tích điểm giống và khác biệt cốt lõi giữa ý tưởng này và các sản phẩm đã có.
3. ĐÁNH GIÁ TÍNH MỚI: Đánh giá khách quan xem ý tưởng này có thực sự "chưa ai làm" không? Điểm nào là cải tiến ĐÁNG GIÁ NHẤT và SÁNG TẠO NHẤT so với cái cũ?
4. ${isKhktContest ? 'CHẤM THEO RUBRIC KHKT: Ước lượng điểm 100, chỉ rõ điểm yếu cần bổ sung ở dữ liệu, nguyên mẫu, phương pháp, poster hoặc nội dung khoa học.' : `TÍNH ỨNG DỤNG: Đánh giá tính ứng dụng thực tế đối với học sinh lớp ${grade}.`}

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

THÔNG TIN CUỘC THI:
- Cuộc thi: ${contestMeta.title}
- Lĩnh vực: ${field}
${isKhktContest ? `- Loại dự án ưu tiên: ${getKhktProjectTypeLabel(khktProjectType)}
- Mục tiêu nâng cấp: tăng điểm Thực hiện/kiểm chứng, Tính sáng tạo và Nội dung khoa học theo phiếu chấm KHKT.` : `- Mục tiêu: ${mucTieu}`}
${buildContestFieldInsightPrompt()}

Ý TƯỞNG BAN ĐẦU:
${sectionContent}

BẢN SO SÁNH THỰC TẾ:
${comparisonContent}

YÊU CẦU NÂNG CẤP (Đóng vai trò GPT-4o/GPT-5 để sáng tạo):
1. ĐỀ XUẤT ĐỘT PHÁ (Chưa ai làm): Đưa ra 3-5 tính năng/cải tiến MỚI TOANH, độc đáo, mang yếu tố "WOW" (bất ngờ, thú vị). Hãy suy nghĩ vượt ra ngoài các giải pháp thông thường.
2. TÍNH ỨNG DỤNG THỰC TẾ CAO: Các tính năng này phải giải quyết được vấn đề thực tế một cách hiệu quả, không viển vông, có thể áp dụng ngay vào đời sống.
3. TÍNH KHẢ THI: Đảm bảo các tính năng này ĐƠN GIẢN, phù hợp với trình độ học sinh lớp ${grade} (có thể làm được với công nghệ: ${techLimit}).
4. GIẢI THÍCH SỰ KHÁC BIỆT: Giải thích rõ tại sao các cải tiến này lại làm cho dự án trở nên "vô đối" và hữu ích hơn nhiều so với các sản phẩm cũ trên mạng.
5. ${isKhktContest ? 'BÁM RUBRIC KHKT: Với mỗi cải tiến, nêu rõ nó giúp tăng điểm mục nào trong phiếu chấm 100 điểm và cần dữ liệu/thử nghiệm nào để chứng minh.' : 'VẬN DỤNG KIẾN THỨC: Gợi ý cách học sinh vận dụng kiến thức môn học để làm các tính năng mới này.'}

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

    const prompt = `Bạn là giáo viên hướng dẫn học sinh làm ${isKhktContest ? 'dự án nghiên cứu khoa học, kỹ thuật dự thi cấp quốc gia' : 'dự án sáng tạo khoa học kỹ thuật'}. Hãy viết HƯỚNG DẪN CÁCH LÀM CHI TIẾT cho đúng ý tưởng dưới đây.

Ý TƯỞNG CẦN HƯỚNG DẪN:
${sectionContent}

THÔNG TIN HỌC SINH:
- Cuộc thi: ${contestMeta.title}
- Lĩnh vực: ${field}
${isKhktContest ? `- Loại dự án ưu tiên: ${getKhktProjectTypeLabel(khktProjectType)}` : ''}
- Cấp học: ${capHoc}
- Lớp: ${grade}
- Giới hạn công nghệ: ${techLimit}
- Nguồn lực đang có: ${resources || 'Không có'}
- Bối cảnh: ${context || 'Không có'}
- Vấn đề muốn giải quyết: ${problem || 'Không có'}
- Ý tưởng đã có/không muốn trùng: ${avoidIdeas || 'Không có'}
- Điểm riêng địa phương/trường/lớp: ${localTraits || 'Không có'}
${buildContestFieldInsightPrompt()}

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
3. ${isKhktContest ? 'Bổ sung riêng cho KHKT: câu hỏi/giả thuyết hoặc vấn đề kỹ thuật, biến số/tiêu chí giải pháp, bảng dữ liệu cần đo, cách phân tích kết quả, giới hạn kết luận và gợi ý bố cục poster/báo cáo theo phiếu chấm 100 điểm.' : `Bổ sung riêng cho Sáng tạo TTNND: mô hình/sản phẩm hoặc video vận hành, vật liệu an toàn, cách thuyết minh tính mới - tính sáng tạo - khả năng áp dụng, và cách kiểm chứng hiệu quả phù hợp học sinh ${grade}.`}
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

    const prompt = `Bạn là chuyên gia đánh giá ${isKhktContest ? 'dự án nghiên cứu khoa học, kỹ thuật học sinh trung học' : 'dự án khoa học kỹ thuật và đổi mới sáng tạo'}.
Hãy phân tích khách quan ý tưởng sau bằng cách so sánh với sản phẩm/dự án tương tự đã có trên thị trường hoặc trên mạng.
Hãy dùng DeepSeek V4 Pro để đối chiếu thực tế; chỉ nêu nguồn/link khi chắc chắn, không bịa nguồn.

THÔNG TIN CUỘC THI:
- Cuộc thi: ${contestMeta.title}
- Lĩnh vực: ${field}
${isKhktContest ? `- Loại dự án ưu tiên: ${getKhktProjectTypeLabel(khktProjectType)}
- Rubric KHKT: Vấn đề/Câu hỏi 10; Thiết kế & phương pháp 15; Thực hiện/kiểm chứng 20; Sáng tạo 20; Báo cáo 10; Nội dung khoa học 25.` : `- Mục tiêu: ${mucTieu}`}
${buildContestFieldInsightPrompt()}

Ý TƯỞNG CẦN SO SÁNH:
${ideaContent}

YÊU CẦU:
1. Nêu 2-3 sản phẩm/dự án tương tự đã có, kèm link tham khảo nếu tìm thấy.
2. So sánh điểm giống và khác biệt cốt lõi.
3. ${isKhktContest ? 'Ước lượng điểm theo phiếu chấm KHKT 100 điểm, nêu rõ mục nào còn yếu.' : 'Đánh giá tính mới, tính sáng tạo và khả năng đạt giải.'}
4. Gợi ý cách cải tiến để ý tưởng khác biệt hơn nhưng vẫn phù hợp với học sinh ${grade}${isKhktContest ? ' và có dữ liệu/thử nghiệm đủ bảo vệ trước giám khảo' : ''}.

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

  const generateAndDownloadWordDoc = async (content: string, title: string, summaryOnly = false) => {
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
      return cleaned || (isKhktContest ? 'ideagpt-de-tai-khkt' : 'ideagpt-y-tuong-sang-tao');
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

    const cleanSummaryText = (value: string) => value
      .replace(/\*\*/g, '')
      .replace(/`/g, '')
      .replace(/\s+/g, ' ')
      .trim();

    const truncateSummary = (value: string, maxLength = 260) => {
      const cleaned = cleanSummaryText(value);
      if (cleaned.length <= maxLength) return cleaned;
      return `${cleaned.slice(0, maxLength - 1).replace(/\s+\S*$/, '')}…`;
    };

    const extractBulletValue = (section: string, labels: string[]) => {
      const lines = section.split('\n').map(line => line.trim());

      for (const label of labels) {
        const target = label.toLowerCase();
        const line = lines.find(item => {
          const normalized = item
            .replace(/^[-*]\s*/, '')
            .replace(/^\*\*/, '')
            .toLowerCase();
          return normalized.startsWith(`${target}:`);
        });

        if (line) {
          return cleanSummaryText(line
            .replace(/^[-*]\s*/, '')
            .replace(/^\*\*/, '')
            .replace(new RegExp(`^${label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}:\\*?\\*?\\s*`, 'i'), '')
          );
        }
      }

      return '';
    };

    const extractIdeaSummaries = () => {
      const sections = content
        .split(/(?=^###\s*(?:💡\s*)?Ý TƯỞNG\s+\d+\s*:)/m)
        .filter(section => /^###\s*(?:💡\s*)?Ý TƯỞNG\s+\d+\s*:/m.test(section));

      return sections.map(section => {
        const cleanedSection = section.split(/^##\s+🏆/m)[0];
        const titleMatch = cleanedSection.match(/^###\s*(?:💡\s*)?Ý TƯỞNG\s+(\d+)\s*:\s*(.+)$/m);
        const number = titleMatch?.[1] || '';
        const ideaTitle = cleanSummaryText(titleMatch?.[2] || 'Không có tiêu đề');
        const scoreMatch = cleanedSection.match(/\*\*(?:Điểm đánh giá|Điểm rubric KHKT):\*\*[\s\S]{0,100}?(\d{1,3})\s*\/\s*100/i);
        const score = scoreMatch ? `${scoreMatch[1]}/100` : '';
        const summaryLabels = isKhktContest
          ? [
              ['Loại dự án & lĩnh vực'],
              ['Câu hỏi/Vấn đề nghiên cứu'],
              ['Mục tiêu & tiêu chí thành công'],
              ['Giải pháp/Giả thuyết cốt lõi'],
              ['Thực hiện và kiểm chứng'],
            ]
          : [
              ['Lĩnh vực & dạng sản phẩm'],
              ['Vấn đề & người hưởng lợi', 'Vấn đề & Ý nghĩa thực tiễn'],
              ['Tính mới & sáng tạo ở khâu nào'],
              ['Tính năng nổi bật duy nhất'],
              ['Cơ chế hoạt động & mô hình', 'Cơ chế hoạt động & Giải pháp'],
              ['Cách kiểm chứng hiệu quả'],
              ['Kiến thức vận dụng & vật liệu', 'Kiến thức vận dụng'],
              ['Tính khả thi & Bền vững'],
            ];
        const summary = summaryLabels
          .map(labels => extractBulletValue(cleanedSection, labels))
          .filter(Boolean)
          .slice(0, 3)
          .map(value => truncateSummary(value, 170))
          .join(' ');

        return {
          number,
          title: ideaTitle,
          summary: summary || truncateSummary(cleanedSection.replace(/^###.*$/m, ''), 260),
          score,
        };
      });
    };

    const inputRows = [
      ['Cuộc thi', contestMeta.title],
      ['Lĩnh vực', field],
      ...(isKhktContest ? [['Loại dự án ưu tiên', getKhktProjectTypeLabel(khktProjectType)]] : []),
      ['Cấp học', capHoc],
      ['Lớp', grade],
      ['Giới hạn công nghệ', techLimit],
      ['Mục tiêu', mucTieu],
      ['Bối cảnh', context || 'Không có'],
      ['Nguồn lực', resources || 'Không có'],
      ['Vấn đề muốn giải quyết', problem || 'Không có'],
      ['Ý tưởng đã có/không muốn trùng', avoidIdeas || 'Không có'],
      ['Điểm riêng địa phương/trường/lớp', localTraits || 'Không có'],
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

    if (summaryOnly) {
      const ideaSummaries = extractIdeaSummaries();
      const summaryCell = (
        text: string,
        options: { width: number; bold?: boolean; fill?: string; color?: string; size?: number; center?: boolean } = { width: 25 }
      ) => new TableCell({
        width: { size: options.width, type: WidthType.PERCENTAGE },
        margins: { top: 120, bottom: 120, left: 130, right: 130 },
        shading: options.fill ? { type: ShadingType.CLEAR, fill: options.fill, color: 'auto' } : undefined,
        children: [
          new Paragraph({
            alignment: options.center ? AlignmentType.CENTER : undefined,
            children: [new TextRun({
              text,
              bold: options.bold || false,
              color: options.color || colors.text,
              font: 'Arial',
              size: options.size || 20,
            })],
          }),
        ],
      });

      bodyBlocks.push(
        new Paragraph({
          children: [new TextRun({
            text: isKhktContest ? 'DANH SÁCH ĐỀ TÀI KHKT TÓM TẮT' : 'DANH SÁCH Ý TƯỞNG TÓM TẮT',
            bold: true,
            color: colors.primary,
            font: 'Arial',
            size: 32,
          })],
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 120, after: 240 },
          alignment: AlignmentType.CENTER,
        })
      );

      if (ideaSummaries.length === 0) {
        bodyBlocks.push(new Paragraph({
          children: [new TextRun({ text: 'Không tìm thấy danh sách ý tưởng trong nội dung hiện tại.', ...normalRun })],
        }));
      } else {
        bodyBlocks.push(new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: border,
            bottom: border,
            left: border,
            right: border,
            insideHorizontal: border,
            insideVertical: border,
          },
          rows: [
            new TableRow({
              children: [
                summaryCell('STT', { width: 8, bold: true, fill: colors.primaryDark, color: colors.white, center: true }),
                summaryCell(isKhktContest ? 'Đề tài' : 'Ý tưởng', { width: 30, bold: true, fill: colors.primaryDark, color: colors.white }),
                summaryCell('Tóm tắt ngắn gọn', { width: 50, bold: true, fill: colors.primaryDark, color: colors.white }),
                summaryCell('Điểm', { width: 12, bold: true, fill: colors.primaryDark, color: colors.white, center: true }),
              ],
            }),
            ...ideaSummaries.map((idea, index) => new TableRow({
              children: [
                summaryCell(idea.number || String(index + 1), { width: 8, center: true, bold: true, color: colors.primaryDark }),
                summaryCell(idea.title, { width: 30, bold: true, color: colors.primaryDark }),
                summaryCell(idea.summary, { width: 50 }),
                summaryCell(idea.score || '-', { width: 12, center: true, bold: true, color: colors.primaryDark }),
              ],
            })),
          ],
        }));
      }
    } else {
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
                  children: [new TextRun({ text: `IdeaGPT - ${contestMeta.docTitle}`, color: colors.primary, font: 'Arial', size: 18, bold: true })],
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
              children: [new TextRun({
                text: summaryOnly
                  ? (isKhktContest ? 'DANH SÁCH ĐỀ TÀI KHKT TÓM TẮT' : 'DANH SÁCH Ý TƯỞNG TÓM TẮT')
                  : contestMeta.docTitle.toUpperCase(),
                bold: true,
                color: colors.primaryDark,
                font: 'Arial',
                size: 36,
              })],
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
                  text: summaryOnly
                    ? 'Tài liệu tóm tắt tự động: chỉ gồm danh sách ý tưởng/đề tài, mô tả ngắn và điểm đánh giá.'
                    : 'Tài liệu được định dạng tự động: tiêu đề, phân mục, danh sách, bảng thông tin và số trang.',
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
      saveAs(blob, `${makeFileName(summaryOnly ? `${title} tóm tắt` : title)}.docx`);
    } catch (error) {
      console.error('Error generating or downloading Word document:', error);
      alert('Đã có lỗi xảy ra khi tạo file Word. Vui lòng thử lại.');
    }
  };

  const loadSession = (session: SavedSession) => {
    const nextContestType = session.inputs.contestType || inferContestTypeFromField(session.inputs.field);
    setContestType(nextContestType);
    setKhktProjectType(session.inputs.khktProjectType || 'auto');
    setField(session.inputs.field);
    setCapHoc(session.inputs.capHoc);
    setGrade(session.inputs.grade);
    setTechLimit(session.inputs.techLimit);
    setMucTieu(session.inputs.mucTieu);
    setContext(session.inputs.context);
    setResources(session.inputs.resources);
    setProblem(session.inputs.problem || '');
    setAvoidIdeas(session.inputs.avoidIdeas || '');
    setLocalTraits(session.inputs.localTraits || '');
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
      const scoreMatch = part.match(/\*\*(?:Điểm đánh giá|Điểm rubric KHKT):\*\*[\s\S]{0,80}?(\d{1,3})\s*\/\s*100/i);
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
                  {ideaScore !== null ? `${isKhktContest ? 'Rubric' : 'Điểm'}: ${ideaScore}/100` : 'Chưa có điểm'}
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
                  So sánh {isKhktContest ? 'đề tài' : 'ý tưởng'} này
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
                      inputs: { contestType, khktProjectType, field, capHoc, grade }
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
              title={isSaved ? "Bỏ lưu" : isKhktContest ? "Lưu đề tài này" : "Lưu ý tưởng này"}
            >
              <Heart className={cn("w-4 h-4", isSaved && "fill-current")} />
              {isSaved ? "Đã lưu" : isKhktContest ? "Lưu đề tài" : "Lưu ý tưởng"}
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
    table: ({node, ...props}: any) => (
      <div className="my-8 overflow-x-auto rounded-2xl border border-emerald-700/60 bg-teal-950/50 shadow-lg shadow-teal-950/30">
        <table className="w-full min-w-[760px] border-collapse text-sm text-teal-50" {...props} />
      </div>
    ),
    thead: ({node, ...props}: any) => <thead className="bg-emerald-500/15 text-emerald-200" {...props} />,
    tbody: ({node, ...props}: any) => <tbody className="divide-y divide-teal-700/60" {...props} />,
    tr: ({node, ...props}: any) => <tr className="border-b border-teal-700/60 last:border-b-0 hover:bg-emerald-500/5 transition-colors" {...props} />,
    th: ({node, ...props}: any) => (
      <th
        className="border-r border-teal-700/60 last:border-r-0 px-4 py-3 text-left text-xs font-extrabold uppercase tracking-wide text-emerald-200"
        {...props}
      />
    ),
    td: ({node, ...props}: any) => (
      <td
        className="border-r border-teal-800/60 last:border-r-0 px-4 py-4 align-top leading-relaxed text-teal-50/90 first:font-bold first:text-emerald-300 [&:nth-child(3)]:whitespace-nowrap [&:nth-child(3)]:text-center [&:nth-child(3)]:font-extrabold [&:nth-child(3)]:text-emerald-300"
        {...props}
      />
    ),
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
            <span>IdeaGPT</span>
            <span className="text-[10px] px-2 py-0.5 rounded-md bg-teal-900/70 border border-teal-700/60 text-teal-200 font-extrabold tracking-wide">
              {contestMeta.shortLabel}
            </span>
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
                <Target className="w-4 h-4 text-teal-400" />
                Cuộc thi
              </label>
              <div className="grid grid-cols-2 gap-2">
                {CONTEST_OPTIONS.map(option => (
                  <button
                    key={option.id}
                    onClick={() => {
                      if (contestType === option.id) return;
                      setContestType(option.id);
                      setResult('');
                      setCompareResult('');
                      setInlineComparisons({});
                      setInlineEnhancements({});
                      setInlineDetailedGuides({});
                      setCurrentSessionId(null);
                      setActiveTab('main');
                    }}
                    className={cn(
                      "min-h-[56px] rounded-xl px-3 py-2 text-sm font-bold transition-all border flex flex-col items-start justify-center gap-0.5",
                      contestType === option.id
                        ? "bg-emerald-500 text-teal-950 border-emerald-400 shadow-sm shadow-emerald-500/20"
                        : "bg-teal-900/50 text-teal-200 border-teal-700/50 hover:bg-teal-800/50"
                    )}
                  >
                    <span>{option.shortLabel}</span>
                    <span className="text-[10px] font-semibold opacity-80 leading-tight text-left">{option.id === 'khkt' ? '22 lĩnh vực' : 'Sáng tạo'}</span>
                  </button>
                ))}
              </div>
            </div>

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
                {isKhktContest ? 'Lĩnh vực KHKT' : 'Lĩnh vực'}
              </label>
              <select
                value={field}
                onChange={(e) => setField(e.target.value)}
                className="w-full p-2.5 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all text-teal-50 font-medium"
              >
                {activeFieldOptions.map((opt) => (
                  <option key={opt} value={opt} className="bg-teal-900 text-teal-50">{opt}</option>
                ))}
              </select>
              {(() => {
                const selectedFieldGuide = isKhktContest ? getKhktFieldGuide(field) : getCreativeFieldGuide(field);
                if (!selectedFieldGuide) return null;
                const focusText = 'researchFocus' in selectedFieldGuide
                  ? selectedFieldGuide.researchFocus
                  : selectedFieldGuide.productFocus;

                return (
                  <div className="rounded-xl border border-emerald-400/20 bg-emerald-950/25 px-3 py-2.5 text-xs text-teal-100">
                    <div className="font-extrabold uppercase tracking-wide text-emerald-300">
                      {isKhktContest ? 'Trọng tâm nghiên cứu' : 'Trọng tâm sáng tạo'}
                    </div>
                    <p className="mt-1 leading-relaxed">{focusText}</p>
                    <p className="mt-1.5 leading-relaxed text-teal-200">
                      <span className="font-bold text-emerald-200">Góc đột phá:</span> {selectedFieldGuide.breakthroughAngles[0]}
                    </p>
                  </div>
                );
              })()}
            </div>

            {isKhktContest && (
              <div className="space-y-4 rounded-2xl border border-teal-700/50 bg-teal-900/30 p-4">
                <div className="space-y-2">
                  <label className="flex items-center gap-2 text-sm font-semibold text-teal-200">
                    <GitCompare className="w-4 h-4 text-teal-400" />
                    Loại dự án
                  </label>
                  <select
                    value={khktProjectType}
                    onChange={(e) => setKhktProjectType(e.target.value as KhktProjectType)}
                    className="w-full p-2.5 bg-teal-950/60 border border-teal-700/50 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all text-teal-50 font-medium"
                  >
                    {KHKT_PROJECT_TYPE_OPTIONS.map((opt) => (
                      <option key={opt.value} value={opt.value} className="bg-teal-900 text-teal-50">{opt.label}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <div className="text-xs font-extrabold uppercase tracking-wide text-emerald-300 mb-3">Phiếu chấm KHKT</div>
                  <div className="grid grid-cols-2 gap-2">
                    {KHKT_RUBRIC_ITEMS.map(([label, score]) => (
                      <div key={label} className="rounded-lg border border-teal-700/50 bg-teal-950/40 px-3 py-2">
                        <div className="text-[11px] leading-snug text-teal-200">{label}</div>
                        <div className="text-sm font-extrabold text-emerald-300">{score} điểm</div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}

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
                  {activeCapHocOptions.map((opt) => (
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
                  onChange={(e) => {
                    const nextGrade = e.target.value;
                    setGrade(nextGrade);
                    if (nextGrade.startsWith('THCS')) setCapHoc('THCS');
                    if (nextGrade.startsWith('THPT')) setCapHoc('THPT');
                    if (nextGrade.startsWith('Tiểu học')) setCapHoc('Tiểu học');
                  }}
                  className="w-full p-2.5 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all text-teal-50 font-medium"
                >
                  {activeGradeOptions.map((opt) => (
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
                placeholder={isKhktContest ? "Ví dụ: Trường gần biển, có phòng STEM, cần xử lý nước mặn..." : "Ví dụ: Trường ở vùng nông thôn, gần biển..."}
                className="w-full p-3 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all min-h-[80px] resize-y text-teal-50 placeholder:text-teal-400/50"
              />
            </div>

            <div className="space-y-2">
              <label className="text-sm font-semibold text-teal-200">Nguồn lực <span className="text-teal-400/70 font-normal">(Tùy chọn)</span></label>
              <textarea
                value={resources}
                onChange={(e) => setResources(e.target.value)}
                placeholder={isKhktContest ? "Ví dụ: Có Arduino, cảm biến pH, cân điện tử, Excel, máy in 3D..." : "Ví dụ: Có sẵn bìa carton, chai nhựa, biết lập trình Scratch..."}
                className="w-full p-3 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all min-h-[80px] resize-y text-teal-50 placeholder:text-teal-400/50"
              />
            </div>

            <div className="pt-2 border-t border-teal-800/50 space-y-5">
              <div className="flex items-center gap-2 text-sm font-bold text-emerald-300">
                <Lightbulb className="w-4 h-4" />
                Bối cảnh nâng cao
              </div>

              <div className="space-y-2">
                <label className="flex items-center gap-2 text-sm font-semibold text-teal-200">
                  <Target className="w-4 h-4 text-teal-400" />
                  Vấn đề muốn giải quyết <span className="text-teal-400/70 font-normal">(Tùy chọn)</span>
                </label>
                <textarea
                  value={problem}
                  onChange={(e) => setProblem(e.target.value)}
                  placeholder={isKhktContest ? "Ví dụ: nước uống ở trường có độ đục cao, cây trong vườn trường chết khi nắng nóng..." : "Ví dụ: học sinh dân tộc khó ghi nhớ từ vựng, ngại phát biểu, thiếu Internet..."}
                  className="w-full p-3 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all min-h-[82px] resize-y text-teal-50 placeholder:text-teal-400/50"
                />
              </div>

              <div className="space-y-2">
                <label className="flex items-center gap-2 text-sm font-semibold text-teal-200">
                  <GitCompare className="w-4 h-4 text-teal-400" />
                  Ý tưởng đã có/không muốn trùng <span className="text-teal-400/70 font-normal">(Tùy chọn)</span>
                </label>
                <textarea
                  value={avoidIdeas}
                  onChange={(e) => setAvoidIdeas(e.target.value)}
                  placeholder={isKhktContest ? "Ví dụ: không lấy máy lọc nước mini, robot tưới cây, thùng rác thông minh..." : "Ví dụ: không lấy app nhắc học, thùng rác thông minh, robot tưới cây..."}
                  className="w-full p-3 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all min-h-[82px] resize-y text-teal-50 placeholder:text-teal-400/50"
                />
              </div>

              <div className="space-y-2">
                <label className="flex items-center gap-2 text-sm font-semibold text-teal-200">
                  <Layers className="w-4 h-4 text-teal-400" />
                  Điểm riêng địa phương/trường/lớp <span className="text-teal-400/70 font-normal">(Tùy chọn)</span>
                </label>
                <textarea
                  value={localTraits}
                  onChange={(e) => setLocalTraits(e.target.value)}
                  placeholder="Ví dụ: trường bán trú miền núi, văn hóa dân tộc, mùa mưa kéo dài, mạng yếu..."
                  className="w-full p-3 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition-all min-h-[82px] resize-y text-teal-50 placeholder:text-teal-400/50"
                />
              </div>
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
                {contestMeta.actionLabel}
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
            {contestMeta.resultTab}
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
                    Ý tưởng/Đề tài đã lưu
                  </h2>
                  <span className="text-sm text-teal-200 font-medium bg-teal-900/50 px-3 py-1 rounded-full border border-teal-700/50">
                    {savedIdeas.length} ý tưởng
                  </span>
                </div>

                {savedIdeas.length === 0 ? (
                  <div className="bg-teal-900/40 backdrop-blur-sm rounded-2xl shadow-sm border border-teal-700/50 p-12 text-center">
                    <Heart className="w-12 h-12 text-teal-700 mx-auto mb-4" />
                    <h3 className="text-lg font-bold text-teal-100 mb-2">Chưa có mục nào</h3>
                    <p className="text-teal-300">Hãy nhấn nút "Lưu ý tưởng" bên cạnh mỗi đề xuất để lưu lại.</p>
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
                           <span className="text-xs font-bold text-emerald-300 bg-emerald-900/30 px-2.5 py-1 rounded-lg border border-emerald-700/50">
                             {getContestMeta(idea.inputs.contestType || inferContestTypeFromField(idea.inputs.field)).shortLabel}
                           </span>
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
                        <div className="flex items-center gap-2 mb-1.5">
                          <span className="text-[10px] font-extrabold text-emerald-300 bg-emerald-900/30 px-2 py-0.5 rounded-md border border-emerald-700/50">
                            {getContestMeta(session.inputs.contestType || inferContestTypeFromField(session.inputs.field)).shortLabel}
                          </span>
                          <h4 className="font-bold text-teal-50 line-clamp-1 text-base">{session.inputs.field}</h4>
                        </div>
                        <p className="text-sm text-teal-200 mb-4 flex items-center gap-2">
                          <span className="bg-teal-800/50 px-2 py-0.5 rounded text-xs border border-teal-700/50">{session.inputs.capHoc}</span>
                          <span className="text-teal-600">•</span>
                          <span className="truncate">{session.inputs.grade}</span>
                          {session.inputs.contestType === 'khkt' && (
                            <>
                              <span className="text-teal-600">•</span>
                              <span className="truncate">{getKhktProjectTypeLabel(session.inputs.khktProjectType || 'auto')}</span>
                            </>
                          )}
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
                    <h2 className="text-xl font-bold text-teal-100 mb-2">{contestMeta.emptyTitle}</h2>
                    <p className="text-teal-300 text-center max-w-sm">
                      Hãy điền các thông tin cần thiết ở thanh bên trái và nhấn <strong className="text-teal-50">{contestMeta.actionLabel}</strong> để bắt đầu.
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
                      {isKhktContest
                        ? 'Hệ thống đang phân tích bằng AI nâng cao để đề xuất đề tài bám 22 lĩnh vực và phiếu chấm KHKT 100 điểm. Quá trình này có thể mất thêm chút thời gian.'
                        : 'Hệ thống đang phân tích bằng AI nâng cao để đảm bảo các ý tưởng đề xuất có tính mới và sáng tạo cao nhất. Quá trình này có thể mất thêm chút thời gian.'}
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
                      <p className="text-xs text-teal-300">Tạo lại danh sách mới hoặc so sánh chuyên sâu một {isKhktContest ? 'đề tài' : 'ý tưởng'}.</p>
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
                            generateAndDownloadWordDoc(result, `${isKhktContest ? 'Đề tài KHKT' : 'Ý Tưởng Sáng Tạo'} - ${field}`);
                          }}
                          className="px-5 py-2.5 bg-teal-800/50 border border-teal-600/50 hover:bg-teal-700/50 hover:border-teal-500/50 text-teal-100 rounded-xl font-semibold flex items-center justify-center gap-2 transition-all shadow-sm active:scale-[0.98]"
                        >
                          <Download className="w-4 h-4" />
                          Tải Word
                        </button>
                        <button
                          onClick={() => {
                            generateAndDownloadWordDoc(result, `${isKhktContest ? 'Đề tài KHKT' : 'Ý Tưởng Sáng Tạo'} - ${field}`, true);
                          }}
                          className="px-5 py-2.5 bg-emerald-900/30 text-emerald-300 hover:bg-emerald-800/40 border border-emerald-700/50 rounded-xl font-semibold flex items-center justify-center gap-2 transition-all shadow-sm active:scale-[0.98]"
                        >
                          <Download className="w-4 h-4" />
                          Word tóm tắt
                        </button>
                      </div>
                      
                      <div className="flex items-center gap-2 flex-1 sm:flex-none">
                        <select 
                          id="compareSelect"
                          className="p-2.5 bg-teal-900/50 border border-teal-700/50 rounded-xl text-sm font-medium focus:ring-2 focus:ring-emerald-500/20 outline-none text-teal-50"
                        >
                          {Array.from({ length: 20 }, (_, index) => index + 1).map(num => (
                            <option key={num} value={num} className="bg-teal-900 text-teal-50">{isKhktContest ? 'Đề tài' : 'Ý tưởng'} {num}</option>
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
