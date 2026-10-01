import React, { useState, useMemo } from 'react';
import {
  FileText,
  Download,
  Copy,
  Check,
  Play,
  Layers,
  Palette,
  Clock,
  User,
  Shield,
  UploadCloud,
  ChevronLeft,
  ChevronRight,
  ExternalLink,
  BookOpen,
  Sparkles,
  Info,
  CheckCircle2,
  FileCode,
  FileSpreadsheet,
  AlertCircle
} from 'lucide-react';
import {
  AML_MODULES,
  THEMES,
  AmlModule,
  ThemeConfig,
  SAMPLE_CASE_PRESETS,
} from './amlData';
import { generatePptxFile } from './pptxGenerator';
import { APP_PY_CODE } from './appPyCode';

export default function App() {
  // Sidebar State
  const [lectureTitle, setLectureTitle] = useState(
    '금융기관 임직원을 위한 자금세탁방지(AML/CFT) 핵심 실무'
  );
  const [targetAudience, setTargetAudience] = useState('영업점실무자');
  const [durationMin, setDurationMin] = useState(60);
  const [instructorName, setInstructorName] = useState(
    '준법감시팀장 / AML 공인전문가(CAMS)'
  );
  const [themeKey, setThemeKey] = useState<string>('orange');
  const [customText, setCustomText] = useState('');
  const [uploadedFileName, setUploadedFileName] = useState<string | null>(null);

  // Module Selection
  const [selectedModuleIds, setSelectedModuleIds] = useState<string[]>([
    'aml_basics',
    'cdd_edd',
    'str_ctr',
    'internal_audit',
    'virtual_assets',
    'sanctions',
  ]);

  // Main UI Tab
  const [activeTab, setActiveTab] = useState<'builder' | 'code' | 'guide'>('builder');

  // Slide Preview Navigation
  const [currentSlideIndex, setCurrentSlideIndex] = useState(0);
  const [isGeneratingPptx, setIsGeneratingPptx] = useState(false);
  const [copiedCode, setCopiedCode] = useState(false);
  const [generationSuccessMsg, setGenerationSuccessMsg] = useState<string | null>(null);

  const currentTheme = THEMES[themeKey] || THEMES.orange;

  const selectedModules = useMemo(() => {
    return AML_MODULES.filter((m) => selectedModuleIds.includes(m.id));
  }, [selectedModuleIds]);

  // All Slide Deck Data for Live 16:9 Preview
  interface SlideDeckItem {
    type: 'cover' | 'agenda' | 'module' | 'custom' | 'final';
    title: string;
    badge: string;
    subtitle?: string;
    moduleData?: AmlModule;
  }

  const slideDeck = useMemo<SlideDeckItem[]>(() => {
    const deck: SlideDeckItem[] = [
      {
        type: 'cover',
        title: lectureTitle,
        badge: 'Slide 01 • 표지',
      },
      {
        type: 'agenda',
        title: '강의 진행 목차 및 구성 개요',
        badge: 'Slide 02 • 목차 (Agenda)',
        subtitle: '선택된 모듈 기반 2열 카드 그리드 자동 배치',
      },
    ];

    selectedModules.forEach((mod, idx) => {
      deck.push({
        type: 'module',
        title: mod.name,
        badge: `Slide ${String(idx + 3).padStart(2, '0')} • ${mod.badge}`,
        subtitle: mod.subtitle,
        moduleData: mod,
      });
    });

    if (customText.trim()) {
      deck.push({
        type: 'custom',
        title: '최신 법령 개정 및 현장 이슈 사례 분석',
        badge: `Slide ${String(deck.length + 1).padStart(2, '0')} • 특화 사례 & 법령`,
        subtitle: '사용자 입력 및 업로드 자료 기반 실무 체크리스트',
      });
    }

    deck.push({
      type: 'final',
      title: '질의응답(Q&A) 및 준법 실천 다짐',
      badge: `Slide ${String(deck.length + 1).padStart(2, '0')} • Q&A / 마무리`,
      subtitle: '준법감시실 자금세탁방지팀 핫라인 및 사내 신고 채널 안내',
    });

    return deck;
  }, [lectureTitle, selectedModules, customText]);

  // Handle Module Toggle
  const toggleModule = (id: string) => {
    setSelectedModuleIds((prev) =>
      prev.includes(id) ? prev.filter((m) => m !== id) : [...prev, id]
    );
  };

  const selectAllModules = () => {
    setSelectedModuleIds(AML_MODULES.map((m) => m.id));
  };

  const clearAllModules = () => {
    setSelectedModuleIds([]);
  };

  // Handle File Upload (Client-side text parsing for preview)
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setUploadedFileName(file.name);
    const reader = new FileReader();

    if (file.name.endsWith('.txt')) {
      reader.onload = (event) => {
        const text = event.target?.result as string;
        setCustomText((prev) => (prev ? prev + '\n\n' + text : text));
      };
      reader.readAsText(file, 'utf-8');
    } else {
      // For PDF or DOCX in client preview, simulate extraction summary
      const simulatedText = `[업로드 파일 요약: ${file.name}]\n• 파일 크기: ${(file.size / 1024).toFixed(1)} KB\n• 감지된 주제: 금융위원회 특금법 고시 개정안 및 의심거래보고(STR) 가이드라인\n• 추출 요점: STR 보고 기한(3영업일) 엄수 및 고액 자금세탁 방지 모니터링 강화 의무`;
      setCustomText((prev) => (prev ? prev + '\n\n' + simulatedText : simulatedText));
    }
  };

  // PPTX Export Handler
  const handleGeneratePptx = async () => {
    if (selectedModules.length === 0) {
      alert('최소 1개 이상의 강의 모듈을 선택해 주세요.');
      return;
    }

    setIsGeneratingPptx(true);
    setGenerationSuccessMsg(null);
    try {
      const fileName = await generatePptxFile({
        lectureTitle,
        targetAudience,
        durationMin,
        instructorName,
        theme: currentTheme,
        selectedModules,
        customText,
      });
      setGenerationSuccessMsg(`🎉 "${fileName}" 프레젠테이션 생성이 완료되어 다운로드되었습니다!`);
    } catch (err) {
      console.error(err);
      alert('PPTX 파일 생성 중 문제가 발생했습니다: ' + String(err));
    } finally {
      setIsGeneratingPptx(false);
    }
  };

  // Download Python file
  const handleDownloadAppPy = () => {
    const blob = new Blob([APP_PY_CODE], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'app.py';
    link.click();
    URL.revokeObjectURL(url);
  };

  const handleDownloadRequirements = () => {
    const text = `streamlit>=1.30.0\npython-pptx>=0.6.21\npypdf>=3.0.0\nPyPDF2>=3.0.0\npython-docx>=0.8.11\n`;
    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'requirements.txt';
    link.click();
    URL.revokeObjectURL(url);
  };

  const handleCopyCode = () => {
    navigator.clipboard.writeText(APP_PY_CODE);
    setCopiedCode(true);
    setTimeout(() => setCopiedCode(false), 2000);
  };

  const currentSlide = slideDeck[Math.min(currentSlideIndex, slideDeck.length - 1)] || slideDeck[0];

  return (
    <div className="min-h-screen bg-gradient-to-b from-orange-50/80 via-amber-50/40 to-orange-100/50 flex flex-col font-sans text-stone-800">
      {/* Top Navigation Bar */}
      <header className="bg-stone-900 text-white border-b-2 border-orange-500 sticky top-0 z-30 shadow-md">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-gradient-to-br from-amber-400 to-orange-500 rounded-lg text-stone-950 font-black shadow-md flex items-center justify-center">
              <Shield className="w-5 h-5 text-stone-950" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="font-extrabold text-lg tracking-tight">AML Slide Builder</span>
                <span className="text-xs px-2 py-0.5 bg-orange-950 text-orange-300 border border-orange-800/60 rounded font-mono font-medium">
                  Python-pptx & Streamlit
                </span>
                <span className="text-xs px-2 py-0.5 bg-amber-950 text-amber-300 border border-amber-800/60 rounded font-mono font-medium">
                  16:9 Widescreen
                </span>
              </div>
              <p className="text-xs text-stone-400">자금세탁방지(AML/CFT) 강의 프레젠테이션 자동 생성 솔루션</p>
            </div>
          </div>

          <div className="flex items-center gap-3">
            <button
              onClick={() => setActiveTab('code')}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all ${
                activeTab === 'code'
                  ? 'bg-gradient-to-r from-amber-400 to-orange-500 text-stone-950 font-bold'
                  : 'bg-stone-800 hover:bg-stone-700 text-stone-300'
              }`}
            >
              <FileCode className="w-4 h-4" />
              <span>app.py 파이썬 코드</span>
            </button>

            <button
              onClick={handleGeneratePptx}
              disabled={isGeneratingPptx || selectedModules.length === 0}
              className="px-4 py-1.5 bg-orange-600 hover:bg-orange-500 disabled:opacity-50 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-md shadow-orange-600/20 transition-all active:scale-95"
            >
              <Download className="w-4 h-4" />
              <span>{isGeneratingPptx ? 'PPTX 빌드 중...' : 'PPTX 즉시 다운로드'}</span>
            </button>
          </div>
        </div>
      </header>

      {/* Main Workspace Layout */}
      <div className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 lg:p-8 flex flex-col lg:flex-row gap-6">
        {/* ================================================================= */}
        {/* SIDEBAR: Controls & Inputs */}
        {/* ================================================================= */}
        <aside className="w-full lg:w-96 shrink-0 flex flex-col gap-5">
          {/* Section 1: Lecture Basic Config */}
          <div className="bg-white rounded-xl shadow-xs border border-orange-200/80 p-5">
            <div className="flex items-center gap-2 pb-3 mb-4 border-b border-orange-100">
              <span className="w-6 h-6 rounded-full bg-orange-100 text-orange-800 font-bold text-xs flex items-center justify-center">
                1
              </span>
              <h2 className="font-bold text-stone-900 text-sm">강의 기본 설정</h2>
            </div>

            <div className="space-y-4 text-xs">
              <div>
                <label className="block font-semibold text-stone-700 mb-1">강의 제목</label>
                <input
                  type="text"
                  value={lectureTitle}
                  onChange={(e) => setLectureTitle(e.target.value)}
                  className="w-full px-3 py-2 border border-stone-300 rounded-lg text-xs focus:ring-2 focus:ring-orange-500 focus:outline-none"
                  placeholder="강의 제목을 입력하세요"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-stone-700 mb-1">수강 대상</label>
                  <select
                    value={targetAudience}
                    onChange={(e) => setTargetAudience(e.target.value)}
                    className="w-full px-2.5 py-2 border border-stone-300 rounded-lg text-xs focus:ring-2 focus:ring-orange-500 focus:outline-none bg-white"
                  >
                    <option value="신규입사자">신규입사자</option>
                    <option value="영업점실무자">영업점실무자</option>
                    <option value="AML전담부서">AML전담부서</option>
                    <option value="임원 및 준법감시인">임원 및 준법감시인</option>
                  </select>
                </div>

                <div>
                  <label className="block font-semibold text-stone-700 mb-1">
                    강의 시간: <span className="text-orange-600 font-bold">{durationMin}분</span>
                  </label>
                  <input
                    type="range"
                    min="30"
                    max="120"
                    step="10"
                    value={durationMin}
                    onChange={(e) => setDurationMin(Number(e.target.value))}
                    className="w-full h-2 bg-orange-100 rounded-lg appearance-none cursor-pointer accent-orange-600 mt-2"
                  />
                </div>
              </div>

              <div>
                <label className="block font-semibold text-stone-700 mb-1">강사명 및 직책</label>
                <div className="relative">
                  <User className="w-3.5 h-3.5 text-stone-400 absolute left-3 top-2.5" />
                  <input
                    type="text"
                    value={instructorName}
                    onChange={(e) => setInstructorName(e.target.value)}
                    className="w-full pl-8 pr-3 py-2 border border-stone-300 rounded-lg text-xs focus:ring-2 focus:ring-orange-500 focus:outline-none"
                    placeholder="홍길동 준법감시팀장 / CAMS"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Section 2: Theme Config */}
          <div className="bg-white rounded-xl shadow-xs border border-orange-200/80 p-5">
            <div className="flex items-center gap-2 pb-3 mb-4 border-b border-orange-100">
              <span className="w-6 h-6 rounded-full bg-orange-100 text-orange-800 font-bold text-xs flex items-center justify-center">
                2
              </span>
              <h2 className="font-bold text-stone-900 text-sm">프레젠테이션 테마</h2>
            </div>

            <div className="space-y-2">
              {Object.entries(THEMES).map(([key, item]) => {
                const isSelected = themeKey === key;
                return (
                  <button
                    key={key}
                    type="button"
                    onClick={() => setThemeKey(key)}
                    className={`w-full text-left p-3 rounded-lg border text-xs flex items-center justify-between transition-all ${
                      isSelected
                        ? 'border-orange-500 bg-orange-50/80 shadow-xs ring-1 ring-orange-400'
                        : 'border-stone-200 hover:bg-orange-50/40'
                    }`}
                  >
                    <div>
                      <div className="font-bold text-stone-800 flex items-center gap-1.5">
                        <span>{item.name}</span>
                        {key === 'orange' && (
                          <span className="px-1.5 py-0.2 rounded text-[10px] bg-orange-200 text-orange-900 font-medium">
                            추천
                          </span>
                        )}
                      </div>
                      <div className="text-[11px] text-stone-500 mt-0.5">
                        {key === 'orange'
                          ? '신뢰와 활력의 앰버 오렌지 (따뜻한 피치 오렌지 슬라이드 배경)'
                          : key === 'brightOrange'
                          ? '선명하고 세련된 모던 선셋 오렌지 배경'
                          : key === 'terracotta'
                          ? '클래식 테라코타 러스티 오렌지'
                          : key === 'red'
                          ? '프리미엄 딥 크림슨 레드 (경각심 & 제재)'
                          : key === 'navy'
                          ? '신뢰도 높은 정통 컴플라이언스'
                          : '모던 핀테크 및 금융 혁신 스타일'}
                      </div>
                    </div>
                    <div className="flex items-center gap-1.5 shrink-0 ml-2">
                      <span
                        className="w-5 h-5 rounded-full border border-stone-300 shadow-inner"
                        style={{ backgroundColor: item.primary }}
                        title="Primary Color"
                      />
                      <span
                        className="w-5 h-5 rounded-full border border-stone-300 shadow-inner"
                        style={{ backgroundColor: item.accent }}
                        title="Accent Color"
                      />
                    </div>
                  </button>
                );
              })}
            </div>
          </div>

          {/* Section 3: Custom Materials / Upload */}
          <div className="bg-white rounded-xl shadow-xs border border-orange-200/80 p-5">
            <div className="flex items-center justify-between pb-3 mb-4 border-b border-orange-100">
              <div className="flex items-center gap-2">
                <span className="w-6 h-6 rounded-full bg-orange-100 text-orange-800 font-bold text-xs flex items-center justify-center">
                  3
                </span>
                <h2 className="font-bold text-stone-900 text-sm">최신 자료 및 사례 반영</h2>
              </div>
              <span className="text-[11px] text-stone-400">선택 사항</span>
            </div>

            <div className="space-y-3 text-xs">
              <p className="text-stone-500 text-[11px] leading-relaxed">
                법령 개정안, 감독원 제재 사례를 입력하거나 파일을 업로드하면 전용 슬라이드가 자동으로 추가됩니다.
              </p>

              <div>
                <textarea
                  rows={4}
                  value={customText}
                  onChange={(e) => setCustomText(e.target.value)}
                  placeholder="예: 2024년 KoFIU 검사 결과: STR 보고 지연 및 가상자산 연계 차명계좌 거래 미식별로 시중은행 과태료 3.2억원 부과..."
                  className="w-full p-2.5 border border-stone-300 rounded-lg text-xs focus:ring-2 focus:ring-orange-500 focus:outline-none"
                />
              </div>

              {/* File Uploader */}
              <div>
                <label className="flex flex-col items-center justify-center p-3 border-2 border-dashed border-orange-200 hover:border-orange-400 rounded-lg cursor-pointer bg-orange-50/30 hover:bg-orange-50/70 transition-all text-center">
                  <UploadCloud className="w-5 h-5 text-orange-400 mb-1" />
                  <span className="text-[11px] font-semibold text-stone-700">
                    파일 업로드 (.txt, .pdf, .docx)
                  </span>
                  <span className="text-[10px] text-stone-400">클릭하여 파일 선택</span>
                  <input
                    type="file"
                    accept=".txt,.pdf,.docx"
                    onChange={handleFileUpload}
                    className="hidden"
                  />
                </label>
                {uploadedFileName && (
                  <div className="mt-1.5 text-[11px] text-emerald-600 flex items-center gap-1 font-medium">
                    <CheckCircle2 className="w-3.5 h-3.5" />
                    <span>업로드 완료: {uploadedFileName}</span>
                  </div>
                )}
              </div>

              {/* Case Preset Quick Loaders */}
              <div className="pt-2 border-t border-orange-100">
                <span className="text-[11px] font-semibold text-stone-600 block mb-1.5">
                  ⚡ 실무 사례 프리셋 불러오기
                </span>
                <div className="space-y-1">
                  {SAMPLE_CASE_PRESETS.map((preset, idx) => (
                    <button
                      key={idx}
                      type="button"
                      onClick={() => setCustomText(preset.text)}
                      className="w-full text-left px-2 py-1.5 bg-orange-50/60 hover:bg-orange-100 hover:text-orange-900 rounded text-[11px] text-stone-700 truncate transition-colors"
                      title={preset.title}
                    >
                      • {preset.title}
                    </button>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </aside>

        {/* ================================================================= */}
        {/* MAIN STAGE: Tabs & Presentation Builder */}
        {/* ================================================================= */}
        <main className="flex-1 flex flex-col gap-5 min-w-0">
          {/* Navigation Tabs */}
          <div className="bg-white rounded-xl shadow-xs border border-orange-200/80 p-1 flex items-center gap-1">
            <button
              onClick={() => setActiveTab('builder')}
              className={`flex-1 py-2.5 px-4 rounded-lg text-xs font-bold flex items-center justify-center gap-2 transition-all ${
                activeTab === 'builder'
                  ? 'bg-orange-600 text-white shadow-xs'
                  : 'text-stone-700 hover:bg-orange-50'
              }`}
            >
              <Layers className="w-4 h-4" />
              <span>슬라이드 빌더 & 16:9 실시간 프리뷰</span>
              <span className={`ml-1 px-1.5 py-0.5 rounded-full text-[10px] ${
                activeTab === 'builder' ? 'bg-orange-700 text-white' : 'bg-orange-100 text-orange-900'
              }`}>
                {slideDeck.length}장
              </span>
            </button>

            <button
              onClick={() => setActiveTab('code')}
              className={`flex-1 py-2.5 px-4 rounded-lg text-xs font-bold flex items-center justify-center gap-2 transition-all ${
                activeTab === 'code'
                  ? 'bg-stone-900 text-amber-400 shadow-xs border border-stone-800'
                  : 'text-stone-700 hover:bg-orange-50'
              }`}
            >
              <FileCode className="w-4 h-4" />
              <span>Python 소스 코드 (app.py)</span>
            </button>

            <button
              onClick={() => setActiveTab('guide')}
              className={`flex-1 py-2.5 px-4 rounded-lg text-xs font-bold flex items-center justify-center gap-2 transition-all ${
                activeTab === 'guide'
                  ? 'bg-orange-600 text-white shadow-xs'
                  : 'text-stone-700 hover:bg-orange-50'
              }`}
            >
              <BookOpen className="w-4 h-4" />
              <span>AML 핵심 실무 가이드</span>
            </button>
          </div>

          {/* TAB 1: SLIDE BUILDER & LIVE 16:9 PREVIEW */}
          {activeTab === 'builder' && (
            <div className="space-y-6">
              {/* Notification Message if generated */}
              {generationSuccessMsg && (
                <div className="p-4 bg-emerald-50 border border-emerald-300 rounded-xl flex items-center justify-between text-xs text-emerald-800">
                  <div className="flex items-center gap-2">
                    <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                    <span className="font-semibold">{generationSuccessMsg}</span>
                  </div>
                  <button
                    onClick={() => setGenerationSuccessMsg(null)}
                    className="text-emerald-700 hover:text-emerald-900 font-bold"
                  >
                    닫기
                  </button>
                </div>
              )}

              {/* Step 1: Module Selector (Curriculum Checkboxes) */}
              <div className="bg-white rounded-xl shadow-xs border border-orange-200/80 p-5">
                <div className="flex items-center justify-between mb-3 pb-3 border-b border-orange-100">
                  <div>
                    <h3 className="font-bold text-stone-900 text-sm flex items-center gap-2">
                      <span>1. 강의 모듈 선택 (Curriculum Selection)</span>
                      <span className="text-xs font-normal text-stone-500">
                        ({selectedModules.length} / {AML_MODULES.length}개 선택됨)
                      </span>
                    </h3>
                    <p className="text-xs text-stone-500 mt-0.5">
                      강의 목적과 수강 대상 난이도에 맞게 모듈을 자유롭게 켜고 끌 수 있습니다.
                    </p>
                  </div>
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={selectAllModules}
                      className="text-xs font-semibold text-orange-600 hover:text-orange-800 px-2 py-1 rounded bg-orange-50 hover:bg-orange-100"
                    >
                      전체 선택
                    </button>
                    <button
                      type="button"
                      onClick={clearAllModules}
                      className="text-xs font-semibold text-stone-500 hover:text-stone-700 px-2 py-1 rounded bg-stone-100 hover:bg-stone-200"
                    >
                      선택 해제
                    </button>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  {AML_MODULES.map((mod) => {
                    const isChecked = selectedModuleIds.includes(mod.id);
                    return (
                      <label
                        key={mod.id}
                        className={`p-3 rounded-xl border text-xs cursor-pointer flex items-start gap-3 transition-all ${
                          isChecked
                            ? 'border-orange-500 bg-orange-50/60 shadow-xs'
                            : 'border-stone-200 hover:bg-orange-50/20 opacity-75'
                        }`}
                      >
                        <input
                          type="checkbox"
                          checked={isChecked}
                          onChange={() => toggleModule(mod.id)}
                          className="mt-0.5 rounded border-stone-300 text-orange-600 focus:ring-orange-500 w-4 h-4 cursor-pointer"
                        />
                        <div className="min-w-0 flex-1">
                          <div className="flex items-center gap-2 mb-1">
                            <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-orange-100 text-orange-900">
                              {mod.badge}
                            </span>
                            <span className="font-bold text-stone-900 truncate">{mod.name}</span>
                          </div>
                          <p className="text-[11px] text-stone-500 line-clamp-2 leading-relaxed">
                            {mod.subtitle}
                          </p>
                        </div>
                      </label>
                    );
                  })}
                </div>
              </div>

              {/* Step 2: 16:9 Presentation Live Viewer */}
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
                <div className="flex items-center justify-between mb-4">
                  <div>
                    <h3 className="font-bold text-slate-900 text-sm flex items-center gap-2">
                      <span>2. 슬라이드 실시간 16:9 와이드스크린 프리뷰</span>
                      <span className="px-2 py-0.5 rounded text-[11px] font-bold bg-slate-100 text-slate-700">
                        {currentSlideIndex + 1} / {slideDeck.length} 슬라이드
                      </span>
                    </h3>
                    <p className="text-xs text-slate-500">
                      실제 생성될 파워포인트 슬라이드와 동일한 비율(16:9) 및 디자인 레이아웃을 확인하세요.
                    </p>
                  </div>

                  <div className="flex items-center gap-1.5">
                    <button
                      onClick={() => setCurrentSlideIndex((prev) => Math.max(0, prev - 1))}
                      disabled={currentSlideIndex === 0}
                      className="p-1.5 rounded-lg border border-slate-200 hover:bg-slate-100 disabled:opacity-40 transition-colors"
                      title="이전 슬라이드"
                    >
                      <ChevronLeft className="w-4 h-4" />
                    </button>
                    <button
                      onClick={() =>
                        setCurrentSlideIndex((prev) =>
                          Math.min(slideDeck.length - 1, prev + 1)
                        )
                      }
                      disabled={currentSlideIndex === slideDeck.length - 1}
                      className="p-1.5 rounded-lg border border-slate-200 hover:bg-slate-100 disabled:opacity-40 transition-colors"
                      title="다음 슬라이드"
                    >
                      <ChevronRight className="w-4 h-4" />
                    </button>
                  </div>
                </div>

                {/* 16:9 Canvas Container */}
                <div className="relative w-full aspect-[16/9] bg-slate-950 rounded-xl overflow-hidden shadow-xl border border-slate-300">
                  {/* SLIDE 1: COVER */}
                  {currentSlide.type === 'cover' && (
                    <div
                      className="w-full h-full relative p-6 sm:p-10 flex flex-col justify-between text-white"
                      style={{ backgroundColor: currentTheme.primary }}
                    >
                      {/* Top Accent Strip */}
                      <div
                        className="absolute top-0 left-0 right-0 h-2 sm:h-3"
                        style={{ backgroundColor: currentTheme.accent }}
                      />

                      {/* Cover Content */}
                      <div className="mt-4 sm:mt-6">
                        <span
                          className="inline-block text-[10px] sm:text-xs font-extrabold tracking-wider uppercase mb-2 px-2.5 py-1 rounded bg-white/10"
                          style={{ color: currentTheme.accent }}
                        >
                          🛡️ Compliance & Financial Crime Prevention
                        </span>
                        <h1 className="text-xl sm:text-3xl lg:text-4xl font-black text-white leading-tight max-w-3xl drop-shadow-sm">
                          {lectureTitle}
                        </h1>
                        <p className="text-xs sm:text-sm text-orange-100 mt-2 sm:mt-3 max-w-2xl font-light">
                          자금세탁방지(AML) 및 공중협박자금조달금지(CFT) 제도 실무 역량 강화를 위한 핵심 교육
                        </p>
                      </div>

                      {/* Cover Meta Box */}
                      <div
                        className="p-3 sm:p-4 rounded-lg border border-white/10 flex flex-wrap items-center justify-between gap-3 text-[11px] sm:text-xs"
                        style={{ backgroundColor: currentTheme.secondary }}
                      >
                        <div className="flex items-center gap-4">
                          <div>
                            <span className="text-orange-200 block text-[10px]">수강 대상</span>
                            <span className="font-bold text-white">{targetAudience}</span>
                          </div>
                          <div className="h-6 w-px bg-white/20" />
                          <div>
                            <span className="text-orange-200 block text-[10px]">강의 시간</span>
                            <span className="font-bold text-white">{durationMin}분</span>
                          </div>
                        </div>

                        <div className="flex items-center gap-4">
                          <div>
                            <span className="text-orange-200 block text-[10px]">전담 강사</span>
                            <span className="font-bold text-white">{instructorName}</span>
                          </div>
                          <div className="h-6 w-px bg-white/20" />
                          <div>
                            <span className="text-orange-200 block text-[10px]">시행일자</span>
                            <span className="font-bold text-slate-100">
                              {new Date().toLocaleDateString('ko-KR')}
                            </span>
                          </div>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* SLIDE 2: AGENDA */}
                  {currentSlide.type === 'agenda' && (
                    <div
                      className="w-full h-full p-6 sm:p-8 flex flex-col justify-between text-stone-800"
                      style={{ backgroundColor: currentTheme.slideBg }}
                    >
                      <div>
                        <div
                          className="text-[10px] sm:text-xs font-black tracking-wider uppercase"
                          style={{ color: currentTheme.accent }}
                        >
                          TABLE OF CONTENTS
                        </div>
                        <h2
                          className="text-lg sm:text-2xl font-black mt-0.5"
                          style={{ color: currentTheme.primary }}
                        >
                          강의 진행 목차 및 구성 개요
                        </h2>
                      </div>

                      {/* 2-Column Grid */}
                      <div className="grid grid-cols-2 gap-2.5 sm:gap-3 my-2 overflow-y-auto max-h-[68%] pr-1">
                        {selectedModules.map((mod, idx) => (
                          <div
                            key={mod.id}
                            className="p-2 sm:p-2.5 rounded-lg border flex items-center gap-2.5"
                            style={{
                              backgroundColor: currentTheme.cardBg,
                              borderColor: currentTheme.cardBorder,
                            }}
                          >
                            <span
                              className="w-7 h-7 sm:w-8 sm:h-8 rounded font-black text-white text-xs sm:text-sm flex items-center justify-center shrink-0 shadow-xs"
                              style={{ backgroundColor: currentTheme.primary }}
                            >
                              {String(idx + 1).padStart(2, '0')}
                            </span>
                            <div className="min-w-0">
                              <h4 className="font-bold text-[11px] sm:text-xs text-slate-900 truncate">
                                {mod.name}
                              </h4>
                              <p className="text-[9px] sm:text-[10px] text-slate-500 truncate">
                                {mod.subtitle}
                              </p>
                            </div>
                          </div>
                        ))}

                        {customText.trim() && (
                          <div
                            className="p-2 sm:p-2.5 rounded-lg border flex items-center gap-2.5"
                            style={{
                              backgroundColor: currentTheme.cardBg,
                              borderColor: currentTheme.cardBorder,
                            }}
                          >
                            <span
                              className="w-7 h-7 sm:w-8 sm:h-8 rounded font-black text-white text-xs sm:text-sm flex items-center justify-center shrink-0 shadow-xs"
                              style={{ backgroundColor: currentTheme.accent }}
                            >
                              {String(selectedModules.length + 1).padStart(2, '0')}
                            </span>
                            <div className="min-w-0">
                              <h4 className="font-bold text-[11px] sm:text-xs text-slate-900 truncate">
                                최신 법령 개정 및 현장 이슈 사례
                              </h4>
                              <p className="text-[9px] sm:text-[10px] text-slate-500 truncate">
                                업로드 및 입력 특화 데이터 기반 실무 케이스
                              </p>
                            </div>
                          </div>
                        )}

                        <div
                          className="p-2 sm:p-2.5 rounded-lg border flex items-center gap-2.5"
                          style={{
                            backgroundColor: currentTheme.cardBg,
                            borderColor: currentTheme.cardBorder,
                          }}
                        >
                          <span
                            className="w-7 h-7 sm:w-8 sm:h-8 rounded font-black text-white text-xs sm:text-sm flex items-center justify-center shrink-0 shadow-xs"
                            style={{ backgroundColor: currentTheme.secondary }}
                          >
                            {String(selectedModules.length + (customText.trim() ? 2 : 1)).padStart(2, '0')}
                          </span>
                          <div className="min-w-0">
                            <h4 className="font-bold text-[11px] sm:text-xs text-slate-900 truncate">
                              Q&A 및 컴플라이언스 실천 다짐
                            </h4>
                            <p className="text-[9px] sm:text-[10px] text-slate-500 truncate">
                              질의응답 및 준법지원 핫라인 안내
                            </p>
                          </div>
                        </div>
                      </div>

                      {/* Footer */}
                      <div className="text-[9px] sm:text-[10px] text-slate-400 border-t border-slate-100 pt-1 flex justify-between">
                        <span>AML/CFT Compliance Education  |  {lectureTitle}</span>
                        <span>Slide {currentSlideIndex + 1} of {slideDeck.length}</span>
                      </div>
                    </div>
                  )}

                  {/* SLIDES 3..N: MODULE CONTENT SLIDES */}
                  {currentSlide.type === 'module' && currentSlide.moduleData && (
                    <div
                      className="w-full h-full p-5 sm:p-7 flex flex-col justify-between text-stone-800"
                      style={{ backgroundColor: currentTheme.slideBg }}
                    >
                      {/* Header */}
                      <div>
                        <span
                          className="text-[10px] sm:text-xs font-black uppercase"
                          style={{ color: currentTheme.accent }}
                        >
                          {currentSlide.moduleData.badge}
                        </span>
                        <h2
                          className="text-base sm:text-xl font-black"
                          style={{ color: currentTheme.primary }}
                        >
                          {currentSlide.moduleData.name}
                        </h2>
                        <p className="text-[10px] sm:text-xs text-slate-500 truncate">
                          {currentSlide.moduleData.subtitle}
                        </p>
                      </div>

                      {/* 3 Columns Section Cards */}
                      <div className="grid grid-cols-3 gap-2.5 sm:gap-3 my-2 flex-1">
                        {currentSlide.moduleData.sections.map((sec, sIdx) => (
                          <div
                            key={sIdx}
                            className="rounded-lg p-2.5 sm:p-3 border flex flex-col justify-start relative overflow-hidden"
                            style={{
                              backgroundColor: currentTheme.cardBg,
                              borderColor: currentTheme.cardBorder,
                            }}
                          >
                            <div
                              className="absolute top-0 left-0 right-0 h-1"
                              style={{
                                backgroundColor:
                                  sIdx === 1 ? currentTheme.accent : currentTheme.secondary,
                              }}
                            />
                            <h4
                              className="font-bold text-[11px] sm:text-xs mb-2 mt-0.5 truncate"
                              style={{ color: currentTheme.primary }}
                            >
                              {sec.title}
                            </h4>
                            <ul className="space-y-1.5 text-[9px] sm:text-[11px] text-slate-700 overflow-y-auto">
                              {sec.bullets.map((b, bIdx) => (
                                <li key={bIdx} className="flex items-start gap-1 leading-relaxed">
                                  <span className="text-slate-400 shrink-0">•</span>
                                  <span>{b}</span>
                                </li>
                              ))}
                            </ul>
                          </div>
                        ))}
                      </div>

                      {/* Footer */}
                      <div className="text-[9px] sm:text-[10px] text-slate-400 border-t border-slate-100 pt-1 flex justify-between">
                        <span>AML/CFT Compliance Education  |  {lectureTitle}</span>
                        <span>Slide {currentSlideIndex + 1} of {slideDeck.length}</span>
                      </div>
                    </div>
                  )}

                  {/* SPECIAL SLIDE: CUSTOM MATERIAL / CASES */}
                  {currentSlide.type === 'custom' && (
                    <div
                      className="w-full h-full p-5 sm:p-7 flex flex-col justify-between text-stone-800"
                      style={{ backgroundColor: currentTheme.slideBg }}
                    >
                      <div>
                        <span
                          className="text-[10px] sm:text-xs font-black uppercase"
                          style={{ color: currentTheme.accent }}
                        >
                          SPECIAL TOPIC • 현장 맞춤형 자료
                        </span>
                        <h2
                          className="text-base sm:text-xl font-black"
                          style={{ color: currentTheme.primary }}
                        >
                          최신 법령 개정 및 현장 이슈 사례 심층 분석
                        </h2>
                        <p className="text-[10px] sm:text-xs text-slate-500">
                          강의 담당자 입력 자료 및 최근 금융감독원/KoFIU 검사 동향 반영
                        </p>
                      </div>

                      <div className="grid grid-cols-5 gap-3 my-2 flex-1">
                        {/* Left Card: 3 cols */}
                        <div
                          className="col-span-3 rounded-lg p-3 border flex flex-col"
                          style={{
                            backgroundColor: currentTheme.cardBg,
                            borderColor: currentTheme.cardBorder,
                          }}
                        >
                          <h4
                            className="font-bold text-xs sm:text-sm mb-2"
                            style={{ color: currentTheme.primary }}
                          >
                            📑 입력 및 업로드 자료 핵심 요약
                          </h4>
                          <div className="text-[10px] sm:text-xs text-slate-700 whitespace-pre-line leading-relaxed overflow-y-auto pr-1">
                            {customText || '입력된 자료가 없습니다.'}
                          </div>
                        </div>

                        {/* Right Card: 2 cols */}
                        <div
                          className="col-span-2 rounded-lg p-3 text-white flex flex-col justify-between"
                          style={{ backgroundColor: currentTheme.secondary }}
                        >
                          <div>
                            <h4
                              className="font-bold text-xs sm:text-sm mb-2"
                              style={{ color: currentTheme.accent }}
                            >
                              💡 실무 점검 가이드
                            </h4>
                            <ul className="space-y-2 text-[10px] sm:text-xs text-slate-100">
                              <li className="flex items-start gap-1.5">
                                <span className="text-emerald-400 font-bold">✔</span>
                                <span>관련 업무 지침 및 내부규정 즉시 업데이트 여부 점검</span>
                              </li>
                              <li className="flex items-start gap-1.5">
                                <span className="text-emerald-400 font-bold">✔</span>
                                <span>이상거래탐지(FDS/STR) 룰셋 고도화 및 임계치 검토</span>
                              </li>
                              <li className="flex items-start gap-1.5">
                                <span className="text-emerald-400 font-bold">✔</span>
                                <span>영업점 대상 즉시 공지 및 현장 전파 교육 시행</span>
                              </li>
                              <li className="flex items-start gap-1.5">
                                <span className="text-emerald-400 font-bold">✔</span>
                                <span>금융감독원 정기 종합검사 사전 점검 항목 반영</span>
                              </li>
                            </ul>
                          </div>

                          <div className="text-[9px] text-slate-300 border-t border-white/10 pt-2">
                            * 본 가이드는 특금법 준법감시 지침에 의거하여 편성되었습니다.
                          </div>
                        </div>
                      </div>

                      {/* Footer */}
                      <div className="text-[9px] sm:text-[10px] text-slate-400 border-t border-slate-100 pt-1 flex justify-between">
                        <span>AML/CFT Compliance Education  |  {lectureTitle}</span>
                        <span>Slide {currentSlideIndex + 1} of {slideDeck.length}</span>
                      </div>
                    </div>
                  )}

                  {/* FINAL SLIDE: Q&A */}
                  {currentSlide.type === 'final' && (
                    <div
                      className="w-full h-full relative p-6 sm:p-10 flex flex-col justify-between text-white"
                      style={{ backgroundColor: currentTheme.primary }}
                    >
                      <div
                        className="absolute top-0 left-0 right-0 h-2 sm:h-3"
                        style={{ backgroundColor: currentTheme.accent }}
                      />

                      <div className="mt-4">
                        <span
                          className="text-[10px] sm:text-xs font-black uppercase tracking-wider"
                          style={{ color: currentTheme.accent }}
                        >
                          COMPLIANCE MINDSET
                        </span>
                        <h1 className="text-xl sm:text-3xl font-black text-white mt-1">
                          Question & Answer  |  준법 실천 다짐
                        </h1>
                        <p className="text-xs sm:text-sm text-slate-300 mt-2 font-light">
                          “자금세탁방지는 형식적 의무가 아닌, 금융기관의 생존과 고객의 신뢰를 지키는 최후의 보루입니다.”
                        </p>
                      </div>

                      <div
                        className="p-3 sm:p-5 rounded-lg border border-white/10"
                        style={{ backgroundColor: currentTheme.secondary }}
                      >
                        <h4
                          className="font-bold text-xs sm:text-sm mb-2"
                          style={{ color: currentTheme.accent }}
                        >
                          📞 준법지원 및 의심거래 자문 핫라인
                        </h4>
                        <div className="space-y-1 text-[11px] sm:text-xs text-slate-200">
                          <p>• 담당 부서: 준법감시실 자금세탁방지팀 (내선: 02-XXXX-XXXX / aml-support@company.com)</p>
                          <p>• 보고 채널: 사내 인트라넷 익명 제보 시스템 및 KoFIU 전자보고 포털</p>
                          <p className="text-slate-400 text-[10px] pt-1">
                            • 수고 많으셨습니다. 본 강의 교안은 사내 LMS에서 상시 다운로드 및 복습 가능합니다.
                          </p>
                        </div>
                      </div>

                      <div className="text-[9px] sm:text-[10px] text-slate-400 border-t border-white/10 pt-1 flex justify-between">
                        <span>AML/CFT Compliance Education  |  {lectureTitle}</span>
                        <span>Slide {currentSlideIndex + 1} of {slideDeck.length}</span>
                      </div>
                    </div>
                  )}
                </div>

                {/* Thumbnail Strip */}
                <div className="flex items-center gap-2 mt-4 overflow-x-auto pb-2">
                  {slideDeck.map((s, idx) => (
                    <button
                      key={idx}
                      onClick={() => setCurrentSlideIndex(idx)}
                      className={`px-3 py-1.5 rounded-lg text-xs font-medium whitespace-nowrap transition-all flex items-center gap-1.5 ${
                        currentSlideIndex === idx
                          ? 'bg-orange-600 text-white font-bold shadow-xs'
                          : 'bg-orange-100/70 hover:bg-orange-200/80 text-orange-950 font-medium'
                      }`}
                    >
                      <span className={`w-4 h-4 rounded-full text-[10px] flex items-center justify-center font-bold ${
                        currentSlideIndex === idx ? 'bg-black/20 text-white' : 'bg-orange-200/80 text-orange-900'
                      }`}>
                        {idx + 1}
                      </span>
                      <span>
                        {s.type === 'cover'
                          ? '표지'
                          : s.type === 'agenda'
                          ? '목차'
                          : s.type === 'custom'
                          ? '특화사례'
                          : s.type === 'final'
                          ? '마무리'
                          : s.title.slice(0, 8) + '...'}
                      </span>
                    </button>
                  ))}
                </div>
              </div>

              {/* Step 3: PPTX Generation CTA Banner */}
              <div className="bg-gradient-to-r from-orange-900 via-amber-900 to-orange-950 text-white rounded-xl p-6 shadow-md border border-orange-800/60 flex flex-col sm:flex-row items-center justify-between gap-4">
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <Sparkles className="w-5 h-5 text-amber-400" />
                    <h3 className="font-extrabold text-base sm:text-lg">
                      16:9 파워포인트 (.pptx) 파일 생성 준비 완료
                    </h3>
                  </div>
                  <p className="text-xs text-orange-200">
                    선택한 {selectedModules.length}개 모듈과 테마({currentTheme.name})가 적용된 16:9 와이드스크린 슬라이드 {slideDeck.length}장을 생성합니다.
                  </p>
                </div>

                <button
                  onClick={handleGeneratePptx}
                  disabled={isGeneratingPptx || selectedModules.length === 0}
                  className="px-6 py-3 bg-gradient-to-r from-amber-400 to-orange-500 hover:from-amber-300 hover:to-orange-400 disabled:opacity-50 text-stone-950 font-black rounded-xl text-sm shadow-lg shadow-orange-500/20 flex items-center gap-2 transition-all active:scale-95 shrink-0 cursor-pointer"
                >
                  <Download className="w-5 h-5 text-stone-950" />
                  <span>{isGeneratingPptx ? '슬라이드 빌드 중...' : '🎯 PPTX 생성 및 다운로드'}</span>
                </button>
              </div>
            </div>
          )}

          {/* TAB 2: PYTHON SOURCE CODE (app.py) */}
          {activeTab === 'code' && (
            <div className="bg-white rounded-xl shadow-xs border border-orange-200/80 p-6 space-y-5">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-4 border-b border-orange-100">
                <div>
                  <h3 className="font-bold text-stone-900 text-base flex items-center gap-2">
                    <FileCode className="w-5 h-5 text-orange-500" />
                    <span>Streamlit 기반 "AML Slide Builder" 단일 파이썬 코드 (app.py)</span>
                  </h3>
                  <p className="text-xs text-stone-500 mt-0.5">
                    요청하신 요구사항(Streamlit + python-pptx 16:9 + PyPDF2/docx 파싱 + 다운로드 스트림)이 100% 구현된 실행 파일입니다.
                  </p>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={handleCopyCode}
                    className="px-3 py-1.5 rounded-lg border border-stone-300 hover:bg-stone-50 text-stone-700 text-xs font-bold flex items-center gap-1.5 transition-all"
                  >
                    {copiedCode ? <Check className="w-4 h-4 text-emerald-600" /> : <Copy className="w-4 h-4" />}
                    <span>{copiedCode ? '복사 완료!' : '전체 코드 복사'}</span>
                  </button>

                  <button
                    onClick={handleDownloadAppPy}
                    className="px-3.5 py-1.5 bg-orange-600 hover:bg-orange-500 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-xs transition-all"
                  >
                    <Download className="w-4 h-4" />
                    <span>app.py 다운로드</span>
                  </button>

                  <button
                    onClick={handleDownloadRequirements}
                    className="px-3 py-1.5 bg-stone-800 hover:bg-stone-700 text-stone-200 rounded-lg text-xs font-semibold flex items-center gap-1.5"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>requirements.txt</span>
                  </button>
                </div>
              </div>

              {/* Local Execution Guide Card */}
              <div className="bg-slate-900 text-slate-200 rounded-xl p-4 text-xs font-mono space-y-2">
                <div className="flex items-center justify-between text-slate-400 font-sans font-bold text-xs pb-1 border-b border-slate-800">
                  <span className="flex items-center gap-1.5 text-amber-400">
                    <Play className="w-3.5 h-3.5" /> 로컬 PC(Windows/Mac/Linux)에서 실행하는 방법
                  </span>
                  <span>Python 3.10+</span>
                </div>
                <div className="space-y-1 text-[11px] leading-relaxed">
                  <p className="text-slate-400"># 1. 의존성 라이브러리 설치</p>
                  <p className="text-emerald-400 font-bold bg-slate-950 px-2.5 py-1.5 rounded border border-slate-800">
                    pip install streamlit python-pptx pypdf python-docx
                  </p>
                  <p className="text-slate-400 mt-2"># 2. Streamlit 웹앱 실행 (브라우저 자동 실행)</p>
                  <p className="text-emerald-400 font-bold bg-slate-950 px-2.5 py-1.5 rounded border border-slate-800">
                    streamlit run app.py
                  </p>
                </div>
              </div>

              {/* Code Box */}
              <div className="relative">
                <div className="bg-slate-950 text-slate-200 rounded-xl p-4 font-mono text-xs overflow-x-auto max-h-[600px] leading-relaxed border border-slate-800 shadow-inner">
                  <pre className="text-slate-300">
                    <code>{APP_PY_CODE}</code>
                  </pre>
                </div>
              </div>
            </div>
          )}

          {/* TAB 3: AML COMPLIANCE GUIDELINES & CHEAT SHEET */}
          {activeTab === 'guide' && (
            <div className="bg-white rounded-xl shadow-xs border border-orange-200/80 p-6 space-y-6">
              <div>
                <h3 className="font-bold text-stone-900 text-base flex items-center gap-2">
                  <BookOpen className="w-5 h-5 text-orange-600" />
                  <span>AML/CFT 실무 법령 및 핵심 제도 요약 치트시트</span>
                </h3>
                <p className="text-xs text-stone-500 mt-0.5">
                  강의 교안 작성 시 슬라이드 본문에 바로 복사하여 활용할 수 있는 법적 기준입니다.
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
                {/* CDD / EDD Box */}
                <div className="p-4 rounded-xl border border-orange-200/80 bg-orange-50/30 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-stone-900 text-sm">고객확인의무 (CDD / EDD)</span>
                    <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-orange-100 text-orange-900">
                      특금법 제5조의2
                    </span>
                  </div>
                  <ul className="space-y-1.5 text-stone-600">
                    <li>• **CDD 대상**: 신규 계좌 개설, 1회 2,000만원(외화 1만불) 이상 1회성 거래</li>
                    <li>• **EDD 고위험군**: FATF 비협조국가 거주자, 고위공직자(PEP), 가상자산/환전상</li>
                    <li>• **실제소유자(BO)**: 지분 25% 이상 보유한 자연인 검증 필수</li>
                  </ul>
                  <button
                    onClick={() =>
                      setCustomText(
                        '[CDD/EDD 점검 기준]\n1. 1회성 2,000만원 이상 거래 시 실명 및 신분증 진위확인\n2. 고위험 고객(PEP)의 경우 자금원천 및 거래목적 추가 소명 징구\n3. 법인 고객의 경우 25% 이상 지분율 보유 실제소유자(BO) 검증'
                      )
                    }
                    className="mt-2 text-orange-600 hover:text-orange-800 font-semibold text-[11px]"
                  >
                    + 이 내용을 맞춤 사례 입력창에 적용
                  </button>
                </div>

                {/* STR / CTR Box */}
                <div className="p-4 rounded-xl border border-orange-200/80 bg-orange-50/30 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-stone-900 text-sm">STR / CTR 보고 의무</span>
                    <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-amber-100 text-amber-900">
                      특금법 제4조 / 제4조의2
                    </span>
                  </div>
                  <ul className="space-y-1.5 text-stone-600">
                    <li>• **STR**: 자금세탁 의심에 '합당한 근거'가 있을 때 (금액 제한 없음, 3영업일 내)</li>
                    <li>• **CTR**: 1일 동일 금융회사 현금 1,000만원 이상 입출금 (전산 자동 보고)</li>
                    <li>• **비밀누설 금지**: 고객에게 보고 사실 누설 시 3년 이하 징역 / 3천만원 벌금</li>
                  </ul>
                  <button
                    onClick={() =>
                      setCustomText(
                        '[STR/CTR 보고 실무 지침]\n1. STR은 금액에 무관하게 이상징후 인지 시 3영업일 이내 KoFIU 보고\n2. 1천만원 미만으로 쪼개기(Structuring) 분할 입출금은 즉시 STR 검토\n3. 고객에게 STR 검토 사실 누설(Tipping-off) 시 형사처벌 대상'
                      )
                    }
                    className="mt-2 text-orange-600 hover:text-orange-800 font-semibold text-[11px]"
                  >
                    + 이 내용을 맞춤 사례 입력창에 적용
                  </button>
                </div>

                {/* Travel Rule Box */}
                <div className="p-4 rounded-xl border border-orange-200/80 bg-orange-50/30 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-stone-900 text-sm">가상자산 트래블룰 (Travel Rule)</span>
                    <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-orange-200 text-orange-950">
                      FATF Rec. 16
                    </span>
                  </div>
                  <ul className="space-y-1.5 text-stone-600">
                    <li>• **적용 기준**: 100만원 상당 이상의 가상자산 외부 이전 시</li>
                    <li>• **제공 정보**: 송·수신인의 성명, 가상자산 주소, 주민등록번호 등</li>
                    <li>• **개인지갑(Un-hosted)**: 검증되지 않은 외부 개인지갑 전송 차단 원칙</li>
                  </ul>
                  <button
                    onClick={() =>
                      setCustomText(
                        '[가상자산 트래블룰 준수 지침]\n1. 100만원 이상 코인 전송 시 송수인 신원정보 실시간 교환\n2. 비인가 개인지갑 및 해외 미신고 거래소와의 직접 전송 통제\n3. 토네이도캐시 등 믹서(Mixer) 연계 지갑 자동 동결 조치'
                      )
                    }
                    className="mt-2 text-orange-600 hover:text-orange-800 font-semibold text-[11px]"
                  >
                    + 이 내용을 맞춤 사례 입력창에 적용
                  </button>
                </div>

                {/* OFAC Sanctions Box */}
                <div className="p-4 rounded-xl border border-orange-200/80 bg-orange-50/30 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-stone-900 text-sm">글로벌 금융제재 (OFAC / UN)</span>
                    <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-red-100 text-red-900">
                      외환거래법 및 OFAC
                    </span>
                  </div>
                  <ul className="space-y-1.5 text-stone-600">
                    <li>• **SDN 리스트**: 미국 OFAC 특별지정제재대상자와의 거래 전면 금지</li>
                    <li>• **50% Rule**: 제재 대상자가 50% 이상 지분을 보유한 모든 자회사 포함</li>
                    <li>• **세컨더리 보이콧**: 제재 위반 거래를 중개한 제3국 금융기관도 미국시장 퇴출</li>
                  </ul>
                  <button
                    onClick={() =>
                      setCustomText(
                        '[글로벌 금융제재 점검 수칙]\n1. 외환 송금 및 무역금융 거래 전 전산 워치리스트 스크리닝(WLS) 100% 수행\n2. OFAC 50% 룰에 따라 복수 제재 대상자의 합산 지분 50% 이상 법인 필터링\n3. Wire Stripping(제재 대상자명 삭제 송금) 방지 감사 절차 구축'
                      )
                    }
                    className="mt-2 text-orange-600 hover:text-orange-800 font-semibold text-[11px]"
                  >
                    + 이 내용을 맞춤 사례 입력창에 적용
                  </button>
                </div>
              </div>
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
