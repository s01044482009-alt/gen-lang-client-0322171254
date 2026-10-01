export interface AmlSection {
  title: string;
  bullets: string[];
}

export interface AmlModule {
  id: string;
  name: string;
  badge: string;
  subtitle: string;
  sections: AmlSection[];
}

export interface ThemeConfig {
  name: string;
  primary: string;       // Hex with '#'
  secondary: string;     // Hex with '#'
  accent: string;        // Hex with '#'
  cardBg: string;        // Hex with '#'
  cardBorder: string;    // Hex with '#'
  slideBg: string;       // Slide background color hex with '#'
  textDark: string;      // Hex with '#'
  textMuted: string;     // Hex with '#'
  pptxPrimary: string;   // Hex without '#'
  pptxSecondary: string;
  pptxAccent: string;
  pptxCardBg: string;
  pptxCardBorder: string;
  pptxSlideBg: string;   // Slide background for PPTX without '#'
  pptxTextDark: string;
  pptxTextMuted: string;
}

export const THEMES: Record<string, ThemeConfig> = {
  orange: {
    name: "Amber Orange (신뢰와 활력의 앰버 오렌지)",
    primary: "#C2410C",       // 비비드 딥 오렌지 (#C2410C)
    secondary: "#9A3412",     // 다크 번트 오렌지 (#9A3412)
    accent: "#FBBF24",        // 웜 골든 앰버 (#FBBF24)
    cardBg: "#FFFFFF",        // 카드 배경: 화이트로 대비 강화
    cardBorder: "#FED7AA",    // 부드러운 오렌지 보더 (#FED7AA)
    slideBg: "#FFF7ED",       // 슬라이드 전체 따뜻한 오렌지 배경 (Orange-50, #FFF7ED)
    textDark: "#1C1917",      // 짙은 본문 (#1C1917)
    textMuted: "#9A3412",     // 보조 오렌지 텍스트
    pptxPrimary: "C2410C",
    pptxSecondary: "9A3412",
    pptxAccent: "FBBF24",
    pptxCardBg: "FFFFFF",
    pptxCardBorder: "FED7AA",
    pptxSlideBg: "FFF7ED",    // PPTX 슬라이드 전체 오렌지 배경
    pptxTextDark: "1C1917",
    pptxTextMuted: "9A3412",
  },
  brightOrange: {
    name: "Sunset Orange (선명하고 모던한 선셋 오렌지)",
    primary: "#EA580C",       // 브라이트 오렌지 (#EA580C)
    secondary: "#C2410C",     // 딥 오렌지 (#C2410C)
    accent: "#F59E0B",        // 골드 앰버 (#F59E0B)
    cardBg: "#FFFBF7",
    cardBorder: "#FDBA74",
    slideBg: "#FFEDD5",       // 더욱 짙은 선셋 오렌지 배경 (Orange-100, #FFEDD5)
    textDark: "#18181B",
    textMuted: "#C2410C",
    pptxPrimary: "EA580C",
    pptxSecondary: "C2410C",
    pptxAccent: "F59E0B",
    pptxCardBg: "FFFBF7",
    pptxCardBorder: "FDBA74",
    pptxSlideBg: "FFEDD5",
    pptxTextDark: "18181B",
    pptxTextMuted: "C2410C",
  },
  terracotta: {
    name: "Terracotta Rust (클래식 테라코타 러스티)",
    primary: "#7C2D12",       // 딥 러스티 오렌지 (#7C2D12)
    secondary: "#9A3412",     // 미디엄 테라코타 (#9A3412)
    accent: "#FB923C",        // 파스텔 오렌지 (#FB923C)
    cardBg: "#FFFFFF",
    cardBorder: "#E7D7C9",
    slideBg: "#FAF5F0",
    textDark: "#1C1917",
    textMuted: "#7C2D12",
    pptxPrimary: "7C2D12",
    pptxSecondary: "9A3412",
    pptxAccent: "FB923C",
    pptxCardBg: "FFFFFF",
    pptxCardBorder: "E7D7C9",
    pptxSlideBg: "FAF5F0",
    pptxTextDark: "1C1917",
    pptxTextMuted: "7C2D12",
  },
  red: {
    name: "Crimson Red (프리미엄 딥 크림슨 레드)",
    primary: "#7F1D1D",
    secondary: "#991B1B",
    accent: "#F59E0B",
    cardBg: "#FFFFFF",
    cardBorder: "#FECACA",
    slideBg: "#FEF2F2",
    textDark: "#18181B",
    textMuted: "#991B1B",
    pptxPrimary: "7F1D1D",
    pptxSecondary: "991B1B",
    pptxAccent: "F59E0B",
    pptxCardBg: "FFFFFF",
    pptxCardBorder: "FECACA",
    pptxSlideBg: "FEF2F2",
    pptxTextDark: "18181B",
    pptxTextMuted: "991B1B",
  },
  navy: {
    name: "Navy (신뢰와 규준 준수)",
    primary: "#0F2537",
    secondary: "#1B365D",
    accent: "#D97706",
    cardBg: "#FFFFFF",
    cardBorder: "#CBD5E1",
    slideBg: "#F1F5F9",
    textDark: "#1E293B",
    textMuted: "#64748B",
    pptxPrimary: "0F2537",
    pptxSecondary: "1B365D",
    pptxAccent: "D97706",
    pptxCardBg: "FFFFFF",
    pptxCardBorder: "CBD5E1",
    pptxSlideBg: "F1F5F9",
    pptxTextDark: "1E293B",
    pptxTextMuted: "64748B",
  },
  blue: {
    name: "Blue (모던 금융 & 테크)",
    primary: "#1E40AF",
    secondary: "#2563EB",
    accent: "#0EA5E9",
    cardBg: "#FFFFFF",
    cardBorder: "#BFDBFE",
    slideBg: "#EFF6FF",
    textDark: "#0F172A",
    textMuted: "#475569",
    pptxPrimary: "1E40AF",
    pptxSecondary: "2563EB",
    pptxAccent: "0EA5E9",
    pptxCardBg: "FFFFFF",
    pptxCardBorder: "BFDBFE",
    pptxSlideBg: "EFF6FF",
    pptxTextDark: "0F172A",
    pptxTextMuted: "475569",
  }
};

export const AML_MODULES: AmlModule[] = [
  {
    id: "aml_basics",
    name: "AML/CFT 기본 개념 및 자금세탁 3단계 구조",
    badge: "Module 1 • 기초 이론",
    subtitle: "자금세탁의 정의, 법적 근거(특금법) 및 자금세탁의 3단계 메커니즘",
    sections: [
      {
        title: "📌 자금세탁(Money Laundering)의 정의",
        bullets: [
          "불법 수익의 취득·처분 사실을 가장하거나 은닉하여 정당한 재산으로 위장하는 일련의 행위",
          "국내 근거법률: 「특정 금융거래정보의 보고 및 이용 등에 관한 법률(특금법)」 및 범죄수익은닉규제법",
          "금융기관의 사회적 신뢰 유지 및 글로벌 금융제재 리스크 방지를 위한 필수 법적 의무"
        ]
      },
      {
        title: "🔄 자금세탁의 전형적인 3단계 프로세스",
        bullets: [
          "1단계 [배치(Placement)]: 불법 취득 현금을 소액 분할하여 금융시스템 내부로 최초 진입",
          "2단계 [은닉(Layering)]: 수많은 복잡한 이체, 환전, 해외 송금, 차명계좌 거래를 통해 자금 출처 단절",
          "3단계 [통합(Integration)]: 부동산, 합법 사업체 투자, 가상자산 인수 등을 통해 완전한 합법 자금으로 둔갑"
        ]
      },
      {
        title: "⚖️ 금융기관 임직원의 실무 책임",
        bullets: [
          "의심스러운 거래 인지 시 보고 지연 및 묵인은 형사처벌 및 금융당국 행정제재 대상",
          "임직원의 AML 인지도 결여는 기관 전체의 영업정지 및 거액 과태료 부과로 직결",
          "정기 교육 이수 및 업무 지침서 준수는 컴플라이언스 방어권의 핵심"
        ]
      }
    ]
  },
  {
    id: "cdd_edd",
    name: "고객확인의무 (CDD / EDD) 및 실제소유자(BO) 확인",
    badge: "Module 2 • 고객실사",
    subtitle: "KYC(Know Your Customer)의 체계적 이행 및 고위험 고객 심층평가 기준",
    sections: [
      {
        title: "🔍 기본 고객확인 (CDD, Customer Due Diligence)",
        bullets: [
          "계좌 개설, 1회성 2천만원(외화 1만불) 이상 거래, 자금세탁 우려 시 신원 확인 필수",
          "개인: 실명, 주민등록번호, 주소, 연락처, 직업 검증 (신분증 진위확인 시스템 연동)",
          "법인: 법인명, 사업자등록번호, 본점 소재지, 대표자 및 업종 확인 (사업자등록증 원본 등)"
        ]
      },
      {
        title: "🛡️ 강화된 고객확인 (EDD, Enhanced Due Diligence)",
        bullets: [
          "적용 대상: 고위험 국가 거주자, 정치적 주요인물(PEP), 카지노/환전업/가상자산사업자 등",
          "추가 확인 사항: 거래 목적, 거래자금의 구체적 원천(Source of Funds) 및 재산 현황",
          "승인 절차: 고위험 거래 체결 전 준법감시부서 또는 관할 지점장 서면 승인 필수"
        ]
      },
      {
        title: "👥 실제소유자 (Beneficial Owner) 확인 규정",
        bullets: [
          "1단계: 법인 지분 25% 이상 소유한 사실상 지배 주주 확인",
          "2단계: 지분율 기준 파악 불가 시 최대주주 또는 대표자 사실상 지배 여부 검증",
          "실소유자 확인 거부 또는 위조 의심 시 거래 거절(Decline) 및 STR 보고 검토"
        ]
      }
    ]
  },
  {
    id: "str_ctr",
    name: "의심거래보고(STR) 및 고액현금거래보고(CTR) 실무",
    badge: "Module 3 • 보고 제도",
    subtitle: "금융정보분석원(KoFIU) 보고 기준, 보고 기한 및 거래 누설 금지 원칙",
    sections: [
      {
        title: "🚨 의심거래보고 (STR, Suspicious Transaction Report)",
        bullets: [
          "보고 기준: 금융거래와 관련하여 불법자금 또는 자금세탁 의심에 대한 '합당한 근거'가 있는 경우",
          "주관적 판단 존중: 금액 제한 없음(1원이라도 의심되면 보고 가능), 영업점 실무자의 직관과 정황 고려",
          "보고 기한: 의심 정황 보고서 작성 완료 후 3영업일 이내 KoFIU에 전산 보고"
        ]
      },
      {
        title: "💵 고액현금거래보고 (CTR, Currency Transaction Report)",
        bullets: [
          "보고 기준: 1일 동일 금융회사에서 현금 1천만원 이상을 입금하거나 출금하는 거래",
          "자동화 보고: 전산 시스템에 의해 30영업일 이내 KoFIU로 자동 전송",
          "분할 거래(Structuring) 주의: 1천만원 회피 목적으로 950만원 분할 거래 시 STR 보고 대상"
        ]
      },
      {
        title: "🔒 비밀누설 금지 원칙 (Tipping-Off 금지)",
        bullets: [
          "임직원은 고객에게 STR 보고 사실이나 검토 중인 정황을 일절 누설할 수 없음",
          "특금법 위반 시 3년 이하의 징역 또는 3천만원 이하의 벌금형 (중대 형사처벌)",
          "고객이 의심스러운 질의를 하더라도 평상시 매뉴얼대로 자연스럽게 응대"
        ]
      }
    ]
  },
  {
    id: "internal_audit",
    name: "내부통제체계 구축 및 독립적 감사(Audit) 운영",
    badge: "Module 4 • 내부통제",
    subtitle: "3선 방어체계(3 Lines of Defense)와 최고경영진·이사회 법적 책무",
    sections: [
      {
        title: "🏛️ AML 3선 방어체계 (3 Lines of Defense)",
        bullets: [
          "제1선 (영업점/비즈니스): 고객 대면 창구에서 직접 CDD 수행, 이상징후 1차 필터링",
          "제2선 (준법감시부서/AML팀): 위험평가(RBA) 모델링, 정책 제정, STR 2차 심사 및 KoFIU 전송",
          "제3선 (내부감사실/외부감사인): AML 운영체계의 실효성 및 준수 여부에 대한 독립적 감사 수행"
        ]
      },
      {
        title: "📋 위험기반 접근법 (RBA, Risk-Based Approach)",
        bullets: [
          "고객, 상품, 서비스, 유통채널, 지리적 위치의 위험도를 종합 평가하여 위험도별 차등 통제",
          "고위험군 자원에 전사적 감시 역량을 집중 배분하여 통제 효율 극대화",
          "정기적인 전사 위험평가(Enterprise-Wide Risk Assessment) 의무 수행"
        ]
      },
      {
        title: "👔 경영진의 책무구조 및 징계 기준",
        bullets: [
          "이사회 및 CEO의 AML 최종 감독 책임 명문화 (금융판 책무구조도 반영)",
          "통제 부실로 인한 대규모 자금세탁 방조 시 대표이사 및 감사 직무정지 등 강력 제재",
          "전 임직원 대상 연 1회 이상 정기 AML 직무교육 의무 수료 관리"
        ]
      }
    ]
  },
  {
    id: "virtual_assets",
    name: "가상자산(VASP) 트렌드 및 FATF 트래블룰(Travel Rule)",
    badge: "Module 5 • 신종 리스크",
    subtitle: "블록체인 탈중앙화 금융(DeFi), 믹서(Mixer) 기술 악용과 글로벌 규제 동향",
    sections: [
      {
        title: "🌐 가상자산사업자(VASP)의 특금법상 규제 의무",
        bullets: [
          "FIU 정식 신고 수리 및 ISMS 인증 취득 사업자만 합법 영업 가능",
          "시중은행의 실명확인 입출금계정(실명계좌) 연동 의무화 및 위험평가 통과 필수",
          "코인간 거래, 스테이블코인 환전, OTC 장외거래에 대한 모니터링 강화"
        ]
      },
      {
        title: "✈️ FATF 트래블룰 (Travel Rule) 핵심 요건",
        bullets: [
          "100만원 상당 이상의 가상자산 전송 시 송·수신인의 신원 정보(성명, 지갑주소 등) 제공 의무",
          "인가받지 않은 개인지갑(Un-hosted Wallet)으로의 대규모 전송 엄격 제한",
          "국내 주요 거래소 간 화이트리스트 검증 솔루션(CODE, VerifyVASP 등) 연계 운용"
        ]
      },
      {
        title: "⚠️ 신종 자금세탁 기법 및 대응",
        bullets: [
          "토네이도캐시 등 가상자산 믹서(Mixer) 및 텀블러를 통한 거래 추적 방해 행위 차단",
          "디파이(DeFi), NFT 마켓플레이스, P2P 거래를 통한 불법자금 차명 분산 주의",
          "온체인 분석 툴(Chainalysis, Elliptic 등)을 활용한 블랙리스트 지갑 실시간 차단"
        ]
      }
    ]
  },
  {
    id: "sanctions",
    name: "글로벌 금융제재(Sanctions) 및 주요 위반 사례 분석",
    badge: "Module 6 • 제재 사례",
    subtitle: "OFAC, UN 제재 리스트 스크리닝 및 글로벌 금융사 수조원대 벌금 교훈",
    sections: [
      {
        title: "🌍 주요 글로벌 제재 기구 및 감시 목록",
        bullets: [
          "미국 재무부 해외자산통제국(OFAC): SDN(특별지정제재대상자) 리스트 운영 (2차 제재/세컨더리 보이콧)",
          "UN 안보리 제재, EU 금융제재 및 한국 기획재정부/금융위 금융거래제한대상자 공고",
          "모든 외환 송금 및 무역금융 거래 시 실시간 워치리스트 스크리닝(WLS) 필수"
        ]
      },
      {
        title: "💥 글로벌 금융사 주요 제재 위반 사례",
        bullets: [
          "글로벌 B사 사례: 북한·이란 관련 제재 회피 거래 방조로 1조원 이상의 천문학적 벌금 부과",
          "국내 은행 해외지점 사례: AML 모니터링 인력 및 시스템 미비로 美 뉴욕주 금융청(DFS) 수천만 달러 과징금",
          "수법: 송금 전문에서 제재 대상자 이름 고의 삭제(Wire Stripping) 등 악의적 우회 거래 적발"
        ]
      },
      {
        title: "🎯 실무자의 리스크 예방 행동수칙",
        bullets: [
          "송금 사유와 실 수령인이 모호한 중동/제3국 환적 무역거래 철저 재확인",
          "거래 상대방의 실소유자가 제재 대상자 지분 50% 이상 보유 시 제재 대상과 동일 취급 (50% Rule)",
          "조금이라도 의심되는 제재 대상 거래는 실행 전 본점 준법감시팀 공식 사전 질의 필수"
        ]
      }
    ]
  }
];

export const SAMPLE_CASE_PRESETS = [
  {
    title: "2024년 KoFIU 검사 중점사항 (내부통제 및 STR 적시성)",
    text: `[2024년도 금융정보분석원(KoFIU) 검사 중점 방향]
1. STR 보고의 적시성(3영업일) 준수 여부 및 장기 미보고 건에 대한 제재 강화
2. 고위험 고객(PEP, 카지노/환전업 등)에 대한 EDD 주기적 갱신 이행 실태
3. 가상자산 연계 실명확인 입출금계좌에 대한 이상거래탐지 룰셋 고도화 여부
4. 경영진 및 이사회의 AML 내부통제 이행 점검 결과 보고 체계 준수 여부`
  },
  {
    title: "무역기반 자금세탁(TBML) 적발 사례 및 주의사항",
    text: `[무역기반 자금세탁(TBML) 주요 적발 수법]
1. 허위 수출입 송장(Invoice) 발행을 통한 물품 가격 조작(Over/Under Invoicing)
2. 실물 이동이 없는 유령 환적(Ghost Shipment)을 통한 해외 불법 자금 유출
3. 동일 선하증권(B/L) 번호의 다중 금융기관 중복 대출 신청
4. 대응 수칙: 관세청 환적 데이터 대조 및 고액 무역 대금 분할 송금 시 선적서류 심층 검증`
  },
  {
    title: "글로벌 제재 위반(OFAC) 50% 룰 적용 주의 사례",
    text: `[미 재무부 OFAC 50% Rule 실무 적용 지침]
1. 단일 SDN 대상자가 직접 50% 이상 지분을 보유한 법인은 물론,
2. 복수의 SDN 대상자가 합산하여 50% 이상 지분을 보유한 해외 합작법인도 제재 대상
3. 차명 주주 위장 여부 파악을 위해 최종 실소유자(BO) 확인서 징구 및 신용평가기관 온체인/오프체인 실사 필수`
  }
];
