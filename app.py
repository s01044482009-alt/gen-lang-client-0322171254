"""
AML 강의 프레젠테이션 자동 생성 웹앱 (AML Slide Builder)
--------------------------------------------------
자금세탁방지(AML/CFT) 강의 담당자를 위한 원클릭 맞춤형 PPTX 자동 생성 Streamlit 애플리케이션
- Tech Stack: Python 3.10+, Streamlit, python-pptx, PyPDF2 / pypdf, python-docx
- Presentation Ratio: 16:9 Widescreen (13.333" x 7.5" = 33.867cm x 19.05cm)
- Features: 주황색 계열 앰버 오렌지 테마 기본 적용, 16:9 정밀 규격, 메모리 스트림 기반 PPTX 다운로드
"""

import io
import os
from datetime import datetime
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# PDF 파서 라이브러리 (pypdf 또는 PyPDF2 호환성 처리)
try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        PdfReader = None

# DOCX 파서 라이브러리
try:
    import docx
except ImportError:
    docx = None

# ==============================================================================
# 1. 테마 색상 정의 (Color Palette) - 주황색 계열 기본
# ==============================================================================
THEME_PALETTES = {
    "Amber Orange (신뢰와 활력의 앰버 오렌지)": {
        "primary": RGBColor(194, 65, 12),      # 딥 번트 오렌지 (#C2410C)
        "secondary": RGBColor(154, 52, 18),    # 다크 테라코타 오렌지 (#9A3412)
        "accent": RGBColor(251, 191, 36),      # 웜 골든 앰버 (#FBBF24)
        "card_bg": RGBColor(255, 255, 255),    # 카드: 화이트 (대비 확보)
        "card_border": RGBColor(254, 215, 170),# 부드러운 오렌지 보더 (#FED7AA)
        "slide_bg": RGBColor(255, 247, 237),   # 슬라이드 전체 따뜻한 오렌지 배경 (#FFF7ED)
        "text_dark": RGBColor(28, 25, 23),     # 짙은 본문 (#1C1917)
        "text_muted": RGBColor(154, 52, 18),   # 보조 오렌지 (#9A3412)
        "white": RGBColor(255, 255, 255),
        "hex_primary": "#C2410C",
        "hex_secondary": "#9A3412",
    },
    "Sunset Orange (선명하고 모던한 선셋 오렌지)": {
        "primary": RGBColor(234, 88, 12),      # 브라이트 오렌지 (#EA580C)
        "secondary": RGBColor(194, 65, 12),    # 딥 오렌지 (#C2410C)
        "accent": RGBColor(245, 158, 11),      # 밝은 골드 앰버 (#F59E0B)
        "card_bg": RGBColor(255, 255, 255),
        "card_border": RGBColor(253, 186, 116),
        "slide_bg": RGBColor(255, 237, 213),   # 선셋 오렌지 배경 (#FFEDD5)
        "text_dark": RGBColor(24, 24, 27),
        "text_muted": RGBColor(194, 65, 12),
        "white": RGBColor(255, 255, 255),
        "hex_primary": "#EA580C",
        "hex_secondary": "#C2410C",
    },
    "Terracotta Rust (클래식 테라코타 러스티)": {
        "primary": RGBColor(124, 45, 18),      # 딥 러스티 오렌지 (#7C2D12)
        "secondary": RGBColor(154, 52, 18),    # 미디엄 테라코타 (#9A3412)
        "accent": RGBColor(251, 146, 60),      # 파스텔 오렌지 (#FB923C)
        "card_bg": RGBColor(255, 255, 255),
        "card_border": RGBColor(231, 215, 201),
        "slide_bg": RGBColor(250, 245, 240),
        "text_dark": RGBColor(28, 25, 23),
        "text_muted": RGBColor(124, 45, 18),
        "white": RGBColor(255, 255, 255),
        "hex_primary": "#7C2D12",
        "hex_secondary": "#9A3412",
    },
    "Crimson Red (프리미엄 딥 크림슨 레드)": {
        "primary": RGBColor(127, 29, 29),
        "secondary": RGBColor(153, 27, 27),
        "accent": RGBColor(245, 158, 11),
        "card_bg": RGBColor(255, 255, 255),
        "card_border": RGBColor(254, 202, 202),
        "slide_bg": RGBColor(254, 242, 242),
        "text_dark": RGBColor(24, 24, 27),
        "text_muted": RGBColor(153, 27, 27),
        "white": RGBColor(255, 255, 255),
        "hex_primary": "#7F1D1D",
        "hex_secondary": "#991B1B",
    },
    "Navy (신뢰와 규준 준수)": {
        "primary": RGBColor(15, 37, 55),
        "secondary": RGBColor(27, 54, 93),
        "accent": RGBColor(217, 119, 6),
        "card_bg": RGBColor(255, 255, 255),
        "card_border": RGBColor(203, 213, 225),
        "slide_bg": RGBColor(241, 245, 249),
        "text_dark": RGBColor(30, 41, 59),
        "text_muted": RGBColor(100, 116, 139),
        "white": RGBColor(255, 255, 255),
        "hex_primary": "#0F2537",
        "hex_secondary": "#1B365D",
    },
    "Blue (모던 금융 & 테크)": {
        "primary": RGBColor(30, 64, 175),
        "secondary": RGBColor(37, 99, 235),
        "accent": RGBColor(14, 165, 233),
        "card_bg": RGBColor(255, 255, 255),
        "card_border": RGBColor(191, 219, 254),
        "slide_bg": RGBColor(239, 246, 255),
        "text_dark": RGBColor(15, 23, 42),
        "text_muted": RGBColor(71, 85, 105),
        "white": RGBColor(255, 255, 255),
        "hex_primary": "#1E40AF",
        "hex_secondary": "#2563EB",
    }
}

# ==============================================================================
# 2. AML 표준 강의 모듈 사전 (Curated AML Curriculum)
# ==============================================================================
AML_MODULE_DATABASE = {
    "aml_basics": {
        "id": "aml_basics",
        "name": "AML/CFT 기본 개념 및 자금세탁 3단계 구조",
        "badge": "Module 1 • 기초 이론",
        "subtitle": "자금세탁의 정의, 법적 근거(특금법) 및 자금세탁의 3단계 메커니즘",
        "sections": [
            {
                "title": "📌 자금세탁(Money Laundering)의 정의",
                "bullets": [
                    "불법 수익의 취득·처분 사실을 가장하거나 은닉하여 정당한 재산으로 위장하는 일련의 행위",
                    "국내 근거법률: 「특정 금융거래정보의 보고 및 이용 등에 관한 법률(특금법)」 및 범죄수익은닉규제법",
                    "금융기관의 사회적 신뢰 유지 및 글로벌 금융제재 리스크 방지를 위한 필수 법적 의무"
                ]
            },
            {
                "title": "🔄 자금세탁의 전형적인 3단계 프로세스",
                "bullets": [
                    "1단계 [배치(Placement)]: 불법 취득 현금을 소액 분할하여 금융시스템 내부로 최초 진입",
                    "2단계 [은닉(Layering)]: 수많은 복잡한 이체, 환전, 해외 송금, 차명계좌 거래를 통해 자금 출처 단절",
                    "3단계 [통합(Integration)]: 부동산, 합법 사업체 투자, 가상자산 인수 등을 통해 완전한 합법 자금으로 둔갑"
                ]
            },
            {
                "title": "⚖️ 금융기관 임직원의 실무 책임",
                "bullets": [
                    "의심스러운 거래 인지 시 보고 지연 및 묵인은 형사처벌 및 금융당국 행정제재 대상",
                    "임직원의 AML 인지도 결여는 기관 전체의 영업정지 및 거액 과태료 부과로 직결",
                    "정기 교육 이수 및 업무 지침서 준수는 컴플라이언스 방어권의 핵심"
                ]
            }
        ]
    },
    "cdd_edd": {
        "id": "cdd_edd",
        "name": "고객확인의무 (CDD / EDD) 및 실제소유자(BO) 확인",
        "badge": "Module 2 • 고객실사",
        "subtitle": "KYC(Know Your Customer)의 체계적 이행 및 고위험 고객 심층평가 기준",
        "sections": [
            {
                "title": "🔍 기본 고객확인 (CDD, Customer Due Diligence)",
                "bullets": [
                    "계좌 개설, 1회성 2천만원(외화 1만불) 이상 거래, 자금세탁 우려 시 신원 확인 필수",
                    "개인: 실명, 주민등록번호, 주소, 연락처, 직업 검증 (신분증 진위확인 시스템 연동)",
                    "법인: 법인명, 사업자등록번호, 본점 소재지, 대표자 및 업종 확인 (사업자등록증 원본 등)"
                ]
            },
            {
                "title": "🛡️ 강화된 고객확인 (EDD, Enhanced Due Diligence)",
                "bullets": [
                    "적용 대상: 고위험 국가 거주자, 정치적 주요인물(PEP), 카지노/환전업/가상자산사업자 등",
                    "추가 확인 사항: 거래 목적, 거래자금의 구체적 원천(Source of Funds) 및 재산 현황",
                    "승인 절차: 고위험 거래 체결 전 준법감시부서 또는 관할 지점장 서면 승인 필수"
                ]
            },
            {
                "title": "👥 실제소유자 (Beneficial Owner) 확인 규정",
                "bullets": [
                    "1단계: 법인 지분 25% 이상 소유한 사실상 지배 주주 확인",
                    "2단계: 지분율 기준 파악 불가 시 최대주주 또는 대표자 사실상 지배 여부 검증",
                    "실소유자 확인 거부 또는 위조 의심 시 거래 거절(Decline) 및 STR 보고 검토"
                ]
            }
        ]
    },
    "str_ctr": {
        "id": "str_ctr",
        "name": "의심거래보고(STR) 및 고액현금거래보고(CTR) 실무",
        "badge": "Module 3 • 보고 제도",
        "subtitle": "금융정보분석원(KoFIU) 보고 기준, 보고 기한 및 거래 누설 금지 원칙",
        "sections": [
            {
                "title": "🚨 의심거래보고 (STR, Suspicious Transaction Report)",
                "bullets": [
                    "보고 기준: 금융거래와 관련하여 불법자금 또는 자금세탁 의심에 대한 '합당한 근거'가 있는 경우",
                    "주관적 판단 존중: 금액 제한 없음(1원이라도 의심되면 보고 가능), 영업점 실무자의 직관과 정황 고려",
                    "보고 기한: 의심 정황 보고서 작성 완료 후 3영업일 이내 KoFIU에 전산 보고"
                ]
            },
            {
                "title": "💵 고액현금거래보고 (CTR, Currency Transaction Report)",
                "bullets": [
                    "보고 기준: 1일 동일 금융회사에서 현금 1천만원 이상을 입금하거나 출금하는 거래",
                    "자동화 보고: 전산 시스템에 의해 30영업일 이내 KoFIU로 자동 전송",
                    "분할 거래(Structuring) 주의: 1천만원 회피 목적으로 950만원 분할 거래 시 STR 보고 대상"
                ]
            },
            {
                "title": "🔒 비밀누설 금지 원칙 (Tipping-Off 금지)",
                "bullets": [
                    "임직원은 고객에게 STR 보고 사실이나 검토 중인 정황을 일절 누설할 수 없음",
                    "특금법 위반 시 3년 이하의 징역 또는 3천만원 이하의 벌금형 (중대 형사처벌)",
                    "고객이 의심스러운 질의를 하더라도 평상시 매뉴얼대로 자연스럽게 응대"
                ]
            }
        ]
    },
    "internal_audit": {
        "id": "internal_audit",
        "name": "내부통제체계 구축 및 독립적 감사(Audit) 운영",
        "badge": "Module 4 • 내부통제",
        "subtitle": "3선 방어체계(3 Lines of Defense)와 최고경영진·이사회 법적 책무",
        "sections": [
            {
                "title": "🏛️ AML 3선 방어체계 (3 Lines of Defense)",
                "bullets": [
                    "제1선 (영업점/비즈니스): 고객 대면 창구에서 직접 CDD 수행, 이상징후 1차 필터링",
                    "제2선 (준법감시부서/AML팀): 위험평가(RBA) 모델링, 정책 제정, STR 2차 심사 및 KoFIU 전송",
                    "제3선 (내부감사실/외부감사인): AML 운영체계의 실효성 및 준수 여부에 대한 독립적 감사 수행"
                ]
            },
            {
                "title": "📋 위험기반 접근법 (RBA, Risk-Based Approach)",
                "bullets": [
                    "고객, 상품, 서비스, 유통채널, 지리적 위치의 위험도를 종합 평가하여 위험도별 차등 통제",
                    "고위험군 자원에 전사적 감시 역량을 집중 배분하여 통제 효율 극대화",
                    "정기적인 전사 위험평가(Enterprise-Wide Risk Assessment) 의무 수행"
                ]
            },
            {
                "title": "👔 경영진의 책무구조 및 징계 기준",
                "bullets": [
                    "이사회 및 CEO의 AML 최종 감독 책임 명문화 (금융판 책무구조도 반영)",
                    "통제 부실로 인한 대규모 자금세탁 방조 시 대표이사 및 감사 직무정지 등 강력 제재",
                    "전 임직원 대상 연 1회 이상 정기 AML 직무교육 의무 수료 관리"
                ]
            }
        ]
    },
    "virtual_assets": {
        "id": "virtual_assets",
        "name": "가상자산(VASP) 트렌드 및 FATF 트래블룰(Travel Rule)",
        "badge": "Module 5 • 신종 리스크",
        "subtitle": "블록체인 탈중앙화 금융(DeFi), 믹서(Mixer) 기술 악용과 글로벌 규제 동향",
        "sections": [
            {
                "title": "🌐 가상자산사업자(VASP)의 특금법상 규제 의무",
                "bullets": [
                    "FIU 정식 신고 수리 및 ISMS 인증 취득 사업자만 합법 영업 가능",
                    "시중은행의 실명확인 입출금계정(실명계좌) 연동 의무화 및 위험평가 통과 필수",
                    "코인간 거래, 스테이블코인 환전, OTC 장외거래에 대한 모니터링 강화"
                ]
            },
            {
                "title": "✈️ FATF 트래블룰 (Travel Rule) 핵심 요건",
                "bullets": [
                    "100만원 상당 이상의 가상자산 전송 시 송·수신인의 신원 정보(성명, 지갑주소 등) 제공 의무",
                    "인가받지 않은 개인지갑(Un-hosted Wallet)으로의 대규모 전송 엄격 제한",
                    "국내 주요 거래소 간 화이트리스트 검증 솔루션(CODE, VerifyVASP 등) 연계 운용"
                ]
            },
            {
                "title": "⚠️ 신종 자금세탁 기법 및 대응",
                "bullets": [
                    "토네이도캐시 등 가상자산 믹서(Mixer) 및 텀블러를 통한 거래 추적 방해 행위 차단",
                    "디파이(DeFi), NFT 마켓플레이스, P2P 거래를 통한 불법자금 차명 분산 주의",
                    "온체인 분석 툴(Chainalysis, Elliptic 등)을 활용한 블랙리스트 지갑 실시간 차단"
                ]
            }
        ]
    },
    "sanctions": {
        "id": "sanctions",
        "name": "글로벌 금융제재(Sanctions) 및 주요 위반 사례 분석",
        "badge": "Module 6 • 제재 사례",
        "subtitle": "OFAC, UN 제재 리스트 스크리닝 및 글로벌 금융사 수조원대 벌금 교훈",
        "sections": [
            {
                "title": "🌍 주요 글로벌 제재 기구 및 감시 목록",
                "bullets": [
                    "미국 재무부 해외자산통제국(OFAC): SDN(특별지정제재대상자) 리스트 운영 (2차 제재/세컨더리 보이콧)",
                    "UN 안보리 제재, EU 금융제재 및 한국 기획재정부/금융위 금융거래제한대상자 공고",
                    "모든 외환 송금 및 무역금융 거래 시 실시간 워치리스트 스크리닝(WLS) 필수"
                ]
            },
            {
                "title": "💥 글로벌 금융사 주요 제재 위반 사례",
                "bullets": [
                    "글로벌 B사 사례: 북한·이란 관련 제재 회피 거래 방조로 1조원 이상의 천문학적 벌금 부과",
                    "국내 은행 해외지점 사례: AML 모니터링 인력 및 시스템 미비로 美 뉴욕주 금융청(DFS) 수천만 달러 과징금",
                    "수법: 송금 전문에서 제재 대상자 이름 고의 삭제(Wire Stripping) 등 악의적 우회 거래 적발"
                ]
            },
            {
                "title": "🎯 실무자의 리스크 예방 행동수칙",
                "bullets": [
                    "송금 사유와 실 수령인이 모호한 중동/제3국 환적 무역거래 철저 재확인",
                    "거래 상대방의 실소유자가 제재 대상자 지분 50% 이상 보유 시 제재 대상과 동일 취급 (50% Rule)",
                    "조금이라도 의심되는 제재 대상 거래는 실행 전 본점 준법감시팀 공식 사전 질의 필수"
                ]
            }
        ]
    }
}

# ==============================================================================
# 3. 문서 파싱 헬퍼 함수
# ==============================================================================
def parse_uploaded_document(uploaded_file) -> str:
    if uploaded_file is None:
        return ""
    
    file_name = uploaded_file.name.lower()
    extracted_text = ""
    
    try:
        if file_name.endswith(".txt"):
            raw_data = uploaded_file.read()
            for encoding in ["utf-8", "cp949", "euc-kr", "latin1"]:
                try:
                    extracted_text = raw_data.decode(encoding)
                    break
                except UnicodeDecodeError:
                    continue
            if not extracted_text:
                extracted_text = raw_data.decode("utf-8", errors="ignore")
                
        elif file_name.endswith(".pdf"):
            if PdfReader is None:
                raise ImportError("PDF 파싱 라이브러리(pypdf 또는 PyPDF2)가 설치되어 있지 않습니다.")
            
            reader = PdfReader(uploaded_file)
            page_texts = []
            for i, page in enumerate(reader.pages):
                txt = page.extract_text() or ""
                if txt.strip():
                    page_texts.append(f"--- [Page {i+1}] ---\n{txt.strip()}")
            extracted_text = "\n\n".join(page_texts)
            
        elif file_name.endswith(".docx"):
            if docx is None:
                raise ImportError("DOCX 파싱 라이브러리(python-docx)가 설치되어 있지 않습니다.")
            
            doc = docx.Document(uploaded_file)
            doc_paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            for table in doc.tables:
                for row in table.rows:
                    row_data = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                    if row_data:
                        doc_paragraphs.append(" | ".join(row_data))
            extracted_text = "\n".join(doc_paragraphs)
            
        else:
            raise ValueError(f"지원하지 않는 파일 형식입니다: {uploaded_file.name}")

    except Exception as e:
        st.error(f"⚠️ 파일 '{uploaded_file.name}' 파싱 중 오류가 발생했습니다: {str(e)}")
        return ""

    return extracted_text.strip()


def chunk_custom_text(text: str, max_chars: int = 500) -> list[str]:
    if not text:
        return []
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    cleaned_items = []
    for line in lines:
        if len(line) > 120:
            cleaned_items.append(line[:115] + "...")
        else:
            cleaned_items.append(line)
        if len(cleaned_items) >= 6:
            break
    return cleaned_items


# ==============================================================================
# 4. python-pptx 슬라이드 생성 엔진 (주황색 계열 풀 배경 + 16:9 와이드스크린)
# ==============================================================================
def create_aml_presentation(
    lecture_title: str,
    target_audience: str,
    duration_min: int,
    instructor_name: str,
    theme_key: str,
    selected_module_keys: list[str],
    custom_content_text: str
) -> io.BytesIO:
    prs = Presentation()
    
    # 16:9 와이드스크린 표준 비율 (13.333333 x 7.5 inches)
    prs.slide_width = Inches(13.333333)
    prs.slide_height = Inches(7.5)
    
    theme = THEME_PALETTES.get(theme_key, THEME_PALETTES["Amber Orange (신뢰와 활력의 앰버 오렌지)"])
    blank_layout = prs.slide_layouts[6]
    today_str = datetime.now().strftime("%Y년 %m월 %d일")
    
    def add_slide_footer(slide, current_idx: int, total_slides: int):
        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.35))
        tf = footer_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = f"AML/CFT Compliance Education  |  {lecture_title}  |  Slide {current_idx} of {total_slides}"
        p.font.size = Pt(10)
        p.font.color.rgb = theme["text_muted"]
        p.font.name = "Malgun Gothic"

    total_slide_count = 2 + len(selected_module_keys) + (1 if custom_content_text.strip() else 0) + 1
    current_slide_num = 1

    # ==========================================================================
    # SLIDE 1: 표지 슬라이드 (전체 주황색 계열 풀 블리드 배경)
    # ==========================================================================
    cover_slide = prs.slides.add_slide(blank_layout)
    
    # 13.333" x 7.5" 풀 블리드 오렌지 배경
    bg_fill = cover_slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333333), Inches(7.5)
    )
    bg_fill.fill.solid()
    bg_fill.fill.fore_color.rgb = theme["primary"]
    bg_fill.line.fill.background()

    # 상단 골드 앰버 스트립
    top_band = cover_slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333333), Inches(0.35)
    )
    top_band.fill.solid()
    top_band.fill.fore_color.rgb = theme["accent"]
    top_band.line.fill.background()

    # 표지 텍스트 프레임 (좌우 1.4" 안전 여백)
    cover_tx = cover_slide.shapes.add_textbox(Inches(1.4), Inches(1.3), Inches(10.5), Inches(3.2))
    ctf = cover_tx.text_frame
    ctf.word_wrap = True

    p_tag = ctf.paragraphs[0]
    p_tag.text = "🛡️ COMPLIANCE & FINANCIAL CRIME PREVENTION"
    p_tag.font.size = Pt(13)
    p_tag.font.bold = True
    p_tag.font.color.rgb = theme["accent"]
    p_tag.font.name = "Malgun Gothic"
    p_tag.space_after = Pt(14)

    p_title = ctf.add_paragraph()
    p_title.text = lecture_title
    p_title.font.size = Pt(34)
    p_title.font.bold = True
    p_title.font.color.rgb = theme["white"]
    p_title.font.name = "Malgun Gothic"
    p_title.space_after = Pt(18)

    p_desc = ctf.add_paragraph()
    p_desc.text = "자금세탁방지(AML) 및 공중협박자금조달금지(CFT) 제도 실무 역량 강화를 위한 핵심 교육"
    p_desc.font.size = Pt(16)
    p_desc.font.color.rgb = RGBColor(255, 237, 213)  # 부드러운 피치 틴트
    p_desc.font.name = "Malgun Gothic"

    # 하단 메타정보 박스 (Secondary 다크 테라코타 오렌지)
    info_card = cover_slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.4), Inches(4.6), Inches(10.5), Inches(1.4)
    )
    info_card.fill.solid()
    info_card.fill.fore_color.rgb = theme["secondary"]
    info_card.line.fill.background()

    info_tx = cover_slide.shapes.add_textbox(Inches(1.7), Inches(4.75), Inches(9.9), Inches(1.1))
    itf = info_tx.text_frame
    itf.word_wrap = True
    
    p_info1 = itf.paragraphs[0]
    p_info1.text = f"• 수강 대상:  {target_audience}          • 강의 소요시간:  {duration_min}분"
    p_info1.font.size = Pt(14)
    p_info1.font.bold = True
    p_info1.font.color.rgb = theme["white"]
    p_info1.font.name = "Malgun Gothic"
    p_info1.space_after = Pt(8)

    p_info2 = itf.add_paragraph()
    p_info2.text = f"• 전담 강사:  {instructor_name}          • 작성 및 시행일자:  {today_str}"
    p_info2.font.size = Pt(13)
    p_info2.font.color.rgb = RGBColor(254, 215, 170)
    p_info2.font.name = "Malgun Gothic"

    current_slide_num += 1

    # ==========================================================================
    # SLIDE 2: 목차 슬라이드 (오렌지 배경)
    # ==========================================================================
    agenda_slide = prs.slides.add_slide(blank_layout)
    bg_agenda = agenda_slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333333), Inches(7.5)
    )
    bg_agenda.fill.solid()
    bg_agenda.fill.fore_color.rgb = theme.get("slide_bg", RGBColor(255, 247, 237))
    bg_agenda.line.fill.background()

    title_box = agenda_slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.733), Inches(1.0))
    atf = title_box.text_frame
    atf.word_wrap = True
    
    p_sub = atf.paragraphs[0]
    p_sub.text = "TABLE OF CONTENTS"
    p_sub.font.size = Pt(12)
    p_sub.font.bold = True
    p_sub.font.color.rgb = theme["accent"]
    p_sub.font.name = "Malgun Gothic"

    p_main = atf.add_paragraph()
    p_main.text = "강의 진행 목차 및 구성 개요"
    p_main.font.size = Pt(28)
    p_main.font.bold = True
    p_main.font.color.rgb = theme["primary"]
    p_main.font.name = "Malgun Gothic"

    agenda_items = []
    for idx, mod_key in enumerate(selected_module_keys, 1):
        mod_info = AML_MODULE_DATABASE.get(mod_key)
        if mod_info:
            agenda_items.append((f"{idx:02d}", mod_info["name"], mod_info["subtitle"]))

    if custom_content_text.strip():
        agenda_items.append((
            f"{len(agenda_items)+1:02d}",
            "최신 법령 개정 및 현장 이슈 사례 분석",
            "업로드 및 입력된 특화 데이터 기반 실전 케이스 스터디"
        ))

    agenda_items.append((
        f"{len(agenda_items)+1:02d}",
        "Q&A 및 컴플라이언스 실천 다짐",
        "자주 묻는 질문(FAQ) 및 준법감시부서 핫라인 안내"
    ))

    col_width = Inches(5.6)
    row_height = Inches(1.05)
    start_x = Inches(0.8)
    start_y = Inches(1.8)
    spacing_x = Inches(0.533)
    spacing_y = Inches(0.18)

    for i, (num_str, title_text, desc_text) in enumerate(agenda_items):
        col = i % 2
        row = i // 2
        cur_x = start_x + (col * (col_width + spacing_x))
        cur_y = start_y + (row * (row_height + spacing_y))

        card = agenda_slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, cur_x, cur_y, col_width, row_height
        )
        card.fill.solid()
        card.fill.fore_color.rgb = theme["card_bg"]
        card.line.color.rgb = theme["card_border"]
        card.line.width = Pt(1)

        badge = agenda_slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, cur_x + Inches(0.18), cur_y + Inches(0.18), Inches(0.7), Inches(0.69)
        )
        badge.fill.solid()
        badge.fill.fore_color.rgb = theme["primary"]
        badge.line.fill.background()
        btf = badge.text_frame
        btf.vertical_anchor = MSO_ANCHOR.MIDDLE
        bp = btf.paragraphs[0]
        bp.text = num_str
        bp.alignment = PP_ALIGN.CENTER
        bp.font.size = Pt(15)
        bp.font.bold = True
        bp.font.color.rgb = theme["white"]
        bp.font.name = "Malgun Gothic"

        tx_box = agenda_slide.shapes.add_textbox(
            cur_x + Inches(1.0), cur_y + Inches(0.1), col_width - Inches(1.15), row_height - Inches(0.2)
        )
        ttf = tx_box.text_frame
        ttf.word_wrap = True
        ttf.margin_left = ttf.margin_right = ttf.margin_top = ttf.margin_bottom = 0
        
        tp1 = ttf.paragraphs[0]
        tp1.text = title_text
        tp1.font.size = Pt(13)
        tp1.font.bold = True
        tp1.font.color.rgb = theme["text_dark"]
        tp1.font.name = "Malgun Gothic"

        tp2 = ttf.add_paragraph()
        tp2.text = desc_text if len(desc_text) <= 38 else desc_text[:36] + "..."
        tp2.font.size = Pt(10)
        tp2.font.color.rgb = theme["text_muted"]
        tp2.font.name = "Malgun Gothic"

    add_slide_footer(agenda_slide, current_slide_num, total_slide_count)
    current_slide_num += 1

    # ==========================================================================
    # SLIDES 3..N: 모듈별 슬라이드 (3열 카드 정밀 레이아웃)
    # ==========================================================================
    for mod_key in selected_module_keys:
        mod_info = AML_MODULE_DATABASE.get(mod_key)
        if not mod_info:
            continue

        mod_slide = prs.slides.add_slide(blank_layout)
        bg_mod = mod_slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333333), Inches(7.5)
        )
        bg_mod.fill.solid()
        bg_mod.fill.fore_color.rgb = theme.get("slide_bg", RGBColor(255, 247, 237))
        bg_mod.line.fill.background()

        header_box = mod_slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(11.733), Inches(1.2))
        htf = header_box.text_frame
        htf.word_wrap = True

        p_mod_badge = htf.paragraphs[0]
        p_mod_badge.text = mod_info["badge"].upper()
        p_mod_badge.font.size = Pt(11)
        p_mod_badge.font.bold = True
        p_mod_badge.font.color.rgb = theme["accent"]
        p_mod_badge.font.name = "Malgun Gothic"

        p_mod_title = htf.add_paragraph()
        p_mod_title.text = mod_info["name"]
        p_mod_title.font.size = Pt(24)
        p_mod_title.font.bold = True
        p_mod_title.font.color.rgb = theme["primary"]
        p_mod_title.font.name = "Malgun Gothic"

        p_mod_sub = htf.add_paragraph()
        p_mod_sub.text = mod_info["subtitle"]
        p_mod_sub.font.size = Pt(12)
        p_mod_sub.font.color.rgb = theme["text_muted"]
        p_mod_sub.font.name = "Malgun Gothic"

        sec_list = mod_info.get("sections", [])
        num_sections = len(sec_list)
        if num_sections == 0:
            continue

        card_w = Inches(3.68)
        gap = Inches(0.346)
        card_h = Inches(4.75)
        card_y = Inches(1.85)

        for s_idx, sec in enumerate(sec_list):
            card_x = Inches(0.8) + (s_idx * (card_w + gap))
            sec_card = mod_slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, card_x, card_y, card_w, card_h
            )
            sec_card.fill.solid()
            sec_card.fill.fore_color.rgb = theme["card_bg"]
            sec_card.line.color.rgb = theme["card_border"]
            sec_card.line.width = Pt(1)

            top_line = mod_slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, card_x, card_y, card_w, Inches(0.1)
            )
            top_line.fill.solid()
            top_line.fill.fore_color.rgb = theme["secondary"] if s_idx != 1 else theme["accent"]
            top_line.line.fill.background()

            sec_tx = mod_slide.shapes.add_textbox(
                card_x + Inches(0.2), card_y + Inches(0.25), card_w - Inches(0.4), card_h - Inches(0.4)
            )
            stf = sec_tx.text_frame
            stf.word_wrap = True

            p_stitle = stf.paragraphs[0]
            p_stitle.text = sec["title"]
            p_stitle.font.size = Pt(15)
            p_stitle.font.bold = True
            p_stitle.font.color.rgb = theme["primary"]
            p_stitle.font.name = "Malgun Gothic"
            p_stitle.space_after = Pt(12)

            for bullet in sec.get("bullets", []):
                p_bullet = stf.add_paragraph()
                p_bullet.text = f"• {bullet}"
                p_bullet.font.size = Pt(12)
                p_bullet.font.color.rgb = theme["text_dark"]
                p_bullet.font.name = "Malgun Gothic"
                p_bullet.space_after = Pt(8)

        add_slide_footer(mod_slide, current_slide_num, total_slide_count)
        current_slide_num += 1

    # ==========================================================================
    # CUSTOM SLIDE: 최신 법령 및 사례 슬라이드
    # ==========================================================================
    if custom_content_text.strip():
        custom_slide = prs.slides.add_slide(blank_layout)
        bg_custom = custom_slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333333), Inches(7.5)
        )
        bg_custom.fill.solid()
        bg_custom.fill.fore_color.rgb = theme.get("slide_bg", RGBColor(255, 247, 237))
        bg_custom.line.fill.background()

        c_title_box = custom_slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(11.733), Inches(1.2))
        ctf = c_title_box.text_frame
        ctf.word_wrap = True

        cp_badge = ctf.paragraphs[0]
        cp_badge.text = "SPECIAL TOPIC • 현장 맞춤형 자료"
        cp_badge.font.size = Pt(11)
        cp_badge.font.bold = True
        cp_badge.font.color.rgb = theme["accent"]
        cp_badge.font.name = "Malgun Gothic"

        cp_title = ctf.add_paragraph()
        cp_title.text = "최신 법령 개정 및 현장 이슈 사례 심층 분석"
        cp_title.font.size = Pt(24)
        cp_title.font.bold = True
        cp_title.font.color.rgb = theme["primary"]
        cp_title.font.name = "Malgun Gothic"

        cp_sub = ctf.add_paragraph()
        cp_sub.text = "강의 담당자 입력 자료 및 최근 금융감독원/KoFIU 검사 동향 반영"
        cp_sub.font.size = Pt(12)
        cp_sub.font.color.rgb = theme["text_muted"]
        cp_sub.font.name = "Malgun Gothic"

        left_w = Inches(7.4)
        right_w = Inches(3.98)
        card_h = Inches(4.75)
        card_y = Inches(1.85)

        left_card = custom_slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), card_y, left_w, card_h
        )
        left_card.fill.solid()
        left_card.fill.fore_color.rgb = theme["card_bg"]
        left_card.line.color.rgb = theme["card_border"]

        ltx = custom_slide.shapes.add_textbox(Inches(1.05), card_y + Inches(0.25), left_w - Inches(0.5), card_h - Inches(0.5))
        ltf = ltx.text_frame
        ltf.word_wrap = True

        lp_head = ltf.paragraphs[0]
        lp_head.text = "📑 입력/업로드 자료 핵심 내용"
        lp_head.font.size = Pt(16)
        lp_head.font.bold = True
        lp_head.font.color.rgb = theme["primary"]
        lp_head.font.name = "Malgun Gothic"
        lp_head.space_after = Pt(12)

        custom_bullets = chunk_custom_text(custom_content_text)
        if not custom_bullets:
            custom_bullets = [custom_content_text[:300]]

        for b in custom_bullets:
            bp = ltf.add_paragraph()
            bp.text = f"• {b}"
            bp.font.size = Pt(12)
            bp.font.color.rgb = theme["text_dark"]
            bp.font.name = "Malgun Gothic"
            bp.space_after = Pt(8)

        right_card = custom_slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.55), card_y, right_w, card_h
        )
        right_card.fill.solid()
        right_card.fill.fore_color.rgb = theme["secondary"]
        right_card.line.color.rgb = theme["primary"]

        rtx = custom_slide.shapes.add_textbox(Inches(8.8), card_y + Inches(0.25), right_w - Inches(0.5), card_h - Inches(0.5))
        rtf = rtx.text_frame
        rtf.word_wrap = True

        rp_head = rtf.paragraphs[0]
        rp_head.text = "💡 실무 점검 가이드"
        rp_head.font.size = Pt(16)
        rp_head.font.bold = True
        rp_head.font.color.rgb = theme["accent"]
        rp_head.font.name = "Malgun Gothic"
        rp_head.space_after = Pt(14)

        action_points = [
            "관련 업무 지침 및 내부규정 즉시 업데이트 여부 점검",
            "동일 유형 거래에 대한 이상거래탐지시스템(FDS/STR) 시나리오 룰셋 고도화",
            "관련 실무 부서 및 영업점 대상 즉시 공지 및 현장 전파 교육 시행",
            "금융감독원 정기 종합검사 전 사전 자체 점검 리스트 항목에 추가 반영"
        ]
        for ap in action_points:
            app = rtf.add_paragraph()
            app.text = f"✔ {ap}"
            app.font.size = Pt(12)
            app.font.color.rgb = theme["white"]
            app.font.name = "Malgun Gothic"
            app.space_after = Pt(10)

        add_slide_footer(custom_slide, current_slide_num, total_slide_count)
        current_slide_num += 1

    # ==========================================================================
    # FINAL SLIDE: 마무리 슬라이드 (전체 주황색 계열 풀 블리드 배경)
    # ==========================================================================
    final_slide = prs.slides.add_slide(blank_layout)
    bg_fill_final = final_slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333333), Inches(7.5)
    )
    bg_fill_final.fill.solid()
    bg_fill_final.fill.fore_color.rgb = theme["primary"]
    bg_fill_final.line.fill.background()

    fin_tx = final_slide.shapes.add_textbox(Inches(1.4), Inches(1.3), Inches(10.5), Inches(4.8))
    ftf = fin_tx.text_frame
    ftf.word_wrap = True

    fp_tag = ftf.paragraphs[0]
    fp_tag.text = "COMPLIANCE MINDSET"
    fp_tag.font.size = Pt(13)
    fp_tag.font.bold = True
    fp_tag.font.color.rgb = theme["accent"]
    fp_tag.font.name = "Malgun Gothic"
    fp_tag.space_after = Pt(10)

    fp_title = ftf.add_paragraph()
    fp_title.text = "Question & Answer  |  준법 실천 다짐"
    fp_title.font.size = Pt(32)
    fp_title.font.bold = True
    fp_title.font.color.rgb = theme["white"]
    fp_title.font.name = "Malgun Gothic"
    fp_title.space_after = Pt(16)

    fp_sub = ftf.add_paragraph()
    fp_sub.text = "“자금세탁방지는 형식적 의무가 아닌, 금융기관의 생존과 신뢰를 지키는 최후의 보루입니다.”"
    fp_sub.font.size = Pt(16)
    fp_sub.font.color.rgb = RGBColor(255, 237, 213)
    fp_sub.font.name = "Malgun Gothic"
    fp_sub.space_after = Pt(26)

    c_box = final_slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.4), Inches(3.9), Inches(10.5), Inches(2.1)
    )
    c_box.fill.solid()
    c_box.fill.fore_color.rgb = theme["secondary"]
    c_box.line.fill.background()

    ctx = final_slide.shapes.add_textbox(Inches(1.7), Inches(4.05), Inches(9.9), Inches(1.8))
    ctf_box = ctx.text_frame
    ctf_box.word_wrap = True

    cp1 = ctf_box.paragraphs[0]
    cp1.text = "📞 준법지원 및 의심거래 자문 핫라인"
    cp1.font.size = Pt(14)
    cp1.font.bold = True
    cp1.font.color.rgb = theme["accent"]
    cp1.font.name = "Malgun Gothic"
    cp1.space_after = Pt(6)

    cp2 = ctf_box.add_paragraph()
    cp2.text = "• 담당 부서: 준법감시실 자금세탁방지팀 (내선: 02-XXXX-XXXX / aml-support@company.com)\n" \
               "• 보고 채널: 사내 인트라넷 익명 제보 시스템 및 KoFIU 전자보고 포털\n" \
               "• 수고 많으셨습니다. 본 강의 자료는 사내 학습관리시스템(LMS)에서 언제든지 복습 가능합니다."
    cp2.font.size = Pt(12)
    cp2.font.color.rgb = theme["white"]
    cp2.font.name = "Malgun Gothic"

    add_slide_footer(final_slide, current_slide_num, total_slide_count)

    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream


# ==============================================================================
# 5. Streamlit 메인 애플리케이션 UI (Layout & Interactivity)
# ==============================================================================
def main():
    st.set_page_config(
        page_title="AML Slide Builder - 자금세탁방지 강의 프레젠테이션 자동 생성기",
        page_icon="💼",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    st.markdown("""
    <style>
    .stApp {
        background-color: #FFF7ED;
    }
    [data-testid="stSidebar"] {
        background-color: #FFEDD5;
        border-right: 1px solid #FED7AA;
    }
    .main-title {
        font-size: 2.1rem;
        font-weight: 800;
        color: #9A3412;
        margin-bottom: 0.2rem;
    }
    .sub-text {
        font-size: 1.05rem;
        color: #475569;
        margin-bottom: 1.5rem;
    }
    .preview-card {
        background-color: #FFF7ED;
        border: 1px solid #FED7AA;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 12px;
    }
    .preview-badge {
        display: inline-block;
        background-color: #9A3412;
        color: white;
        font-size: 0.75rem;
        font-weight: 700;
        padding: 2px 8px;
        border-radius: 4px;
        margin-bottom: 6px;
    }
    .badge-accent {
        background-color: #F59E0B;
    }
    </style>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("### ⚙️ AML Slide Builder")
        st.caption("자금세탁방지 강의 프레젠테이션 자동 생성기")
        st.markdown("---")

        st.markdown("#### 1️⃣ 강의 기본 설정")
        lecture_title = st.text_input(
            "강의 제목",
            value="금융기관 임직원을 위한 자금세탁방지(AML/CFT) 핵심 실무",
            help="생성될 PPTX 표지 및 바닥글에 반영됩니다."
        )

        target_audience = st.selectbox(
            "수강 대상",
            options=["신규입사자", "영업점실무자", "AML전담부서", "임원 및 준법감시인"],
            index=1,
            help="수강 대상에 맞춘 표지 메타정보 및 권장 슬라이드 구성에 활용됩니다."
        )

        duration_min = st.slider(
            "강의 시간 (분)",
            min_value=30,
            max_value=120,
            value=60,
            step=10,
            help="강의 소요 시간에 따라 표지에 표기됩니다."
        )

        instructor_name = st.text_input(
            "강사명 및 직책",
            value="준법감시팀장 / AML 공인전문가(CAMS)",
            help="표지 슬라이드 강사 소개 영역에 표기됩니다."
        )

        st.markdown("---")

        st.markdown("#### 2️⃣ 프레젠테이션 테마")
        theme_choice = st.selectbox(
            "메인 포인트 컬러",
            options=list(THEME_PALETTES.keys()),
            index=0,  # Amber Orange가 기본값
            help="슬라이드 전반의 헤더, 카드, 강조색상에 일괄 적용됩니다."
        )

        selected_theme = THEME_PALETTES[theme_choice]
        st.markdown(
            f"""
            <div style="display:flex; align-items:center; gap:8px; margin-top:4px; margin-bottom:12px;">
                <span style="font-size:0.85rem; color:#64748B;">테마 프리뷰:</span>
                <span style="width:24px; height:24px; border-radius:4px; background-color:{selected_theme['hex_primary']}; border:1px solid #FED7AA;" title="Primary"></span>
                <span style="width:24px; height:24px; border-radius:4px; background-color:{selected_theme['hex_secondary']}; border:1px solid #FED7AA;" title="Secondary"></span>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown("---")

        st.markdown("#### 3️⃣ 최신 자료 및 사례 반영")
        st.caption("최신 법령 개정안, 금융감독원 제재 사례, 사내 지침 등을 입력하세요.")

        custom_text_input = st.text_area(
            "텍스트 직접 입력",
            value="",
            placeholder="예시: 2024년 FIU 제재 사례: 가상자산 연계 의심거래보고(STR) 지연으로 과태료 2.5억원 부과...",
            height=120
        )

        uploaded_file = st.file_uploader(
            "관련 파일 업로드 (.txt, .pdf, .docx)",
            type=["txt", "pdf", "docx"],
            help="텍스트를 자동 파싱하여 '최신 법령 및 사례 분석' 슬라이드에 반영합니다."
        )

        parsed_file_text = ""
        if uploaded_file is not None:
            with st.spinner("파일 내용을 파싱하는 중입니다..."):
                parsed_file_text = parse_uploaded_document(uploaded_file)
                if parsed_file_text:
                    st.success(f"✅ '{uploaded_file.name}' 파싱 완료 ({len(parsed_file_text):,}자 추출)")

        combined_custom_text = (custom_text_input.strip() + "\n\n" + parsed_file_text.strip()).strip()

    st.markdown('<div class="main-title">💼 AML 강의 프레젠테이션 자동 생성 웹앱</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-text">자금세탁방지(AML) 강의 담당자가 16:9 와이드스크린 규격의 고품질 PPTX를 단 1초 만에 자동 빌드합니다.</div>', unsafe_allow_html=True)

    st.subheader("1. 강의 모듈 선택 (Curriculum Selection)")
    st.caption("발표 목적과 수강 대상에 맞춰 포함할 교육 모듈을 체크하세요.")

    col1, col2 = st.columns(2)
    selected_modules = []

    module_items = list(AML_MODULE_DATABASE.items())
    half = len(module_items) // 2

    with col1:
        for key, data in module_items[:half]:
            checked = st.checkbox(
                f"**{data['name']}**",
                value=True,
                key=f"mod_chk_{key}",
                help=data["subtitle"]
            )
            st.caption(f"↳ {data['subtitle']}")
            if checked:
                selected_modules.append(key)

    with col2:
        for key, data in module_items[half:]:
            checked = st.checkbox(
                f"**{data['name']}**",
                value=True,
                key=f"mod_chk_{key}",
                help=data["subtitle"]
            )
            st.caption(f"↳ {data['subtitle']}")
            if checked:
                selected_modules.append(key)

    st.markdown("---")

    st.subheader("2. 슬라이드 구성 및 목차 미리보기")
    total_slides_est = 2 + len(selected_modules) + (1 if combined_custom_text else 0) + 1
    
    st.info(
        f"📊 **예상 생성 슬라이드 수:** 총 **{total_slides_est}장** "
        f"(표지 1장 + 목차 1장 + 선택 모듈 {len(selected_modules)}장 + "
        f"{'특화 사례 1장 + ' if combined_custom_text else ''}마무리 1장)"
    )

    with st.expander("🔍 슬라이드별 상세 구성 개요 펼쳐보기", expanded=True):
        st.markdown(
            f"""
            <div class="preview-card">
                <span class="preview-badge">Slide 01 • 표지</span>
                <strong style="font-size:1.1rem; color:#9A3412;">{lecture_title}</strong>
                <p style="margin:4px 0 0 0; font-size:0.9rem; color:#64748B;">
                    • 수강 대상: {target_audience} | 강사: {instructor_name} | 시간: {duration_min}분 | 테마: {theme_choice.split(' ')[0]}
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown(
            """
            <div class="preview-card">
                <span class="preview-badge">Slide 02 • 목차 (Agenda)</span>
                <strong style="font-size:1.05rem; color:#9A3412;">강의 진행 아젠다 및 프레임워크</strong>
                <p style="margin:4px 0 0 0; font-size:0.9rem; color:#64748B;">선택된 모듈 기반 2열 카드 그리드 자동 배치</p>
            </div>
            """,
            unsafe_allow_html=True
        )

        for idx, m_key in enumerate(selected_modules, 3):
            mod = AML_MODULE_DATABASE[m_key]
            st.markdown(
                f"""
                <div class="preview-card">
                    <span class="preview-badge">Slide {idx:02d} • {mod['badge']}</span>
                    <strong style="font-size:1.05rem; color:#9A3412;">{mod['name']}</strong>
                    <p style="margin:4px 0 0 0; font-size:0.9rem; color:#475569;">{mod['subtitle']}</p>
                    <ul style="margin:6px 0 0 16px; font-size:0.85rem; color:#64748B;">
                        {''.join([f"<li>{s['title']}</li>" for s in mod['sections']])}
                    </ul>
                </div>
                """,
                unsafe_allow_html=True
            )

        if combined_custom_text:
            custom_slide_idx = 2 + len(selected_modules) + 1
            st.markdown(
                f"""
                <div class="preview-card" style="border-left: 4px solid #F59E0B;">
                    <span class="preview-badge badge-accent">Slide {custom_slide_idx:02d} • 특화 사례 & 법령</span>
                    <strong style="font-size:1.05rem; color:#9A3412;">최신 법령 개정 및 현장 이슈 사례 분석</strong>
                    <p style="margin:4px 0 0 0; font-size:0.9rem; color:#475569;">
                        사용자 입력 텍스트 및 업로드 파일 기반 실무 체크리스트 생성 ({len(combined_custom_text):,}자)
                    </p>
                </div>
                """,
                unsafe_allow_html=True
            )

        st.markdown(
            f"""
            <div class="preview-card">
                <span class="preview-badge">Slide {total_slides_est:02d} • Q&A / 마무리</span>
                <strong style="font-size:1.05rem; color:#9A3412;">질의응답(Q&A) 및 준법 실천 다짐</strong>
                <p style="margin:4px 0 0 0; font-size:0.9rem; color:#64748B;">준법감시실 자금세탁방지팀 핫라인 및 사내 신고 채널 안내</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("---")
    st.subheader("3. PPTX 프레젠테이션 생성 및 다운로드")

    if not selected_modules:
        st.warning("⚠️ 최소 1개 이상의 강의 모듈을 선택해 주세요.")
        return

    generate_clicked = st.button("🎯 PPTX 프레젠테이션 자동 생성하기", type="primary", use_container_width=True)

    if generate_clicked:
        with st.spinner("16:9 와이드스크린 PPTX를 빌드 중입니다..."):
            try:
                pptx_stream = create_aml_presentation(
                    lecture_title=lecture_title,
                    target_audience=target_audience,
                    duration_min=duration_min,
                    instructor_name=instructor_name,
                    theme_key=theme_choice,
                    selected_module_keys=selected_modules,
                    custom_content_text=combined_custom_text
                )
                
                clean_title = "".join(c for c in lecture_title if c.isalnum() or c in (" ", "_", "-")).rstrip()
                file_name = f"AML_Presentation_{clean_title[:15]}_{datetime.now().strftime('%Y%m%d_%H%M')}.pptx"

                st.success("🎉 PPTX 프레젠테이션 생성이 성공적으로 완료되었습니다!")
                
                st.download_button(
                    label="📥 생성된 PPTX 파일 다운로드 (.pptx)",
                    data=pptx_stream,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary",
                    use_container_width=True
                )

                st.balloons()

            except Exception as e:
                st.error(f"❌ PPTX 생성 중 예외가 발생했습니다: {str(e)}")
                st.exception(e)


if __name__ == "__main__":
    main()
