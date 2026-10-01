import PptxGenJS from 'pptxgenjs';
import { AmlModule, ThemeConfig } from './amlData';

export interface PptxOptions {
  lectureTitle: string;
  targetAudience: string;
  durationMin: number;
  instructorName: string;
  theme: ThemeConfig;
  selectedModules: AmlModule[];
  customText: string;
}

export async function generatePptxFile(options: PptxOptions): Promise<string> {
  const {
    lectureTitle,
    targetAudience,
    durationMin,
    instructorName,
    theme,
    selectedModules,
    customText,
  } = options;

  const pres = new PptxGenJS();

  // 16:9 와이드스크린 규격 등록 및 적용 (13.333333" x 7.5")
  pres.defineLayout({ name: 'LAYOUT_16X9_WIDE', width: 13.333333, height: 7.5 });
  pres.layout = 'LAYOUT_16X9_WIDE';

  const todayStr = new Date().toLocaleDateString('ko-KR', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
  });

  const totalSlides =
    2 +
    selectedModules.length +
    (customText.trim() ? 1 : 0) +
    1;
  let currentSlide = 1;

  // --------------------------------------------------------------------------
  // Helper: Slide Footer (바닥글)
  // --------------------------------------------------------------------------
  const addFooter = (slide: PptxGenJS.Slide, slideNum: number) => {
    slide.addText(
      `AML/CFT Compliance Education  |  ${lectureTitle}  |  Slide ${slideNum} of ${totalSlides}`,
      {
        x: 0.8,
        y: 6.9,
        w: 11.733,
        h: 0.35,
        fontSize: 10,
        fontFace: 'Malgun Gothic',
        color: theme.pptxTextMuted,
        align: 'left',
        valign: 'middle',
      }
    );
  };

  // ==========================================================================
  // SLIDE 1: Cover Slide (표지 슬라이드 - 주황색 계열 풀 블리드 배경)
  // ==========================================================================
  const coverSlide = pres.addSlide();

  // 슬라이드 전체 배경을 테마 주 색상(비비드/딥 오렌지)으로 설정
  coverSlide.background = { color: theme.pptxPrimary };

  // 상단 골드/액센트 포인트 스트립 (13.333" 전체 폭)
  coverSlide.addShape(pres.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.35,
    fill: { color: theme.pptxAccent },
    line: { color: theme.pptxAccent },
  });

  // 메인 비주얼 카드 (0.8" 좌우 여백, 깊이감 있는 2차 오렌지 톤)
  coverSlide.addShape(pres.ShapeType.roundRect, {
    x: 0.8,
    y: 0.85,
    w: 11.733,
    h: 5.75,
    fill: { color: theme.pptxPrimary },
    line: { color: theme.pptxSecondary, width: 2 },
    rectRadius: 0.12,
  });

  // 카테고리 태그 뱃지
  coverSlide.addText('🛡️ COMPLIANCE & FINANCIAL CRIME PREVENTION', {
    x: 1.4,
    y: 1.3,
    w: 10.5,
    h: 0.4,
    fontSize: 13,
    fontFace: 'Malgun Gothic',
    bold: true,
    color: theme.pptxAccent,
  });

  // 메인 강의 제목
  coverSlide.addText(lectureTitle, {
    x: 1.4,
    y: 1.8,
    w: 10.5,
    h: 1.6,
    fontSize: 34,
    fontFace: 'Malgun Gothic',
    bold: true,
    color: 'FFFFFF',
    valign: 'top',
    wrap: true,
  });

  // 강의 개요 서브타이틀
  coverSlide.addText(
    '자금세탁방지(AML) 및 공중협박자금조달금지(CFT) 제도 실무 역량 강화를 위한 핵심 교육',
    {
      x: 1.4,
      y: 3.55,
      w: 10.5,
      h: 0.6,
      fontSize: 16,
      fontFace: 'Malgun Gothic',
      color: 'FFEDD5',
      wrap: true,
    }
  );

  // 하단 메타정보 박스 (수강대상, 강의시간, 강사명, 일자)
  coverSlide.addShape(pres.ShapeType.roundRect, {
    x: 1.4,
    y: 4.6,
    w: 10.5,
    h: 1.4,
    fill: { color: theme.pptxSecondary },
    line: { color: theme.pptxSecondary },
    rectRadius: 0.08,
  });

  coverSlide.addText(
    [
      {
        text: `• 수강 대상:  ${targetAudience}          • 강의 소요시간:  ${durationMin}분\n`,
        options: { bold: true, fontSize: 14, color: 'FFFFFF' },
      },
      {
        text: `• 전담 강사:  ${instructorName}          • 작성 및 시행일자:  ${todayStr}`,
        options: { bold: false, fontSize: 13, color: 'FED7AA' },
      },
    ],
    {
      x: 1.7,
      y: 4.8,
      w: 9.9,
      h: 1.0,
      fontFace: 'Malgun Gothic',
      valign: 'middle',
      wrap: true,
    }
  );

  currentSlide++;

  // ==========================================================================
  // SLIDE 2: Agenda Slide (목차 슬라이드 - 오렌지 배경)
  // ==========================================================================
  const agendaSlide = pres.addSlide();
  agendaSlide.background = { color: theme.pptxSlideBg || 'FFF7ED' };

  // Header Title
  agendaSlide.addText('TABLE OF CONTENTS', {
    x: 0.8,
    y: 0.6,
    w: 11.733,
    h: 0.3,
    fontSize: 12,
    fontFace: 'Malgun Gothic',
    bold: true,
    color: theme.pptxAccent,
  });

  agendaSlide.addText('강의 진행 목차 및 구성 개요', {
    x: 0.8,
    y: 0.9,
    w: 11.733,
    h: 0.6,
    fontSize: 28,
    fontFace: 'Malgun Gothic',
    bold: true,
    color: theme.pptxPrimary,
  });

  const agendaItems: Array<{ num: string; title: string; subtitle: string }> = [];
  selectedModules.forEach((mod, idx) => {
    agendaItems.push({
      num: String(idx + 1).padStart(2, '0'),
      title: mod.name,
      subtitle: mod.subtitle,
    });
  });

  if (customText.trim()) {
    agendaItems.push({
      num: String(agendaItems.length + 1).padStart(2, '0'),
      title: '최신 법령 개정 및 현장 이슈 사례 분석',
      subtitle: '업로드 및 입력된 특화 데이터 기반 실무 케이스 스터디',
    });
  }

  agendaItems.push({
    num: String(agendaItems.length + 1).padStart(2, '0'),
    title: 'Q&A 및 컴플라이언스 실천 다짐',
    subtitle: '자주 묻는 질문(FAQ) 및 준법감시부서 핫라인 안내',
  });

  const colWidth = 5.6;
  const rowHeight = 1.05;
  const startX = 0.8;
  const startY = 1.8;
  const gapX = 0.533;
  const gapY = 0.18;

  agendaItems.forEach((item, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = startX + col * (colWidth + gapX);
    const y = startY + row * (rowHeight + gapY);

    agendaSlide.addShape(pres.ShapeType.roundRect, {
      x,
      y,
      w: colWidth,
      h: rowHeight,
      fill: { color: theme.pptxCardBg },
      line: { color: theme.pptxCardBorder, width: 1 },
      rectRadius: 0.08,
    });

    agendaSlide.addShape(pres.ShapeType.rect, {
      x: x + 0.18,
      y: y + 0.18,
      w: 0.7,
      h: 0.69,
      fill: { color: theme.pptxPrimary },
      line: { color: theme.pptxPrimary },
    });

    agendaSlide.addText(item.num, {
      x: x + 0.18,
      y: y + 0.18,
      w: 0.7,
      h: 0.69,
      fontSize: 15,
      fontFace: 'Malgun Gothic',
      bold: true,
      color: 'FFFFFF',
      align: 'center',
      valign: 'middle',
    });

    agendaSlide.addText(
      [
        {
          text: `${item.title}\n`,
          options: { bold: true, fontSize: 13, color: theme.pptxTextDark },
        },
        {
          text: item.subtitle.length > 38 ? item.subtitle.slice(0, 36) + '...' : item.subtitle,
          options: { bold: false, fontSize: 10, color: theme.pptxTextMuted },
        },
      ],
      {
        x: x + 1.0,
        y: y + 0.1,
        w: colWidth - 1.15,
        h: rowHeight - 0.2,
        fontFace: 'Malgun Gothic',
        valign: 'middle',
        wrap: true,
      }
    );
  });

  addFooter(agendaSlide, currentSlide);
  currentSlide++;

  // ==========================================================================
  // SLIDES 3..N: Module Slides (모듈별 3단 카드 슬라이드 - 오렌지 배경)
  // ==========================================================================
  for (const mod of selectedModules) {
    const slide = pres.addSlide();
    slide.background = { color: theme.pptxSlideBg || 'FFF7ED' };

    slide.addText(mod.badge.toUpperCase(), {
      x: 0.8,
      y: 0.45,
      w: 11.733,
      h: 0.3,
      fontSize: 11,
      fontFace: 'Malgun Gothic',
      bold: true,
      color: theme.pptxAccent,
    });

    slide.addText(mod.name, {
      x: 0.8,
      y: 0.75,
      w: 11.733,
      h: 0.55,
      fontSize: 24,
      fontFace: 'Malgun Gothic',
      bold: true,
      color: theme.pptxPrimary,
      wrap: true,
    });

    slide.addText(mod.subtitle, {
      x: 0.8,
      y: 1.3,
      w: 11.733,
      h: 0.35,
      fontSize: 12,
      fontFace: 'Malgun Gothic',
      color: theme.pptxTextMuted,
      wrap: true,
    });

    const cardW = 3.68;
    const gap = 0.346;
    const cardH = 4.75;
    const cardY = 1.85;

    mod.sections.forEach((sec, sIdx) => {
      const cardX = 0.8 + sIdx * (cardW + gap);

      slide.addShape(pres.ShapeType.roundRect, {
        x: cardX,
        y: cardY,
        w: cardW,
        h: cardH,
        fill: { color: theme.pptxCardBg },
        line: { color: theme.pptxCardBorder, width: 1 },
        rectRadius: 0.08,
      });

      slide.addShape(pres.ShapeType.rect, {
        x: cardX,
        y: cardY,
        w: cardW,
        h: 0.1,
        fill: { color: sIdx === 1 ? theme.pptxAccent : theme.pptxSecondary },
        line: { color: sIdx === 1 ? theme.pptxAccent : theme.pptxSecondary },
      });

      slide.addText(sec.title, {
        x: cardX + 0.2,
        y: cardY + 0.25,
        w: cardW - 0.4,
        h: 0.45,
        fontSize: 15,
        fontFace: 'Malgun Gothic',
        bold: true,
        color: theme.pptxPrimary,
        wrap: true,
      });

      const bulletParagraphs = sec.bullets.map((b) => ({
        text: `• ${b}\n\n`,
        options: {
          fontSize: 12,
          color: theme.pptxTextDark,
          fontFace: 'Malgun Gothic',
        },
      }));

      slide.addText(bulletParagraphs, {
        x: cardX + 0.2,
        y: cardY + 0.75,
        w: cardW - 0.4,
        h: cardH - 0.9,
        fontFace: 'Malgun Gothic',
        valign: 'top',
        wrap: true,
      });
    });

    addFooter(slide, currentSlide);
    currentSlide++;
  }

  // ==========================================================================
  // CUSTOM SLIDE: 최신 법령 및 사례 슬라이드
  // ==========================================================================
  if (customText.trim()) {
    const customSlide = pres.addSlide();
    customSlide.background = { color: theme.pptxSlideBg || 'FFF7ED' };

    customSlide.addText('SPECIAL TOPIC • 현장 맞춤형 자료', {
      x: 0.8,
      y: 0.45,
      w: 11.733,
      h: 0.3,
      fontSize: 11,
      fontFace: 'Malgun Gothic',
      bold: true,
      color: theme.pptxAccent,
    });

    customSlide.addText('최신 법령 개정 및 현장 이슈 사례 심층 분석', {
      x: 0.8,
      y: 0.75,
      w: 11.733,
      h: 0.55,
      fontSize: 24,
      fontFace: 'Malgun Gothic',
      bold: true,
      color: theme.pptxPrimary,
      wrap: true,
    });

    customSlide.addText('강의 담당자 입력 자료 및 최근 금융감독원/KoFIU 검사 동향 반영', {
      x: 0.8,
      y: 1.3,
      w: 11.733,
      h: 0.35,
      fontSize: 12,
      fontFace: 'Malgun Gothic',
      color: theme.pptxTextMuted,
      wrap: true,
    });

    const leftW = 7.4;
    const rightW = 3.98;
    const cardH = 4.75;
    const cardY = 1.85;

    customSlide.addShape(pres.ShapeType.roundRect, {
      x: 0.8,
      y: cardY,
      w: leftW,
      h: cardH,
      fill: { color: theme.pptxCardBg },
      line: { color: theme.pptxCardBorder, width: 1 },
      rectRadius: 0.08,
    });

    customSlide.addText('📑 입력/업로드 자료 핵심 내용', {
      x: 1.05,
      y: cardY + 0.25,
      w: leftW - 0.5,
      h: 0.4,
      fontSize: 16,
      fontFace: 'Malgun Gothic',
      bold: true,
      color: theme.pptxPrimary,
      wrap: true,
    });

    const lines = customText
      .split('\n')
      .map((l) => l.trim())
      .filter((l) => l.length > 0)
      .slice(0, 6);

    const customBullets = (lines.length > 0 ? lines : [customText.slice(0, 300)]).map(
      (line) => ({
        text: `• ${line.length > 115 ? line.slice(0, 110) + '...' : line}\n\n`,
        options: {
          fontSize: 12,
          color: theme.pptxTextDark,
          fontFace: 'Malgun Gothic',
        },
      })
    );

    customSlide.addText(customBullets, {
      x: 1.05,
      y: cardY + 0.75,
      w: leftW - 0.5,
      h: cardH - 0.9,
      fontFace: 'Malgun Gothic',
      valign: 'top',
      wrap: true,
    });

    customSlide.addShape(pres.ShapeType.roundRect, {
      x: 8.55,
      y: cardY,
      w: rightW,
      h: cardH,
      fill: { color: theme.pptxSecondary },
      line: { color: theme.pptxPrimary, width: 1 },
      rectRadius: 0.08,
    });

    customSlide.addText('💡 실무 점검 가이드', {
      x: 8.8,
      y: cardY + 0.25,
      w: rightW - 0.5,
      h: 0.4,
      fontSize: 16,
      fontFace: 'Malgun Gothic',
      bold: true,
      color: theme.pptxAccent,
      wrap: true,
    });

    const guides = [
      '관련 업무 지침 및 내부규정 즉시 업데이트 여부 점검',
      '동일 유형 거래에 대한 이상거래탐지(FDS/STR) 룰셋 고도화',
      '관련 실무 부서 및 영업점 대상 즉시 공지 및 현장 전파 교육',
      '금융감독원 정기 종합검사 사전 자체 점검 리스트 항목 반영',
    ];

    customSlide.addText(
      guides.map((g) => ({
        text: `✔ ${g}\n\n`,
        options: {
          fontSize: 12,
          color: 'FFFFFF',
          fontFace: 'Malgun Gothic',
        },
      })),
      {
        x: 8.8,
        y: cardY + 0.75,
        w: rightW - 0.5,
        h: cardH - 0.9,
        fontFace: 'Malgun Gothic',
        valign: 'top',
        wrap: true,
      }
    );

    addFooter(customSlide, currentSlide);
    currentSlide++;
  }

  // ==========================================================================
  // FINAL SLIDE: Q&A Slide (마무리 - 주황색 풀 블리드 배경)
  // ==========================================================================
  const finalSlide = pres.addSlide();
  finalSlide.background = { color: theme.pptxPrimary };

  finalSlide.addShape(pres.ShapeType.roundRect, {
    x: 0.8,
    y: 0.85,
    w: 11.733,
    h: 5.75,
    fill: { color: theme.pptxPrimary },
    line: { color: theme.pptxSecondary, width: 2 },
    rectRadius: 0.12,
  });

  finalSlide.addText('COMPLIANCE MINDSET', {
    x: 1.4,
    y: 1.3,
    w: 10.5,
    h: 0.4,
    fontSize: 13,
    fontFace: 'Malgun Gothic',
    bold: true,
    color: theme.pptxAccent,
  });

  finalSlide.addText('Question & Answer  |  준법 실천 다짐', {
    x: 1.4,
    y: 1.8,
    w: 10.5,
    h: 1.0,
    fontSize: 32,
    fontFace: 'Malgun Gothic',
    bold: true,
    color: 'FFFFFF',
    wrap: true,
  });

  finalSlide.addText(
    '“자금세탁방지는 형식적 의무가 아닌, 금융기관의 생존과 신뢰를 지키는 최후의 보루입니다.”',
    {
      x: 1.4,
      y: 2.85,
      w: 10.5,
      h: 0.6,
      fontSize: 16,
      fontFace: 'Malgun Gothic',
      color: 'FFEDD5',
      wrap: true,
    }
  );

  finalSlide.addShape(pres.ShapeType.roundRect, {
    x: 1.4,
    y: 3.9,
    w: 10.5,
    h: 2.1,
    fill: { color: theme.pptxSecondary },
    line: { color: theme.pptxSecondary },
    rectRadius: 0.08,
  });

  finalSlide.addText(
    [
      {
        text: '📞 준법지원 및 의심거래 자문 핫라인\n',
        options: { bold: true, fontSize: 14, color: theme.pptxAccent },
      },
      {
        text: '• 담당 부서: 준법감시실 자금세탁방지팀 (내선: 02-XXXX-XXXX / aml-support@company.com)\n',
        options: { fontSize: 12, color: 'FFFFFF' },
      },
      {
        text: '• 보고 채널: 사내 인트라넷 익명 제보 시스템 및 KoFIU 전자보고 포털\n',
        options: { fontSize: 12, color: 'FFFFFF' },
      },
      {
        text: '• 수고 많으셨습니다. 본 강의 자료는 사내 학습관리시스템(LMS)에서 언제든지 복습 가능합니다.',
        options: { fontSize: 12, color: 'FED7AA' },
      },
    ],
    {
      x: 1.7,
      y: 4.05,
      w: 9.9,
      h: 1.8,
      fontFace: 'Malgun Gothic',
      valign: 'middle',
      wrap: true,
    }
  );

  addFooter(finalSlide, currentSlide);

  const cleanTitle = lectureTitle.replace(/[^a-zA-Z0-9가-힣_-]/g, '_').slice(0, 18);
  const fileName = `AML_Presentation_${cleanTitle}_${Date.now()}.pptx`;
  await pres.writeFile({ fileName });
  return fileName;
}
