const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3" x 7.5"
pres.author = "Trip Planner";
pres.title = "Kyoto-Osaka Trip Itinerary";

// === Palette: Warm Terracotta (Japan-evoking) ===
const C = {
  primary: "B85042",    // terracotta
  secondary: "E7E8D1",  // sand
  accent: "A7BEAE",     // sage
  dark: "2C2C2C",
  light: "FAF7F2",
  muted: "6B6B6B",
  white: "FFFFFF",
  gold: "C9A961",
};

const FONT_H = "Georgia";
const FONT_B = "Calibri";

const SLIDE_W = 13.3;
const SLIDE_H = 7.5;

// ---------- Helpers ----------
function addFooter(slide, pageNum, totalPages) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: SLIDE_H - 0.35, w: SLIDE_W, h: 0.35,
    fill: { color: C.dark }, line: { color: C.dark },
  });
  slide.addText("교토·오사카 여행 · 2026.07.10–14", {
    x: 0.5, y: SLIDE_H - 0.33, w: 8, h: 0.3,
    fontSize: 10, fontFace: FONT_B, color: C.secondary, valign: "middle", margin: 0,
  });
  slide.addText(`${pageNum} / ${totalPages}`, {
    x: SLIDE_W - 1.5, y: SLIDE_H - 0.33, w: 1, h: 0.3,
    fontSize: 10, fontFace: FONT_B, color: C.secondary, align: "right", valign: "middle", margin: 0,
  });
}

function addHeader(slide, dayLabel, title) {
  // Top bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: SLIDE_W, h: 1.2,
    fill: { color: C.primary }, line: { color: C.primary },
  });
  // Day label box
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.25, w: 1.6, h: 0.7,
    fill: { color: C.gold }, line: { color: C.gold },
  });
  slide.addText(dayLabel, {
    x: 0.5, y: 0.25, w: 1.6, h: 0.7,
    fontSize: 20, fontFace: FONT_H, bold: true, color: C.white,
    align: "center", valign: "middle", margin: 0,
  });
  slide.addText(title, {
    x: 2.3, y: 0.25, w: 10.5, h: 0.7,
    fontSize: 26, fontFace: FONT_H, bold: true, color: C.white,
    valign: "middle", margin: 0,
  });
}

// =====================================================
// Slide 1: Title (cover)
// =====================================================
const TOTAL = 13;
{
  const s = pres.addSlide();
  s.background = { color: C.dark };

  // Decorative side bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 4.5, h: SLIDE_H,
    fill: { color: C.primary }, line: { color: C.primary },
  });
  // Gold accent strip
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.5, y: 0, w: 0.08, h: SLIDE_H,
    fill: { color: C.gold }, line: { color: C.gold },
  });

  // Vertical Japanese-style title (in side bar)
  s.addText("京都", {
    x: 0.5, y: 0.6, w: 3.5, h: 2,
    fontSize: 110, fontFace: FONT_H, bold: true, color: C.white,
    align: "center", valign: "middle", margin: 0,
  });
  s.addText("大阪", {
    x: 0.5, y: 3.0, w: 3.5, h: 2,
    fontSize: 110, fontFace: FONT_H, bold: true, color: C.gold,
    align: "center", valign: "middle", margin: 0,
  });
  s.addText("KYOTO · OSAKA", {
    x: 0.5, y: 5.3, w: 3.5, h: 0.5,
    fontSize: 16, fontFace: FONT_B, charSpacing: 8, color: C.secondary,
    align: "center", valign: "middle", margin: 0,
  });

  // Right side: main title
  s.addText("교토 3일 · 오사카 1일", {
    x: 5.2, y: 2.0, w: 7.5, h: 0.7,
    fontSize: 28, fontFace: FONT_H, italic: true, color: C.secondary,
    valign: "middle", margin: 0,
  });
  s.addText("여행 일정표", {
    x: 5.2, y: 2.7, w: 7.5, h: 1.0,
    fontSize: 54, fontFace: FONT_H, bold: true, color: C.white,
    valign: "middle", margin: 0,
  });
  // Divider
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 3.95, w: 1.5, h: 0.05,
    fill: { color: C.gold }, line: { color: C.gold },
  });
  // Meta info
  s.addText([
    { text: "기간  ", options: { color: C.gold, bold: true, fontSize: 14 } },
    { text: "2026. 07. 10 (금) ~ 07. 14 (화)  ·  4박 5일", options: { color: C.white, fontSize: 14, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "인원  ", options: { color: C.gold, bold: true, fontSize: 14 } },
    { text: "성인 8명 (대학 교수·직원)", options: { color: C.white, fontSize: 14, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "항공  ", options: { color: C.gold, bold: true, fontSize: 14 } },
    { text: "인천(ICN) ↔ 간사이(KIX) 직항", options: { color: C.white, fontSize: 14, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "숙소  ", options: { color: C.gold, bold: true, fontSize: 14 } },
    { text: "교토 3박 (가와라마치/기온) + 오사카 1박 (난바/우메다)", options: { color: C.white, fontSize: 14 } },
  ], {
    x: 5.2, y: 4.2, w: 7.5, h: 2.4, fontFace: FONT_B, valign: "top", margin: 0, paraSpaceAfter: 4,
  });

  s.addText("Prepared 2026.05.27", {
    x: 5.2, y: SLIDE_H - 0.8, w: 7.5, h: 0.3,
    fontSize: 10, fontFace: FONT_B, color: C.muted, italic: true, margin: 0,
  });
}

// =====================================================
// Slide 2: Overview (5-day at a glance)
// =====================================================
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  addHeader(s, "개요", "5일간의 여정 한눈에 보기");

  const days = [
    { d: "DAY 1", date: "07.10 금", title: "입국 & 교토", desc: "인천 → 간사이공항\n하루카 특급 → 교토\n기온·폰토초 저녁", color: C.primary },
    { d: "DAY 2", date: "07.11 토", title: "교토 동부", desc: "기요미즈데라\n산넨자카·고다이지\n야사카·기온 시라카와", color: C.gold },
    { d: "DAY 3", date: "07.12 일", title: "교토 북서부", desc: "금각사·료안지\n아라시야마 대나무숲\n도게쓰교·텐류지", color: C.accent },
    { d: "DAY 4", date: "07.13 월", title: "교토 → 오사카", desc: "후시미 이나리\n우지·뵤도인\n오사카성·도톤보리", color: C.primary },
    { d: "DAY 5", date: "07.14 화", title: "오사카 & 귀국", desc: "우메다·구로몬\n간사이공항\n인천 귀국", color: C.gold },
  ];

  const cardW = 2.32, cardH = 4.5, gap = 0.15;
  const startX = (SLIDE_W - (cardW * 5 + gap * 4)) / 2;
  const startY = 1.7;

  days.forEach((day, i) => {
    const x = startX + i * (cardW + gap);
    // Card background
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: startY, w: cardW, h: cardH,
      fill: { color: C.white }, line: { color: "DDDDDD", width: 0.5 },
      shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 90, opacity: 0.12 },
    });
    // Top color bar
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: startY, w: cardW, h: 0.5,
      fill: { color: day.color }, line: { color: day.color },
    });
    s.addText(day.d, {
      x, y: startY, w: cardW, h: 0.5,
      fontSize: 14, fontFace: FONT_H, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 0, charSpacing: 4,
    });
    // Date
    s.addText(day.date, {
      x, y: startY + 0.65, w: cardW, h: 0.35,
      fontSize: 12, fontFace: FONT_B, color: C.muted,
      align: "center", valign: "middle", margin: 0,
    });
    // Title
    s.addText(day.title, {
      x: x + 0.1, y: startY + 1.05, w: cardW - 0.2, h: 0.7,
      fontSize: 18, fontFace: FONT_H, bold: true, color: C.dark,
      align: "center", valign: "middle", margin: 0,
    });
    // Divider
    s.addShape(pres.shapes.RECTANGLE, {
      x: x + cardW / 2 - 0.3, y: startY + 1.85, w: 0.6, h: 0.03,
      fill: { color: day.color }, line: { color: day.color },
    });
    // Description
    s.addText(day.desc, {
      x: x + 0.15, y: startY + 2.05, w: cardW - 0.3, h: 2.3,
      fontSize: 12, fontFace: FONT_B, color: C.dark,
      align: "center", valign: "top", margin: 0, paraSpaceAfter: 4,
    });
  });

  addFooter(s, 2, TOTAL);
}

// =====================================================
// Slides 3-7: Day 1-5 details
// =====================================================
const dayDetails = [
  {
    dayLabel: "DAY 1",
    title: "07.10 (금) · 입국 & 교토 이동",
    schedule: [
      { time: "09:30", place: "인천공항 출발", note: "대한항공/아시아나" },
      { time: "11:30", place: "간사이공항 도착", note: "입국 수속 약 1시간" },
      { time: "13:30", place: "JR 하루카 → 교토역", note: "단체석 사전예약" },
      { time: "15:30", place: "숙소 체크인 (가와라마치)", note: "" },
      { time: "17:00", place: "기온(祇園) 거리 산책", note: "워밍업" },
      { time: "19:00", place: "저녁: 폰토초 가이세키", note: "★ 사전예약 필수" },
    ],
    mapUrl: "Kansai공항 → 교토역 → 기온 → 폰토초",
    highlights: ["하루카 단체석 미리 예매", "가이세키 식당 2주 전 예약", "첫날은 가볍게"],
  },
  {
    dayLabel: "DAY 2",
    title: "07.11 (토) · 교토 동부 세계유산 코스",
    schedule: [
      { time: "09:00", place: "기요미즈데라(清水寺)", note: "입장료 ¥400" },
      { time: "11:00", place: "산넨자카·니넨자카", note: "기념품·말차" },
      { time: "12:30", place: "점심: 유도후 오쿠탄(奥丹)", note: "★ 8인 단체예약" },
      { time: "14:00", place: "고다이지 + 네네노미치", note: "" },
      { time: "15:30", place: "야사카 신사 → 마루야마 공원", note: "" },
      { time: "17:00", place: "기온 시라카와 산책", note: "게이샤 거리" },
      { time: "19:00", place: "저녁: 교토 와규 데판야키", note: "★ 사전예약" },
    ],
    mapUrl: "기요미즈데라 → 산넨자카 → 고다이지 → 야사카 → 시라카와",
    highlights: ["전 일정 도보 가능 (동선 효율 ↑)", "오후 햇볕 강함 → 양산 필수", "기온 골목 사진 명소"],
  },
  {
    dayLabel: "DAY 3",
    title: "07.12 (일) · 교토 북서부 & 아라시야마",
    schedule: [
      { time: "09:00", place: "금각사(金閣寺)", note: "입장료 ¥500" },
      { time: "10:30", place: "료안지(龍安寺) 석정원", note: "도보 20분" },
      { time: "12:00", place: "아라시야마 이동 + 점심", note: "가이세키" },
      { time: "13:30", place: "대나무숲(竹林の道)", note: "사진 명소" },
      { time: "14:30", place: "도게쓰교 + 호즈가와", note: "" },
      { time: "15:30", place: "텐류지(天龍寺) 정원", note: "세계유산" },
      { time: "17:00", place: "시내 복귀 → 니시키 시장", note: "자유시간" },
    ],
    mapUrl: "금각사 → 료안지 → 아라시야마 → 도게쓰교 → 텐류지 → 니시키",
    highlights: ["MK 점보택시 2대 추천", "아라시야마는 오전이 한산", "대나무숲은 이른 오후가 빛 좋음"],
  },
  {
    dayLabel: "DAY 4",
    title: "07.13 (월) · 교토 남부 → 오사카",
    schedule: [
      { time: "09:00", place: "후시미 이나리 신사", note: "천 개의 도리이" },
      { time: "11:30", place: "우지 → 뵤도인(平等院)", note: "10엔 동전 도안" },
      { time: "13:00", place: "점심: 우지 말차 가이세키", note: "★ 예약" },
      { time: "15:00", place: "JR로 오사카 이동", note: "약 50분" },
      { time: "16:30", place: "숙소 체크인 (난바/우메다)", note: "" },
      { time: "17:00", place: "오사카 성 외관", note: "" },
      { time: "19:00", place: "도톤보리·신사이바시", note: "쿠시카츠·타코야키" },
    ],
    mapUrl: "후시미 이나리 → 뵤도인 → 오사카성 → 도톤보리",
    highlights: ["짐은 호텔 → 호텔 택배 서비스 활용", "후시미는 이른 아침이 시원", "도톤보리 저녁은 자유식"],
  },
  {
    dayLabel: "DAY 5",
    title: "07.14 (화) · 오사카 & 귀국",
    schedule: [
      { time: "09:00", place: "우메다 공중정원 / 구로몬 시장", note: "택1" },
      { time: "11:00", place: "체크아웃 + 짐 보관", note: "" },
      { time: "12:00", place: "점심: 한큐백화점 데파치카", note: "도시락·스시" },
      { time: "14:00", place: "난카이 라피트 → 간사이공항", note: "약 45분" },
      { time: "17:00", place: "KIX 출발", note: "" },
      { time: "19:00", place: "인천공항 도착", note: "" },
    ],
    mapUrl: "우메다 → 구로몬 → 난바 → 간사이공항",
    highlights: ["면세 쇼핑은 데파치카 이후", "라피트 좌석 사전 지정", "비행기 3시간 전 도착"],
  },
];

dayDetails.forEach((day, idx) => {
  const s = pres.addSlide();
  s.background = { color: C.light };
  addHeader(s, day.dayLabel, day.title);

  // Left: timeline
  s.addText("일정 (Timeline)", {
    x: 0.5, y: 1.45, w: 6, h: 0.4,
    fontSize: 14, fontFace: FONT_H, bold: true, color: C.primary,
    valign: "middle", margin: 0, charSpacing: 4,
  });
  // Timeline vertical line
  s.addShape(pres.shapes.RECTANGLE, {
    x: 1.3, y: 1.95, w: 0.04, h: 4.8,
    fill: { color: C.secondary }, line: { color: C.secondary },
  });

  day.schedule.forEach((item, i) => {
    const yPos = 1.95 + i * 0.72;
    // Time dot
    s.addShape(pres.shapes.OVAL, {
      x: 1.18, y: yPos + 0.08, w: 0.28, h: 0.28,
      fill: { color: C.primary }, line: { color: C.white, width: 2 },
    });
    // Time
    s.addText(item.time, {
      x: 0.5, y: yPos, w: 0.65, h: 0.42,
      fontSize: 12, fontFace: FONT_B, bold: true, color: C.primary,
      align: "right", valign: "middle", margin: 0,
    });
    // Place
    s.addText(item.place, {
      x: 1.6, y: yPos, w: 5.0, h: 0.4,
      fontSize: 13, fontFace: FONT_B, bold: true, color: C.dark,
      valign: "middle", margin: 0,
    });
    // Note
    if (item.note) {
      s.addText(item.note, {
        x: 1.6, y: yPos + 0.32, w: 5.0, h: 0.3,
        fontSize: 10, fontFace: FONT_B, italic: true,
        color: item.note.startsWith("★") ? C.primary : C.muted,
        valign: "middle", margin: 0,
      });
    }
  });

  // Right column: map + highlights
  // Map card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.5, y: 1.45, w: 5.3, h: 2.6,
    fill: { color: C.white }, line: { color: "DDDDDD", width: 0.5 },
    shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 90, opacity: 0.1 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.5, y: 1.45, w: 0.08, h: 2.6,
    fill: { color: C.primary }, line: { color: C.primary },
  });
  s.addText("🗺  동선 (Google Maps)", {
    x: 7.75, y: 1.55, w: 5.0, h: 0.4,
    fontSize: 14, fontFace: FONT_H, bold: true, color: C.primary,
    valign: "middle", margin: 0,
  });
  s.addText(day.mapUrl, {
    x: 7.75, y: 2.05, w: 5.0, h: 1.5,
    fontSize: 13, fontFace: FONT_B, color: C.dark,
    valign: "top", margin: 0,
  });
  s.addText("👉 MD 파일의 링크로 바로 열기", {
    x: 7.75, y: 3.55, w: 5.0, h: 0.35,
    fontSize: 10, fontFace: FONT_B, italic: true, color: C.muted,
    valign: "middle", margin: 0,
  });

  // Highlights card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.5, y: 4.25, w: 5.3, h: 2.5,
    fill: { color: C.white }, line: { color: "DDDDDD", width: 0.5 },
    shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 90, opacity: 0.1 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.5, y: 4.25, w: 0.08, h: 2.5,
    fill: { color: C.gold }, line: { color: C.gold },
  });
  s.addText("💡  포인트 / 유의사항", {
    x: 7.75, y: 4.35, w: 5.0, h: 0.4,
    fontSize: 14, fontFace: FONT_H, bold: true, color: C.gold,
    valign: "middle", margin: 0,
  });
  s.addText(
    day.highlights.map((h, i) => ({
      text: h,
      options: { bullet: { code: "25A0" }, breakLine: i < day.highlights.length - 1, color: C.dark },
    })),
    {
      x: 7.85, y: 4.85, w: 4.9, h: 1.8,
      fontSize: 12, fontFace: FONT_B, valign: "top", margin: 0, paraSpaceAfter: 6,
    }
  );

  addFooter(s, 3 + idx, TOTAL);
});

// =====================================================
// Slide 8: Flight & Transport
// =====================================================
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  addHeader(s, "교통", "항공편 & 현지 교통 패스");

  // Flight section title
  s.addText("✈  항공편 (인천 ↔ 간사이 직항)", {
    x: 0.5, y: 1.5, w: 12, h: 0.5,
    fontSize: 18, fontFace: FONT_H, bold: true, color: C.primary,
    valign: "middle", margin: 0,
  });

  const flights = [
    [
      { text: "구분", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center", valign: "middle" } },
      { text: "일시", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center", valign: "middle" } },
      { text: "항공편 (예시)", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center", valign: "middle" } },
      { text: "비행 시간", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center", valign: "middle" } },
    ],
    ["출국", "07/10 (금) 09:30 → 11:30", "KE723 / OZ112", "약 1시간 50분"],
    ["귀국", "07/14 (화) 17:30 → 19:30", "KE724 / OZ113", "약 1시간 50분"],
  ];
  s.addTable(flights, {
    x: 0.5, y: 2.1, w: 12.3, colW: [1.5, 3.8, 4.0, 3.0],
    fontSize: 13, fontFace: FONT_B, color: C.dark,
    border: { type: "solid", pt: 1, color: "DDDDDD" },
    rowH: 0.5, align: "center", valign: "middle",
  });

  // Transport pass
  s.addText("🚆  교통 패스 & 현지 이동", {
    x: 0.5, y: 4.3, w: 12, h: 0.5,
    fontSize: 18, fontFace: FONT_H, bold: true, color: C.primary,
    valign: "middle", margin: 0,
  });

  const passes = [
    { title: "JR 간사이 와이드 패스 5일권", desc: "하루카·우지·오사카·후시미 모두 커버", price: "약 ¥12,000" },
    { title: "MK 점보택시 (8인승)", desc: "교토 시내 이동 시 2대 분산 추천", price: "시간제 ¥6,000~/h" },
    { title: "ICOCA 카드 / 모바일 스이카", desc: "지하철·버스·편의점 통합 결제", price: "충전식" },
  ];

  passes.forEach((p, i) => {
    const x = 0.5 + i * 4.27;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 4.9, w: 4.1, h: 1.7,
      fill: { color: C.white }, line: { color: "DDDDDD", width: 0.5 },
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 90, opacity: 0.1 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 4.9, w: 4.1, h: 0.08,
      fill: { color: C.gold }, line: { color: C.gold },
    });
    s.addText(p.title, {
      x: x + 0.2, y: 5.05, w: 3.7, h: 0.45,
      fontSize: 13, fontFace: FONT_H, bold: true, color: C.dark, valign: "middle", margin: 0,
    });
    s.addText(p.desc, {
      x: x + 0.2, y: 5.5, w: 3.7, h: 0.6,
      fontSize: 11, fontFace: FONT_B, color: C.muted, valign: "top", margin: 0,
    });
    s.addText(p.price, {
      x: x + 0.2, y: 6.15, w: 3.7, h: 0.4,
      fontSize: 12, fontFace: FONT_B, bold: true, color: C.primary, valign: "middle", margin: 0,
    });
  });

  addFooter(s, 8, TOTAL);
}

// =====================================================
// Slide 9: Reservation list
// =====================================================
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  addHeader(s, "예약", "식당 사전예약 리스트 (필수)");

  s.addText("8인 단체석은 반드시 사전예약 — 카이세키/유도후/와규 식당은 1~2주 전 확정", {
    x: 0.5, y: 1.45, w: 12.3, h: 0.5,
    fontSize: 13, fontFace: FONT_B, italic: true, color: C.muted,
    valign: "middle", margin: 0,
  });

  const reservations = [
    [
      { text: "일자", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center", valign: "middle" } },
      { text: "시간", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center", valign: "middle" } },
      { text: "식당 / 메뉴", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center", valign: "middle" } },
      { text: "예약 시점", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center", valign: "middle" } },
      { text: "비고", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center", valign: "middle" } },
    ],
    ["Day 1", "19:00", "폰토초 가이세키", "출발 2주 전", "강변 가와도코"],
    ["Day 2", "12:30", "오쿠탄(奥丹) 유도후", "출발 1주 전", "전통 두부 코스"],
    ["Day 2", "19:00", "교토 와규 데판야키", "출발 1주 전", "예산 ¥15,000~"],
    ["Day 3", "12:00", "아라시야마 가이세키", "출발 1주 전", "정원 뷰"],
    ["Day 4", "13:00", "우지 말차 가이세키", "출발 1주 전", "디저트 포함"],
    ["Day 4", "19:00", "도톤보리 자유식", "—", "쿠시카츠·타코야키"],
  ];
  s.addTable(reservations, {
    x: 0.5, y: 2.1, w: 12.3, colW: [1.2, 1.2, 4.3, 2.6, 3.0],
    fontSize: 12, fontFace: FONT_B, color: C.dark,
    border: { type: "solid", pt: 1, color: "DDDDDD" },
    rowH: 0.55, align: "center", valign: "middle",
  });

  // Bottom tip box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 6.5, w: 12.3, h: 0.45,
    fill: { color: C.secondary }, line: { color: C.secondary },
  });
  s.addText("💡 카이세키 식당 대부분 한국어 예약 불가 → 호텔 컨시어지 또는 OMAKASE 사이트 이용", {
    x: 0.5, y: 6.5, w: 12.3, h: 0.45,
    fontSize: 12, fontFace: FONT_B, bold: true, color: C.dark,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(s, 9, TOTAL);
}

// =====================================================
// Slide 10: Preparation
// =====================================================
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  addHeader(s, "준비", "7월 중순 일본 — 준비물 체크리스트");

  const categories = [
    {
      title: "날씨 대비",
      items: ["우산 / 우비 (장마 끝물)", "휴대용 선풍기·쿨토시", "양산·모자 (자외선)", "휴대용 물병"],
      color: C.primary,
    },
    {
      title: "복장 / 신발",
      items: ["통기성 좋은 의류", "편한 운동화 (1.5만보+)", "료칸용 양말·실내복", "겉옷 1벌 (실내 냉방)"],
      color: C.gold,
    },
    {
      title: "결제 / 통신",
      items: ["ICOCA 또는 모바일 스이카", "엔화 현금 (¥3만 권장)", "신용카드 (해외 사용 가능)", "포켓 와이파이 또는 eSIM"],
      color: C.accent,
    },
    {
      title: "서류 / 기타",
      items: ["여권 (잔여 6개월+)", "여행자 보험 증서", "전자 비자 신청 (Visit Japan Web)", "비상약·자외선 차단제"],
      color: C.primary,
    },
  ];

  const cardW = 2.95, cardH = 4.6, gap = 0.2;
  const startX = (SLIDE_W - (cardW * 4 + gap * 3)) / 2;
  const startY = 1.7;

  categories.forEach((cat, i) => {
    const x = startX + i * (cardW + gap);
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: startY, w: cardW, h: cardH,
      fill: { color: C.white }, line: { color: "DDDDDD", width: 0.5 },
      shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 90, opacity: 0.12 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: startY, w: cardW, h: 0.7,
      fill: { color: cat.color }, line: { color: cat.color },
    });
    s.addText(cat.title, {
      x, y: startY, w: cardW, h: 0.7,
      fontSize: 16, fontFace: FONT_H, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(
      cat.items.map((item, j) => ({
        text: item,
        options: { bullet: { code: "2713" }, breakLine: j < cat.items.length - 1, color: C.dark },
      })),
      {
        x: x + 0.2, y: startY + 0.9, w: cardW - 0.4, h: cardH - 1.1,
        fontSize: 12, fontFace: FONT_B, valign: "top", margin: 0, paraSpaceAfter: 10,
      }
    );
  });

  addFooter(s, 10, TOTAL);
}

// =====================================================
// Slide 11: Budget
// =====================================================
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  addHeader(s, "예산", "1인 예상 경비 (참고)");

  // Pie chart
  s.addChart(pres.charts.DOUGHNUT, [{
    name: "Budget",
    labels: ["항공권", "숙박 (4박)", "식사", "JR 패스", "입장료·기타"],
    values: [525, 550, 400, 120, 100],
  }], {
    x: 0.5, y: 1.6, w: 5.5, h: 5.3,
    chartColors: [C.primary, C.gold, C.accent, "8B7355", C.muted],
    showLegend: true, legendPos: "b", legendFontSize: 11, legendFontFace: FONT_B,
    showTitle: true, title: "비용 구성 (단위: 천원)",
    titleFontSize: 14, titleFontFace: FONT_H, titleColor: C.dark,
    chartArea: { fill: { color: C.white }, roundedCorners: true },
    showPercent: true, dataLabelColor: C.white, dataLabelFontBold: true,
  });

  // Right: detailed table
  const budgetTable = [
    [
      { text: "항목", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center" } },
      { text: "1인 (₩)", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center" } },
      { text: "8인 합계 (₩)", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center" } },
    ],
    ["항공권", "525,000", "4,200,000"],
    ["숙박 (4박)", "550,000", "4,400,000"],
    ["식사 (5일)", "400,000", "3,200,000"],
    ["JR 패스", "120,000", "960,000"],
    ["입장료·기타", "100,000", "800,000"],
    [
      { text: "합계 (예상)", options: { bold: true, fill: { color: C.secondary } } },
      { text: "1,695,000", options: { bold: true, color: C.primary, fill: { color: C.secondary } } },
      { text: "13,560,000", options: { bold: true, color: C.primary, fill: { color: C.secondary } } },
    ],
  ];
  s.addTable(budgetTable, {
    x: 6.5, y: 1.7, w: 6.3, colW: [2.3, 2.0, 2.0],
    fontSize: 12, fontFace: FONT_B, color: C.dark,
    border: { type: "solid", pt: 1, color: "DDDDDD" },
    rowH: 0.55, align: "center", valign: "middle",
  });
  s.addText("* 항공권은 6월 초 예약 기준 / 숙박은 4성급 트윈 기준", {
    x: 6.5, y: 6.45, w: 6.3, h: 0.4,
    fontSize: 10, fontFace: FONT_B, italic: true, color: C.muted,
    valign: "middle", margin: 0,
  });

  addFooter(s, 11, TOTAL);
}

// =====================================================
// Slide 12: Map links summary
// =====================================================
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  addHeader(s, "지도", "일자별 Google Maps 동선 링크");

  const links = [
    { day: "Day 1", route: "간사이공항 → 교토역 → 기온 → 폰토초", url: "https://www.google.com/maps/dir/Kansai+International+Airport/Kyoto+Station/Gion/Pontocho" },
    { day: "Day 2", route: "기요미즈데라 → 산넨자카 → 고다이지 → 야사카신사 → 시라카와", url: "https://www.google.com/maps/dir/Kiyomizu-dera/Sannenzaka/Kodai-ji/Yasaka+Shrine/Shirakawa+Gion" },
    { day: "Day 3", route: "금각사 → 료안지 → 아라시야마 → 도게쓰교 → 텐류지 → 니시키시장", url: "https://www.google.com/maps/dir/Kinkaku-ji/Ryoan-ji/Arashiyama+Bamboo+Grove/Togetsukyo/Tenryu-ji/Nishiki+Market" },
    { day: "Day 4", route: "후시미 이나리 → 뵤도인 → 오사카성 → 도톤보리", url: "https://www.google.com/maps/dir/Fushimi+Inari+Taisha/Byodo-in/Osaka+Castle/Dotonbori" },
    { day: "Day 5", route: "우메다 공중정원 → 구로몬시장 → 난바 → 간사이공항", url: "https://www.google.com/maps/dir/Umeda+Sky+Building/Kuromon+Ichiba/Namba+Station/Kansai+International+Airport" },
  ];

  links.forEach((l, i) => {
    const y = 1.7 + i * 1.0;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 12.3, h: 0.85,
      fill: { color: C.white }, line: { color: "DDDDDD", width: 0.5 },
      shadow: { type: "outer", color: "000000", blur: 5, offset: 1, angle: 90, opacity: 0.08 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 0.08, h: 0.85,
      fill: { color: C.primary }, line: { color: C.primary },
    });
    // Day badge
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y: y + 0.18, w: 1.0, h: 0.5,
      fill: { color: C.primary }, line: { color: C.primary },
    });
    s.addText(l.day, {
      x: 0.8, y: y + 0.18, w: 1.0, h: 0.5,
      fontSize: 14, fontFace: FONT_H, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 0,
    });
    // Route
    s.addText(l.route, {
      x: 2.0, y: y + 0.05, w: 10.5, h: 0.4,
      fontSize: 13, fontFace: FONT_B, bold: true, color: C.dark,
      valign: "middle", margin: 0,
    });
    // Link
    s.addText([
      { text: "🗺  ", options: { color: C.primary } },
      { text: "Google Maps에서 열기", options: { color: C.primary, underline: { style: "sng" }, hyperlink: { url: l.url } } },
    ], {
      x: 2.0, y: y + 0.43, w: 10.5, h: 0.35,
      fontSize: 11, fontFace: FONT_B, italic: true, valign: "middle", margin: 0,
    });
  });

  addFooter(s, 12, TOTAL);
}

// =====================================================
// Slide 13: Closing
// =====================================================
{
  const s = pres.addSlide();
  s.background = { color: C.dark };

  // Side bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: SLIDE_W, h: 0.15,
    fill: { color: C.gold }, line: { color: C.gold },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: SLIDE_H - 0.15, w: SLIDE_W, h: 0.15,
    fill: { color: C.gold }, line: { color: C.gold },
  });

  s.addText("ありがとうございました", {
    x: 0.5, y: 1.8, w: 12.3, h: 1.0,
    fontSize: 42, fontFace: FONT_H, italic: true, color: C.gold,
    align: "center", valign: "middle", margin: 0,
  });
  s.addText("좋은 여행 되세요", {
    x: 0.5, y: 2.9, w: 12.3, h: 1.2,
    fontSize: 60, fontFace: FONT_H, bold: true, color: C.white,
    align: "center", valign: "middle", margin: 0,
  });

  // Divider
  s.addShape(pres.shapes.RECTANGLE, {
    x: SLIDE_W / 2 - 1.0, y: 4.3, w: 2.0, h: 0.04,
    fill: { color: C.primary }, line: { color: C.primary },
  });

  s.addText("교토 3일 · 오사카 1일 · 8명의 특별한 5일", {
    x: 0.5, y: 4.5, w: 12.3, h: 0.5,
    fontSize: 18, fontFace: FONT_H, italic: true, color: C.secondary,
    align: "center", valign: "middle", margin: 0,
  });
  s.addText("2026. 07. 10 — 07. 14", {
    x: 0.5, y: 5.1, w: 12.3, h: 0.5,
    fontSize: 16, fontFace: FONT_B, charSpacing: 6, color: C.gold,
    align: "center", valign: "middle", margin: 0,
  });

  s.addText("문의 / 일정 변경 요청은 언제든지", {
    x: 0.5, y: 6.4, w: 12.3, h: 0.4,
    fontSize: 12, fontFace: FONT_B, italic: true, color: C.muted,
    align: "center", valign: "middle", margin: 0,
  });
}

// === Write ===
pres.writeFile({ fileName: "c:/Users/ohsew/Desktop/huhsame-trip/교토-오사카-여행일정.pptx" })
  .then((fn) => console.log("Generated:", fn));
