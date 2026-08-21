const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.layout = "LAYOUT_WIDE"; // 13.3 x 7.5

// ---- palette (dark, code/GPU theme) ----
const BG     = "13161C"; // page background
const CARD   = "1E232C"; // surface
const CODEBG = "0E1116"; // code panel
const TEXT   = "E8EDF2"; // primary text
const MUTED  = "9AA6B2"; // secondary text
const GREEN  = "8CC63F"; // accent (CUDA-ish)
const TEAL   = "4FB0C6"; // secondary accent
const AMBER  = "E0A458"; // warm accent for warnings
const LINE   = "2C3440"; // hairline

const KFONT = "Malgun Gothic"; // Korean-capable
const MONO  = "Courier New";

const W = 13.3, H = 7.5, M = 0.6;

function bg(slide, color) { slide.background = { color: color || BG }; }

// section-number chip + title (repeated motif, no accent stripes)
function header(slide, num, title, color) {
  slide.addText(num, {
    x: M, y: 0.45, w: 0.7, h: 0.7, align: "center", valign: "middle",
    fontFace: MONO, fontSize: 22, bold: true, color: "13161C",
    fill: { color: color || GREEN }, rectRadius: 0.09, shape: p.ShapeType.roundRect,
  });
  slide.addText(title, {
    x: 1.45, y: 0.42, w: W - 1.45 - M, h: 0.76, align: "left", valign: "middle",
    fontFace: KFONT, fontSize: 30, bold: true, color: TEXT, margin: 0,
  });
}

function codePanel(slide, x, y, w, h, lines, opts) {
  opts = opts || {};
  slide.addShape(p.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06, fill: { color: CODEBG }, line: { color: LINE, width: 1 },
  });
  slide.addText(lines, {
    x: x + 0.22, y: y + 0.16, w: w - 0.44, h: h - 0.32, align: "left", valign: "top",
    fontFace: MONO, fontSize: opts.fontSize || 13, color: opts.color || TEXT, margin: 0,
    lineSpacingMultiple: 1.08,
  });
}

// build code as rich runs for simple highlighting
function ln(text, color, opts) {
  return Object.assign({ text, options: Object.assign({ color: color || TEXT, breakLine: true }, opts || {}) });
}

// =====================================================================
// Slide 1 — Title (dark)
// =====================================================================
(() => {
  const s = p.addSlide(); bg(s);
  s.addText("CUDA 세미나 · 세션 1", {
    x: M, y: 2.05, w: 8, h: 0.5, fontFace: KFONT, fontSize: 18, bold: true, color: GREEN, margin: 0,
  });
  s.addText("CUDA 스레드 모델과\n첫 번째 커널(kernel)", {
    x: M, y: 2.5, w: 11.5, h: 2.0, fontFace: KFONT, fontSize: 46, bold: true, color: TEXT,
    lineSpacingMultiple: 1.05, margin: 0,
  });
  s.addText("cuda_memtest 실제 코드로 배우는 GPU 프로그래밍", {
    x: M, y: 4.7, w: 11, h: 0.5, fontFace: KFONT, fontSize: 20, color: MUTED, margin: 0,
  });
  // decorative kernel signature as motif (not a stripe)
  codePanel(s, M, 5.6, 7.6, 1.0, [
    ln("__global__ void kernel_move_inv_write(", GREEN),
    ln("        char* _ptr, char* end_ptr, unsigned int pattern);", TEAL),
  ], { fontSize: 12.5 });
  s.addText("입문 과정 · C를 아는 개발자 대상", {
    x: 8.7, y: 6.05, w: 4.0, h: 0.4, align: "right", fontFace: KFONT, fontSize: 13, color: MUTED, margin: 0,
  });
  s.addNotes("세션 1 도입. 오늘은 CUDA의 가장 기본인 커널과 스레드 모델을 cuda_memtest의 실제 커널로 배웁니다. 장난감 예제가 아니라 실무 코드라는 점을 강조하세요.");
})();

// =====================================================================
// Slide 2 — Learning objectives (icon rows)
// =====================================================================
(() => {
  const s = p.addSlide(); bg(s);
  header(s, "1", "오늘의 목표");
  const items = [
    ["커널이란 무엇인가", "__global__ 함수가 일반 C 함수와 어떻게 다른지 설명할 수 있다"],
    ["실행 구성 <<<grid, block>>>", "그리드(grid)·블록(block)·스레드(thread)의 3단 계층을 이해한다"],
    ["스레드는 자기 데이터를 어떻게 찾나", "blockIdx·threadIdx·blockDim·gridDim의 역할을 안다"],
    ["첫 커널 완독", "kernel_move_inv_write를 한 줄씩 설명할 수 있다"],
  ];
  let y = 1.7;
  items.forEach((it, i) => {
    s.addShape(p.ShapeType.roundRect, { x: M, y, w: W - 2*M, h: 1.06, rectRadius: 0.08, fill: { color: CARD }, line: { type: "none" } });
    s.addText(String(i + 1), { x: M + 0.22, y: y + 0.2, w: 0.66, h: 0.66, align: "center", valign: "middle",
      fontFace: MONO, fontSize: 22, bold: true, color: BG, fill: { color: i % 2 ? TEAL : GREEN }, shape: p.ShapeType.roundRect, rectRadius: 0.33 });
    s.addText(it[0], { x: M + 1.15, y: y + 0.14, w: W - 2*M - 1.4, h: 0.44, fontFace: KFONT, fontSize: 19, bold: true, color: TEXT, margin: 0, valign: "middle" });
    s.addText(it[1], { x: M + 1.15, y: y + 0.55, w: W - 2*M - 1.4, h: 0.42, fontFace: KFONT, fontSize: 14, color: MUTED, margin: 0, valign: "middle" });
    y += 1.24;
  });
  s.addNotes("네 가지 목표를 미리 보여주고, 세션 끝에 다시 돌아와 자가 점검하겠다고 예고하세요.");
})();

// =====================================================================
// Slide 3 — CPU vs GPU recap (two column compare)
// =====================================================================
(() => {
  const s = p.addSlide(); bg(s);
  header(s, "2", "복습 · GPU는 왜 다른가");
  const colW = (W - 2*M - 0.4) / 2;
  const mkCol = (x, title, color, rows) => {
    s.addShape(p.ShapeType.roundRect, { x, y: 1.7, w: colW, h: 4.7, rectRadius: 0.08, fill: { color: CARD }, line: { type: "none" } });
    s.addText(title, { x: x + 0.3, y: 1.95, w: colW - 0.6, h: 0.6, fontFace: KFONT, fontSize: 22, bold: true, color, margin: 0 });
    s.addText(rows.map((r, i) => ln(r, TEXT, { breakLine: true, bullet: { code: "2022" }, paraSpaceAfter: 10 })), {
      x: x + 0.3, y: 2.65, w: colW - 0.6, h: 3.5, fontFace: KFONT, fontSize: 15.5, color: TEXT, valign: "top", margin: 0,
    });
  };
  mkCol(M, "CPU — 지연시간(latency) 최적화", TEAL, [
    "소수의 강력한 코어 (수 개~수십 개)",
    "복잡한 분기·순차 작업에 강함",
    "코어 하나가 한 작업을 빠르게",
    "큰 캐시, 정교한 예측",
  ]);
  mkCol(M + colW + 0.4, "GPU — 처리량(throughput) 최적화", GREEN, [
    "수천 개의 단순한 코어",
    "같은 연산을 대량 데이터에 동시에",
    "SIMT — Single Instruction, Multiple Threads",
    "메모리 대역폭이 핵심 (수백 GB/s)",
  ]);
  s.addText("cuda_memtest는 4 GB 메모리를 검사합니다 — 수천 스레드가 동시에 일해야 현실적인 시간에 끝납니다.", {
    x: M, y: 6.55, w: W - 2*M, h: 0.5, fontFace: KFONT, fontSize: 14, italic: true, color: MUTED, align: "center", margin: 0,
  });
  s.addNotes("SIMT 개념만 각인시키면 됩니다. GPU는 '같은 명령을 수천 스레드가 서로 다른 데이터에' 실행한다.");
})();

// =====================================================================
// Slide 4 — What is a kernel (concept + host/device flow)
// =====================================================================
(() => {
  const s = p.addSlide(); bg(s);
  header(s, "3", "커널(kernel)이란?");
  s.addText([
    ln("커널", GREEN, { bold: true, breakLine: false }),
    ln("은 GPU에서 실행되는 함수입니다. CPU(호스트)가 ", TEXT, { breakLine: false }),
    ln("호출", TEAL, { bold: true, breakLine: false }),
    ln("하고, GPU(디바이스)의 수많은 스레드가 ", TEXT, { breakLine: false }),
    ln("동시에", TEAL, { bold: true, breakLine: false }),
    ln(" 실행합니다.", TEXT, { breakLine: true }),
  ], { x: M, y: 1.65, w: W - 2*M, h: 0.8, fontFace: KFONT, fontSize: 18, color: TEXT, margin: 0, valign: "top" });

  codePanel(s, M, 2.55, W - 2*M, 1.5, [
    ln("__global__", GREEN, { breakLine: false }),
    ln(" void kernel_move_inv_write(char* _ptr, char* end_ptr, unsigned int pattern)", TEXT, { breakLine: true }),
    ln("//  ▲ 한정자        ▲ 반환형은 항상 void      ▲ 인자는 값 또는 디바이스 포인터", MUTED, { breakLine: true }),
  ], { fontSize: 13.5 });

  // three cards: __global__ / __device__ / __host__
  const cards = [
    ["__global__", GREEN, "CPU가 호출 → GPU가 실행\n= 커널. <<<>>>로 부른다"],
    ["__device__", TEAL, "GPU에서 GPU만 호출\n커널이 쓰는 보조 함수"],
    ["__host__", MUTED, "일반 CPU 함수 (기본값)\n생략하면 host"],
  ];
  const cw = (W - 2*M - 0.8) / 3;
  cards.forEach((c, i) => {
    const x = M + i * (cw + 0.4);
    s.addShape(p.ShapeType.roundRect, { x, y: 4.35, w: cw, h: 2.0, rectRadius: 0.08, fill: { color: CARD }, line: { type: "none" } });
    s.addText(c[0], { x: x + 0.25, y: 4.55, w: cw - 0.5, h: 0.5, fontFace: MONO, fontSize: 18, bold: true, color: c[1], margin: 0 });
    s.addText(c[2], { x: x + 0.25, y: 5.1, w: cw - 0.5, h: 1.1, fontFace: KFONT, fontSize: 14, color: TEXT, margin: 0, valign: "top", lineSpacingMultiple: 1.1 });
  });
  s.addNotes("반환형이 항상 void라는 점, 결과는 인자로 받은 포인터를 통해 메모리에 남긴다는 점을 강조하세요. tests.cpp:245.");
})();

// =====================================================================
// Slide 5 — Execution configuration <<<grid, block>>> (hierarchy diagram)
// =====================================================================
(() => {
  const s = p.addSlide(); bg(s);
  header(s, "4", "실행 구성 <<<grid, block>>>");
  s.addText("커널을 부를 때 '스레드를 몇 개, 어떻게 묶어' 실행할지 지정합니다.", {
    x: M, y: 1.6, w: W - 2*M, h: 0.5, fontFace: KFONT, fontSize: 17, color: MUTED, margin: 0,
  });
  codePanel(s, M, 2.15, 7.0, 0.75, [
    ln("kernel_move_inv_write", TEXT, { breakLine: false }),
    ln("<<<128, 1>>>", GREEN, { bold: true, breakLine: false }),
    ln("(ptr, end_ptr, p1);", TEXT, { breakLine: true }),
  ], { fontSize: 14 });

  // grid diagram on right
  const gx = 8.2, gy = 2.15, gw = 4.5, gh = 4.6;
  s.addShape(p.ShapeType.roundRect, { x: gx, y: gy, w: gw, h: gh, rectRadius: 0.08, fill: { color: CARD }, line: { type: "none" } });
  s.addText("그리드(grid) = 블록 128개", { x: gx + 0.25, y: gy + 0.15, w: gw - 0.5, h: 0.4, fontFace: KFONT, fontSize: 14, bold: true, color: GREEN, margin: 0 });
  // draw a few blocks
  for (let r = 0; r < 3; r++) {
    for (let c = 0; c < 4; c++) {
      const bx = gx + 0.35 + c * 1.0, by = gy + 0.7 + r * 1.15;
      s.addShape(p.ShapeType.roundRect, { x: bx, y: by, w: 0.85, h: 0.95, rectRadius: 0.05, fill: { color: CODEBG }, line: { color: TEAL, width: 1 } });
      s.addText("블록", { x: bx, y: by + 0.08, w: 0.85, h: 0.3, align: "center", fontFace: KFONT, fontSize: 9, color: TEAL, margin: 0 });
      // single thread dot
      s.addShape(p.ShapeType.ovalType || "ellipse", { x: bx + 0.34, y: by + 0.42, w: 0.18, h: 0.18, fill: { color: GREEN }, line: { type: "none" } });
    }
  }
  s.addText("● = 스레드 1개  (블록당 1개)", { x: gx + 0.35, y: gy + gh - 0.5, w: gw - 0.7, h: 0.35, fontFace: KFONT, fontSize: 12, color: MUTED, margin: 0 });

  // left explanation
  s.addText([
    ln("첫 번째 숫자 = 블록 수", GREEN, { bold: true, breakLine: true, paraSpaceAfter: 4 }),
    ln("그리드 안에 블록이 몇 개인가", MUTED, { breakLine: true, paraSpaceAfter: 14 }),
    ln("두 번째 숫자 = 블록당 스레드 수", TEAL, { bold: true, breakLine: true, paraSpaceAfter: 4 }),
    ln("각 블록 안에 스레드가 몇 개인가", MUTED, { breakLine: true, paraSpaceAfter: 14 }),
    ln("총 스레드 = 128 × 1 = 128개", TEXT, { bold: true, breakLine: true }),
  ], { x: M, y: 3.2, w: 7.0, h: 3.2, fontFace: KFONT, fontSize: 16, color: TEXT, margin: 0, valign: "top" });
  s.addNotes("이 저장소의 Test 0~9는 블록당 스레드 1개(<<<grid,1>>>)라는 아주 단순한 구성을 씁니다. 왜 그런지는 뒤에서.");
})();

// =====================================================================
// Slide 6 — built-in variables (table/cards)
// =====================================================================
(() => {
  const s = p.addSlide(); bg(s);
  header(s, "5", "내장 변수 · 스레드의 신분증");
  s.addText("커널 안에서 자동으로 주어지는 값들. 이것으로 '내가 몇 번째 스레드인지'를 계산합니다.", {
    x: M, y: 1.6, w: W - 2*M, h: 0.5, fontFace: KFONT, fontSize: 16, color: MUTED, margin: 0,
  });
  const rows = [
    ["blockIdx", "지금 이 블록의 번호 (그리드 안에서)", GREEN],
    ["threadIdx", "지금 이 스레드의 번호 (블록 안에서)", TEAL],
    ["blockDim", "블록당 스레드 수 (블록 크기)", TEAL],
    ["gridDim", "그리드의 블록 수 (그리드 크기)", GREEN],
  ];
  let y = 2.3;
  rows.forEach((r) => {
    s.addShape(p.ShapeType.roundRect, { x: M, y, w: W - 2*M, h: 0.92, rectRadius: 0.07, fill: { color: CARD }, line: { type: "none" } });
    s.addText(r[0], { x: M + 0.3, y: y + 0.12, w: 3.0, h: 0.68, valign: "middle", fontFace: MONO, fontSize: 20, bold: true, color: r[2], margin: 0 });
    s.addText(r[1], { x: M + 3.5, y: y + 0.12, w: W - 2*M - 3.8, h: 0.68, valign: "middle", fontFace: KFONT, fontSize: 16, color: TEXT, margin: 0 });
    y += 1.06;
  });
  s.addText("이 저장소는 .x 만 씁니다 (1차원). 실전에선 .x/.y/.z 로 2·3차원도 가능합니다.", {
    x: M, y: 6.75, w: W - 2*M, h: 0.4, fontFace: KFONT, fontSize: 13, italic: true, color: MUTED, margin: 0,
  });
  s.addNotes("blockIdx/threadIdx는 '읽기 전용, 스레드마다 다름', blockDim/gridDim은 '모든 스레드가 같은 값'이라는 대비를 짚어주세요.");
})();

// =====================================================================
// Slide 7 — repo convention: blockIdx.x * BLOCKSIZE
// =====================================================================
(() => {
  const s = p.addSlide(); bg(s);
  header(s, "6", "이 저장소의 관례 · 블록 1개 = 1 MB");
  codePanel(s, M, 1.7, W - 2*M, 1.35, [
    ln("#define BLOCKSIZE ((unsigned long)(1024*1024))   // = 1 MB", MUTED, { breakLine: true }),
    ln("unsigned int* ptr = (unsigned int*)(_ptr + ", TEXT, { breakLine: false }),
    ln("blockIdx.x * BLOCKSIZE", GREEN, { bold: true, breakLine: false }),
    ln(");", TEXT, { breakLine: true }),
  ], { fontSize: 14 });

  // memory strip diagram
  const mx = M, my = 3.4, mw = W - 2*M, mh = 1.0;
  const nseg = 6;
  const segw = mw / nseg;
  for (let i = 0; i < nseg; i++) {
    s.addShape(p.ShapeType.roundRect, { x: mx + i * segw + 0.03, y: my, w: segw - 0.06, h: mh, rectRadius: 0.04,
      fill: { color: i < nseg - 1 ? CODEBG : "1A1414" }, line: { color: i % 2 ? TEAL : GREEN, width: 1 } });
    const label = i < nseg - 1 ? `블록 ${i}` : "· · ·";
    s.addText(label, { x: mx + i * segw, y: my + 0.12, w: segw, h: 0.34, align: "center", fontFace: KFONT, fontSize: 12, bold: true, color: TEXT, margin: 0 });
    s.addText(i < nseg - 1 ? `${i} MB~` : "", { x: mx + i * segw, y: my + 0.5, w: segw, h: 0.34, align: "center", fontFace: MONO, fontSize: 11, color: MUTED, margin: 0 });
  }
  s.addText("전체 GPU 메모리", { x: mx, y: my + mh + 0.12, w: mw, h: 0.35, align: "center", fontFace: KFONT, fontSize: 12, italic: true, color: MUTED, margin: 0 });

  s.addText([
    ln("블록 번호(blockIdx.x)에 1 MB를 곱해 자기 구역을 정합니다. ", TEXT, { breakLine: false }),
    ln("블록 0은 0~1 MB, 블록 1은 1~2 MB … 서로 겹치지 않습니다.", TEXT, { breakLine: true }),
  ], { x: M, y: 5.5, w: W - 2*M, h: 1.0, fontFace: KFONT, fontSize: 17, color: TEXT, margin: 0, valign: "top" });
  s.addNotes("핵심 통찰: 스레드/블록이 겹치지 않는 자기 구역을 갖는 것이 병렬 프로그래밍의 기본. cuda_memtest.h:91.");
})();

// =====================================================================
// Slide 8 — full first kernel read (annotated code)
// =====================================================================
(() => {
  const s = p.addSlide(); bg(s);
  header(s, "7", "첫 커널 완독 · kernel_move_inv_write");
  codePanel(s, M, 1.7, W - 2*M, 4.15, [
    ln("__global__ void", GREEN, { breakLine: true }),
    ln("kernel_move_inv_write(char* _ptr, char* end_ptr, unsigned int pattern) {", TEXT, { breakLine: true }),
    ln("    unsigned int i;", TEXT, { breakLine: true }),
    ln("    unsigned int* ptr = (unsigned int*)(_ptr + blockIdx.x*BLOCKSIZE);", TEXT, { breakLine: true }),
    ln("    if (ptr >= (unsigned int*) end_ptr) {", TEXT, { breakLine: true }),
    ln("        return;                       // 범위 밖이면 아무 것도 안 함(guard)", AMBER, { breakLine: true }),
    ln("    }", TEXT, { breakLine: true }),
    ln("    for (i = 0; i < BLOCKSIZE/sizeof(unsigned int); i++) {", TEXT, { breakLine: true }),
    ln("        ptr[i] = pattern;             // 1 MB를 4바이트씩 pattern으로 채움", GREEN, { breakLine: true }),
    ln("    }", TEXT, { breakLine: true }),
    ln("    return;", TEXT, { breakLine: true }),
    ln("}", TEXT, { breakLine: true }),
  ], { fontSize: 13.5 });
  s.addText("tests.cpp:245  —  블록당 스레드 1개라 threadIdx가 등장하지 않습니다. 스레드 하나가 1 MB 전체를 순회합니다.", {
    x: M, y: 6.05, w: W - 2*M, h: 0.7, fontFace: KFONT, fontSize: 15, color: MUTED, margin: 0, valign: "top",
  });
  s.addNotes("한 줄씩: 포인터 계산 → 경계 검사 → 루프로 채우기. sizeof로 나눠 4바이트 단위 인덱싱하는 부분을 짚어주세요.");
})();

// =====================================================================
// Slide 9 — why <<<grid, 1>>>? (design rationale)
// =====================================================================
(() => {
  const s = p.addSlide(); bg(s);
  header(s, "8", "왜 블록당 스레드 1개일까?");
  const colW = (W - 2*M - 0.4) / 2;
  s.addShape(p.ShapeType.roundRect, { x: M, y: 1.75, w: colW, h: 4.5, rectRadius: 0.08, fill: { color: CARD }, line: { type: "none" } });
  s.addText("의도된 단순함", { x: M + 0.3, y: 2.0, w: colW - 0.6, h: 0.5, fontFace: KFONT, fontSize: 20, bold: true, color: GREEN, margin: 0 });
  s.addText([
    ln("Test 0~9의 목적은 속도가 아니라 정확한 진단", TEXT, { breakLine: true, bullet: { code: "2022" }, paraSpaceAfter: 10 }),
    ln("주소 배선(address wire) 오류를 찾으려면 예측 가능한 접근 패턴이 유리", TEXT, { breakLine: true, bullet: { code: "2022" }, paraSpaceAfter: 10 }),
    ln("코드가 단순 = 버그가 적고 신뢰 가능", TEXT, { breakLine: true, bullet: { code: "2022" }, paraSpaceAfter: 10 }),
    ln("입문자가 읽기에 완벽", GREEN, { breakLine: true, bullet: { code: "2022" } }),
  ], { x: M + 0.3, y: 2.6, w: colW - 0.6, h: 3.4, fontFace: KFONT, fontSize: 15.5, color: TEXT, margin: 0, valign: "top" });

  const x2 = M + colW + 0.4;
  s.addShape(p.ShapeType.roundRect, { x: x2, y: 1.75, w: colW, h: 4.5, rectRadius: 0.08, fill: { color: CARD }, line: { type: "none" } });
  s.addText("하지만 느리다", { x: x2 + 0.3, y: 2.0, w: colW - 0.6, h: 0.5, fontFace: KFONT, fontSize: 20, bold: true, color: AMBER, margin: 0 });
  s.addText([
    ln("GPU의 수천 코어를 거의 놀림", TEXT, { breakLine: true, bullet: { code: "2022" }, paraSpaceAfter: 10 }),
    ln("블록당 스레드 1개 → 병렬성 최소", TEXT, { breakLine: true, bullet: { code: "2022" }, paraSpaceAfter: 10 }),
    ln("Test 10은 정반대로 설계됨:", TEXT, { breakLine: true, bullet: { code: "2022" }, paraSpaceAfter: 6 }),
    ln("<<<32768, 64>>>  — 대역폭 최적화", TEAL, { breakLine: true, paraSpaceAfter: 10 }),
    ln("→ 세션 3에서 두 방식을 직접 비교합니다", GREEN, { breakLine: true, italic: true }),
  ], { x: x2 + 0.3, y: 2.6, w: colW - 0.6, h: 3.4, fontFace: KFONT, fontSize: 15.5, color: TEXT, margin: 0, valign: "top" });
  s.addNotes("이 대비가 세미나 전체의 뼈대입니다. 같은 문제를 느린 방식과 빠른 방식으로 푸는 코드가 한 저장소에 있다는 점.");
})();

// =====================================================================
// Slide 10 — host launches the kernel (move_inv_test loop)
// =====================================================================
(() => {
  const s = p.addSlide(); bg(s);
  header(s, "9", "호스트가 커널을 부르는 법");
  codePanel(s, M, 1.7, W - 2*M, 3.0, [
    ln("// move_inv_test  (tests.cpp:312)  — 호스트(CPU) 함수", MUTED, { breakLine: true }),
    ln("for (i = 0; i < tot_num_blocks; i += GRIDSIZE) {", TEXT, { breakLine: true }),
    ln("    dim3 grid;  grid.x = GRIDSIZE;                 // GRIDSIZE = 128", TEXT, { breakLine: true }),
    ln("    kernel_move_inv_write", GREEN, { breakLine: false }),
    ln("<<<grid, 1>>>", TEAL, { bold: true, breakLine: false }),
    ln("(ptr + i*BLOCKSIZE, end_ptr, p1);", TEXT, { breakLine: true }),
    ln("}", TEXT, { breakLine: true }),
  ], { fontSize: 13.5 });
  s.addText([
    ln("한 번에 128 MB(블록 128개)를 처리하고, ", TEXT, { breakLine: false }),
    ln("바깥 for 루프", GREEN, { bold: true, breakLine: false }),
    ln("가 전체 메모리를 128 MB씩 훑습니다. 4 GB이면 약 32번 반복됩니다.", TEXT, { breakLine: true }),
  ], { x: M, y: 5.0, w: W - 2*M, h: 1.4, fontFace: KFONT, fontSize: 17, color: TEXT, margin: 0, valign: "top" });
  s.addNotes("커널 launch는 CPU 코드 안의 한 줄이라는 점, dim3로 그리드 차원을 지정한다는 점을 설명하세요.");
})();

// =====================================================================
// Slide 11 — Lab preview (4 exercise cards, 2x2)
// =====================================================================
(() => {
  const s = p.addSlide(); bg(s);
  header(s, "10", "실습 미리보기 · Lab 1");
  const labs = [
    ["Ex 1", "스레드는 자기가 누구인지 안다", "커널 안에서 blockIdx.x를 printf로 출력해 확인", GREEN],
    ["Ex 2", "그리드 크기를 바꿔보기", "GRIDSIZE를 32/128/512로 바꿔 실행 시간 측정", TEAL],
    ["Ex 3", "패턴을 내 마음대로", "--pattern 0xDEADBEEF 로 커널 인자 전달 관찰", GREEN],
    ["Ex 4", "경계 검사를 지우면?", "guard를 없애 illegal memory access 재현", AMBER],
  ];
  const cw = (W - 2*M - 0.5) / 2, ch = 2.15;
  labs.forEach((l, i) => {
    const x = M + (i % 2) * (cw + 0.5);
    const y = 1.7 + Math.floor(i / 2) * (ch + 0.35);
    s.addShape(p.ShapeType.roundRect, { x, y, w: cw, h: ch, rectRadius: 0.08, fill: { color: CARD }, line: { type: "none" } });
    s.addText(l[0], { x: x + 0.28, y: y + 0.22, w: 1.3, h: 0.55, align: "center", valign: "middle", fontFace: MONO, fontSize: 17, bold: true, color: BG, fill: { color: l[3] }, shape: p.ShapeType.roundRect, rectRadius: 0.08 });
    s.addText(l[1], { x: x + 1.75, y: y + 0.24, w: cw - 2.0, h: 0.9, fontFace: KFONT, fontSize: 17, bold: true, color: TEXT, margin: 0, valign: "top" });
    s.addText(l[2], { x: x + 0.28, y: y + 1.2, w: cw - 0.56, h: 0.8, fontFace: KFONT, fontSize: 14, color: MUTED, margin: 0, valign: "top", lineSpacingMultiple: 1.1 });
  });
  s.addNotes("실습 랩 문서(세션1_실습랩.md)를 나눠주고 진행합니다. Ex 4는 반드시 원복하도록 강조하세요.");
})();

// =====================================================================
// Slide 12 — recap / common mistakes
// =====================================================================
(() => {
  const s = p.addSlide(); bg(s);
  header(s, "11", "핵심 정리 & 자주 하는 실수");
  const colW = (W - 2*M - 0.4) / 2;
  s.addShape(p.ShapeType.roundRect, { x: M, y: 1.75, w: colW, h: 4.5, rectRadius: 0.08, fill: { color: CARD }, line: { type: "none" } });
  s.addText("기억할 것", { x: M + 0.3, y: 2.0, w: colW - 0.6, h: 0.5, fontFace: KFONT, fontSize: 20, bold: true, color: GREEN, margin: 0 });
  s.addText([
    ln("커널 = __global__, 반환형 void", TEXT, { breakLine: true, bullet: { code: "2713" }, paraSpaceAfter: 12 }),
    ln("<<<블록수, 블록당스레드수>>>", TEXT, { breakLine: true, bullet: { code: "2713" }, paraSpaceAfter: 12 }),
    ln("blockIdx로 자기 구역 계산", TEXT, { breakLine: true, bullet: { code: "2713" }, paraSpaceAfter: 12 }),
    ln("경계 검사(guard)는 필수", TEXT, { breakLine: true, bullet: { code: "2713" } }),
  ], { x: M + 0.3, y: 2.65, w: colW - 0.6, h: 3.4, fontFace: KFONT, fontSize: 16, color: TEXT, margin: 0, valign: "top" });

  const x2 = M + colW + 0.4;
  s.addShape(p.ShapeType.roundRect, { x: x2, y: 1.75, w: colW, h: 4.5, rectRadius: 0.08, fill: { color: CARD }, line: { type: "none" } });
  s.addText("입문자 함정", { x: x2 + 0.3, y: 2.0, w: colW - 0.6, h: 0.5, fontFace: KFONT, fontSize: 20, bold: true, color: AMBER, margin: 0 });
  s.addText([
    ln("커널을 일반 함수처럼 () 로 호출 → <<<>>> 필수", TEXT, { breakLine: true, bullet: { code: "2717" }, paraSpaceAfter: 12 }),
    ln("커널에서 값을 return 하려 함 → void, 결과는 메모리로", TEXT, { breakLine: true, bullet: { code: "2717" }, paraSpaceAfter: 12 }),
    ln("경계 검사 생략 → 이웃 메모리 침범", TEXT, { breakLine: true, bullet: { code: "2717" }, paraSpaceAfter: 12 }),
    ln("호스트 포인터를 커널에 그대로 전달 → 디바이스 포인터여야", TEXT, { breakLine: true, bullet: { code: "2717" } }),
  ], { x: x2 + 0.3, y: 2.65, w: colW - 0.6, h: 3.4, fontFace: KFONT, fontSize: 16, color: TEXT, margin: 0, valign: "top" });
  s.addNotes("네 가지 목표(슬라이드 2)로 돌아가 자가 점검을 유도하세요.");
})();

// =====================================================================
// Slide 13 — wrap up / next (dark)
// =====================================================================
(() => {
  const s = p.addSlide(); bg(s);
  s.addText("다음 세션 예고", { x: M, y: 1.6, w: 10, h: 0.5, fontFace: KFONT, fontSize: 18, bold: true, color: GREEN, margin: 0 });
  s.addText("세션 2 · 메모리 모델과 데이터 이동", { x: M, y: 2.15, w: 12, h: 1.0, fontFace: KFONT, fontSize: 38, bold: true, color: TEXT, margin: 0 });
  s.addText([
    ln("오늘 값을 쓴 커널이, 다음엔 그 값을 ", TEXT, { breakLine: false }),
    ln("읽어서 검사", TEAL, { bold: true, breakLine: false }),
    ln("합니다.", TEXT, { breakLine: true }),
  ], { x: M, y: 3.3, w: 12, h: 0.6, fontFace: KFONT, fontSize: 20, color: TEXT, margin: 0 });
  s.addText([
    ln("cudaMalloc · cudaMemcpy · cudaMemset", MONO, { breakLine: true, color: GREEN, paraSpaceAfter: 8 }),
    ln("왜 쓰기 커널을 끝내야 읽기 커널이 올바른 값을 볼까? (메모리 플러시)", MUTED, { breakLine: true }),
  ], { x: M, y: 4.1, w: 12, h: 1.2, fontFace: KFONT, fontSize: 16, color: TEXT, margin: 0, valign: "top" });

  codePanel(s, M, 5.5, 8.5, 1.1, [
    ln("kernel_move_inv_readwrite  →  읽어서 검사하고, 보수(complement)를 쓴다", TEAL),
    ln("kernel_move_inv_read       →  다시 읽어서 검사한다", GREEN),
  ], { fontSize: 12.5 });
  s.addText("실습 랩 1을 먼저 완료하고 오세요.", { x: M, y: 6.75, w: 12, h: 0.4, fontFace: KFONT, fontSize: 14, italic: true, color: MUTED, margin: 0 });
  s.addNotes("오늘의 write 커널이 다음 세션 read 커널로 이어진다는 연속성을 강조하며 마무리하세요.");
})();

p.writeFile({ fileName: "세션1_CUDA_스레드모델과_첫커널.pptx" }).then(f => console.log("wrote", f));
