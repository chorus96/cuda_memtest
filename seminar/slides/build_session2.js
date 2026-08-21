const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.layout = "LAYOUT_WIDE"; // 13.3 x 7.5

// ---- palette (shared with Session 1) ----
const BG="13161C", CARD="1E232C", CODEBG="0E1116", TEXT="E8EDF2", MUTED="9AA6B2";
const GREEN="8CC63F", TEAL="4FB0C6", AMBER="E0A458", RED="E0645B", LINE="2C3440";
const KFONT="Malgun Gothic", MONO="Courier New";
const W=13.3, H=7.5, M=0.6;

function bg(s,c){ s.background={color:c||BG}; }
function header(s,num,title,color){
  s.addText(num,{x:M,y:0.45,w:0.7,h:0.7,align:"center",valign:"middle",fontFace:MONO,fontSize:22,bold:true,color:"13161C",fill:{color:color||GREEN},rectRadius:0.09,shape:p.ShapeType.roundRect});
  s.addText(title,{x:1.45,y:0.42,w:W-1.45-M,h:0.76,align:"left",valign:"middle",fontFace:KFONT,fontSize:30,bold:true,color:TEXT,margin:0});
}
function codePanel(s,x,y,w,h,lines,opts){
  opts=opts||{};
  s.addShape(p.ShapeType.roundRect,{x,y,w,h,rectRadius:0.06,fill:{color:CODEBG},line:{color:LINE,width:1}});
  s.addText(lines,{x:x+0.22,y:y+0.16,w:w-0.44,h:h-0.32,align:"left",valign:"top",fontFace:MONO,fontSize:opts.fontSize||13,color:opts.color||TEXT,margin:0,lineSpacingMultiple:1.08});
}
function ln(text,color,opts){ return {text,options:Object.assign({color:color||TEXT,breakLine:true},opts||{})}; }

// =====================================================================
// Slide 1 — Title
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  s.addText("CUDA 세미나 · 세션 2",{x:M,y:2.05,w:8,h:0.5,fontFace:KFONT,fontSize:18,bold:true,color:GREEN,margin:0});
  s.addText("메모리 모델과\n데이터 이동(data movement)",{x:M,y:2.5,w:11.8,h:2.0,fontFace:KFONT,fontSize:44,bold:true,color:TEXT,lineSpacingMultiple:1.05,margin:0});
  s.addText("쓰고 · 끝내고 · 읽어서 · 검사한다 — cuda_memtest의 심장부",{x:M,y:4.7,w:11.5,h:0.5,fontFace:KFONT,fontSize:20,color:MUTED,margin:0});
  codePanel(s,M,5.6,8.6,1.0,[
    ln("cudaMalloc  →  cudaMemset  →  커널 실행  →  cudaMemcpy(D→H)",GREEN),
    ln("//  GPU에 할당      0으로 초기화        검사           오류를 CPU로",TEAL),
  ],{fontSize:12.5});
  s.addText("입문 과정 · 세션 1에 이어서",{x:9.4,y:6.05,w:3.3,h:0.4,align:"right",fontFace:KFONT,fontSize:13,color:MUTED,margin:0});
  s.addNotes("세션 1의 쓰기 커널에 이어, 오늘은 읽어서 검사하기 + GPU↔CPU 데이터 이동을 배웁니다.");
})();

// =====================================================================
// Slide 2 — objectives
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"1","오늘의 목표");
  const items=[
    ["GPU 메모리 3종 API","cudaMalloc · cudaMemset · cudaMemcpy를 구분해 쓸 수 있다"],
    ["쓰기-종료-읽기 패턴","왜 쓰기 커널을 끝내야 읽기가 올바른 값을 보는지 설명한다"],
    ["이동 반전(moving inversion)","커널 3개의 협업으로 p1→p2 검사가 이뤄짐을 이해한다"],
    ["오류가 CPU로 돌아오는 길","error_checking이 cudaMemcpy로 오류를 가져오는 흐름을 안다"],
  ];
  let y=1.7;
  items.forEach((it,i)=>{
    s.addShape(p.ShapeType.roundRect,{x:M,y,w:W-2*M,h:1.06,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(String(i+1),{x:M+0.22,y:y+0.2,w:0.66,h:0.66,align:"center",valign:"middle",fontFace:MONO,fontSize:22,bold:true,color:BG,fill:{color:i%2?TEAL:GREEN},shape:p.ShapeType.roundRect,rectRadius:0.33});
    s.addText(it[0],{x:M+1.15,y:y+0.14,w:W-2*M-1.4,h:0.44,fontFace:KFONT,fontSize:19,bold:true,color:TEXT,margin:0,valign:"middle"});
    s.addText(it[1],{x:M+1.15,y:y+0.55,w:W-2*M-1.4,h:0.42,fontFace:KFONT,fontSize:14,color:MUTED,margin:0,valign:"middle"});
    y+=1.24;
  });
  s.addNotes("세션 1의 write 커널이 오늘의 read 커널과 어떻게 이어지는지 예고하세요.");
})();

// =====================================================================
// Slide 3 — GPU memory hierarchy (concept)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"2","GPU 메모리 계층 한눈에");
  s.addText("cuda_memtest가 검사하는 것은 이 중 가장 크고 느린 전역 메모리(global memory)입니다.",{x:M,y:1.6,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:16,color:MUTED,margin:0});
  const rows=[
    ["레지스터(register)","스레드 전용 · 가장 빠름","커널 안의 지역 변수 (i, ptr)",GREEN],
    ["공유 메모리(shared)","블록 내 스레드가 공유 · 빠름","이 저장소는 거의 안 씀 (세션 3에서 개념만)",TEAL],
    ["전역 메모리(global)","모든 스레드가 접근 · 크고 느림","검사 대상! cudaMalloc으로 할당한 4 GB",AMBER],
  ];
  let y=2.35;
  rows.forEach(r=>{
    s.addShape(p.ShapeType.roundRect,{x:M,y,w:W-2*M,h:1.28,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(r[0],{x:M+0.3,y:y+0.18,w:3.5,h:0.5,fontFace:KFONT,fontSize:20,bold:true,color:r[3],margin:0,valign:"middle"});
    s.addText(r[1],{x:M+0.3,y:y+0.68,w:3.5,h:0.45,fontFace:KFONT,fontSize:13,color:MUTED,margin:0,valign:"middle"});
    s.addText(r[2],{x:M+4.1,y:y+0.18,w:W-2*M-4.4,h:0.95,fontFace:KFONT,fontSize:16,color:TEXT,margin:0,valign:"middle"});
    y+=1.44;
  });
  s.addNotes("입문 단계에선 '전역 메모리 = 크고 느림 = 검사 대상' 하나만 확실히. 나머지는 세션 3.");
})();

// =====================================================================
// Slide 3b — memory hierarchy block diagram (nested scopes)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"2","메모리 계층 구조 · 범위(scope)로 보기");
  s.addText("메모리는 '누가 접근할 수 있는가(범위)'로 나뉩니다. 범위가 넓을수록 크지만 느립니다.",{x:M,y:1.5,w:W-2*M,h:0.45,fontFace:KFONT,fontSize:15,color:MUTED,margin:0});

  // nested boxes: Global > Shared > Register  (left ~8.4 wide)
  const gx=M, gy=2.05, gw=8.0, gh=4.15;
  s.addShape(p.ShapeType.roundRect,{x:gx,y:gy,w:gw,h:gh,rectRadius:0.07,fill:{color:"14201A"},line:{color:GREEN,width:1.5}});
  s.addText("전역 메모리(Global)  ·  모든 스레드 접근  ·  크고 느림 (예: 4 GB)",{x:gx+0.25,y:gy+0.14,w:gw-0.5,h:0.4,fontFace:KFONT,fontSize:14,bold:true,color:GREEN,margin:0});
  s.addText("← cuda_memtest가 검사하는 대상",{x:gx+0.25,y:gy+0.52,w:gw-0.5,h:0.32,fontFace:KFONT,fontSize:11.5,italic:true,color:GREEN,margin:0});

  const sx=gx+0.4, sy=gy+0.95, sw=gw-0.8, sh=gh-1.35;
  s.addShape(p.ShapeType.roundRect,{x:sx,y:sy,w:sw,h:sh,rectRadius:0.06,fill:{color:"12212A"},line:{color:TEAL,width:1.5}});
  s.addText("공유 메모리(Shared)  ·  블록 내 스레드끼리 공유  ·  빠름",{x:sx+0.25,y:sy+0.14,w:sw-0.5,h:0.35,fontFace:KFONT,fontSize:13.5,bold:true,color:TEAL,margin:0});
  s.addText("이 저장소는 거의 쓰지 않음 (개념만)",{x:sx+0.25,y:sy+0.5,w:sw-0.5,h:0.3,fontFace:KFONT,fontSize:11,italic:true,color:MUTED,margin:0});

  const rx=sx+0.4, ry=sy+0.9, rw=sw-0.8, rh=sh-1.25;
  s.addShape(p.ShapeType.roundRect,{x:rx,y:ry,w:rw,h:rh,rectRadius:0.05,fill:{color:CODEBG},line:{color:AMBER,width:1.5}});
  s.addText("레지스터(Register)  ·  스레드 전용  ·  가장 빠름·가장 작음",{x:rx+0.25,y:ry+0.18,w:rw-0.5,h:0.4,fontFace:KFONT,fontSize:13.5,bold:true,color:AMBER,margin:0});
  s.addText("커널 안의 지역 변수 (i, ptr 등)",{x:rx+0.25,y:ry+0.6,w:rw-0.5,h:0.4,fontFace:MONO,fontSize:12,color:TEXT,margin:0});

  // Host memory box (right) + cudaMemcpy arrow
  const mx=gx+gw+0.55, my=2.8, mw=W-M-(gx+gw+0.55), mh=2.6;
  s.addShape(p.ShapeType.roundRect,{x:mx,y:my,w:mw,h:mh,rectRadius:0.07,fill:{color:CARD},line:{color:MUTED,width:1.5}});
  s.addText("호스트 메모리",{x:mx+0.2,y:my+0.25,w:mw-0.4,h:0.4,align:"center",fontFace:KFONT,fontSize:15,bold:true,color:TEXT,margin:0});
  s.addText("(CPU RAM)",{x:mx+0.2,y:my+0.68,w:mw-0.4,h:0.35,align:"center",fontFace:KFONT,fontSize:12,color:MUTED,margin:0});
  s.addText("GPU와 별개",{x:mx+0.2,y:my+1.15,w:mw-0.4,h:0.35,align:"center",fontFace:KFONT,fontSize:12,color:MUTED,margin:0});
  s.addText("cudaMemcpy로\n오류 정보를 여기로",{x:mx+0.2,y:my+1.6,w:mw-0.4,h:0.7,align:"center",fontFace:KFONT,fontSize:11.5,italic:true,color:TEAL,margin:0,lineSpacingMultiple:1.05});
  // arrow between global(gx+gw) and host(mx)
  s.addShape(p.ShapeType.line,{x:gx+gw+0.03,y:my+1.05,w:mx-(gx+gw)-0.06,h:0,line:{color:TEAL,width:2.5,beginArrowType:"triangle",endArrowType:"triangle"}});

  s.addText("범위: 레지스터(스레드) ⊂ 공유(블록) ⊂ 전역(전체 GPU). 호스트 메모리는 GPU 밖 — cudaMemcpy로만 오갑니다.",{x:M,y:6.4,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:13,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("중첩 박스로 '범위'를 시각화. 안쪽=빠르고 작음/좁은 범위, 바깥=느리고 큼/넓은 범위. 전역이 검사 대상.");
})();

// =====================================================================
// Slide 3c — address mapping & address-bus test
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"2","메모리 주소 매핑 · 주소 버스 검사");
  s.addText("스레드가 계산한 주소가 실제 어느 물리 셀로 가는지, 그 '주소 배선'이 멀쩡한지가 메모리 검사의 핵심입니다.",{x:M,y:1.45,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:14.5,color:MUTED,margin:0});

  const colW=(W-2*M-0.4)/2, y0=1.95, bh=2.55;
  // LEFT: software address calc
  s.addShape(p.ShapeType.roundRect,{x:M,y:y0,w:colW,h:bh,rectRadius:0.07,fill:{color:CARD},line:{color:TEAL,width:1.5}});
  s.addText("① 소프트웨어: 주소 계산",{x:M+0.25,y:y0+0.15,w:colW-0.5,h:0.4,fontFace:KFONT,fontSize:15,bold:true,color:TEAL,margin:0});
  codePanel(s,M+0.25,y0+0.6,colW-0.5,0.6,[
    ln("addr = _ptr + blockIdx.x*BLOCKSIZE + i*sizeof(uint)",GREEN),
  ],{fontSize:11.5});
  const chips=[["_ptr","기준 주소",MUTED],["+ blockIdx×1MB","블록 구역",TEAL],["+ i×4B","원소 위치",GREEN]];
  const chw=(colW-0.5-0.3)/3;
  chips.forEach((c,i)=>{
    const x=M+0.25+i*(chw+0.15);
    s.addShape(p.ShapeType.roundRect,{x,y:y0+1.4,w:chw,h:0.85,rectRadius:0.05,fill:{color:CODEBG},line:{color:c[2],width:1}});
    s.addText(c[0],{x:x+0.05,y:y0+1.5,w:chw-0.1,h:0.4,align:"center",fontFace:MONO,fontSize:10.5,bold:true,color:c[2],margin:0});
    s.addText(c[1],{x:x+0.05,y:y0+1.9,w:chw-0.1,h:0.3,align:"center",fontFace:KFONT,fontSize:10,color:MUTED,margin:0});
  });
  s.addText("→ 평탄한 선형 주소(flat address)",{x:M+0.25,y:y0+bh-0.32,w:colW-0.5,h:0.28,fontFace:KFONT,fontSize:11,italic:true,color:MUTED,margin:0});

  // RIGHT: hardware decoding
  const rx=M+colW+0.4;
  s.addShape(p.ShapeType.roundRect,{x:rx,y:y0,w:colW,h:bh,rectRadius:0.07,fill:{color:CARD},line:{color:AMBER,width:1.5}});
  s.addText("② 하드웨어: 주소 디코딩",{x:rx+0.25,y:y0+0.15,w:colW-0.5,h:0.4,fontFace:KFONT,fontSize:15,bold:true,color:AMBER,margin:0});
  s.addText("주소 디코더(주소 배선)가 주소 비트를 나눠 물리 셀을 고릅니다.",{x:rx+0.25,y:y0+0.58,w:colW-0.5,h:0.32,fontFace:KFONT,fontSize:11.5,color:MUTED,margin:0});
  const segs=[["뱅크(bank)",TEAL],["행(row)",GREEN],["열(column)",AMBER]];
  const sw=(colW-0.5-0.2)/3;
  segs.forEach((sg,i)=>{
    const x=rx+0.25+i*(sw+0.1);
    s.addShape(p.ShapeType.roundRect,{x,y:y0+1.05,w:sw,h:0.55,rectRadius:0.04,fill:{color:CODEBG},line:{color:sg[1],width:1}});
    s.addText(sg[0],{x,y:y0+1.05,w:sw,h:0.55,align:"center",valign:"middle",fontFace:KFONT,fontSize:11,bold:true,color:sg[1],margin:0});
  });
  s.addText("주소 비트 → (뱅크 · 행 · 열)로 분해",{x:rx+0.25,y:y0+1.68,w:colW-0.5,h:0.3,fontFace:MONO,fontSize:10.5,color:MUTED,margin:0});
  s.addText([
    ln("주소 배선 하나가 고착되면 ",TEXT,{breakLine:false}),
    ln("A로 쓴 값이 엉뚱한 A'에 저장(에일리어싱)",RED,{bold:true,breakLine:false}),
    ln(" — 패턴만 봐선 놓칠 수 있습니다.",TEXT,{breakLine:true}),
  ],{x:rx+0.25,y:y0+2.05,w:colW-0.5,h:0.42,fontFace:KFONT,fontSize:11.5,color:TEXT,margin:0,valign:"top",lineSpacingMultiple:1.0});

  // bottom: test 0/1 band
  s.addShape(p.ShapeType.roundRect,{x:M,y:4.7,w:W-2*M,h:1.35,rectRadius:0.08,fill:{color:"14201A"},line:{color:GREEN,width:1}});
  s.addText("주소 배선을 겨냥한 검사 (원논문 Table I: address bus test)",{x:M+0.3,y:4.82,w:W-2*M-0.6,h:0.4,fontFace:KFONT,fontSize:14,bold:true,color:GREEN,margin:0});
  s.addText([
    ln("Test 0 [Walking 1 bit]: ",GREEN,{bold:true,breakLine:false}),
    ln("주소의 비트를 하나씩 바꿔가며 정말 다른 셀로 가는지 확인 → 주소 배선 검사",TEXT,{breakLine:true,paraSpaceAfter:4}),
    ln("Test 1 [Own address]: ",GREEN,{bold:true,breakLine:false}),
    ln("각 셀에 '자기 주소'를 써 넣고 되읽어, 그 값이 여전히 자기 주소인지 확인 → 잘못 매핑되면 불일치",TEXT,{breakLine:true}),
  ],{x:M+0.3,y:5.22,w:W-2*M-0.6,h:0.8,fontFace:KFONT,fontSize:12.5,color:TEXT,margin:0,valign:"top",lineSpacingMultiple:1.1});

  s.addText("이 저장소: Test 0·1은 데이터가 아니라 '주소 자체'를 검사합니다. Test 2~8(이동 반전)은 데이터 셀 값을 검사합니다.",{x:M,y:6.15,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:12,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("주소 매핑: SW가 선형 주소 계산 → HW 디코더가 뱅크/행/열로 분해. 주소 배선 고장=aliasing. Test0(walking 1비트)·Test1(own address)이 이를 검사. 원논문 Table I.");
})();

// =====================================================================
// Slide 4 — three memory APIs (cards)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"3","메모리 3종 API");
  const cards=[
    ["cudaMalloc","GPU에 메모리 할당",GREEN,["(void**)&ptr 로 포인터를 받음","CPU가 아니라 GPU 메모리","tests.cpp:1573 (오류 버퍼)"]],
    ["cudaMemset","GPU 메모리 초기화",TEAL,["보통 0으로 채움","할당 직후 깨끗이","error_checking에서 재초기화"]],
    ["cudaMemcpy","GPU ↔ CPU 복사",AMBER,["방향 지정 필수","DeviceToHost = GPU→CPU","오류 정보를 CPU로 가져옴"]],
  ];
  const cw=(W-2*M-0.8)/3;
  cards.forEach((c,i)=>{
    const x=M+i*(cw+0.4);
    s.addShape(p.ShapeType.roundRect,{x,y:1.75,w:cw,h:4.5,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(c[0],{x:x+0.28,y:1.98,w:cw-0.56,h:0.5,fontFace:MONO,fontSize:19,bold:true,color:c[2],margin:0});
    s.addText(c[1],{x:x+0.28,y:2.5,w:cw-0.56,h:0.45,fontFace:KFONT,fontSize:15,bold:true,color:TEXT,margin:0});
    s.addText(c[3].map(t=>ln(t,TEXT,{breakLine:true,bullet:{code:"2022"},paraSpaceAfter:9})),{x:x+0.28,y:3.05,w:cw-0.56,h:3.0,fontFace:KFONT,fontSize:14,color:TEXT,margin:0,valign:"top"});
  });
  s.addText("세 API 모두 allocate_small_mem(할당·초기화)과 error_checking(복사)에서 실제로 볼 수 있습니다.",{x:M,y:6.5,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("cudaMalloc은 (void**)&ptr 이중 포인터가 처음엔 헷갈립니다. '포인터의 주소를 넘겨 GPU 주소를 받아온다'로 설명.");
})();

// =====================================================================
// Slide 5 — why split write/read (flush) — the key idea
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"4","핵심 · 왜 커널을 끝내고 다시 시작할까");
  s.addText("같은 커널에서 쓰고 바로 읽으면, 아직 메모리에 반영 안 된 값을 읽을 수 있습니다.",{x:M,y:1.6,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:17,color:TEXT,margin:0});

  // flow: 3 steps with arrows
  const steps=[
    ["① 쓰기 커널",GREEN,"kernel_move_inv_write","전체 메모리에 p1을 쓴다"],
    ["② 커널 종료 = flush",AMBER,"(커널이 끝나야)","쓴 값이 메모리에 확실히 반영"],
    ["③ 읽기 커널",TEAL,"kernel_move_inv_readwrite","p1인지 검사하고 p2를 쓴다"],
  ];
  const bw=3.66, gap=0.55, x0=M, y0=2.35, bh=2.35;
  steps.forEach((st,i)=>{
    const x=x0+i*(bw+gap);
    s.addShape(p.ShapeType.roundRect,{x,y:y0,w:bw,h:bh,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(st[0],{x:x+0.25,y:y0+0.22,w:bw-0.5,h:0.55,fontFace:KFONT,fontSize:18,bold:true,color:st[1],margin:0});
    s.addText(st[2],{x:x+0.25,y:y0+0.85,w:bw-0.5,h:0.5,fontFace:MONO,fontSize:12.5,color:TEXT,margin:0});
    s.addText(st[3],{x:x+0.25,y:y0+1.4,w:bw-0.5,h:0.8,fontFace:KFONT,fontSize:14,color:MUTED,margin:0,valign:"top"});
    if(i<2) s.addText("→",{x:x+bw+0.02,y:y0+0.8,w:gap-0.04,h:0.7,align:"center",valign:"middle",fontFace:MONO,fontSize:26,bold:true,color:MUTED,margin:0});
  });
  s.addText([
    ln("이 분리 덕분에 하드웨어 메모리 오류를 신뢰성 있게 잡습니다. ",TEXT,{breakLine:false}),
    ln("쓰기와 읽기가 서로 다른 커널 실행이어야",GREEN,{bold:true,breakLine:false}),
    ln(" 메모리를 실제로 왕복한 값을 검사하게 됩니다.",TEXT,{breakLine:true}),
  ],{x:M,y:5.15,w:W-2*M,h:1.2,fontFace:KFONT,fontSize:17,color:TEXT,margin:0,valign:"top"});
  s.addNotes("이 슬라이드가 세션 2의 핵심. '커널 종료 = 메모리 반영(flush) 보장'을 각인시키세요.");
})();

// =====================================================================
// Slide 6 — the three-loop structure (code)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"5","move_inv_test · 세 개의 for 루프");
  codePanel(s,M,1.7,W-2*M,4.35,[
    ln("// tests.cpp:305   호스트(CPU) 함수",MUTED),
    ln("① 쓰기 루프  ─ 전체 메모리에 p1을 쓴다",GREEN),
    ln("   for (...) kernel_move_inv_write<<<grid,1>>>(... p1);",TEXT),
    ln("",TEXT),
    ln("② 읽고-쓰기 루프 ─ p1 검사 후 보수 p2를 쓴다",TEAL),
    ln("   for (...) {",TEXT),
    ln("       kernel_move_inv_readwrite<<<grid,1>>>(... p1, p2 ...);",TEXT),
    ln("       err += error_checking(\"move_inv_readwrite\", i);  // 오류를 CPU로",AMBER),
    ln("   }",TEXT),
    ln("",TEXT),
    ln("③ 읽기 루프 ─ p2 인지 다시 검사한다",TEAL),
    ln("   for (...) {",TEXT),
    ln("       kernel_move_inv_read<<<grid,1>>>(... p2 ...);",TEXT),
    ln("       err += error_checking(\"move_inv_read\", i);",AMBER),
    ln("   }",TEXT),
  ],{fontSize:13});
  s.addText("각 커널 실행 뒤 error_checking으로 오류를 확인합니다. 커널→검사→커널→검사 리듬을 기억하세요.",{x:M,y:6.2,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:14,color:MUTED,margin:0});
  s.addNotes("세 루프가 각각 write/readwrite/read 커널에 대응. Test2는 이 함수를 (p1,p2)와 (p2,p1)로 두 번 부릅니다.");
})();

// =====================================================================
// Slide 6b — host<->device sequence diagram
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"5","호스트↔디바이스 시퀀스(sequence)");
  s.addText("move_inv_test가 CPU와 GPU 사이에서 시간순으로 주고받는 상호작용입니다.",{x:M,y:1.45,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:15,color:MUTED,margin:0});

  const xh=3.0, xd=10.3, headY=1.95, headH=0.6;
  // lifeline headers
  s.addShape(p.ShapeType.roundRect,{x:xh-1.4,y:headY,w:2.8,h:headH,rectRadius:0.06,fill:{color:CARD},line:{color:TEAL,width:1.5}});
  s.addText("호스트(CPU)",{x:xh-1.4,y:headY,w:2.8,h:headH,align:"center",valign:"middle",fontFace:KFONT,fontSize:15,bold:true,color:TEAL,margin:0});
  s.addShape(p.ShapeType.roundRect,{x:xd-1.4,y:headY,w:2.8,h:headH,rectRadius:0.06,fill:{color:CARD},line:{color:GREEN,width:1.5}});
  s.addText("디바이스(GPU)",{x:xd-1.4,y:headY,w:2.8,h:headH,align:"center",valign:"middle",fontFace:KFONT,fontSize:15,bold:true,color:GREEN,margin:0});
  // lifelines (dashed vertical)
  const llTop=headY+headH, llBot=6.35;
  s.addShape(p.ShapeType.line,{x:xh,y:llTop,w:0,h:llBot-llTop,line:{color:LINE,width:1.5,dashType:"dash"}});
  s.addShape(p.ShapeType.line,{x:xd,y:llTop,w:0,h:llBot-llTop,line:{color:LINE,width:1.5,dashType:"dash"}});

  // messages
  const msg=(y,dir,color,dashed,label)=>{
    const opt={color,width:2.2}; if(dashed) opt.dashType="dash";
    if(dir>0) opt.endArrowType="triangle"; else opt.beginArrowType="triangle";
    s.addShape(p.ShapeType.line,{x:xh,y,w:xd-xh,h:0,line:opt});
    s.addText(label,{x:xh+0.1,y:y-0.34,w:xd-xh-0.2,h:0.3,align:"center",valign:"middle",fontFace:KFONT,fontSize:12.5,bold:true,color:color,margin:0});
  };
  msg(2.95, 1, GREEN, false, "① write<<<>>> — 전체 메모리에 p1 기록");
  msg(3.75, -1, TEAL, true, "커널 종료 = 메모리 반영(flush) — DeviceSynchronize");
  msg(4.55, 1, GREEN, false, "② readwrite<<<>>> — p1 검사 후 p2 기록");
  msg(5.35, -1, TEAL, true, "error_checking: cudaMemcpy(D→H) — 오류 정보 회수");
  msg(6.15, 1, GREEN, false, "③ read<<<>>> — p2 다시 검사");

  // note
  s.addText("핵심: ①과 ② 사이의 '커널 종료(flush)'가 있어야 ②가 메모리에 반영된 값을 읽습니다.",{x:M,y:6.6,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:12.5,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("UML 시퀀스풍. 실선=호스트가 GPU에 커널 요청, 점선=GPU→호스트 반환(동기화/데이터). flush가 정확성의 열쇠.");
})();

// =====================================================================
// Slide 7 — moving inversion algorithm (visual)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"6","이동 반전(moving inversion) 알고리즘");
  s.addText("패턴 p1을 썼다가, 확인 후 그 보수 p2 = ~p1로 뒤집어 다시 확인합니다. 비트를 양방향으로 검사해 고착 오류를 잡습니다.",{x:M,y:1.6,w:W-2*M,h:0.75,fontFace:KFONT,fontSize:16,color:MUTED,margin:0,valign:"top"});

  // bit strip: p1 then p2
  const drawBits=(y,label,bits,color)=>{
    s.addText(label,{x:M,y:y+0.05,w:2.4,h:0.6,fontFace:KFONT,fontSize:16,bold:true,color:color,margin:0,valign:"middle"});
    const bx=M+2.6, bw=0.7, n=bits.length;
    for(let i=0;i<n;i++){
      s.addShape(p.ShapeType.roundRect,{x:bx+i*(bw+0.12),y,w:bw,h:0.7,rectRadius:0.05,fill:{color:bits[i]==="1"?color:CODEBG},line:{color:color,width:1}});
      s.addText(bits[i],{x:bx+i*(bw+0.12),y,w:bw,h:0.7,align:"center",valign:"middle",fontFace:MONO,fontSize:18,bold:true,color:bits[i]==="1"?BG:MUTED,margin:0});
    }
  };
  drawBits(2.65,"p1 =",  "10110100".split(""),GREEN);
  drawBits(3.75,"p2 = ~p1 =","01001011".split(""),TEAL);
  s.addText("↓ 모든 비트 반전",{x:M+2.6,y:3.42,w:5,h:0.3,fontFace:KFONT,fontSize:12,color:MUTED,margin:0});

  s.addText([
    ln("1로 고착된 비트",GREEN,{bold:true,breakLine:false}),
    ln("는 p2(0을 기대) 검사에서 걸리고, ",TEXT,{breakLine:false}),
    ln("0으로 고착된 비트",TEAL,{bold:true,breakLine:false}),
    ln("는 p1(1을 기대) 검사에서 걸립니다. 그래서 양방향 검사가 필요합니다.",TEXT,{breakLine:true}),
  ],{x:M,y:4.9,w:W-2*M,h:1.0,fontFace:KFONT,fontSize:16,color:TEXT,margin:0,valign:"top"});
  s.addText("Test 2는 p1=0x00000000, p2=0xFFFFFFFF (모두 0 / 모두 1)로 이 검사를 수행합니다. (tests.cpp:705)",{x:M,y:6.1,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:13,italic:true,color:MUTED,margin:0});
  s.addNotes("고착(stuck-at) 오류 개념을 비트 그림으로. 왜 한 방향 검사로는 부족한지 강조.");
})();

// =====================================================================
// Slide 7b — test patterns & fault types
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"6","메모리 테스트 패턴과 결함 유형(fault types)");
  s.addText("패턴마다 노리는 결함이 다릅니다. 여러 패턴을 조합해 다양한 고장을 잡아냅니다.",{x:M,y:1.45,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:14.5,color:MUTED,margin:0});

  // pattern strips
  const strip=(y,label,bits,tag,tagcol)=>{
    s.addText(label,{x:M,y:y,w:2.35,h:0.5,valign:"middle",fontFace:KFONT,fontSize:12.5,bold:true,color:TEXT,margin:0});
    const bx=M+2.45, cw=0.55, gap=0.08;
    for(let i=0;i<bits.length;i++){
      const on=bits[i]===1;
      s.addShape(p.ShapeType.roundRect,{x:bx+i*(cw+gap),y,w:cw,h:0.5,rectRadius:0.04,fill:{color:on?GREEN:CODEBG},line:{color:on?GREEN:LINE,width:1}});
      s.addText(String(bits[i]),{x:bx+i*(cw+gap),y,w:cw,h:0.5,align:"center",valign:"middle",fontFace:MONO,fontSize:13,bold:true,color:on?BG:MUTED,margin:0});
    }
    const tx=bx+bits.length*(cw+gap)+0.15;
    s.addText(tag,{x:tx,y:y,w:W-M-tx,h:0.5,valign:"middle",fontFace:KFONT,fontSize:11.5,color:tagcol||MUTED,margin:0});
  };
  strip(2.3,"워킹 1비트",[0,0,0,0,0,0,0,1],"한 비트씩 이동 → 주소·배선 검사 (Test 0·6)",TEAL);
  strip(2.9,"전부 0 / 전부 1",[1,1,1,1,1,1,1,1],"0x00000000 ↔ 0xFFFFFFFF → 고착 검사 (Test 2)",GREEN);
  strip(3.5,"8비트 반복",[1,0,0,0,0,0,0,0],"0x80808080 (8비트 폭) → 미세 고착 (Test 3)",GREEN);
  strip(4.1,"난수 + 보수",[1,0,1,1,0,1,0,0],"무작위 값과 그 보수 → 데이터 민감 오류 (Test 4·7·8·10)",AMBER);

  // fault type chips
  s.addShape(p.ShapeType.roundRect,{x:M,y:4.85,w:W-2*M,h:1.55,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
  s.addText("결함 유형별 담당 테스트 (원논문 Table I)",{x:M+0.3,y:4.97,w:W-2*M-0.6,h:0.35,fontFace:KFONT,fontSize:13,bold:true,color:TEXT,margin:0});
  const faults=[
    ["주소 버스","address bus","Test 0·1",TEAL],
    ["스턱-앳 고착","stuck-at","Test 2·3",GREEN],
    ["데이터 민감","data sensitive","Test 4~8",AMBER],
    ["데이터 보존","retention","Test 9",TEAL],
    ["소프트 오류","soft error","Test 10",RED],
  ];
  const cw=(W-2*M-0.6-4*0.2)/5;
  faults.forEach((f,i)=>{
    const x=M+0.3+i*(cw+0.2);
    s.addShape(p.ShapeType.roundRect,{x,y:5.42,w:cw,h:0.85,rectRadius:0.06,fill:{color:CODEBG},line:{color:f[3],width:1.2}});
    s.addText(f[0],{x:x+0.1,y:5.5,w:cw-0.2,h:0.32,align:"center",fontFace:KFONT,fontSize:12.5,bold:true,color:f[3],margin:0});
    s.addText(f[1],{x:x+0.1,y:5.8,w:cw-0.2,h:0.24,align:"center",fontFace:KFONT,fontSize:9,color:MUTED,margin:0});
    s.addText(f[2],{x:x+0.1,y:6.02,w:cw-0.2,h:0.24,align:"center",fontFace:MONO,fontSize:10.5,bold:true,color:TEXT,margin:0});
  });
  s.addText("초록 칸 = 비트 1, 어두운 칸 = 비트 0. Test 9(비트 페이드)는 값을 쓰고 오래 두어 '지워지는지' 봅니다.",{x:M,y:6.55,w:W-2*M,h:0.35,fontFace:KFONT,fontSize:11.5,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("패턴→결함 매핑. 워킹=주소/배선, 0·1=고착, 8비트=미세고착, 난수=데이터민감, 비트페이드=보존, 난수반복=소프트. 원논문 Table I 근거.");
})();

// =====================================================================
// Slide 8 — error path back to CPU (code + explanation)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"7","오류가 CPU로 돌아오는 길");
  codePanel(s,M,1.7,W-2*M,2.5,[
    ln("// error_checking()  tests.cpp:108",MUTED),
    ln("unsigned int err = 0;",TEXT),
    ln("cudaMemcpy(&err, err_count, sizeof(unsigned int),",TEXT),
    ln("           cudaMemcpyDeviceToHost);   // GPU 카운터 → CPU 변수",GREEN),
    ln("if (err) {  /* 오류 주소·기대값·현재값을 출력·기록 */  }",TEXT),
  ],{fontSize:13.5});
  s.addText([
    ln("err_count",GREEN,{bold:true,breakLine:false}),
    ln("는 GPU에 있는 디바이스 포인터입니다. CPU는 그 값을 ",TEXT,{breakLine:false}),
    ln("직접 읽을 수 없어",AMBER,{bold:true,breakLine:false}),
    ln(" 반드시 cudaMemcpy로 가져와야 합니다.",TEXT,{breakLine:true}),
  ],{x:M,y:4.4,w:W-2*M,h:0.9,fontFace:KFONT,fontSize:17,color:TEXT,margin:0,valign:"top"});
  // small info cards
  const cards=[
    ["기록되는 정보",GREEN,"오류 주소 · 기대값 · 실제값 · 재읽기값"],
    ["최근 10개만",TEAL,"MAX_ERR_RECORD_COUNT=10, 순환 기록"],
  ];
  const cw=(W-2*M-0.4)/2;
  cards.forEach((c,i)=>{
    const x=M+i*(cw+0.4);
    s.addShape(p.ShapeType.roundRect,{x,y:5.35,w:cw,h:1.1,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(c[0],{x:x+0.25,y:5.5,w:cw-0.5,h:0.4,fontFace:KFONT,fontSize:15,bold:true,color:c[1],margin:0});
    s.addText(c[2],{x:x+0.25,y:5.9,w:cw-0.5,h:0.45,fontFace:KFONT,fontSize:14,color:TEXT,margin:0});
  });
  s.addNotes("디바이스 포인터를 호스트에서 역참조하면 안 된다는 점이 입문자 핵심 함정. cudaMemcpy가 다리 역할.");
})();

// =====================================================================
// Slide 9 — allocation strategy (robust real-world code)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"8","실전 코드 · 할 수 있는 만큼 할당하기");
  codePanel(s,M,1.7,W-2*M,2.9,[
    ln("// cuda_memtest.cpp:157   가용 메모리를 조회하고,",MUTED),
    ln("// 할당 실패하면 블록을 하나씩 줄여 재시도한다",MUTED),
    ln("tot_num_blocks = free / BLOCKSIZE;",TEXT),
    ln("do {",TEXT),
    ln("    err = cudaMalloc(&ptr, tot_num_blocks * BLOCKSIZE);",GREEN),
    ln("    if (err != cudaSuccess) --tot_num_blocks;   // 줄여서 재시도",AMBER),
    ln("} while (cudaGetLastError() != cudaSuccess);",TEXT),
  ],{fontSize:13.5});
  s.addText([
    ln("드라이버가 보고한 여유 메모리를 다 할당하지 못할 때가 있습니다. 실전 코드는 ",TEXT,{breakLine:false}),
    ln("실패를 가정하고 우아하게 물러섭니다",GREEN,{bold:true,breakLine:false}),
    ln(". 장난감 예제와 프로덕션 코드의 차이입니다.",TEXT,{breakLine:true}),
  ],{x:M,y:4.85,w:W-2*M,h:1.0,fontFace:KFONT,fontSize:17,color:TEXT,margin:0,valign:"top"});
  s.addText("cudaMemGetInfo로 여유 메모리를 조회 → BLOCKSIZE(1 MB)로 나눠 블록 수 결정.",{x:M,y:6.0,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,margin:0});
  s.addNotes("에러 처리와 재시도 루프는 세션 4(에러 핸들링)의 예고편이기도 합니다.");
})();

// =====================================================================
// Slide 10 — Lab preview (4 exercises)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"9","실습 미리보기 · Lab 2");
  const labs=[
    ["Ex 1","세 커널의 협업 보기","--verbose --verbose 로 write→readwrite→read 순서 확인",GREEN],
    ["Ex 2","오류가 CPU로 돌아오는 길","error_checking의 cudaMemcpy 직후 출력 추가",TEAL],
    ["Ex 3","메모리를 훼손해 오류 재현","읽기 커널에서 비트 하나를 뒤집어 검출 확인",AMBER],
    ["Ex 4","디바이스 버퍼 할당 보기","allocate_small_mem의 cudaMalloc/Memset 관찰",GREEN],
  ];
  const cw=(W-2*M-0.5)/2, ch=2.15;
  labs.forEach((l,i)=>{
    const x=M+(i%2)*(cw+0.5), y=1.7+Math.floor(i/2)*(ch+0.35);
    s.addShape(p.ShapeType.roundRect,{x,y,w:cw,h:ch,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(l[0],{x:x+0.28,y:y+0.22,w:1.3,h:0.55,align:"center",valign:"middle",fontFace:MONO,fontSize:17,bold:true,color:BG,fill:{color:l[3]},shape:p.ShapeType.roundRect,rectRadius:0.08});
    s.addText(l[1],{x:x+1.75,y:y+0.24,w:cw-2.0,h:0.9,fontFace:KFONT,fontSize:17,bold:true,color:TEXT,margin:0,valign:"top"});
    s.addText(l[2],{x:x+0.28,y:y+1.2,w:cw-0.56,h:0.8,fontFace:KFONT,fontSize:14,color:MUTED,margin:0,valign:"top",lineSpacingMultiple:1.1});
  });
  s.addNotes("Ex 3(fault injection)이 하이라이트. 실제 불량 칩이 내는 증상을 소프트웨어로 재현한다는 점을 강조. 반드시 원복.");
})();

// =====================================================================
// Slide 11 — recap & pitfalls
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"10","핵심 정리 & 자주 하는 실수");
  const colW=(W-2*M-0.4)/2;
  s.addShape(p.ShapeType.roundRect,{x:M,y:1.75,w:colW,h:4.5,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
  s.addText("기억할 것",{x:M+0.3,y:2.0,w:colW-0.6,h:0.5,fontFace:KFONT,fontSize:20,bold:true,color:GREEN,margin:0});
  s.addText([
    ln("Malloc(할당)·Memset(초기화)·Memcpy(복사)",TEXT,{breakLine:true,bullet:{code:"2713"},paraSpaceAfter:12}),
    ln("Memcpy 방향: DeviceToHost = GPU→CPU",TEXT,{breakLine:true,bullet:{code:"2713"},paraSpaceAfter:12}),
    ln("쓰기 커널 종료 = 메모리 반영 보장",TEXT,{breakLine:true,bullet:{code:"2713"},paraSpaceAfter:12}),
    ln("이동 반전: p1 → p2(보수) 양방향 검사",TEXT,{breakLine:true,bullet:{code:"2713"}}),
  ],{x:M+0.3,y:2.65,w:colW-0.6,h:3.4,fontFace:KFONT,fontSize:15.5,color:TEXT,margin:0,valign:"top"});

  const x2=M+colW+0.4;
  s.addShape(p.ShapeType.roundRect,{x:x2,y:1.75,w:colW,h:4.5,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
  s.addText("입문자 함정",{x:x2+0.3,y:2.0,w:colW-0.6,h:0.5,fontFace:KFONT,fontSize:20,bold:true,color:AMBER,margin:0});
  s.addText([
    ln("디바이스 포인터를 CPU에서 *ptr로 역참조 → 오류",TEXT,{breakLine:true,bullet:{code:"2717"},paraSpaceAfter:12}),
    ln("cudaMemcpy 방향을 반대로 지정 → 값이 안 옴",TEXT,{breakLine:true,bullet:{code:"2717"},paraSpaceAfter:12}),
    ln("같은 커널에서 쓰고 바로 읽기 → 반영 전 값 읽음",TEXT,{breakLine:true,bullet:{code:"2717"},paraSpaceAfter:12}),
    ln("cudaMalloc에 (void**) 이중 포인터를 안 넘김",TEXT,{breakLine:true,bullet:{code:"2717"}}),
  ],{x:x2+0.3,y:2.65,w:colW-0.6,h:3.4,fontFace:KFONT,fontSize:15.5,color:TEXT,margin:0,valign:"top"});
  s.addNotes("네 가지 목표(슬라이드 2)로 돌아가 자가 점검을 유도하세요.");
})();

// =====================================================================
// Slide 12 — next session (dark)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  s.addText("다음 세션 예고",{x:M,y:1.6,w:10,h:0.5,fontFace:KFONT,fontSize:18,bold:true,color:GREEN,margin:0});
  s.addText("세션 3 · 병렬성 설계와 성능",{x:M,y:2.15,w:12,h:1.0,fontFace:KFONT,fontSize:38,bold:true,color:TEXT,margin:0});
  s.addText([
    ln("지금까지의 ",TEXT,{breakLine:false}),
    ln("<<<grid, 1>>>",TEAL,{bold:true,breakLine:false}),
    ln(" (블록당 스레드 1개, 느림)을, Test 10의 ",TEXT,{breakLine:false}),
    ln("<<<32768, 64>>>",GREEN,{bold:true,breakLine:false}),
    ln(" (빠름)와 비교합니다.",TEXT,{breakLine:true}),
  ],{x:M,y:3.3,w:12,h:0.9,fontFace:KFONT,fontSize:19,color:TEXT,margin:0,valign:"top"});
  s.addText([
    ln("threadIdx · blockDim 로 수천 스레드가 협업",MONO,{breakLine:true,color:GREEN,paraSpaceAfter:8}),
    ln("CUDA 이벤트(event)로 시간 측정 → 대역폭(GB/s) 계산",MUTED,{breakLine:true,paraSpaceAfter:8}),
    ln("메모리 병합(coalescing)이 왜 속도를 좌우하는가",MUTED,{breakLine:true}),
  ],{x:M,y:4.35,w:12,h:1.4,fontFace:KFONT,fontSize:16,color:TEXT,margin:0,valign:"top"});
  s.addText("실습 랩 2를 먼저 완료하고 오세요.",{x:M,y:6.4,w:12,h:0.4,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,margin:0});
  s.addNotes("느린 커널을 충분히 이해했으니, 다음엔 같은 문제를 빠르게 푸는 커널로 성능을 배운다는 흐름.");
})();

p.writeFile({fileName:"세션2_메모리모델과_데이터이동.pptx"}).then(f=>console.log("wrote",f));
