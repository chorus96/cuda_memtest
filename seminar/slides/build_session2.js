const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.layout = "LAYOUT_WIDE"; // 13.3 x 7.5

// ---- palette (shared with Session 1) ----
const BG="13161C", CARD="1E232C", CODEBG="0E1116", TEXT="E8EDF2", MUTED="9AA6B2";
const GREEN="8CC63F", TEAL="4FB0C6", AMBER="E0A458", LINE="2C3440";
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
