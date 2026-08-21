const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.layout = "LAYOUT_WIDE"; // 13.3 x 7.5

// ---- palette (shared with Sessions 1-4) ----
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
  s.addText("CUDA 세미나 · 세션 5 (마지막 회차)",{x:M,y:2.05,w:9,h:0.5,fontFace:KFONT,fontSize:18,bold:true,color:GREEN,margin:0});
  s.addText("멀티 GPU · 이식성 · 프로젝트\n(multi-GPU, portability, capstone)",{x:M,y:2.5,w:12.0,h:2.0,fontFace:KFONT,fontSize:40,bold:true,color:TEXT,lineSpacingMultiple:1.05,margin:0});
  s.addText("한 GPU에서 여러 GPU로, 그리고 나만의 커널을 만들어 시리즈를 마무리합니다",{x:M,y:4.7,w:12.0,h:0.5,fontFace:KFONT,fontSize:19,color:MUTED,margin:0});
  codePanel(s,M,5.6,9.2,1.0,[
    ln("pthread_create(&pid[i], NULL, thread_func, ...);  // GPU마다 스레드",GREEN),
    ln("MEMTEST_API_PREFIX(Malloc)  →  cudaMalloc  또는  hipMalloc",TEAL),
  ],{fontSize:12.5});
  s.addText("입문 과정 · 세션 4에 이어서",{x:9.9,y:6.05,w:2.8,h:0.4,align:"right",fontFace:KFONT,fontSize:13,color:MUTED,margin:0});
  s.addNotes("본편의 마지막 회차. 멀티 GPU와 이식성을 배우고, 캡스톤으로 배운 것을 종합합니다.");
})();

// =====================================================================
// Slide 2 — objectives
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"1","오늘의 목표");
  const items=[
    ["멀티 GPU 패턴","GPU 하나당 CPU 스레드 하나 — pthread + cudaSetDevice"],
    ["스레드 지역 저장소","__thread gpu_idx가 왜 스레드마다 달라야 하나"],
    ["이식성(portability)","하나의 코드로 CUDA와 HIP(AMD)를 모두 빌드하는 매크로"],
    ["캡스톤(capstone)","배운 것을 종합해 나만의 메모리 테스트를 추가한다"],
  ];
  let y=1.7;
  items.forEach((it,i)=>{
    s.addShape(p.ShapeType.roundRect,{x:M,y,w:W-2*M,h:1.06,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(String(i+1),{x:M+0.22,y:y+0.2,w:0.66,h:0.66,align:"center",valign:"middle",fontFace:MONO,fontSize:22,bold:true,color:BG,fill:{color:i%2?TEAL:GREEN},shape:p.ShapeType.roundRect,rectRadius:0.33});
    s.addText(it[0],{x:M+1.15,y:y+0.14,w:W-2*M-1.4,h:0.44,fontFace:KFONT,fontSize:19,bold:true,color:TEXT,margin:0,valign:"middle"});
    s.addText(it[1],{x:M+1.15,y:y+0.55,w:W-2*M-1.4,h:0.42,fontFace:KFONT,fontSize:14,color:MUTED,margin:0,valign:"middle"});
    y+=1.24;
  });
  s.addNotes("오늘은 관찰 실습 + 직접 구현(캡스톤)입니다. 마지막 회차인 만큼 종합에 집중.");
})();

// =====================================================================
// Slide 3 — multi-GPU model (diagram)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"2","멀티 GPU · GPU 하나당 스레드 하나");
  s.addText("main이 GPU 개수만큼 pthread를 만들고, 각 스레드가 자기 GPU를 cudaSetDevice로 선택합니다.",{x:M,y:1.6,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:16,color:MUTED,margin:0});
  // main box
  s.addShape(p.ShapeType.roundRect,{x:M,y:2.35,w:2.7,h:3.4,rectRadius:0.08,fill:{color:"25303F"},line:{type:"none"}});
  s.addText("main()",{x:M,y:2.55,w:2.7,h:0.5,align:"center",fontFace:MONO,fontSize:18,bold:true,color:TEXT,margin:0});
  s.addText("GPU 개수만큼\npthread 생성",{x:M+0.2,y:3.2,w:2.3,h:1.0,align:"center",fontFace:KFONT,fontSize:14,color:MUTED,margin:0,valign:"top",lineSpacingMultiple:1.2});
  // three thread->gpu lanes
  const lanes=[["스레드 0","cudaSetDevice(0)","GPU 0",GREEN],["스레드 1","cudaSetDevice(1)","GPU 1",TEAL],["스레드 2","cudaSetDevice(2)","GPU 2",AMBER]];
  const lx=M+3.2, lw=W-M-(M+3.2), laneH=1.0, y0=2.35;
  lanes.forEach((L,i)=>{
    const y=y0+i*(laneH+0.2);
    s.addShape(p.ShapeType.roundRect,{x:lx,y,w:lw,h:laneH,rectRadius:0.06,fill:{color:CARD},line:{type:"none"}});
    s.addText(L[0],{x:lx+0.25,y:y+0.1,w:2.6,h:laneH-0.2,valign:"middle",fontFace:KFONT,fontSize:16,bold:true,color:L[3],margin:0});
    s.addText(L[1],{x:lx+3.0,y:y+0.1,w:3.6,h:laneH-0.2,valign:"middle",fontFace:MONO,fontSize:14,color:TEXT,margin:0});
    s.addText(L[2],{x:lx+lw-2.2,y:y+0.1,w:2.0,h:laneH-0.2,align:"right",valign:"middle",fontFace:MONO,fontSize:16,bold:true,color:L[3],margin:0});
  });
  s.addText("각 스레드는 자기 GPU에 오류 버퍼를 할당(allocate_small_mem)하고 모든 테스트를 독립 실행합니다. (cuda_memtest.cpp:543)",{x:M,y:6.05,w:W-2*M,h:0.6,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,margin:0,valign:"top"});
  s.addNotes("스레드↔GPU 1:1 매핑. 각 스레드가 독립적으로 run_tests를 돕니다. 로그가 섞이는 이유이기도.");
})();

// =====================================================================
// Slide 3b — multi-GPU system architecture (parallel per-GPU stacks)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"2","멀티 GPU 시스템 구조 · 병렬 스택");
  s.addText("하나의 프로세스가 GPU마다 독립된 스레드 스택을 띄우고, 모두 동시에 검사를 실행합니다.",{x:M,y:1.45,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:15,color:MUTED,margin:0});

  // Host process box (spans width)
  const hx=M, hy=1.95, hw=W-2*M, hh=0.72;
  s.addShape(p.ShapeType.roundRect,{x:hx,y:hy,w:hw,h:hh,rectRadius:0.06,fill:{color:"25303F"},line:{color:MUTED,width:1.5}});
  s.addText("main() 프로세스  ·  GPU 개수만큼 pthread_create (cuda_memtest.cpp:543)",{x:hx+0.3,y:hy,w:hw-0.6,h:hh,valign:"middle",fontFace:KFONT,fontSize:15,bold:true,color:TEXT,margin:0});

  const cols=[GREEN,TEAL,AMBER];
  const colW=(W-2*M-2*0.45)/3;
  const cells=[
    ["스레드 (pthread)","pthread_create로 생성"],
    ["cudaSetDevice(i)","gpu_idx = i (thread-local)"],
    ["allocate_small_mem","오류 버퍼 할당 (이 GPU)"],
    ["run_tests","11개 테스트 실행"],
    ["GPU i · VRAM","독립 메모리 (검사 대상)"],
  ];
  for(let ci=0;ci<3;ci++){
    const col=cols[ci], cx=M+ci*(colW+0.45);
    // down arrow from host to column
    s.addShape(p.ShapeType.line,{x:cx+colW/2,y:hy+hh+0.02,w:0,h:0.28,line:{color:col,width:2.5,endArrowType:"triangle"}});
    let cy=hy+hh+0.4;
    for(let ri=0;ri<cells.length;ri++){
      const last=ri===cells.length-1;
      const h= 0.54;
      s.addShape(p.ShapeType.roundRect,{x:cx,y:cy,w:colW,h,rectRadius:0.05,fill:{color:last?"1A2418":CODEBG},line:{color:last?col:LINE,width:last?1.5:1}});
      const label = ri===0 ? ("스레드 "+ci+" (pthread)") : (ri===4 ? ("GPU "+ci+" · VRAM") : cells[ri][0]);
      s.addText(label,{x:cx+0.18,y:cy+0.06,w:colW-0.36,h:0.28,fontFace:MONO,fontSize:12,bold:true,color:last?col:TEXT,margin:0,valign:"middle"});
      s.addText(cells[ri][1],{x:cx+0.18,y:cy+0.31,w:colW-0.36,h:0.21,fontFace:KFONT,fontSize:10.5,color:MUTED,margin:0,valign:"middle"});
      // connector between cells
      if(!last) s.addShape(p.ShapeType.line,{x:cx+colW/2,y:cy+h+0.003,w:0,h:0.1,line:{color:LINE,width:1.5,endArrowType:"triangle"}});
      cy+=h+0.11;
    }
  }
  s.addText("각 열은 완전히 독립·병렬 — 한 GPU의 오류가 다른 GPU를 멈추지 않습니다. (예: S1070 = 노드당 4 GPU)",{x:M,y:6.62,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:13,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("GPU마다 pthread→cudaSetDevice→버퍼할당→테스트→VRAM의 독립 스택. 병렬·독립 실행이 핵심. 논문의 S1070(4 GPU/노드)과 연결.");
})();

// =====================================================================
// Slide 4 — thread-local gpu_idx (code)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"3","스레드 지역 저장소 · __thread gpu_idx");
  codePanel(s,M,1.7,W-2*M,2.6,[
    ln("// cuda_memtest.h:81",MUTED),
    ln("extern __thread unsigned int gpu_idx;   // 스레드마다 독립된 복사본",GREEN),
    ln("",TEXT),
    ln("// thread_func  cuda_memtest.cpp:122",MUTED),
    ln("gpu_idx = device;              // 이 스레드의 GPU 번호",TEAL),
    ln("cudaSetDevice(device);         // 이 스레드는 이 GPU 담당",TEXT),
  ],{fontSize:13.5});
  s.addText([
    ln("__thread",GREEN,{bold:true,breakLine:false}),
    ln(" (스레드 지역 저장소, TLS)를 붙이면 같은 이름의 변수라도 ",TEXT,{breakLine:false}),
    ln("스레드마다 별도 값",TEAL,{bold:true,breakLine:false}),
    ln("을 가집니다.",TEXT,{breakLine:true}),
  ],{x:M,y:4.5,w:W-2*M,h:0.7,fontFace:KFONT,fontSize:17,color:TEXT,margin:0,valign:"top"});
  s.addShape(p.ShapeType.roundRect,{x:M,y:5.3,w:W-2*M,h:1.1,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
  s.addText([
    ln("만약 일반 전역 변수였다면? ",AMBER,{bold:true,breakLine:false}),
    ln("여러 스레드가 같은 gpu_idx를 덮어써 로그의 GPU 번호가 뒤죽박죽됩니다. TLS가 이를 막습니다.",TEXT,{breakLine:true}),
  ],{x:M+0.3,y:5.5,w:W-2*M-0.6,h:0.75,fontFace:KFONT,fontSize:15,color:TEXT,margin:0,valign:"middle"});
  s.addNotes("TLS 개념: 로그 접두사 [host][gpu_idx]가 스레드마다 올바른 GPU를 가리키는 이유.");
})();

// =====================================================================
// Slide 5 — portability macro (code)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"4","이식성 · 하나의 코드, 두 백엔드");
  codePanel(s,M,1.7,W-2*M,2.9,[
    ln("// cuda_memtest.h:53                    // 토큰 붙이기(##)로 이름 생성",MUTED),
    ln("#if defined(__CUDACC__)                  // NVIDIA CUDA 컴파일러",TEXT),
    ln("#  define MEMTEST_API_PREFIX(name) cuda##name   // Malloc → cudaMalloc",GREEN),
    ln("#elif defined(__HIP__)                   // AMD HIP 컴파일러",TEXT),
    ln("#  define MEMTEST_API_PREFIX(name) hip##name    // Malloc → hipMalloc",TEAL),
    ln("#endif",TEXT),
  ],{fontSize:13.5});
  const cw=(W-2*M-0.4)/2;
  s.addShape(p.ShapeType.roundRect,{x:M,y:4.85,w:cw,h:1.5,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
  s.addText("코드는 이렇게 한 벌만",{x:M+0.25,y:5.0,w:cw-0.5,h:0.4,fontFace:KFONT,fontSize:15,bold:true,color:GREEN,margin:0});
  s.addText("MEMTEST_API_PREFIX(Malloc)(&ptr, n);",{x:M+0.25,y:5.45,w:cw-0.5,h:0.5,fontFace:MONO,fontSize:12.5,color:TEXT,margin:0});
  s.addText("cudaMalloc을 직접 쓴 곳은 거의 없습니다.",{x:M+0.25,y:5.9,w:cw-0.5,h:0.4,fontFace:KFONT,fontSize:13,color:MUTED,margin:0});
  s.addShape(p.ShapeType.roundRect,{x:M+cw+0.4,y:4.85,w:cw,h:1.5,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
  s.addText("컴파일러가 결정",{x:M+cw+0.65,y:5.0,w:cw-0.5,h:0.4,fontFace:KFONT,fontSize:15,bold:true,color:TEAL,margin:0});
  s.addText("NVIDIA → cudaMalloc\nAMD → hipMalloc",{x:M+cw+0.65,y:5.45,w:cw-0.5,h:0.8,fontFace:MONO,fontSize:13,color:TEXT,margin:0,valign:"top",lineSpacingMultiple:1.2});
  s.addNotes("## 토큰 페이스팅으로 cuda/hip 접두사를 붙임. 코드를 두 벌 관리하지 않는 이식성의 핵심.");
})();

// =====================================================================
// Slide 6 — capstone options (3 cards)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"5","캡스톤 · 나만의 테스트 추가");
  s.addText("배운 것을 종합해 새 메모리 테스트를 구현합니다. 세 옵션 중 하나를 골라 발표하세요.",{x:M,y:1.6,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:16,color:MUTED,margin:0});
  const opts=[
    ["A","체커보드 패턴 테스트","★☆☆","0xAAAA / 0x5555 교차 패턴 새 테스트 추가","세션 1·2 활용",GREEN],
    ["B","Test 3 성능 개선","★★☆","느린 Test 3을 Test 10 스타일로 재작성·대역폭 비교","세션 3 활용",TEAL],
    ["C","오류 히스토그램","★★★","오류 주소의 비트 패턴을 집계해 출력","세션 4 활용",AMBER],
  ];
  const cw=(W-2*M-0.8)/3;
  opts.forEach((o,i)=>{
    const x=M+i*(cw+0.4);
    s.addShape(p.ShapeType.roundRect,{x,y:2.35,w:cw,h:3.9,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(o[0],{x:x+0.28,y:2.6,w:0.9,h:0.8,align:"center",valign:"middle",fontFace:MONO,fontSize:26,bold:true,color:BG,fill:{color:o[5]},shape:p.ShapeType.roundRect,rectRadius:0.1});
    s.addText(o[2],{x:x+cw-1.6,y:2.6,w:1.35,h:0.8,align:"right",valign:"middle",fontFace:KFONT,fontSize:16,color:o[5],margin:0});
    s.addText(o[1],{x:x+0.28,y:3.55,w:cw-0.5,h:0.8,fontFace:KFONT,fontSize:18,bold:true,color:TEXT,margin:0,valign:"top"});
    s.addText(o[3],{x:x+0.28,y:4.35,w:cw-0.5,h:1.2,fontFace:KFONT,fontSize:14,color:MUTED,margin:0,valign:"top",lineSpacingMultiple:1.15});
    s.addText(o[4],{x:x+0.28,y:5.7,w:cw-0.5,h:0.4,fontFace:KFONT,fontSize:13,bold:true,color:o[5],margin:0});
  });
  s.addText("실습 가이드는 가장 쉬운 옵션 A(체커보드)를 단계별로 제공합니다 — move_inv_test를 재사용합니다.",{x:M,y:6.45,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:13,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("옵션 A는 40~50분이면 충분. 자신 있는 참석자에게만 B·C 권장.");
})();

// =====================================================================
// Slide 7 — capstone A walkthrough (steps)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"6","옵션 A · 체커보드 테스트 4단계");
  const steps=[
    ["1","새 테스트 함수 작성","move_inv_test에 p1=0xAAAAAAAA, p2=0x55555555 전달 (tests.cpp)",GREEN],
    ["2","테스트 등록","cuda_memtests[] 배열 끝에 {함수, 설명, 1} 추가 (tests.cpp:1555)",TEAL],
    ["3","배열 크기 상수 고치기","extern cuda_memtests[11] → [12] (cuda_memtest.cpp:57) ⚠️",RED],
    ["4","빌드 & 실행","--list_tests 로 확인 후 --enable_test 11 로 실행",AMBER],
  ];
  let y=1.75;
  steps.forEach(st=>{
    s.addShape(p.ShapeType.roundRect,{x:M,y,w:W-2*M,h:1.02,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(st[0],{x:M+0.25,y:y+0.19,w:0.64,h:0.64,align:"center",valign:"middle",fontFace:MONO,fontSize:22,bold:true,color:BG,fill:{color:st[3]},shape:p.ShapeType.roundRect,rectRadius:0.1});
    s.addText(st[1],{x:M+1.15,y:y+0.12,w:W-2*M-1.4,h:0.44,fontFace:KFONT,fontSize:18,bold:true,color:TEXT,margin:0,valign:"middle"});
    s.addText(st[2],{x:M+1.15,y:y+0.54,w:W-2*M-1.4,h:0.42,fontFace:KFONT,fontSize:14,color:MUTED,margin:0,valign:"middle"});
    y+=1.16;
  });
  s.addText("⚠️ 3단계를 빠뜨리면 새 테스트가 목록에 안 보입니다 — main 쪽 DIM()이 이 상수를 셉니다(cuda_memtest.cpp:223).",{x:M,y:6.5,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:13,italic:true,color:AMBER,margin:0});
  s.addNotes("3단계 하드코딩 상수가 가장 흔한 실수. 세션 2의 move_inv_test 재사용이 핵심 — 새 커널을 짤 필요 없음.");
})();

// =====================================================================
// Slide 8 — capstone A code sketch
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"7","옵션 A · 코드 스케치");
  codePanel(s,M,1.7,W-2*M,2.35,[
    ln("// 1) tests.cpp — test2 아래에 새 함수",MUTED),
    ln("void test_checkerboard(char* ptr, unsigned int tot_num_blocks) {",TEXT),
    ln("    unsigned int p1 = 0xAAAAAAAA;   // 1010...  ",GREEN),
    ln("    unsigned int p2 = ~p1;          // 0101... = 0x55555555",GREEN),
    ln("    move_inv_test(ptr, tot_num_blocks, p1, p2);  // 세션 2 재사용!",TEAL),
    ln("    move_inv_test(ptr, tot_num_blocks, p2, p1);",TEAL),
    ln("}",TEXT),
  ],{fontSize:13});
  codePanel(s,M,4.2,W-2*M,1.6,[
    ln("// 2) tests.cpp:1555 — 배열에 등록",MUTED),
    ln("{test_checkerboard, (char*)\"Test11 [Checkerboard]\", 1},",GREEN),
    ln("// 3) cuda_memtest.cpp:57 — 크기 상수",MUTED),
    ln("extern cuda_memtest_t cuda_memtests[12];   // 11 → 12",RED),
  ],{fontSize:13});
  s.addText("새 커널을 짤 필요가 없습니다 — 세션 2의 move_inv_test(write→readwrite→read)가 검사를 다 해줍니다. 재사용의 힘!",{x:M,y:5.95,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,margin:0});
  s.addNotes("코드는 놀랄 만큼 짧습니다. 지금까지 배운 조각을 조립하는 것이 핵심 메시지.");
})();

// =====================================================================
// Slide 9 — whole series recap (journey)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"8","전체 여정 되짚기");
  const rows=[
    ["세션 1","스레드 모델 · 첫 커널","__global__, <<<grid,block>>>, blockIdx",GREEN],
    ["세션 2","메모리 모델 · 데이터 이동","cudaMalloc/Memcpy, 쓰기-종료-읽기",TEAL],
    ["세션 3","병렬성 · 성능","threadIdx, 병합, 이벤트로 대역폭",AMBER],
    ["세션 4","원자연산 · 에러","atomicAdd, CUERR, 비동기 오류",RED],
    ["세션 5","멀티 GPU · 이식성 · 프로젝트","pthread, HIP, 나만의 커널",GREEN],
  ];
  let y=1.75;
  rows.forEach(r=>{
    s.addShape(p.ShapeType.roundRect,{x:M,y,w:W-2*M,h:0.86,rectRadius:0.07,fill:{color:CARD},line:{type:"none"}});
    s.addText(r[0],{x:M+0.28,y:y+0.1,w:1.7,h:0.66,valign:"middle",fontFace:MONO,fontSize:17,bold:true,color:r[3],margin:0});
    s.addText(r[1],{x:M+2.1,y:y+0.1,w:4.6,h:0.66,valign:"middle",fontFace:KFONT,fontSize:16,bold:true,color:TEXT,margin:0});
    s.addText(r[2],{x:M+6.9,y:y+0.1,w:W-2*M-7.2,h:0.66,valign:"middle",fontFace:MONO,fontSize:13,color:MUTED,margin:0});
    y+=0.98;
  });
  s.addText("장난감 예제가 아니라 500줄 넘는 실제 프로덕션 코드로 CUDA의 핵심을 모두 훑었습니다.",{x:M,y:6.7,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("한 저장소로 커널부터 멀티 GPU까지 완주. 실제 코드로 배운 것이 이 세미나의 강점.");
})();

// =====================================================================
// Slide 10 — where to go next / closing (dark)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  s.addText("수고하셨습니다 🎉",{x:M,y:1.5,w:11,h:0.7,fontFace:KFONT,fontSize:20,bold:true,color:GREEN,margin:0});
  s.addText("이제 실제 CUDA 코드를 읽고 고칠 수 있습니다",{x:M,y:2.15,w:12.2,h:1.0,fontFace:KFONT,fontSize:32,bold:true,color:TEXT,margin:0});
  s.addText("다음 단계로 나아가려면:",{x:M,y:3.4,w:12,h:0.5,fontFace:KFONT,fontSize:17,color:TEXT,margin:0});
  s.addText([
    ln("공유 메모리(shared memory)·워프(warp) 심화 — 타일링, 리덕션",MUTED,{breakLine:true,bullet:{code:"2022"},paraSpaceAfter:8}),
    ln("Nsight 프로파일러로 실제 병목 분석",MUTED,{breakLine:true,bullet:{code:"2022"},paraSpaceAfter:8}),
    ln("NVIDIA CUDA C++ Programming Guide · cuda-samples 저장소",MUTED,{breakLine:true,bullet:{code:"2022"},paraSpaceAfter:8}),
    ln("이 저장소의 원논문: doc/ 의 SAAHPC 2009 (한국어 EPUB 번역본 포함)",TEAL,{breakLine:true,bullet:{code:"2022"}}),
  ],{x:M,y:3.95,w:12,h:2.2,fontFace:KFONT,fontSize:16,color:TEXT,margin:0,valign:"top"});
  s.addText("캡스톤 프로젝트를 완성해 발표해 주세요.",{x:M,y:6.5,w:12,h:0.4,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,margin:0});
  s.addNotes("마무리. 캡스톤 발표로 세미나를 닫습니다. 심화 학습 경로를 안내하세요.");
})();

p.writeFile({fileName:"세션5_멀티GPU_이식성_프로젝트.pptx"}).then(f=>console.log("wrote",f));
