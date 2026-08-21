const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.layout = "LAYOUT_WIDE"; // 13.3 x 7.5

// ---- palette (shared with Sessions 1-5) ----
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
  s.addText("CUDA 세미나 · 세션 0 (오리엔테이션)",{x:M,y:2.0,w:10,h:0.5,fontFace:KFONT,fontSize:18,bold:true,color:GREEN,margin:0});
  s.addText("환경 구축과 큰 그림",{x:M,y:2.5,w:12,h:1.2,fontFace:KFONT,fontSize:48,bold:true,color:TEXT,margin:0});
  s.addText("cuda_memtest로 배우는 실전 GPU 프로그래밍 — 시작하기",{x:M,y:3.85,w:12,h:0.5,fontFace:KFONT,fontSize:20,color:MUTED,margin:0});
  codePanel(s,M,4.7,9.2,1.5,[
    ln("$ cmake -DCMAKE_CUDA_ARCHITECTURES=86 ..",GREEN),
    ln("$ make -j",GREEN),
    ln("$ ./cuda_memtest --stress --num_iterations 100 --num_passes 1",TEAL),
  ],{fontSize:13.5});
  s.addText("입문 과정 · C를 아는 개발자 대상 · 다회차 시리즈의 출발점",{x:M,y:6.35,w:12,h:0.4,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,margin:0});
  s.addNotes("첫 만남. 오늘은 코드를 몰라도 됩니다. 모두가 빌드·실행에 성공하는 것이 목표.");
})();

// =====================================================================
// Slide 2 — what is this seminar (teaching philosophy + roadmap)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"1","이 세미나는?");
  s.addText([
    ln("장난감 예제가 아니라, ",TEXT,{breakLine:false}),
    ln("실제로 동작하는 500줄+ 프로덕션 코드",GREEN,{bold:true,breakLine:false}),
    ln("(cuda_memtest)를 교재로 CUDA를 배웁니다.",TEXT,{breakLine:true}),
  ],{x:M,y:1.6,w:W-2*M,h:0.7,fontFace:KFONT,fontSize:18,color:TEXT,margin:0,valign:"top"});
  const road=[
    ["1","스레드 모델 · 첫 커널","__global__, <<<grid,block>>>",GREEN],
    ["2","메모리 모델 · 데이터 이동","cudaMalloc/Memcpy, 이동 반전",TEAL],
    ["3","병렬성 · 성능","210만 스레드, 대역폭, 병합",AMBER],
    ["4","원자연산 · 에러 핸들링","atomicAdd, CUERR, 비동기",RED],
    ["5","멀티 GPU · 이식성 · 프로젝트","pthread, HIP, 나만의 커널",GREEN],
  ];
  let y=2.5;
  road.forEach(r=>{
    s.addShape(p.ShapeType.roundRect,{x:M,y,w:W-2*M,h:0.76,rectRadius:0.06,fill:{color:CARD},line:{type:"none"}});
    s.addText("세션 "+r[0],{x:M+0.25,y:y+0.08,w:1.5,h:0.6,valign:"middle",fontFace:MONO,fontSize:15,bold:true,color:r[3],margin:0});
    s.addText(r[1],{x:M+1.9,y:y+0.08,w:5.0,h:0.6,valign:"middle",fontFace:KFONT,fontSize:16,bold:true,color:TEXT,margin:0});
    s.addText(r[2],{x:M+7.0,y:y+0.08,w:W-2*M-7.3,h:0.6,valign:"middle",fontFace:MONO,fontSize:13,color:MUTED,margin:0});
    y+=0.88;
  });
  s.addNotes("전체 로드맵을 먼저 보여줘 학습자가 어디로 가는지 알게 합니다. 세션 0은 그 출발점.");
})();

// =====================================================================
// Slide 3 — GPU computing (why GPUs)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"2","GPU 컴퓨팅이란");
  const colW=(W-2*M-0.4)/2;
  const mkCol=(x,title,color,big,rows)=>{
    s.addShape(p.ShapeType.roundRect,{x,y:1.7,w:colW,h:4.7,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(title,{x:x+0.3,y:1.95,w:colW-0.6,h:0.5,fontFace:KFONT,fontSize:20,bold:true,color:color,margin:0});
    s.addText(big,{x:x+0.3,y:2.5,w:colW-0.6,h:0.6,fontFace:KFONT,fontSize:16,italic:true,color:MUTED,margin:0});
    s.addText(rows.map(r=>ln(r,TEXT,{breakLine:true,bullet:{code:"2022"},paraSpaceAfter:10})),{x:x+0.3,y:3.2,w:colW-0.6,h:3.0,fontFace:KFONT,fontSize:15.5,color:TEXT,valign:"top",margin:0});
  };
  mkCol(M,"CPU",TEAL,"소수 정예",["강력한 코어 수 개~수십 개","복잡한 순차 작업에 최적","지연시간(latency) 최적화"]);
  mkCol(M+colW+0.4,"GPU",GREEN,"대규모 병렬",["단순한 코어 수천 개","같은 연산을 대량 데이터에 동시에","처리량(throughput) 최적화","SIMT — 명령 하나, 스레드 여럿"]);
  s.addText("cuda_memtest는 4 GB 메모리를 검사합니다 — 수천 스레드가 동시에 일해야 현실적 시간에 끝납니다.",{x:M,y:6.55,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("세션 1에서 더 깊이 다룹니다. 여기선 'GPU=대규모 병렬'이라는 큰 그림만.");
})();

// =====================================================================
// Slide 3b — Host <-> Device hardware block diagram
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"2","CUDA 하드웨어 구조 · 호스트와 디바이스");
  s.addText("CPU(호스트)와 GPU(디바이스)는 별개의 메모리를 가지며, 데이터는 PCIe를 통해 오갑니다.",{x:M,y:1.5,w:W-2*M,h:0.45,fontFace:KFONT,fontSize:15,color:MUTED,margin:0});

  // Host box (left)
  const hx=M, hy=2.1, hw=4.7, hh=4.05;
  s.addShape(p.ShapeType.roundRect,{x:hx,y:hy,w:hw,h:hh,rectRadius:0.07,fill:{color:CARD},line:{color:TEAL,width:1.5}});
  s.addText("호스트(Host)",{x:hx+0.25,y:hy+0.15,w:hw-0.5,h:0.4,fontFace:KFONT,fontSize:16,bold:true,color:TEAL,margin:0});
  s.addShape(p.ShapeType.roundRect,{x:hx+0.35,y:hy+0.75,w:hw-0.7,h:1.35,rectRadius:0.05,fill:{color:CODEBG},line:{color:LINE,width:1}});
  s.addText("CPU",{x:hx+0.35,y:hy+0.95,w:hw-0.7,h:0.5,align:"center",fontFace:KFONT,fontSize:20,bold:true,color:TEXT,margin:0});
  s.addText("강력한 코어 수 개~수십 개",{x:hx+0.35,y:hy+1.5,w:hw-0.7,h:0.5,align:"center",fontFace:KFONT,fontSize:12,color:MUTED,margin:0});
  s.addShape(p.ShapeType.roundRect,{x:hx+0.35,y:hy+2.3,w:hw-0.7,h:1.4,rectRadius:0.05,fill:{color:CODEBG},line:{color:LINE,width:1}});
  s.addText("호스트 메모리 (RAM)",{x:hx+0.35,y:hy+2.55,w:hw-0.7,h:0.5,align:"center",fontFace:KFONT,fontSize:16,bold:true,color:TEXT,margin:0});
  s.addText("시스템 메인 메모리",{x:hx+0.35,y:hy+3.1,w:hw-0.7,h:0.4,align:"center",fontFace:KFONT,fontSize:12,color:MUTED,margin:0});

  // Device box (right)
  const dx=W-M-4.7, dy=2.1, dw=4.7, dh=4.05;
  s.addShape(p.ShapeType.roundRect,{x:dx,y:dy,w:dw,h:dh,rectRadius:0.07,fill:{color:CARD},line:{color:GREEN,width:1.5}});
  s.addText("디바이스(Device) · GPU",{x:dx+0.25,y:dy+0.15,w:dw-0.5,h:0.4,fontFace:KFONT,fontSize:16,bold:true,color:GREEN,margin:0});
  // SM small boxes 4개
  const smw=(dw-0.7-3*0.15)/4;
  for(let i=0;i<4;i++){
    const x=dx+0.35+i*(smw+0.15);
    s.addShape(p.ShapeType.roundRect,{x,y:dy+0.75,w:smw,h:1.35,rectRadius:0.04,fill:{color:CODEBG},line:{color:LINE,width:1}});
    s.addText("SM",{x,y:dy+1.05,w:smw,h:0.4,align:"center",fontFace:MONO,fontSize:13,bold:true,color:GREEN,margin:0});
    s.addText("코어\n다수",{x,y:dy+1.45,w:smw,h:0.55,align:"center",fontFace:KFONT,fontSize:9.5,color:MUTED,margin:0,lineSpacingMultiple:1.0});
  }
  s.addText("SM(스트리밍 멀티프로세서) 수십 개 · 코어 수천 개",{x:dx+0.35,y:dy+2.12,w:dw-0.7,h:0.3,align:"center",fontFace:KFONT,fontSize:10.5,italic:true,color:MUTED,margin:0});
  s.addShape(p.ShapeType.roundRect,{x:dx+0.35,y:dy+2.55,w:dw-0.7,h:1.15,rectRadius:0.05,fill:{color:"1A2418"},line:{color:GREEN,width:1}});
  s.addText("디바이스 메모리 (VRAM)",{x:dx+0.35,y:dy+2.72,w:dw-0.7,h:0.45,align:"center",fontFace:KFONT,fontSize:15,bold:true,color:TEXT,margin:0});
  s.addText("cuda_memtest가 검사하는 대상 (예: 4 GB)",{x:dx+0.35,y:dy+3.15,w:dw-0.7,h:0.4,align:"center",fontFace:KFONT,fontSize:11.5,color:GREEN,margin:0});

  // PCIe connector (double arrow) between host and device
  const cx0=hx+hw+0.05, cx1=dx-0.05, cy=dy+1.55;
  s.addShape(p.ShapeType.line,{x:cx0,y:cy,w:cx1-cx0,h:0,line:{color:AMBER,width:3,beginArrowType:"triangle",endArrowType:"triangle"}});
  s.addText("PCIe",{x:cx0,y:cy-0.5,w:cx1-cx0,h:0.35,align:"center",fontFace:MONO,fontSize:13,bold:true,color:AMBER,margin:0});
  s.addText("cudaMemcpy",{x:cx0-0.1,y:cy+0.15,w:cx1-cx0+0.2,h:0.35,align:"center",fontFace:MONO,fontSize:11,color:MUTED,margin:0});

  s.addText("핵심: CPU는 GPU 메모리를 직접 못 읽습니다. 반드시 cudaMemcpy로 복사해야 합니다 (세션 2에서 자세히).",{x:M,y:6.35,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:13,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("호스트/디바이스가 분리된 메모리를 갖는다는 것이 CUDA의 근본 구조. PCIe를 통한 cudaMemcpy가 다리 역할.");
})();

// =====================================================================
// Slide 3c — CUDA software stack (layered)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"2","CUDA 소프트웨어 스택(software stack)");
  s.addText("내 코드는 여러 계층 위에서 동작합니다. cudaMalloc 한 줄이 아래로 내려가 GPU까지 도달합니다.",{x:M,y:1.5,w:W-2*M,h:0.45,fontFace:KFONT,fontSize:15,color:MUTED,margin:0});

  const layers=[
    ["애플리케이션(Application)","cuda_memtest — main(), tests.cpp의 커널·테스트","GREEN","우리가 짜는 코드"],
    ["CUDA 런타임 API(Runtime)","cudaMalloc · cudaMemcpy · cudaSetDevice · <<<>>>","TEAL","libcudart · CMake의 find_package(CUDAToolkit)이 링크"],
    ["CUDA 드라이버 / GPU 드라이버","커널을 GPU에 제출 · 메모리·스케줄 관리","AMBER","libcuda · 커널 모드 드라이버"],
    ["GPU 하드웨어(Hardware)","SM(스트리밍 멀티프로세서) · VRAM","MUTED","실제 연산·저장이 일어나는 곳"],
  ];
  const cmap={GREEN:GREEN,TEAL:TEAL,AMBER:AMBER,MUTED:MUTED};
  const bx=M, bw=8.4, bh=1.02, gap=0.06; let y=2.1;
  layers.forEach((L,i)=>{
    const col=cmap[L[2]];
    s.addShape(p.ShapeType.roundRect,{x:bx,y,w:bw,h:bh,rectRadius:0.06,fill:{color:CARD},line:{color:col,width:1.5}});
    s.addText(L[0],{x:bx+0.3,y:y+0.14,w:bw-0.6,h:0.42,fontFace:KFONT,fontSize:16,bold:true,color:col,margin:0});
    s.addText(L[1],{x:bx+0.3,y:y+0.56,w:bw-0.6,h:0.36,fontFace:MONO,fontSize:11.5,color:TEXT,margin:0});
    // down chevron between layers
    if(i<layers.length-1) s.addShape(p.ShapeType.line,{x:bx+bw/2,y:y+bh,w:0,h:gap+0.02,line:{color:MUTED,width:1.5,endArrowType:"triangle"}});
    y+=bh+gap;
  });
  // side annotations
  const ax=bx+bw+0.4; let ay=2.1;
  layers.forEach(L=>{
    s.addText(L[3],{x:ax,y:ay+0.1,w:W-M-ax,h:bh-0.2,valign:"middle",fontFace:KFONT,fontSize:12.5,color:MUTED,margin:0,lineSpacingMultiple:1.05});
    ay+=bh+gap;
  });
  // direction arrow label
  s.addText("cudaMalloc / <<<>>> 호출이\n이 방향으로 내려갑니다 ↓",{x:ax,y:6.35,w:W-M-ax,h:0.5,fontFace:KFONT,fontSize:11.5,italic:true,color:TEAL,margin:0,lineSpacingMultiple:1.05});

  s.addText("세션 1~4에서 만나는 cudaMalloc·cudaMemcpy·<<<>>>는 모두 '런타임 API' 계층입니다.",{x:M,y:6.5,w:bw,h:0.4,fontFace:KFONT,fontSize:12.5,italic:true,color:MUTED,margin:0});
  s.addNotes("SW 스택: 내 코드→런타임→드라이버→HW. CMake가 런타임(CUDAToolkit)을 링크. cudaMalloc 한 줄이 아래로 내려간다.");
})();

// =====================================================================
// Slide 4 — what cuda_memtest solves (paper background)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"3","cuda_memtest가 푸는 문제");
  s.addText("GPU 메모리도 고장 납니다. 이 도구는 그 고장을 찾아냅니다.",{x:M,y:1.6,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:17,color:MUTED,margin:0});
  const cards=[
    ["하드 오류(hard error)","영구적 고장",GREEN,"제조 결함·노후화로 특정 비트가 항상 틀림. 재부팅해도 그대로."],
    ["소프트 오류(soft error)","일시적 뒤집힘",TEAL,"우주 방사선 등으로 비트가 잠깐 뒤집힘. 다시 쓰면 정상."],
  ];
  const cw=(W-2*M-0.4)/2;
  cards.forEach((c,i)=>{
    const x=M+i*(cw+0.4);
    s.addShape(p.ShapeType.roundRect,{x,y:2.35,w:cw,h:2.6,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(c[0],{x:x+0.3,y:2.6,w:cw-0.6,h:0.5,fontFace:KFONT,fontSize:20,bold:true,color:c[2],margin:0});
    s.addText(c[1],{x:x+0.3,y:3.1,w:cw-0.6,h:0.45,fontFace:KFONT,fontSize:15,bold:true,color:MUTED,margin:0});
    s.addText(c[3],{x:x+0.3,y:3.65,w:cw-0.6,h:1.1,fontFace:KFONT,fontSize:15,color:TEXT,margin:0,valign:"top",lineSpacingMultiple:1.2});
  });
  s.addShape(p.ShapeType.roundRect,{x:M,y:5.2,w:W-2*M,h:1.25,rectRadius:0.08,fill:{color:"1A2418"},line:{type:"none"}});
  s.addText([
    ln("실제 배경: ",AMBER,{bold:true,breakLine:false}),
    ln("이 도구는 NCSA의 대규모 GPU 클러스터(500+ Tesla GPU) 운영에서 나왔습니다. 원논문은 doc/ 폴더에 있고, 한국어 EPUB 번역본도 함께 제공됩니다.",TEXT,{breakLine:true}),
  ],{x:M+0.3,y:5.4,w:W-2*M-0.6,h:0.9,fontFace:KFONT,fontSize:15,color:TEXT,margin:0,valign:"middle"});
  s.addNotes("하드/소프트 오류 구분은 세미나 전체의 배경. doc/의 SAAHPC 2009 논문(한국어 EPUB)을 안내.");
})();

// =====================================================================
// Slide 5 — build system (CMake)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"4","빌드 시스템 · CMake");
  codePanel(s,M,1.7,W-2*M,2.0,[
    ln("mkdir build && cd build",TEXT),
    ln("cmake -DCMAKE_CUDA_ARCHITECTURES=86 ..   # 본인 GPU 번호로 교체",GREEN),
    ln("make -j",TEXT),
    ln("",TEXT),
    ln("# AMD GPU면 HIP 백엔드",MUTED),
    ln("cmake -DCUDA_MEMTEST_BACKEND=hip -DGPU_TARGETS=gfx90a ..",TEAL),
  ],{fontSize:13.5});
  const cards=[
    ["CMAKE_CUDA_ARCHITECTURES","GPU 세대(sm_XX)를 지정. 세대마다 기계어가 다름.",GREEN],
    ["CUDA_MEMTEST_BACKEND","cuda / hip 백엔드 선택 (CMakeLists.txt:22).",TEAL],
  ];
  const cw=(W-2*M-0.4)/2;
  cards.forEach((c,i)=>{
    const x=M+i*(cw+0.4);
    s.addShape(p.ShapeType.roundRect,{x,y:4.05,w:cw,h:1.9,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(c[0],{x:x+0.25,y:4.25,w:cw-0.5,h:0.7,fontFace:MONO,fontSize:14,bold:true,color:c[2],margin:0,valign:"top"});
    s.addText(c[1],{x:x+0.25,y:4.95,w:cw-0.5,h:0.9,fontFace:KFONT,fontSize:14,color:TEXT,margin:0,valign:"top",lineSpacingMultiple:1.2});
  });
  s.addText("본인 GPU의 sm_XX 번호: developer.nvidia.com/cuda-gpus  (예: T4→75, RTX30→86, RTX40→89, A100→80)",{x:M,y:6.15,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:13,italic:true,color:MUTED,margin:0});
  s.addNotes("아키텍처 번호를 왜 지정하는지(세대별 기계어)를 짚어주세요. 빌드 실패의 흔한 원인.");
})();

// =====================================================================
// Slide 6 — first run
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"5","첫 실행");
  const steps=[
    ["테스트 목록 보기","./cuda_memtest --list_tests","11개 메모리 테스트가 나옴",GREEN],
    ["빠른 스트레스 실행","./cuda_memtest --stress --num_iterations 100 --num_passes 1","GPU 메모리를 실제로 검사",TEAL],
    ["상태 점검 스크립트","./sanity_check.sh 0","특정 GPU 건강 빠르게 점검",AMBER],
  ];
  let y=1.75;
  steps.forEach(st=>{
    s.addShape(p.ShapeType.roundRect,{x:M,y,w:W-2*M,h:1.35,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(st[0],{x:M+0.3,y:y+0.18,w:3.6,h:1.0,valign:"middle",fontFace:KFONT,fontSize:17,bold:true,color:st[3],margin:0});
    s.addShape(p.ShapeType.roundRect,{x:M+4.0,y:y+0.22,w:W-2*M-5.9,h:0.55,rectRadius:0.04,fill:{color:CODEBG},line:{color:LINE,width:1}});
    s.addText(st[1],{x:M+4.2,y:y+0.22,w:W-2*M-6.3,h:0.55,valign:"middle",fontFace:MONO,fontSize:12,color:TEXT,margin:0});
    s.addText(st[2],{x:M+4.0,y:y+0.8,w:W-2*M-4.3,h:0.4,valign:"middle",fontFace:KFONT,fontSize:13,color:MUTED,margin:0});
    y+=1.5;
  });
  s.addText("오류 없이 끝나면 성공입니다 — 건강한 GPU는 '아무 일도 안 일어난 것처럼' 조용히 끝납니다.",{x:M,y:6.4,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("'아무 일도 안 일어나는데요?'가 정상. 세션 2에서 일부러 오류를 주입해 검출을 봅니다.");
})();

// =====================================================================
// Slide 7 — 11 tests at a glance
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"6","11개 테스트 한눈에");
  const tests=[
    ["0","Walking 1 bit","주소 배선"],["1","Own address","주소 검사"],
    ["2","Moving inv. (0/1)","고착 오류"],["3","Moving inv. (8bit)","고착 오류"],
    ["4","Moving inv. (random)","데이터 민감"],["5","Block move","데이터 민감"],
    ["6","Moving inv. (shift)","데이터 민감"],["7","Random sequence","데이터 민감"],
    ["8","Modulo 20","데이터 민감"],["9","Bit fade (3h)","데이터 보존"],
    ["10","Memory stress","스트레스·소프트"],
  ];
  const cols=3, cw=(W-2*M-0.6)/cols, rh=0.86;
  tests.forEach((t,i)=>{
    const c=i%cols, r=Math.floor(i/cols);
    const x=M+c*(cw+0.3), y=1.75+r*(rh+0.16);
    const hot = t[0]==="10"||t[0]==="9";
    s.addShape(p.ShapeType.roundRect,{x,y,w:cw,h:rh,rectRadius:0.06,fill:{color:CARD},line:{type:"none"}});
    s.addText(t[0],{x:x+0.2,y:y+0.14,w:0.72,h:0.58,align:"center",valign:"middle",fontFace:MONO,fontSize:16,bold:true,color:BG,fill:{color:hot?AMBER:TEAL},shape:p.ShapeType.roundRect,rectRadius:0.06});
    s.addText(t[1],{x:x+1.05,y:y+0.1,w:cw-1.2,h:0.42,valign:"middle",fontFace:KFONT,fontSize:13.5,bold:true,color:TEXT,margin:0});
    s.addText(t[2],{x:x+1.05,y:y+0.46,w:cw-1.2,h:0.34,valign:"middle",fontFace:KFONT,fontSize:11.5,color:MUTED,margin:0});
  });
  s.addText("Test 9(bit fade)는 3시간 걸려 기본 비활성. Test 10(스트레스)은 소프트 오류·급성 하드 오류에 특히 유용.",{x:M,y:6.5,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:13,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("각 테스트가 서로 다른 종류의 오류를 노림. 세부 알고리즘은 세션 2~3에서 다룹니다.");
})();

// =====================================================================
// Slide 8 — program flow (code map)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"7","프로그램의 큰 흐름");
  const steps=[
    ["main()","GPU 개수 확인","cudaGetDeviceCount (cuda_memtest.cpp:299)",GREEN],
    ["thread_func","GPU마다 스레드","각 GPU를 cudaSetDevice로 선택",TEAL],
    ["run_tests","테스트 실행","cuda_memtests[] 배열을 순회",AMBER],
    ["각 테스트","커널 실행","write → read → check (세션 1~2)",GREEN],
  ];
  const bw=(W-2*M-0.9)/4, gap=0.3, y0=2.0, bh=2.6;
  steps.forEach((st,i)=>{
    const x=M+i*(bw+gap);
    s.addShape(p.ShapeType.roundRect,{x,y:y0,w:bw,h:bh,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(String(i+1),{x:x+0.2,y:y0+0.2,w:0.6,h:0.6,align:"center",valign:"middle",fontFace:MONO,fontSize:18,bold:true,color:BG,fill:{color:st[3]},shape:p.ShapeType.roundRect,rectRadius:0.3});
    s.addText(st[0],{x:x+0.2,y:y0+0.95,w:bw-0.4,h:0.5,fontFace:MONO,fontSize:14,bold:true,color:st[3],margin:0});
    s.addText(st[1],{x:x+0.2,y:y0+1.4,w:bw-0.4,h:0.4,fontFace:KFONT,fontSize:14,bold:true,color:TEXT,margin:0});
    s.addText(st[2],{x:x+0.2,y:y0+1.8,w:bw-0.4,h:0.7,fontFace:KFONT,fontSize:11.5,color:MUTED,margin:0,valign:"top",lineSpacingMultiple:1.1});
    if(i<3) s.addText("→",{x:x+bw+0.01,y:y0+1.0,w:gap-0.02,h:0.6,align:"center",valign:"middle",fontFace:MONO,fontSize:20,bold:true,color:MUTED,margin:0});
  });
  s.addText("이 흐름의 각 단계를 앞으로 5회에 걸쳐 안쪽까지 파고듭니다. 오늘은 지도만 봐 두세요.",{x:M,y:5.2,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:15,color:TEXT,align:"center",margin:0});
  s.addNotes("전체 코드 지도. main→thread_func→run_tests→커널. 각 단계가 앞으로의 세션에 대응.");
})();

// =====================================================================
// Slide 9 — Lab preview
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"8","실습 미리보기 · Lab 0");
  const labs=[
    ["Ex 0","내 환경 점검","nvidia-smi · nvcc · cmake 버전 확인",TEAL],
    ["Ex 1","빌드하기","cmake + make로 cuda_memtest 생성",GREEN],
    ["Ex 2","테스트 목록 보기","--list_tests 로 11개 테스트 확인",AMBER],
    ["Ex 3","첫 실행","--stress 로 GPU 메모리 실제 검사",GREEN],
  ];
  const cw=(W-2*M-0.5)/2, ch=2.15;
  labs.forEach((l,i)=>{
    const x=M+(i%2)*(cw+0.5), y=1.7+Math.floor(i/2)*(ch+0.35);
    s.addShape(p.ShapeType.roundRect,{x,y,w:cw,h:ch,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(l[0],{x:x+0.28,y:y+0.22,w:1.3,h:0.55,align:"center",valign:"middle",fontFace:MONO,fontSize:17,bold:true,color:BG,fill:{color:l[3]},shape:p.ShapeType.roundRect,rectRadius:0.08});
    s.addText(l[1],{x:x+1.75,y:y+0.24,w:cw-2.0,h:0.9,fontFace:KFONT,fontSize:17,bold:true,color:TEXT,margin:0,valign:"top"});
    s.addText(l[2],{x:x+0.28,y:y+1.2,w:cw-0.56,h:0.8,fontFace:KFONT,fontSize:14,color:MUTED,margin:0,valign:"top",lineSpacingMultiple:1.1});
  });
  s.addText("GPU가 없다면 Google Colab(런타임→GPU)으로 진행할 수 있습니다.",{x:M,y:6.5,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:13,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("세션 0의 핵심 목표: 전원 빌드·실행 성공. 환경 문제를 여기서 다 해결.");
})();

// =====================================================================
// Slide 10 — next session (dark)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  s.addText("다음 세션 예고",{x:M,y:1.6,w:10,h:0.5,fontFace:KFONT,fontSize:18,bold:true,color:GREEN,margin:0});
  s.addText("세션 1 · CUDA 스레드 모델과 첫 커널",{x:M,y:2.15,w:12.2,h:1.0,fontFace:KFONT,fontSize:34,bold:true,color:TEXT,margin:0});
  s.addText([
    ln("환경이 준비됐으니, 이제 ",TEXT,{breakLine:false}),
    ln("코드 안으로",TEAL,{bold:true,breakLine:false}),
    ln(" 들어갑니다.",TEXT,{breakLine:true}),
  ],{x:M,y:3.3,w:12,h:0.6,fontFace:KFONT,fontSize:20,color:TEXT,margin:0});
  s.addText([
    ln("__global__ 커널이 무엇인지",MONO,{breakLine:true,color:GREEN,paraSpaceAfter:8}),
    ln("<<<grid, block>>>으로 GPU에 일을 시키는 법",MUTED,{breakLine:true,paraSpaceAfter:8}),
    ln("가장 단순한 커널 kernel_move_inv_write를 한 줄씩 읽기",MUTED,{breakLine:true}),
  ],{x:M,y:4.15,w:12,h:1.4,fontFace:KFONT,fontSize:16,color:TEXT,margin:0,valign:"top"});
  s.addText("실습 랩 0(빌드·실행)을 먼저 완료하고 오세요.",{x:M,y:6.4,w:12,h:0.4,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,margin:0});
  s.addNotes("환경 준비 완료 → 세션 1부터 코드. kernel_move_inv_write로 시작.");
})();

p.writeFile({fileName:"세션0_환경구축과_큰그림.pptx"}).then(f=>console.log("wrote",f));
