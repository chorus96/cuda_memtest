const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.layout = "LAYOUT_WIDE"; // 13.3 x 7.5

// ---- palette (shared with Sessions 1-2) ----
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
  s.addText("CUDA 세미나 · 세션 3",{x:M,y:2.05,w:8,h:0.5,fontFace:KFONT,fontSize:18,bold:true,color:GREEN,margin:0});
  s.addText("병렬성 설계와 성능\n(parallelism & performance)",{x:M,y:2.5,w:11.9,h:2.0,fontFace:KFONT,fontSize:44,bold:true,color:TEXT,lineSpacingMultiple:1.05,margin:0});
  s.addText("같은 문제, 두 개의 커널 — 128 스레드 vs 210만 스레드",{x:M,y:4.7,w:11.5,h:0.5,fontFace:KFONT,fontSize:20,color:MUTED,margin:0});
  codePanel(s,M,5.6,8.9,1.0,[
    ln("test10_kernel_write<<<32768, 64>>>(...)   // 2,097,152 threads",GREEN),
    ln("index = i*blockDim.x + threadIdx.x;       // 이웃 스레드 = 이웃 주소",TEAL),
  ],{fontSize:12.5});
  s.addText("입문 과정 · 세션 2에 이어서",{x:9.7,y:6.05,w:3.0,h:0.4,align:"right",fontFace:KFONT,fontSize:13,color:MUTED,margin:0});
  s.addNotes("Test 0~9의 느린 커널을 충분히 봤으니, 오늘은 같은 문제를 빠르게 푸는 Test 10으로 성능을 배웁니다.");
})();

// =====================================================================
// Slide 2 — objectives
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"1","오늘의 목표");
  const items=[
    ["병렬성의 규모","<<<grid,1>>>(느림)과 <<<32768,64>>>(빠름)의 차이를 설명한다"],
    ["스레드 인덱싱","index = i*blockDim.x + threadIdx.x 패턴을 이해한다"],
    ["시간 측정","CUDA 이벤트(event)로 커널 시간을 재고 대역폭을 계산한다"],
    ["메모리 병합(coalescing)","이웃 스레드가 이웃 주소를 접근할 때 왜 빠른지 안다"],
  ];
  let y=1.7;
  items.forEach((it,i)=>{
    s.addShape(p.ShapeType.roundRect,{x:M,y,w:W-2*M,h:1.06,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(String(i+1),{x:M+0.22,y:y+0.2,w:0.66,h:0.66,align:"center",valign:"middle",fontFace:MONO,fontSize:22,bold:true,color:BG,fill:{color:i%2?TEAL:GREEN},shape:p.ShapeType.roundRect,rectRadius:0.33});
    s.addText(it[0],{x:M+1.15,y:y+0.14,w:W-2*M-1.4,h:0.44,fontFace:KFONT,fontSize:19,bold:true,color:TEXT,margin:0,valign:"middle"});
    s.addText(it[1],{x:M+1.15,y:y+0.55,w:W-2*M-1.4,h:0.42,fontFace:KFONT,fontSize:14,color:MUTED,margin:0,valign:"middle"});
    y+=1.24;
  });
  s.addNotes("오늘은 실습에서 실제로 시간을 재고 대역폭을 측정합니다. 측정이 핵심.");
})();

// =====================================================================
// Slide 3 — the scale contrast (big stat compare)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"2","규모의 대비 · 128 vs 210만");
  const colW=(W-2*M-0.4)/2;
  const mkCol=(x,tag,color,threads,cfg,desc)=>{
    s.addShape(p.ShapeType.roundRect,{x,y:1.7,w:colW,h:4.7,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(tag,{x:x+0.3,y:1.95,w:colW-0.6,h:0.5,fontFace:KFONT,fontSize:18,bold:true,color:color,margin:0});
    s.addText(threads,{x:x+0.3,y:2.5,w:colW-0.6,h:1.1,fontFace:KFONT,fontSize:52,bold:true,color:color,margin:0});
    s.addText("개의 스레드",{x:x+0.3,y:3.6,w:colW-0.6,h:0.4,fontFace:KFONT,fontSize:15,color:MUTED,margin:0});
    s.addText(cfg,{x:x+0.3,y:4.15,w:colW-0.6,h:0.5,fontFace:MONO,fontSize:16,bold:true,color:TEXT,margin:0});
    s.addText(desc,{x:x+0.3,y:4.75,w:colW-0.6,h:1.4,fontFace:KFONT,fontSize:15,color:MUTED,margin:0,valign:"top",lineSpacingMultiple:1.15});
  };
  mkCol(M,"Test 0~9 (세션 1·2)",TEAL,"128","<<<128, 1>>>","블록당 스레드 1개.\n단순·정확하지만 GPU의 수천 코어를 거의 놀림.");
  mkCol(M+colW+0.4,"Test 10 (오늘)",GREEN,"2,097,152","<<<32768, 64>>>","블록 32,768 × 스레드 64.\n수백만 스레드가 동시에 메모리를 두드림.");
  s.addText("약 16,000배 더 많은 스레드 — 같은 \"쓰고·읽어·검사\"를 완전히 다른 규모로 실행합니다.",{x:M,y:6.55,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("이 대비가 오늘의 출발점. 규모만 키운 게 아니라 인덱싱 방식도 바뀝니다(다음 슬라이드).");
})();

// =====================================================================
// Slide 3b — SM / warp hardware block diagram
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"2","실행 하드웨어 · SM과 워프(warp)");
  s.addText("블록은 SM(스트리밍 멀티프로세서)에 배정되고, 스레드는 32개씩 워프(warp)로 묶여 함께 실행됩니다.",{x:M,y:1.5,w:W-2*M,h:0.45,fontFace:KFONT,fontSize:15,color:MUTED,margin:0});

  // GPU box with several SMs (left)
  const gx=M, gy=2.0, gw=6.7, gh=4.3;
  s.addShape(p.ShapeType.roundRect,{x:gx,y:gy,w:gw,h:gh,rectRadius:0.07,fill:{color:CARD},line:{color:GREEN,width:1.5}});
  s.addText("GPU · SM 수십 개",{x:gx+0.25,y:gy+0.14,w:gw-0.5,h:0.4,fontFace:KFONT,fontSize:14,bold:true,color:GREEN,margin:0});
  const smw=1.85, smh=1.5, sgx=0.24, sgy=0.28, sx0=gx+0.32, sy0=gy+0.7;
  let n=0;
  for(let r=0;r<2;r++){
    for(let c=0;c<3;c++){
      const x=sx0+c*(smw+sgx), y=sy0+r*(smh+sgy);
      const hot=(r===0&&c===0);
      s.addShape(p.ShapeType.roundRect,{x,y,w:smw,h:smh,rectRadius:0.05,fill:{color:CODEBG},line:{color:hot?TEAL:LINE,width:hot?2:1}});
      s.addText("SM "+n,{x,y:y+0.14,w:smw,h:0.32,align:"center",fontFace:MONO,fontSize:12,bold:true,color:hot?TEAL:TEXT,margin:0});
      s.addText("워프 스케줄러\n+ 코어 다수",{x,y:y+0.5,w:smw,h:0.85,align:"center",fontFace:KFONT,fontSize:10,color:MUTED,margin:0,lineSpacingMultiple:1.05});
      n++;
    }
  }
  // arrow to warp zoom
  s.addShape(p.ShapeType.line,{x:gx+gw+0.03,y:gy+2.0,w:0.7,h:0,line:{color:TEAL,width:2.5,endArrowType:"triangle"}});
  s.addText("확대",{x:gx+gw-0.05,y:gy+1.55,w:0.8,h:0.35,align:"center",fontFace:KFONT,fontSize:11,color:TEAL,margin:0});

  // warp zoom box (right)
  const zx=gx+gw+0.8, zy=2.0, zw=W-M-(gx+gw+0.8), zh=4.3;
  s.addShape(p.ShapeType.roundRect,{x:zx,y:zy,w:zw,h:zh,rectRadius:0.07,fill:{color:CARD},line:{color:TEAL,width:1.5}});
  s.addText("워프(warp) = 32 스레드",{x:zx+0.25,y:zy+0.14,w:zw-0.5,h:0.4,fontFace:KFONT,fontSize:15,bold:true,color:TEAL,margin:0});
  s.addText("같은 명령을 동시에 실행 (lockstep)",{x:zx+0.25,y:zy+0.56,w:zw-0.5,h:0.32,fontFace:KFONT,fontSize:12,color:MUTED,margin:0});
  // 32 small thread squares (8x4)
  const cols=8, rows=4, cw2=(zw-0.6-(cols-1)*0.1)/cols, ch2=0.42, cgx=0.1, cgy=0.14, cx0=zx+0.3, cy0=zy+1.05;
  for(let i=0;i<32;i++){
    const c=i%cols, r=Math.floor(i/cols);
    const x=cx0+c*(cw2+cgx), y=cy0+r*(ch2+cgy);
    s.addShape(p.ShapeType.roundRect,{x,y,w:cw2,h:ch2,rectRadius:0.03,fill:{color:CODEBG},line:{color:LINE,width:1}});
  }
  s.addText("32개 스레드가 한 덩어리로 움직임",{x:zx+0.3,y:cy0+rows*(ch2+cgy)+0.02,w:zw-0.6,h:0.32,fontFace:KFONT,fontSize:11.5,italic:true,color:MUTED,margin:0});
  s.addShape(p.ShapeType.roundRect,{x:zx+0.3,y:zy+zh-0.95,w:zw-0.6,h:0.78,rectRadius:0.05,fill:{color:"1A2418"},line:{type:"none"}});
  s.addText([
    ln("그래서 blockDim은 32의 배수가 유리",AMBER,{bold:true,breakLine:true}),
    ln("(실습 3에서 블록 크기별 대역폭 측정)",TEXT,{breakLine:true}),
  ],{x:zx+0.45,y:zy+zh-0.85,w:zw-0.9,h:0.6,fontFace:KFONT,fontSize:11.5,color:TEXT,margin:0,valign:"top",lineSpacingMultiple:1.05});
  s.addNotes("SM에 블록 배정 → 워프(32스레드) 단위 실행. blockDim이 32 배수여야 스레드가 낭비 없이 채워짐. 코얼레싱과도 연결.");
})();

// =====================================================================
// Slide 4 — thread self-indexing (code)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"3","스레드가 자기 몫을 계산하는 법");
  codePanel(s,M,1.7,W-2*M,3.3,[
    ln("// test10_kernel_write   tests.cpp:1424",MUTED),
    ln("int avenumber = memsize/(gridDim.x*gridDim.y); // 블록 하나가 맡을 바이트",TEXT),
    ln("TYPE* mybuf = (TYPE*)(ptr + blockIdx.x*avenumber); // 이 블록의 시작",TEXT),
    ln("int n = avenumber/(blockDim.x*sizeof(TYPE));   // 스레드당 반복 횟수",TEXT),
    ln("for (i = 0; i < n; i++){",TEXT),
    ln("    int index = i*blockDim.x + threadIdx.x;   // ★ 이웃 스레드 = 이웃 주소",GREEN),
    ln("    mybuf[index] = p1;",TEXT),
    ln("}",TEXT),
  ],{fontSize:13.5});
  s.addText([
    ln("여기서 처음으로 ",TEXT,{breakLine:false}),
    ln("threadIdx.x · blockDim.x",TEAL,{bold:true,breakLine:false}),
    ln("가 등장합니다. 세션 1의 커널엔 없었죠(스레드가 1개였으니까).",TEXT,{breakLine:true}),
  ],{x:M,y:5.2,w:W-2*M,h:0.7,fontFace:KFONT,fontSize:17,color:TEXT,margin:0,valign:"top"});
  s.addText("TYPE = unsigned long (8바이트). blockIdx로 블록의 구역을, threadIdx로 그 안의 위치를 정합니다.",{x:M,y:6.05,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,margin:0});
  s.addNotes("index = i*blockDim.x + threadIdx.x 를 칠판에 그려가며 설명. 스레드 0,1,2,...가 주소 0,1,2,...를 맡음.");
})();

// =====================================================================
// Slide 5 — coalescing visual (good vs bad)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"4","메모리 병합(coalescing)");
  s.addText("한 워프(warp, 32스레드)가 접근하는 주소가 서로 붙어 있으면, 하드웨어가 하나의 큰 전송으로 묶습니다.",{x:M,y:1.55,w:W-2*M,h:0.6,fontFace:KFONT,fontSize:16,color:MUTED,margin:0,valign:"top"});

  const drawRow=(y,label,color,pattern)=>{
    s.addText(label,{x:M,y:y-0.02,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:15,bold:true,color:color,margin:0});
    const cellW=0.62, x0=M, yy=y+0.42, n=12;
    for(let i=0;i<n;i++){
      s.addShape(p.ShapeType.roundRect,{x:x0+i*(cellW+0.08),y:yy,w:cellW,h:0.55,rectRadius:0.04,fill:{color:CODEBG},line:{color:LINE,width:1}});
    }
    // arrows: which thread touches which cell
    pattern.forEach((cell,thr)=>{
      const cx=x0+cell*(cellW+0.08)+cellW/2;
      s.addText("T"+thr,{x:cx-0.31,y:yy+0.12,w:0.62,h:0.3,align:"center",valign:"middle",fontFace:MONO,fontSize:11,bold:true,color:color,margin:0});
    });
  };
  // good: T0->0, T1->1, T2->2, T3->3
  drawRow(2.3,"✓ 병합됨 (index = i*blockDim.x + threadIdx.x)",GREEN,[0,1,2,3]);
  // bad: T0->0, T1->4, T2->8, T3->? (spread) — indices 0,4,8,11
  drawRow(3.75,"✗ 비병합 (index = threadIdx.x*n + i) — 스레드가 멀리 흩어짐",AMBER,[0,4,8,11]);

  const colW=(W-2*M-0.4)/2;
  s.addText([
    ln("병합 → 큰 전송 1번 → 대역폭 최대",GREEN,{breakLine:true,bullet:{code:"2022"},paraSpaceAfter:8}),
    ln("이웃 스레드가 이웃 주소를 읽음",TEXT,{breakLine:true,bullet:{code:"2022"}}),
  ],{x:M,y:5.15,w:colW,h:1.2,fontFace:KFONT,fontSize:15.5,color:TEXT,margin:0,valign:"top"});
  s.addText([
    ln("비병합 → 작은 전송 여러 번 → 급락",AMBER,{breakLine:true,bullet:{code:"2022"},paraSpaceAfter:8}),
    ln("실습 4에서 직접 깨뜨려 측정합니다",TEXT,{breakLine:true,bullet:{code:"2022"}}),
  ],{x:M+colW+0.4,y:5.15,w:colW,h:1.2,fontFace:KFONT,fontSize:15.5,color:TEXT,margin:0,valign:"top"});
  s.addNotes("병합은 CUDA 성능 최적화의 1순위. T0~T3이 어느 칸을 가리키는지 눈으로 비교시키세요.");
})();

// =====================================================================
// Slide 6 — timing with CUDA events (code + why)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"5","CUDA 이벤트로 시간 재기");
  codePanel(s,M,1.7,W-2*M,3.0,[
    ln("// test10   tests.cpp:1510~1524",MUTED),
    ln("cudaEventRecord(start, stream);",TEXT),
    ln("test10_kernel_write<<<...>>>(...);          // 쓰기",GREEN),
    ln("for (i=0;i<n;i++) test10_kernel_readwrite<<<...>>>(...); // 읽고쓰기 n회",GREEN),
    ln("cudaEventRecord(stop, stream);",TEXT),
    ln("cudaEventSynchronize(stop);                 // GPU가 끝날 때까지 대기",AMBER),
    ln("cudaEventElapsedTime(&elapsedtime, start, stop); // 밀리초 경과시간",TEAL),
  ],{fontSize:13});
  s.addText([
    ln("왜 CPU 타이머가 아니라 GPU 이벤트?",TEXT,{bold:true,breakLine:true,paraSpaceAfter:6}),
    ln("커널은 비동기(asynchronous) 실행 — CPU는 커널을 던지고 바로 다음 줄로 갑니다. GPU 타임라인 위에 ",MUTED,{breakLine:false}),
    ln("이벤트를 찍어야",TEAL,{bold:true,breakLine:false}),
    ln(" 진짜 커널 시간을 잴 수 있습니다.",MUTED,{breakLine:true}),
  ],{x:M,y:5.0,w:W-2*M,h:1.5,fontFace:KFONT,fontSize:16,color:TEXT,margin:0,valign:"top"});
  s.addNotes("비동기 실행 개념이 핵심. cudaEventSynchronize가 왜 필요한지(GPU 완료 대기) 짚어주세요.");
})();

// =====================================================================
// Slide 7 — bandwidth formula breakdown
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"6","대역폭(bandwidth) 계산 뜯어보기");
  codePanel(s,M,1.7,W-2*M,1.3,[
    ln("bandwidth = (2*n + 1) * tot_num_blocks / elapsedtime   // GB/s",GREEN),
    ln("//           └ 접근 횟수 ┘   └ MB 단위 ┘   └ ms 단위 ┘",MUTED),
  ],{fontSize:14});
  const cards=[
    ["(2n + 1)","접근 패스 수",GREEN,"쓰기 1회 + 읽고쓰기 n회(각 읽기+쓰기=2)"],
    ["tot_num_blocks","전송량 (MB)","TEAL","블록 하나 = BLOCKSIZE = 1 MB"],
    ["elapsedtime","경과 시간 (ms)","AMBER","이벤트로 잰 밀리초"],
  ];
  const colorMap={TEAL:TEAL,GREEN:GREEN,AMBER:AMBER};
  const cw=(W-2*M-0.8)/3;
  cards.forEach((c,i)=>{
    const x=M+i*(cw+0.4);
    const col=colorMap[c[2]]||c[2];
    s.addShape(p.ShapeType.roundRect,{x,y:3.3,w:cw,h:2.4,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(c[0],{x:x+0.25,y:3.5,w:cw-0.5,h:0.5,fontFace:MONO,fontSize:17,bold:true,color:col,margin:0});
    s.addText(c[1],{x:x+0.25,y:4.05,w:cw-0.5,h:0.45,fontFace:KFONT,fontSize:15,bold:true,color:TEXT,margin:0});
    s.addText(c[3],{x:x+0.25,y:4.55,w:cw-0.5,h:1.0,fontFace:KFONT,fontSize:14,color:MUTED,margin:0,valign:"top",lineSpacingMultiple:1.15});
  });
  s.addText("MB ÷ ms = GB/s. 실측값을 GPU 스펙의 이론 대역폭과 비교하면 보통 60~90% 수준입니다.",{x:M,y:5.95,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,margin:0});
  s.addNotes("단위 변환(MB/ms=GB/s)을 칠판에서 한 번 유도하면 학생들이 확실히 이해합니다.");
})();

// =====================================================================
// Slide 8 — streams (brief)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"7","스트림(stream) 맛보기");
  codePanel(s,M,1.7,W-2*M,1.5,[
    ln("cudaStreamCreate(&stream);",TEXT),
    ln("test10_kernel_write<<<gridDim, blockDim, 0, stream>>>(...);",GREEN),
    ln("//                                    ▲          ▲",MUTED),
    ln("//                          공유메모리 크기   스트림",MUTED),
  ],{fontSize:13.5});
  s.addText([
    ln("스트림(stream)",TEAL,{bold:true,breakLine:false}),
    ln("은 GPU 작업을 순서대로 처리하는 \"작업 큐\"입니다. Test 10은 하나의 스트림에 쓰기·읽기 커널을 순서대로 넣습니다.",TEXT,{breakLine:true}),
  ],{x:M,y:3.4,w:W-2*M,h:1.0,fontFace:KFONT,fontSize:17,color:TEXT,margin:0,valign:"top"});
  const colW=(W-2*M-0.4)/2;
  s.addShape(p.ShapeType.roundRect,{x:M,y:4.5,w:colW,h:1.7,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
  s.addText("<<<grid, block, 0, stream>>>",{x:M+0.25,y:4.7,w:colW-0.5,h:0.5,fontFace:MONO,fontSize:14,bold:true,color:GREEN,margin:0});
  s.addText("실행 구성의 3·4번째 인자: 공유 메모리 크기(0), 스트림. 세션 1에선 생략했던 부분입니다.",{x:M+0.25,y:5.2,w:colW-0.5,h:0.9,fontFace:KFONT,fontSize:14,color:TEXT,margin:0,valign:"top",lineSpacingMultiple:1.15});
  s.addShape(p.ShapeType.roundRect,{x:M+colW+0.4,y:4.5,w:colW,h:1.7,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
  s.addText("왜 이벤트도 스트림에?",{x:M+colW+0.65,y:4.7,w:colW-0.5,h:0.5,fontFace:KFONT,fontSize:15,bold:true,color:TEAL,margin:0});
  s.addText("이벤트를 같은 스트림에 기록해야 그 스트림의 커널들 시간을 정확히 잽니다.",{x:M+colW+0.65,y:5.2,w:colW-0.5,h:0.9,fontFace:KFONT,fontSize:14,color:TEXT,margin:0,valign:"top",lineSpacingMultiple:1.15});
  s.addNotes("스트림은 심화 주제라 개념만. 여러 스트림으로 겹쳐 실행(overlap)하는 건 향후 주제로 남깁니다.");
})();

// =====================================================================
// Slide 9 — Lab preview
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"8","실습 미리보기 · Lab 3");
  const labs=[
    ["Ex 1","느린 커널 vs 빠른 커널","time으로 Test 3와 Test 10 실행 시간 비교",GREEN],
    ["Ex 2","대역폭을 눈으로 보기","--verbose 로 test10의 bandwidth 로그 확인",TEAL],
    ["Ex 3","블록 크기를 바꿔 측정","STRESS_BLOCKSIZE 32~256 별 대역폭 표 작성",AMBER],
    ["Ex 4","메모리 병합 깨뜨리기","인덱스를 비병합으로 바꿔 대역폭 급락 관찰",GREEN],
  ];
  const cw=(W-2*M-0.5)/2, ch=2.15;
  labs.forEach((l,i)=>{
    const x=M+(i%2)*(cw+0.5), y=1.7+Math.floor(i/2)*(ch+0.35);
    s.addShape(p.ShapeType.roundRect,{x,y,w:cw,h:ch,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(l[0],{x:x+0.28,y:y+0.22,w:1.3,h:0.55,align:"center",valign:"middle",fontFace:MONO,fontSize:17,bold:true,color:BG,fill:{color:l[3]},shape:p.ShapeType.roundRect,rectRadius:0.08});
    s.addText(l[1],{x:x+1.75,y:y+0.24,w:cw-2.0,h:0.9,fontFace:KFONT,fontSize:17,bold:true,color:TEXT,margin:0,valign:"top"});
    s.addText(l[2],{x:x+0.28,y:y+1.2,w:cw-0.56,h:0.8,fontFace:KFONT,fontSize:14,color:MUTED,margin:0,valign:"top",lineSpacingMultiple:1.1});
  });
  s.addText("측정이 핵심입니다 — 대역폭 로그는 --verbose 없이는 출력되지 않습니다.",{x:M,y:6.5,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:13,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("Ex 3의 블록 크기 스윕을 그래프로 그리게 하면 '최적점'을 스스로 발견합니다. 정답을 미리 주지 마세요.");
})();

// =====================================================================
// Slide 10 — recap & pitfalls
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"9","핵심 정리 & 자주 하는 실수");
  const colW=(W-2*M-0.4)/2;
  s.addShape(p.ShapeType.roundRect,{x:M,y:1.75,w:colW,h:4.5,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
  s.addText("기억할 것",{x:M+0.3,y:2.0,w:colW-0.6,h:0.5,fontFace:KFONT,fontSize:20,bold:true,color:GREEN,margin:0});
  s.addText([
    ln("총 스레드 = 블록수 × 블록당 스레드",TEXT,{breakLine:true,bullet:{code:"2713"},paraSpaceAfter:12}),
    ln("index = i*blockDim.x + threadIdx.x → 병합",TEXT,{breakLine:true,bullet:{code:"2713"},paraSpaceAfter:12}),
    ln("커널은 비동기 → 이벤트로 시간 측정",TEXT,{breakLine:true,bullet:{code:"2713"},paraSpaceAfter:12}),
    ln("대역폭 = 전송량 ÷ 시간 (GB/s)",TEXT,{breakLine:true,bullet:{code:"2713"}}),
  ],{x:M+0.3,y:2.65,w:colW-0.6,h:3.4,fontFace:KFONT,fontSize:15.5,color:TEXT,margin:0,valign:"top"});
  const x2=M+colW+0.4;
  s.addShape(p.ShapeType.roundRect,{x:x2,y:1.75,w:colW,h:4.5,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
  s.addText("입문자 함정",{x:x2+0.3,y:2.0,w:colW-0.6,h:0.5,fontFace:KFONT,fontSize:20,bold:true,color:AMBER,margin:0});
  s.addText([
    ln("스레드만 늘리면 무한히 빨라진다고 착각 → 포화점 존재",TEXT,{breakLine:true,bullet:{code:"2717"},paraSpaceAfter:12}),
    ln("CPU 타이머로 커널 시간 측정 → 비동기라 부정확",TEXT,{breakLine:true,bullet:{code:"2717"},paraSpaceAfter:12}),
    ln("접근 패턴을 흩뜨림 → 병합 깨져 대역폭 급락",TEXT,{breakLine:true,bullet:{code:"2717"},paraSpaceAfter:12}),
    ln("--verbose 안 줘서 대역폭 로그가 안 보임",TEXT,{breakLine:true,bullet:{code:"2717"}}),
  ],{x:x2+0.3,y:2.65,w:colW-0.6,h:3.4,fontFace:KFONT,fontSize:15.5,color:TEXT,margin:0,valign:"top"});
  s.addNotes("네 가지 목표(슬라이드 2)로 돌아가 자가 점검을 유도하세요.");
})();

// =====================================================================
// Slide 11 — next session (dark)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  s.addText("다음 세션 예고",{x:M,y:1.6,w:10,h:0.5,fontFace:KFONT,fontSize:18,bold:true,color:GREEN,margin:0});
  s.addText("세션 4 · 원자연산과 에러 핸들링",{x:M,y:2.15,w:12,h:1.0,fontFace:KFONT,fontSize:36,bold:true,color:TEXT,margin:0});
  s.addText([
    ln("210만 스레드가 오류 카운터를 ",TEXT,{breakLine:false}),
    ln("동시에",AMBER,{bold:true,breakLine:false}),
    ln(" 늘리면 무슨 일이 생길까?",TEXT,{breakLine:true}),
  ],{x:M,y:3.3,w:12,h:0.6,fontFace:KFONT,fontSize:20,color:TEXT,margin:0});
  s.addText([
    ln("atomicAdd — 경쟁 조건(race condition) 없이 안전하게 세기",MONO,{breakLine:true,color:GREEN,paraSpaceAfter:8}),
    ln("RECORD_ERR 매크로 · CUERR / SYNC_CUERR_KERNEL 에러 처리",MUTED,{breakLine:true,paraSpaceAfter:8}),
    ln("커널은 비동기 → DeviceSynchronize 후에야 오류를 안다",MUTED,{breakLine:true}),
  ],{x:M,y:4.15,w:12,h:1.4,fontFace:KFONT,fontSize:16,color:TEXT,margin:0,valign:"top"});
  s.addText("실습 랩 3(대역폭 측정)을 먼저 완료하고 오세요.",{x:M,y:6.4,w:12,h:0.4,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,margin:0});
  s.addNotes("오늘 본 병렬 커널에서 여러 스레드가 하나의 카운터를 건드리는 문제가 세션 4의 출발점.");
})();

p.writeFile({fileName:"세션3_병렬성설계와_성능.pptx"}).then(f=>console.log("wrote",f));
