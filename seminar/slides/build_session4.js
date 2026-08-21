const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.layout = "LAYOUT_WIDE"; // 13.3 x 7.5

// ---- palette (shared with Sessions 1-3) ----
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
  s.addText("CUDA 세미나 · 세션 4",{x:M,y:2.05,w:8,h:0.5,fontFace:KFONT,fontSize:18,bold:true,color:GREEN,margin:0});
  s.addText("원자연산과 에러 핸들링\n(atomics & error handling)",{x:M,y:2.5,w:11.9,h:2.0,fontFace:KFONT,fontSize:44,bold:true,color:TEXT,lineSpacingMultiple:1.05,margin:0});
  s.addText("210만 스레드가 카운터 하나를 동시에 늘리면? — 안전하게 세고, 오류를 잡는 법",{x:M,y:4.7,w:11.8,h:0.5,fontFace:KFONT,fontSize:19,color:MUTED,margin:0});
  codePanel(s,M,5.6,9.0,1.0,[
    ln("unsigned int idx = atomicAdd(err, 1);   // 원자적으로 +1, 이전 값 반환",GREEN),
    ln("SYNC_CUERR_KERNEL( kernel<<<...>>>(...) ); // 실행+동기화+오류검사",TEAL),
  ],{fontSize:12.5});
  s.addText("입문 과정 · 세션 3에 이어서",{x:9.7,y:6.05,w:3.0,h:0.4,align:"right",fontFace:KFONT,fontSize:13,color:MUTED,margin:0});
  s.addNotes("세션 3의 병렬 커널에서 여러 스레드가 하나의 카운터를 건드리는 문제가 오늘의 출발점.");
})();

// =====================================================================
// Slide 2 — objectives
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"1","오늘의 목표");
  const items=[
    ["경쟁 조건(race condition)","여러 스레드가 err_count++를 동시에 하면 왜 값이 사라지나"],
    ["원자연산(atomic)","atomicAdd가 어떻게 안전하게 세고 고유 인덱스를 주나"],
    ["비동기 오류","커널은 비동기 → DeviceSynchronize 후에야 오류를 안다"],
    ["방어적 매크로","CUERR·SYNC_CUERR_KERNEL이 파일·줄과 함께 오류를 잡는다"],
  ];
  let y=1.7;
  items.forEach((it,i)=>{
    s.addShape(p.ShapeType.roundRect,{x:M,y,w:W-2*M,h:1.06,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(String(i+1),{x:M+0.22,y:y+0.2,w:0.66,h:0.66,align:"center",valign:"middle",fontFace:MONO,fontSize:22,bold:true,color:BG,fill:{color:i%2?TEAL:GREEN},shape:p.ShapeType.roundRect,rectRadius:0.33});
    s.addText(it[0],{x:M+1.15,y:y+0.14,w:W-2*M-1.4,h:0.44,fontFace:KFONT,fontSize:19,bold:true,color:TEXT,margin:0,valign:"middle"});
    s.addText(it[1],{x:M+1.15,y:y+0.55,w:W-2*M-1.4,h:0.42,fontFace:KFONT,fontSize:14,color:MUTED,margin:0,valign:"middle"});
    y+=1.24;
  });
  s.addNotes("실습에서 atomicAdd를 일반 덧셈으로 바꿔 값이 실제로 틀어지는 것을 재현합니다.");
})();

// =====================================================================
// Slide 3 — race condition timeline (the problem)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"2","문제 · 여럿이 하나를 동시에 건드리면",RED);
  s.addText("err_count = 5 인 상태에서 두 스레드가 동시에 '읽고 → +1 → 쓰기'를 하면?",{x:M,y:1.55,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:16,color:MUTED,margin:0});
  // two-thread interleaving table
  const rows=[
    ["시간","스레드 A","스레드 B"],
    ["t1","err_count 읽음 → 5",""],
    ["t2","","err_count 읽음 → 5"],
    ["t3","5 + 1 = 6 저장",""],
    ["t4","","5 + 1 = 6 저장"],
  ];
  const x0=M, y0=2.25, cw=[2.2,4.5,4.5], rh=0.72;
  rows.forEach((r,ri)=>{
    let x=x0;
    r.forEach((cell,ci)=>{
      const isHead=ri===0;
      s.addShape(p.ShapeType.roundRect,{x,y:y0+ri*rh,w:cw[ci]-0.1,h:rh-0.1,rectRadius:0.04,fill:{color:isHead?"25303F":CARD},line:{type:"none"}});
      const col=isHead?TEXT:(ci===1?TEAL:(ci===2?AMBER:MUTED));
      s.addText(cell,{x:x+0.2,y:y0+ri*rh,w:cw[ci]-0.4,h:rh-0.1,valign:"middle",fontFace:ci===0?MONO:KFONT,fontSize:isHead?15:14,bold:isHead,color:col,margin:0});
      x+=cw[ci];
    });
  });
  s.addText([
    ln("결과: err_count = 6  ",RED,{bold:true,breakLine:false}),
    ln("→ 두 번 늘렸는데 1만 증가! 증가 하나가 사라졌습니다(lost update).",TEXT,{breakLine:true}),
  ],{x:M,y:6.05,w:W-2*M,h:0.7,fontFace:KFONT,fontSize:17,color:TEXT,margin:0,valign:"top"});
  s.addNotes("읽기-수정-쓰기(read-modify-write) 사이에 끼어들기가 일어나 증가가 유실됨. 210만 스레드면 대량 유실.");
})();

// =====================================================================
// Slide 4 — atomicAdd (the solution)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"3","해결 · 원자연산(atomic)");
  s.addText([
    ln("원자연산",GREEN,{bold:true,breakLine:false}),
    ln("은 '읽기 → 수정 → 쓰기'를 ",TEXT,{breakLine:false}),
    ln("쪼갤 수 없는 하나의 동작",TEAL,{bold:true,breakLine:false}),
    ln("으로 실행합니다. 중간에 다른 스레드가 끼어들 수 없습니다.",TEXT,{breakLine:true}),
  ],{x:M,y:1.65,w:W-2*M,h:0.8,fontFace:KFONT,fontSize:18,color:TEXT,margin:0,valign:"top"});
  codePanel(s,M,2.55,W-2*M,1.3,[
    ln("unsigned int idx = atomicAdd(err, 1);",GREEN),
    ln("//  err이 가리키는 값을 원자적으로 +1 하고, +1 하기 '전' 값을 idx로 돌려줌",MUTED),
  ],{fontSize:13.5});
  const cards=[
    ["안전한 카운트",GREEN,"동시에 불러도 증가가 유실되지 않음. 정확한 총 오류 개수."],
    ["공짜 고유 번호",TEAL,"반환값(이전 값)이 스레드마다 다름 → 겹치지 않는 슬롯 인덱스."],
    ["하드웨어 지원",AMBER,"GPU가 직접 지원하는 연산. 락(lock) 없이 빠름."],
  ];
  const cw=(W-2*M-0.8)/3;
  cards.forEach((c,i)=>{
    const x=M+i*(cw+0.4);
    s.addShape(p.ShapeType.roundRect,{x,y:4.15,w:cw,h:2.1,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(c[0],{x:x+0.25,y:4.35,w:cw-0.5,h:0.5,fontFace:KFONT,fontSize:17,bold:true,color:c[1],margin:0});
    s.addText(c[2],{x:x+0.25,y:4.9,w:cw-0.5,h:1.2,fontFace:KFONT,fontSize:14,color:TEXT,margin:0,valign:"top",lineSpacingMultiple:1.15});
  });
  s.addNotes("핵심: atomicAdd는 두 가지를 동시에 준다 — (1) 안전한 카운트 (2) 반환값을 통한 고유 인덱스.");
})();

// =====================================================================
// Slide 4b — atomicAdd serialization timeline (before/after)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"3","원자연산이 경쟁을 없애는 과정");
  s.addText("같은 상황(err=5, 두 스레드가 +1)을 일반 증가와 atomicAdd로 비교합니다.",{x:M,y:1.5,w:W-2*M,h:0.45,fontFace:KFONT,fontSize:15,color:MUTED,margin:0});

  const colW=(W-2*M-0.5)/2, boxY=2.05, boxH=4.35;
  const stepRow=(x,y,w,h,tlabel,tcolor,text,tc)=>{
    s.addShape(p.ShapeType.roundRect,{x,y,w,h,rectRadius:0.04,fill:{color:CODEBG},line:{color:LINE,width:1}});
    s.addText(tlabel,{x:x+0.12,y,w:0.6,h,valign:"middle",align:"center",fontFace:MONO,fontSize:12,bold:true,color:tcolor,margin:0});
    s.addText(text,{x:x+0.78,y,w:w-0.9,h,valign:"middle",fontFace:KFONT,fontSize:13,color:tc||TEXT,margin:0});
  };

  // LEFT — plain ++ (race)
  const lx=M;
  s.addShape(p.ShapeType.roundRect,{x:lx,y:boxY,w:colW,h:boxH,rectRadius:0.07,fill:{color:CARD},line:{color:RED,width:1.5}});
  s.addText("① 일반 ++  (경쟁 조건)",{x:lx+0.25,y:boxY+0.15,w:colW-0.5,h:0.4,fontFace:KFONT,fontSize:16,bold:true,color:RED,margin:0});
  const lrx=lx+0.28, lrw=colW-0.56; let ly=boxY+0.7;
  stepRow(lrx,ly,lrw,0.55,"t1",TEAL,"스레드 A : err 읽음 → 5"); ly+=0.67;
  stepRow(lrx,ly,lrw,0.55,"t2",AMBER,"스레드 B : err 읽음 → 5  (끼어듦!)",AMBER); ly+=0.67;
  stepRow(lrx,ly,lrw,0.55,"t3",TEAL,"스레드 A : 5+1 = 6 저장"); ly+=0.67;
  stepRow(lrx,ly,lrw,0.55,"t4",AMBER,"스레드 B : 5+1 = 6 저장"); ly+=0.75;
  s.addShape(p.ShapeType.roundRect,{x:lrx,y:ly,w:lrw,h:0.55,rectRadius:0.05,fill:{color:"2A1618"},line:{color:RED,width:1}});
  s.addText("결과 err = 6  ·  증가 하나 유실 (틀림)",{x:lrx,y:ly,w:lrw,h:0.55,align:"center",valign:"middle",fontFace:KFONT,fontSize:14,bold:true,color:RED,margin:0});

  // RIGHT — atomicAdd (serialized)
  const rx=M+colW+0.5;
  s.addShape(p.ShapeType.roundRect,{x:rx,y:boxY,w:colW,h:boxH,rectRadius:0.07,fill:{color:CARD},line:{color:GREEN,width:1.5}});
  s.addText("② atomicAdd  (원자적)",{x:rx+0.25,y:boxY+0.15,w:colW-0.5,h:0.4,fontFace:KFONT,fontSize:16,bold:true,color:GREEN,margin:0});
  const rrx=rx+0.28, rrw=colW-0.56; let ry=boxY+0.7;
  stepRow(rrx,ry,rrw,0.9,"t1",TEAL,"A : atomicAdd → 5 읽고 6 저장\n     (읽기+쓰기가 한 덩어리)",TEXT); ry+=1.02;
  stepRow(rrx,ry,rrw,0.9,"t2",GREEN,"B : atomicAdd → 6 읽고 7 저장\n     (A가 끝난 뒤 실행)",TEXT); ry+=1.02;
  s.addText("원자연산은 중간에 끼어들 수 없습니다 (쪼갤 수 없는 하나의 동작)",{x:rrx,y:ry,w:rrw,h:0.4,fontFace:KFONT,fontSize:11.5,italic:true,color:MUTED,margin:0}); ry+=0.42;
  s.addShape(p.ShapeType.roundRect,{x:rrx,y:ry,w:rrw,h:0.55,rectRadius:0.05,fill:{color:"16240F"},line:{color:GREEN,width:1}});
  s.addText("결과 err = 7  ·  정확히 셈 (정확)",{x:rrx,y:ry,w:rrw,h:0.55,align:"center",valign:"middle",fontFace:KFONT,fontSize:14,bold:true,color:GREEN,margin:0});

  s.addText("210만 스레드가 동시에 오류를 세도, atomicAdd 덕분에 총 개수가 정확합니다. (RECORD_ERR, tests.cpp:76)",{x:M,y:6.55,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:13,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("왼쪽=끼어들기로 유실, 오른쪽=원자연산이 읽기-쓰기를 직렬화해 정확. 슬라이드 3의 문제에 대한 해답 타임라인.");
})();

// =====================================================================
// Slide 5 — RECORD_ERR macro dissection (code)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"4","RECORD_ERR 매크로 해부");
  codePanel(s,M,1.7,W-2*M,3.5,[
    ln("// tests.cpp:76",MUTED),
    ln("#define RECORD_ERR(err, p, expect, current) do{           \\",TEXT),
    ln("    unsigned int idx = atomicAdd(err, 1);                 \\",GREEN),
    ln("    idx = idx % MAX_ERR_RECORD_COUNT;   // 최근 10개만 순환  \\",TEAL),
    ln("    err_addr[idx]        = (unsigned long)p;        // 오류 주소     \\",TEXT),
    ln("    err_expect[idx]      = (unsigned long)expect;   // 기대값       \\",TEXT),
    ln("    err_current[idx]     = (unsigned long)current;  // 실제 읽은 값  \\",TEXT),
    ln("    err_second_read[idx] = (unsigned long)(*p);     // 다시 읽은 값  \\",TEXT),
    ln("}while(0)",TEXT),
  ],{fontSize:13});
  s.addText([
    ln("오류가 난 스레드는 이 매크로로 ",TEXT,{breakLine:false}),
    ln("주소·기대값·실제값·재읽기값",GREEN,{bold:true,breakLine:false}),
    ln("을 기록합니다. 재읽기값은 '다시 읽으면 값이 바뀌나?'(소프트 오류 판별)를 봅니다.",TEXT,{breakLine:true}),
  ],{x:M,y:5.35,w:W-2*M,h:1.0,fontFace:KFONT,fontSize:16,color:TEXT,margin:0,valign:"top"});
  s.addNotes("각 필드가 무엇을 담는지 짚어주세요. second_read는 소프트 오류(일시적) vs 하드 오류(영구) 구별에 쓰임.");
})();

// =====================================================================
// Slide 6 — atomic return = unique index (concept)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"5","반환값이 고유 슬롯 번호가 된다");
  s.addText("세 스레드가 동시에 atomicAdd(err,1)을 호출해도, 반환값은 겹치지 않습니다.",{x:M,y:1.6,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:16,color:MUTED,margin:0});
  // three threads -> distinct returns
  const th=[["스레드 X","→ 반환 0","슬롯 0",GREEN],["스레드 Y","→ 반환 1","슬롯 1",TEAL],["스레드 Z","→ 반환 2","슬롯 2",AMBER]];
  const cw=(W-2*M-0.8)/3;
  th.forEach((t,i)=>{
    const x=M+i*(cw+0.4);
    s.addShape(p.ShapeType.roundRect,{x,y:2.35,w:cw,h:1.7,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(t[0],{x:x+0.25,y:2.55,w:cw-0.5,h:0.5,fontFace:KFONT,fontSize:17,bold:true,color:t[3],margin:0});
    s.addText(t[1],{x:x+0.25,y:3.05,w:cw-0.5,h:0.45,fontFace:MONO,fontSize:16,bold:true,color:TEXT,margin:0});
    s.addText(t[2]+" 에 기록",{x:x+0.25,y:3.5,w:cw-0.5,h:0.45,fontFace:KFONT,fontSize:14,color:MUTED,margin:0});
  });
  codePanel(s,M,4.35,W-2*M,1.15,[
    ln("idx = atomicAdd(err, 1);   // X→0, Y→1, Z→2  (순서는 달라도 값은 유일)",GREEN),
    ln("idx = idx % 10;            // 배열 10칸을 순환하며 최근 오류만 보관",TEAL),
  ],{fontSize:13});
  s.addText("MAX_ERR_RECORD_COUNT = 10 (tests.cpp:67). 오류가 100만 개여도 배열은 10칸 — 최근 것만 남깁니다.",{x:M,y:5.7,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,margin:0});
  s.addNotes("atomicAdd의 반환값이 락 없이 고유 인덱스를 나눠주는 우아한 트릭. 병렬 프로그래밍의 흔한 관용구.");
})();

// =====================================================================
// Slide 7 — async errors (the trap)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"6","함정 · 커널은 비동기(asynchronous)",AMBER);
  s.addText("커널 실행 <<<>>>은 오류를 반환하지 않습니다. CPU는 커널을 던지고 바로 다음 줄로 갑니다.",{x:M,y:1.6,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:16,color:MUTED,margin:0});
  // timeline
  const steps=[
    ["① 커널 실행","kernel<<<...>>>(...)","CPU는 던지고 즉시 리턴 (오류 모름)",TEAL],
    ["② 동기화","cudaDeviceSynchronize()","GPU가 끝날 때까지 대기",GREEN],
    ["③ 오류 검사","if (err != cudaSuccess)","이제서야 커널 오류를 알 수 있음",AMBER],
  ];
  const bw=3.66, gap=0.55, y0=2.35, bh=2.2;
  steps.forEach((st,i)=>{
    const x=M+i*(bw+gap);
    s.addShape(p.ShapeType.roundRect,{x,y:y0,w:bw,h:bh,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(st[0],{x:x+0.25,y:y0+0.22,w:bw-0.5,h:0.5,fontFace:KFONT,fontSize:17,bold:true,color:st[3],margin:0});
    s.addText(st[1],{x:x+0.25,y:y0+0.8,w:bw-0.5,h:0.6,fontFace:MONO,fontSize:12,color:TEXT,margin:0});
    s.addText(st[2],{x:x+0.25,y:y0+1.35,w:bw-0.5,h:0.75,fontFace:KFONT,fontSize:13.5,color:MUTED,margin:0,valign:"top",lineSpacingMultiple:1.1});
    if(i<2) s.addText("→",{x:x+bw+0.02,y:y0+0.75,w:gap-0.04,h:0.7,align:"center",valign:"middle",fontFace:MONO,fontSize:26,bold:true,color:MUTED,margin:0});
  });
  s.addText([
    ln("동기화를 건너뛰면",AMBER,{bold:true,breakLine:false}),
    ln(" 커널이 잘못된 메모리를 건드려도 CPU는 모른 채 계속 진행합니다. 그래서 SYNC_CUERR_KERNEL이 필요합니다.",TEXT,{breakLine:true}),
  ],{x:M,y:4.9,w:W-2*M,h:1.0,fontFace:KFONT,fontSize:16,color:TEXT,margin:0,valign:"top"});
  s.addNotes("세션 3의 '커널은 비동기'와 직접 연결. 동기화 없이는 커널 오류를 잡을 수 없다는 점이 핵심.");
})();

// =====================================================================
// Slide 8 — CUERR / SYNC_CUERR_KERNEL macros (code)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"7","방어적 매크로 · CUERR / SYNC_CUERR_KERNEL");
  codePanel(s,M,1.7,W-2*M,3.6,[
    ln("// CUERR — 모든 CUDA 호출을 감싼다   cuda_memtest.h:132",MUTED),
    ln("#define CUERR(...) do{ apiError_t e;                             \\",TEXT),
    ln("    if ((e = __VA_ARGS__) != cudaSuccess) {                     \\",GREEN),
    ln("        FPRINTF(\"ERROR: CUDA error: %s, line %d, file %s\\n\",     \\",TEXT),
    ln("                cudaGetErrorString(e), __LINE__, __FILE__);     \\",TEAL),
    ln("        exit(e); }}while(0)",TEXT),
    ln("",TEXT),
    ln("// SYNC_CUERR_KERNEL — 커널 실행 + 동기화 + 검사   cuda_memtest.h:138",MUTED),
    ln("#define SYNC_CUERR_KERNEL(...) do{ __VA_ARGS__;                  \\",GREEN),
    ln("    e = cudaDeviceSynchronize();  /* GPU 완료 대기 후 검사 */      \\",AMBER),
    ln("    if (e != cudaSuccess){ ... exit(e); } }while(0)",TEXT),
  ],{fontSize:12.5});
  s.addText([
    ln("__FILE__ · __LINE__",TEAL,{bold:true,breakLine:false}),
    ln(" 덕분에 오류가 '어느 파일 몇 번째 줄'에서 났는지 즉시 압니다. 디버깅의 필수 습관입니다.",TEXT,{breakLine:true}),
  ],{x:M,y:5.45,w:W-2*M,h:0.9,fontFace:KFONT,fontSize:16,color:TEXT,margin:0,valign:"top"});
  s.addNotes("CUERR은 동기 API 호출용, SYNC_CUERR_KERNEL은 커널 실행용. 둘 다 실패 시 위치를 찍고 종료.");
})();

// =====================================================================
// Slide 9 — error_checking: graceful handling (defensive)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"8","오류를 우아하게 · error_checking");
  s.addText("오류가 발견되면 단순히 종료하지 않고, 정보를 기록하고 상태를 정리합니다.",{x:M,y:1.6,w:W-2*M,h:0.5,fontFace:KFONT,fontSize:16,color:MUTED,margin:0});
  const rows=[
    ["오류 정보 복사","cudaMemcpy로 err_count·주소·값을 CPU로 (tests.cpp:117)",TEAL],
    ["보고 & 로그","오류 주소·기대값·실제값·diff 출력, 필요시 이메일 알림",GREEN],
    ["카운터 재초기화","cudaMemset으로 err_count 등을 0으로 (다음 검사 준비)",AMBER],
    ["필요시 안전 종료","exit_on_error면 cudaDeviceReset 후 종료 (tests.cpp:196)",RED],
  ];
  let y=2.3;
  rows.forEach(r=>{
    s.addShape(p.ShapeType.roundRect,{x:M,y,w:W-2*M,h:0.95,rectRadius:0.07,fill:{color:CARD},line:{type:"none"}});
    s.addText(r[0],{x:M+0.3,y:y+0.12,w:3.4,h:0.7,valign:"middle",fontFace:KFONT,fontSize:17,bold:true,color:r[2],margin:0});
    s.addText(r[1],{x:M+3.9,y:y+0.12,w:W-2*M-4.2,h:0.7,valign:"middle",fontFace:KFONT,fontSize:15,color:TEXT,margin:0});
    y+=1.09;
  });
  s.addText("장난감 예제와 프로덕션 코드의 차이 — 오류를 만나도 무너지지 않고 다음 노드·다음 검사를 준비합니다.",{x:M,y:6.75,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:13,italic:true,color:MUTED,margin:0});
  s.addNotes("cudaDeviceReset은 GPU 상태를 깨끗이 되돌립니다. 멀티노드 검사에서 한 노드 오류가 전체를 멈추지 않게.");
})();

// =====================================================================
// Slide 10 — Lab preview
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  header(s,"9","실습 미리보기 · Lab 4");
  const labs=[
    ["Ex 1","경쟁 조건 재현","atomicAdd를 일반 덧셈으로 바꿔 카운트가 틀어짐 확인",RED],
    ["Ex 2","반환값 = 고유 인덱스","atomicAdd 반환 idx를 출력해 겹치지 않음 확인",TEAL],
    ["Ex 3","CUERR가 잡는 순간","cudaMemcpy 크기를 틀리게 해 오류 메시지 유발",AMBER],
    ["Ex 4","비동기 오류 & 동기화","커널 guard 제거 → SYNC_CUERR_KERNEL이 잡음",GREEN],
  ];
  const cw=(W-2*M-0.5)/2, ch=2.15;
  labs.forEach((l,i)=>{
    const x=M+(i%2)*(cw+0.5), y=1.7+Math.floor(i/2)*(ch+0.35);
    s.addShape(p.ShapeType.roundRect,{x,y,w:cw,h:ch,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
    s.addText(l[0],{x:x+0.28,y:y+0.22,w:1.3,h:0.55,align:"center",valign:"middle",fontFace:MONO,fontSize:17,bold:true,color:BG,fill:{color:l[3]},shape:p.ShapeType.roundRect,rectRadius:0.08});
    s.addText(l[1],{x:x+1.75,y:y+0.24,w:cw-2.0,h:0.9,fontFace:KFONT,fontSize:17,bold:true,color:TEXT,margin:0,valign:"top"});
    s.addText(l[2],{x:x+0.28,y:y+1.2,w:cw-0.56,h:0.8,fontFace:KFONT,fontSize:14,color:MUTED,margin:0,valign:"top",lineSpacingMultiple:1.1});
  });
  s.addText("Ex 1은 실행마다 결과가 달라집니다 — 그 자체가 경쟁 조건의 증거입니다.",{x:M,y:6.5,w:W-2*M,h:0.4,fontFace:KFONT,fontSize:13,italic:true,color:MUTED,align:"center",margin:0});
  s.addNotes("모든 Ex는 원복이 중요. 특히 Ex 3/4는 오류를 일부러 유발하므로 반드시 되돌리도록.");
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
    ln("동시 접근 카운트엔 atomicAdd",TEXT,{breakLine:true,bullet:{code:"2713"},paraSpaceAfter:12}),
    ln("atomicAdd 반환값 = 고유 인덱스",TEXT,{breakLine:true,bullet:{code:"2713"},paraSpaceAfter:12}),
    ln("커널 오류는 동기화 후에만 잡힌다",TEXT,{breakLine:true,bullet:{code:"2713"},paraSpaceAfter:12}),
    ln("모든 CUDA 호출을 CUERR로 감싸기",TEXT,{breakLine:true,bullet:{code:"2713"}}),
  ],{x:M+0.3,y:2.65,w:colW-0.6,h:3.4,fontFace:KFONT,fontSize:15.5,color:TEXT,margin:0,valign:"top"});
  const x2=M+colW+0.4;
  s.addShape(p.ShapeType.roundRect,{x:x2,y:1.75,w:colW,h:4.5,rectRadius:0.08,fill:{color:CARD},line:{type:"none"}});
  s.addText("입문자 함정",{x:x2+0.3,y:2.0,w:colW-0.6,h:0.5,fontFace:KFONT,fontSize:20,bold:true,color:AMBER,margin:0});
  s.addText([
    ln("공유 변수에 그냥 ++ → 증가 유실(race)",TEXT,{breakLine:true,bullet:{code:"2717"},paraSpaceAfter:12}),
    ln("커널 실행만 하고 동기화 없이 결과 신뢰",TEXT,{breakLine:true,bullet:{code:"2717"},paraSpaceAfter:12}),
    ln("CUDA 호출 반환값을 확인 안 함 → 조용한 실패",TEXT,{breakLine:true,bullet:{code:"2717"},paraSpaceAfter:12}),
    ln("atomic 남용 → 과도하면 성능 저하(꼭 필요한 곳만)",TEXT,{breakLine:true,bullet:{code:"2717"}}),
  ],{x:x2+0.3,y:2.65,w:colW-0.6,h:3.4,fontFace:KFONT,fontSize:15.5,color:TEXT,margin:0,valign:"top"});
  s.addNotes("네 가지 목표(슬라이드 2)로 돌아가 자가 점검을 유도하세요.");
})();

// =====================================================================
// Slide 12 — next session (dark)
// =====================================================================
(()=>{
  const s=p.addSlide(); bg(s);
  s.addText("다음 세션 예고",{x:M,y:1.6,w:10,h:0.5,fontFace:KFONT,fontSize:18,bold:true,color:GREEN,margin:0});
  s.addText("세션 5 · 멀티 GPU · 이식성 · 프로젝트",{x:M,y:2.15,w:12.2,h:1.0,fontFace:KFONT,fontSize:34,bold:true,color:TEXT,margin:0});
  s.addText([
    ln("이제 한 GPU를 이해했으니, ",TEXT,{breakLine:false}),
    ln("여러 GPU",TEAL,{bold:true,breakLine:false}),
    ln("를 동시에 돌리고 배운 것을 종합합니다.",TEXT,{breakLine:true}),
  ],{x:M,y:3.3,w:12,h:0.6,fontFace:KFONT,fontSize:20,color:TEXT,margin:0});
  s.addText([
    ln("pthread로 GPU마다 스레드 · cudaSetDevice",MONO,{breakLine:true,color:GREEN,paraSpaceAfter:8}),
    ln("MEMTEST_API_PREFIX 매크로 — 하나의 코드로 CUDA와 HIP(AMD) 모두",MUTED,{breakLine:true,paraSpaceAfter:8}),
    ln("미니 프로젝트: 나만의 메모리 테스트 커널 작성",MUTED,{breakLine:true}),
  ],{x:M,y:4.15,w:12,h:1.4,fontFace:KFONT,fontSize:16,color:TEXT,margin:0,valign:"top"});
  s.addText("실습 랩 4를 먼저 완료하고 오세요.",{x:M,y:6.4,w:12,h:0.4,fontFace:KFONT,fontSize:14,italic:true,color:MUTED,margin:0});
  s.addNotes("세션 5는 종합 회차. 멀티 GPU 스레딩, 이식성, 그리고 캡스톤 프로젝트로 마무리.");
})();

p.writeFile({fileName:"세션4_원자연산과_에러핸들링.pptx"}).then(f=>console.log("wrote",f));
