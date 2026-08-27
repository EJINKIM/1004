[Uploading TOYU_breakfast_reading.html…]()
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, minimum-scale=0.75, user-scalable=yes">
<title>조식 LIST</title>

<!-- 온라인에서는 실제 .xlsx 파일로 저장됨. -->
<script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js"></script>

<style>
:root{
  --bg:#f4f3ee; --panel:#fff; --ink:#1e2420; --line:#cfd3cf;
  --deep:#253d34; --soft:#eef1ed; --front:#fff0bf; --kitchen:#f4c7c7;
  --green:#e5f1df; --red:#ff3a2e; --blue:#dce8f6;
}
*{box-sizing:border-box}
body{
  margin:0;background:var(--bg);color:var(--ink);
  font-family:Pretendard, "Noto Sans KR", Apple SD Gothic Neo, Arial, sans-serif;
}
.wrap{max-width:1500px;margin:0 auto;padding:24px 16px 60px}
.topbar{
  display:grid;grid-template-columns:1fr auto 1fr;align-items:end;
  margin-bottom:12px;gap:12px
}
.title{text-align:center;font-size:23px;font-weight:900;letter-spacing:-.04em}
.notice{color:#ff2a1c;font-weight:900;font-size:15px;margin-left:10px}
.datebox{text-align:right}
.datebox label{font-size:11px;color:#707872;font-weight:800;display:block;margin-bottom:5px}
.datebox input{border:1px solid #cfd3cf;background:#fff;padding:8px 10px;font-weight:800}
.table-wrap{
  background:#fff;border:1px solid #bfc4c0;overflow:auto;
  box-shadow:0 12px 30px rgba(30,40,35,.06)
}
table{width:100%;border-collapse:collapse;min-width:1260px;table-layout:fixed}
th,td{border:1px solid #bec3bf;padding:5px 4px;text-align:center;height:40px}
th{background:#fafafa;font-size:11px;font-weight:900;white-space:nowrap}
td{font-size:12px}
th.menu-head{background:#f8f9f7}
input,select{
  width:100%;border:0;background:transparent;text-align:center;
  font:inherit;outline:none;padding:5px 2px
}
select{cursor:pointer}
.name-input{text-align:left;padding-left:8px;font-weight:700}
.room-select,.people-select{font-weight:900}
.type-select{font-size:11px}
.time-select{font-size:11px}
.qty{font-weight:900}
.row-total{background:#f5f6f3;font-weight:900}
.total-row td{background:#f2f3f0;font-weight:900}
.total-row .sumcell{background:#dce8f6}
.total-row .grand{background:#f3e5bc}
.menu-label{font-weight:900;color:#334139}
.small{font-size:10px;color:#778078}

.people-cell-wrap,.child-cell-wrap{display:grid;gap:6px;align-items:start}
.people-summary{
  display:inline-flex;align-items:baseline;gap:6px;justify-content:center;
  padding:7px 10px;border:1px solid #d5dbd5;border-radius:12px;background:#f8f9f6;min-width:88px
}
.people-summary small{font-size:10px;color:#69736d;font-weight:800;letter-spacing:-.02em}
.people-summary b{font-size:16px;font-weight:950;color:#223129}
.people-summary.filled{background:#eef5ef;border-color:#b9cbbd}
.inline-summary{font-size:11px;color:#5f6963;line-height:1.35;word-break:keep-all}
.inline-check{display:inline-flex;align-items:center;gap:6px;font-size:11px;font-weight:800;color:#334139}
.inline-check input{width:auto;margin:0}
.stack-editor{display:none;gap:6px}
.stack-editor.open{display:grid}
.editor-row{display:grid;grid-template-columns:1fr 84px auto;gap:6px;align-items:center}
.editor-row.single{grid-template-columns:1fr 84px}
.editor-row select{border:1px solid #d8ddd8;background:#fff;border-radius:10px;padding:7px 8px;font-size:12px;font-weight:800}
.editor-row .row-remove,.add-mini{border:1px solid rgba(60,70,64,.2);background:#fff;border-radius:10px;padding:7px 9px;cursor:pointer;font-weight:900;font-size:11px}
.add-mini{justify-self:start;background:#f7f8f5}
.child-pill-row,.seat-order-wrap{display:flex;flex-wrap:wrap;gap:4px;justify-content:center}
.child-pill,.order-chip{display:inline-flex;align-items:center;justify-content:center;min-width:46px;min-height:46px;padding:4px 8px;border-radius:999px;border:1px solid #bfc8c1;background:#fff;font-size:11px;font-weight:900;color:#223129;line-height:1.15;text-align:center}
.child-pill{min-width:auto;min-height:auto;padding:4px 8px}
.child-placeholder{font-size:11px;color:#8a938d}
.seat-people{display:inline-flex;align-items:baseline;gap:5px;padding:4px 8px;border-radius:999px;background:#f1f5f2;border:1px solid #cad5cd}
.seat-people small{font-size:9px;font-weight:800;color:#6b756f}.seat-people b{font-size:13px;font-weight:950;color:#24352d}
.archive-people{font-size:12px;color:#52605a;margin-top:4px}


/* ===== 레스토랑 좌석 배치도 ===== */
.layout-card{
  margin-top:18px;padding:18px;background:#fff;border:1px solid #bfc4c0;
  box-shadow:0 12px 30px rgba(30,40,35,.055)
}
.layout-head{display:flex;justify-content:space-between;align-items:end;gap:12px;margin-bottom:14px}
.layout-head h2{margin:0;font-size:18px;letter-spacing:-.03em}
.layout-help{font-size:12px;color:#707872;line-height:1.55}
.floorplan-shell{max-width:980px;margin:0 auto;background:#fff;padding:10px}
.floorplan{position:relative;width:100%;aspect-ratio:1.34/1;border:7px solid #111;background:#fff;min-height:570px}
.plan-box{position:absolute;border:2px solid #171717;background:#fff;display:flex;align-items:center;justify-content:center;text-align:center;overflow:hidden}
.fixed-label{font-size:18px;font-weight:850}
.r2{left:3%;top:3%;width:19%;height:30%}.r1{left:22%;top:3%;width:18%;height:30%}
.r3{left:3%;top:33%;width:19%;height:22%}.r4{left:3%;top:55%;width:19%;height:24%}
.r5{left:27%;top:41%;width:13%;height:38%}.wc1{left:3%;top:79%;width:19%;height:18%}.wc2{left:27%;top:79%;width:13%;height:18%}
.s6{left:44%;top:7%;width:13%;height:18%}.s4{left:60%;top:7%;width:13%;height:18%}.s2{left:76%;top:7%;width:13%;height:18%}
.s5{left:44%;top:29%;width:13%;height:18%}.s3{left:60%;top:29%;width:13%;height:18%}.s1{left:76%;top:29%;width:13%;height:18%}
.counter{left:46%;top:66%;width:36%;height:11%}.kitchen-box{left:46%;top:81%;width:47%;height:16%}
.entrance{position:absolute;left:82%;top:49%;width:11%;height:13%;border:2px solid #171717;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:900;background:#fff}
.entrance-arrow{position:absolute;left:88%;top:43%;font-size:34px;font-weight:900}.small-box{left:86%;top:66%;width:7.5%;height:11%}
.seat{background:#fbfbf8;transition:.18s}.seat.occupied{background:#e8f1eb}.seat.duplicate{background:#fff0e4;border-color:#d5533b}
.seat-content{width:100%;height:100%;padding:6px 5px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:2px}
.seat-name{font-size:24px;font-weight:950;line-height:1}.room-tag{margin-top:3px;background:#253d34;color:#fff;border-radius:999px;padding:3px 7px;font-size:11px;font-weight:900}
.seat-guest{font-size:10px;color:#5f6963;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.seat-meal{font-size:10px;line-height:1.25;font-weight:750;max-width:100%;white-space:normal}
.seat-dup{font-size:9px;color:#a53d2b;font-weight:900}
.plan-legend{max-width:980px;margin:12px auto 0;display:flex;gap:16px;flex-wrap:wrap;font-size:12px;color:#646e68}
.legend-dot{display:inline-block;width:11px;height:11px;border:1px solid #aaa;margin-right:5px;vertical-align:-1px}.legend-dot.used{background:#e8f1eb}.legend-dot.dup{background:#fff0e4}

.notes{
  margin-top:18px;border:1px solid #bfc4c0;
}
.note-section{padding:15px 18px}
.note-section.front{background:var(--front)}
.note-section.kitchen{background:var(--kitchen);border-top:1px solid #d4aaaa}
.note-title{font-weight:950;margin-bottom:10px}
.memo-toolbar{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px}
.num-btn,.emo-btn{
  border:1px solid rgba(60,70,64,.25);background:rgba(255,255,255,.7);
  cursor:pointer;font-weight:900
}
.num-btn{width:32px;height:32px;border-radius:50%}
.num-btn.active{background:#263d34;color:#fff;border-color:#263d34}
.num-btn.has{box-shadow:inset 0 0 0 3px rgba(44,73,61,.13)}
.emo-btn{border-radius:8px;padding:6px 8px;font-size:16px}
.emo-btn.active{background:#fff;border-color:#526b5f}
.memo-editor{display:grid;grid-template-columns:1fr;gap:7px}
.memo-editor textarea{
  width:100%;min-height:78px;resize:vertical;border:1px solid rgba(70,75,70,.28);
  background:rgba(255,255,255,.7);padding:10px 12px;font-family:inherit;line-height:1.55
}
.memo-list{margin-top:11px;font-size:13px;line-height:1.8}
.memo-line{display:flex;gap:8px}
.memo-circle{
  min-width:24px;height:24px;border-radius:50%;background:rgba(255,255,255,.75);
  display:inline-grid;place-items:center;font-size:11px;font-weight:900
}

.actions{
  display:flex;flex-wrap:wrap;gap:9px;justify-content:flex-end;margin-top:16px
}
.action{
  border:0;border-radius:10px;padding:12px 17px;font-weight:900;cursor:pointer;
  background:#263d34;color:#fff
}
.action.secondary{background:#e3e8e4;color:#34443b}
.action.excel{background:#166534}
.action.danger{background:#f2e1df;color:#973e36}
.save-state{font-size:11px;color:#67736c;margin-top:8px;text-align:right}
.hint{font-size:11px;color:#7c837e;margin-top:10px}

@media(max-width:700px){
  .floorplan{min-height:500px}.seat-name{font-size:18px}.seat-meal,.seat-guest{font-size:8px}
  .topbar{grid-template-columns:1fr;align-items:start}
  .title{text-align:left}
  .datebox{text-align:left}
  .notice{display:block;margin:4px 0 0}
}
@media print{
  body{background:white}
  .wrap{max-width:none;padding:0}
  .actions,.save-state,.hint,.memo-toolbar,.memo-editor{display:none}
  .table-wrap{box-shadow:none}
}

/* ===== 연박 상세 입력 · 객실별 좌석 메모 ===== */
.stay-detail-row td{
  background:#fffaf0;
  padding:10px 12px 12px;
  text-align:left;
}
.stay-editor{
  display:grid;
  grid-template-columns:150px minmax(0,1fr);
  gap:12px;
  align-items:start;
}
.stay-count-box{
  border:1px solid #e3d7b0;
  background:#fff7d9;
  padding:10px;
}
.stay-count-box label{
  display:block;
  font-size:11px;
  font-weight:900;
  color:#6f6446;
  margin-bottom:6px;
}
.stay-count-box select{
  width:100%;
  min-height:40px;
  border:1px solid #d8cfaf;
  background:#fff;
  border-radius:8px;
  font-weight:900;
}
.stay-breakfast-grid{
  display:grid;
  grid-template-columns:repeat(2,minmax(0,1fr));
  gap:8px;
}
.stay-breakfast-item{
  border:1px solid #dedfd9;
  background:#fff;
  padding:9px 10px;
}
.stay-breakfast-item label{
  display:block;
  margin-bottom:5px;
  color:#59615c;
  font-size:11px;
  font-weight:900;
}
.stay-breakfast-item textarea{
  width:100%;
  min-height:54px;
  resize:vertical;
  border:1px solid #d7dbd7;
  border-radius:8px;
  background:#fbfcfa;
  padding:8px 9px;
  font:inherit;
  font-size:12px;
  line-height:1.5;
  text-align:left;
}
.stay-help{
  grid-column:1/-1;
  font-size:10.5px;
  color:#847b65;
  line-height:1.55;
}

.seat{cursor:pointer}
.seat:focus-visible{outline:4px solid rgba(37,61,52,.18);outline-offset:-4px}
.seat-memo-preview{
  width:94%;
  margin-top:4px;
  padding:3px 5px;
  border-radius:7px;
  background:#fff7c9;
  color:#6d6039;
  font-size:8.5px;
  font-weight:800;
  line-height:1.25;
  overflow:hidden;
  display:-webkit-box;
  -webkit-line-clamp:2;
  -webkit-box-orient:vertical;
}
.seat-memo-empty{
  margin-top:3px;
  font-size:8px;
  color:#89918c;
  font-weight:800;
}
.seat-memo-trigger{
  margin-top:3px;
  border:1px solid #cbd4ce;
  background:#fff;
  border-radius:999px;
  padding:3px 6px;
  font-size:8px;
  font-weight:900;
  color:#496056;
  cursor:pointer;
}

.room-memo-panel{
  max-width:980px;
  margin:12px auto 0;
  border:1px solid #c9cfca;
  background:#f8faf7;
  padding:12px 14px;
  display:none;
}
.room-memo-panel.active{display:block}
.room-memo-head{
  display:flex;
  align-items:center;
  justify-content:space-between;
  gap:10px;
  margin-bottom:8px;
}
.room-memo-head b{font-size:13px;color:#344139}
.room-memo-head span{font-size:10.5px;color:#7b837e}
.room-memo-panel textarea{
  width:100%;
  min-height:76px;
  resize:vertical;
  border:1px solid #cfd5d1;
  border-radius:10px;
  background:#fff;
  padding:10px 12px;
  font:inherit;
  font-size:13px;
  line-height:1.55;
  text-align:left;
}
.room-memo-actions{
  display:flex;
  justify-content:flex-end;
  margin-top:7px;
}
.room-memo-close{
  border:1px solid #d7dcd8;
  background:#fff;
  border-radius:8px;
  padding:7px 10px;
  font-size:11px;
  font-weight:900;
  cursor:pointer;
  color:#56615a;
}

/* 모바일에서 도식 자체 비율을 보존하고, 도식 영역만 가로 스크롤 */
@media(max-width:700px){
  .table-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch}
  .stay-editor{grid-template-columns:1fr}
  .stay-breakfast-grid{grid-template-columns:1fr}
  .stay-count-box select,
  .stay-breakfast-item textarea,
  .room-memo-panel textarea{font-size:16px}

  .floorplan-shell{
    max-width:none;
    overflow-x:auto;
    -webkit-overflow-scrolling:touch;
    padding:10px 10px 14px;
    scroll-snap-type:x proximity;
  }
  .floorplan{
    width:820px;
    min-width:820px;
    min-height:612px;
    aspect-ratio:1.34/1;
    scroll-snap-align:start;
  }
  .seat-name{font-size:21px}
  .seat-guest,.seat-meal{font-size:9px}
  .seat-memo-preview{font-size:8px}
  .layout-help::after{
    content:" · 모바일에서는 좌석도를 좌우로 밀어서 확인 가능.";
    font-weight:800;
    color:#4c6158;
  }
  .room-memo-panel{margin-left:0;margin-right:0}
}


/* ===== 조식 리딩연습 · 행 집중 모드 ===== */
#breakfastTable tbody tr.reading-row,
#breakfastTable tbody tr.stay-detail-row{
  transition:opacity .18s ease, background .18s ease, filter .18s ease, box-shadow .18s ease;
}
#breakfastTable tbody tr.reading-row:nth-of-type(odd) td{
  background-color:#fff;
}
#breakfastTable tbody tr.reading-row:nth-of-type(even) td{
  background-color:#fafbf9;
}
#breakfastTable tbody tr.reading-row:hover td{
  background-color:#f7faf7;
}
#breakfastTable.focus-mode tbody tr[data-row-index]{
  opacity:.24;
  filter:saturate(.6);
}
#breakfastTable.focus-mode tbody tr[data-row-index].row-focused{
  opacity:1;
  filter:none;
}
#breakfastTable.focus-mode tbody tr.reading-row.row-focused td{
  background:#fffdf4;
  border-top-color:#8da096;
  border-bottom-color:#8da096;
}
#breakfastTable.focus-mode tbody tr.reading-row.row-focused td:first-child{
  box-shadow:inset 4px 0 0 #3c5c4d;
  font-weight:950;
  color:#253d34;
}
#breakfastTable.focus-mode tbody tr.stay-detail-row.row-focused td{
  background:#fff8df;
}
#breakfastTable.focus-mode tbody tr[data-row-index]:not(.row-focused):hover{
  opacity:.55;
  filter:none;
}
.focus-guide{
  margin:0 0 10px;
  padding:9px 11px;
  border:1px solid #d8dfda;
  background:#f8faf7;
  color:#68726c;
  font-size:11px;
  line-height:1.55;
}
.focus-guide b{color:#31463b}

/* ===== 날짜별 조식 리딩 저장 ===== */
.action.archive{
  background:#715b7d;
  color:#fff;
}
.reading-archive{
  margin-top:22px;
  border:1px solid #c9ceca;
  background:#fff;
}
.archive-head{
  display:flex;
  justify-content:space-between;
  align-items:flex-end;
  gap:12px;
  padding:16px 18px 13px;
  border-bottom:1px solid #daddda;
  background:linear-gradient(110deg,#f8f0f5,#fbf8e9,#f1f6f2);
}
.archive-head h2{
  margin:0;
  font-size:17px;
  letter-spacing:-.03em;
}
.archive-head p{
  margin:4px 0 0;
  color:#778079;
  font-size:11px;
  line-height:1.5;
}
.archive-count{
  flex:0 0 auto;
  border:1px solid #d4d9d5;
  background:#fff;
  border-radius:999px;
  padding:5px 9px;
  font-size:10px;
  color:#68726c;
  font-weight:900;
}
.archive-list{
  display:grid;
  gap:8px;
  padding:12px;
}
.archive-empty{
  padding:20px 12px;
  text-align:center;
  color:#8a918d;
  font-size:12px;
}
.archive-item{
  width:100%;
  appearance:none;
  border:1px solid #d9ddda;
  background:#fff;
  border-radius:11px;
  padding:12px 13px;
  display:grid;
  grid-template-columns:minmax(0,1fr) auto;
  align-items:center;
  gap:10px;
  text-align:left;
  cursor:pointer;
  transition:.16s ease;
}
.archive-item:hover{
  border-color:#abb7b0;
  background:#fafcf9;
  transform:translateY(-1px);
}
.archive-item b{
  display:block;
  color:#344139;
  font-size:13px;
  margin-bottom:3px;
}
.archive-item small{
  display:block;
  color:#848c87;
  font-size:10px;
}
.archive-item .open-mark{
  font-size:17px;
  color:#78847d;
}
.archive-toast{
  position:fixed;
  left:50%;
  bottom:24px;
  transform:translateX(-50%) translateY(12px);
  z-index:2500;
  opacity:0;
  visibility:hidden;
  background:#263d34;
  color:#fff;
  border-radius:999px;
  padding:9px 13px;
  font-size:11px;
  font-weight:900;
  box-shadow:0 10px 30px rgba(20,30,25,.18);
  transition:.2s ease;
}
.archive-toast.show{
  opacity:1;
  visibility:visible;
  transform:translateX(-50%) translateY(0);
}

/* 저장 기록 보기 팝업 */
.archive-modal{
  position:fixed;
  inset:0;
  z-index:2400;
  display:grid;
  place-items:center;
  padding:18px;
  background:rgba(25,31,28,.46);
  opacity:0;
  visibility:hidden;
  transition:.2s ease;
}
.archive-modal.open{opacity:1;visibility:visible}
.archive-dialog{
  width:min(820px,96vw);
  max-height:88vh;
  display:flex;
  flex-direction:column;
  overflow:hidden;
  border-radius:16px;
  background:#fff;
  border:1px solid #cfd5d1;
  box-shadow:0 26px 80px rgba(18,26,22,.28);
}
.archive-dialog-head{
  display:flex;
  justify-content:space-between;
  gap:14px;
  align-items:flex-start;
  padding:16px 18px;
  border-bottom:1px solid #dde1de;
  background:linear-gradient(110deg,#f8f0f5,#fbf8e9,#f1f6f2);
}
.archive-dialog-head h3{
  margin:0;
  font-size:17px;
  letter-spacing:-.03em;
}
.archive-dialog-head p{
  margin:4px 0 0;
  font-size:10.5px;
  color:#7a837d;
}
.archive-close{
  width:36px;height:36px;
  flex:0 0 auto;
  border:1px solid #d2d7d3;
  border-radius:50%;
  background:#fff;
  cursor:pointer;
  font-size:19px;
  color:#5d6761;
}
.archive-dialog-body{
  overflow:auto;
  padding:14px 16px 18px;
}
.archive-summary{
  display:grid;
  gap:9px;
}
.archive-room{
  border:1px solid #dce0dd;
  border-radius:11px;
  overflow:hidden;
  background:#fff;
}
.archive-room-head{
  display:flex;
  flex-wrap:wrap;
  justify-content:space-between;
  gap:7px;
  padding:10px 12px;
  background:#f7f9f7;
  border-bottom:1px solid #e2e5e3;
}
.archive-room-head b{font-size:12px;color:#314239}
.archive-room-head span{font-size:10px;color:#7c857f}
.archive-room-body{
  padding:10px 12px;
  color:#59635d;
  font-size:11.5px;
  line-height:1.7;
}
.archive-room-body .memo{
  margin-top:7px;
  padding:7px 9px;
  border-left:3px solid #d7c88f;
  background:#fffaf0;
  color:#675f49;
}
.archive-stay{
  margin-top:7px;
  padding:7px 9px;
  background:#f8f2fb;
  border:1px solid #e2d8e8;
}
.archive-notes{
  margin-top:14px;
  display:grid;
  grid-template-columns:1fr 1fr;
  gap:10px;
}
.archive-note-card{
  padding:10px 11px;
  border:1px solid #dce0dd;
  background:#fafbf9;
  font-size:11px;
  line-height:1.65;
}
.archive-note-card b{
  display:block;
  margin-bottom:5px;
  color:#3d4942;
}
.archive-dialog-actions{
  display:flex;
  justify-content:flex-end;
  gap:8px;
  padding:11px 14px;
  border-top:1px solid #dde1de;
  background:#fafbf9;
}
.archive-dialog-actions button{
  border:1px solid #cfd5d1;
  border-radius:9px;
  padding:9px 12px;
  background:#fff;
  color:#455149;
  font-size:11px;
  font-weight:900;
  cursor:pointer;
}
.archive-dialog-actions .load{
  background:#263d34;
  color:#fff;
  border-color:#263d34;
}

@media(max-width:700px){
  #breakfastTable.focus-mode tbody tr[data-row-index]{opacity:.18}
  #breakfastTable.focus-mode tbody tr[data-row-index].row-focused{opacity:1}
  .archive-head{align-items:flex-start}
  .archive-list{padding:9px}
  .archive-item{padding:11px}
  .archive-modal{padding:7px}
  .archive-dialog{width:100%;max-height:94vh;border-radius:13px}
  .archive-notes{grid-template-columns:1fr}
}


/* ===== 입력 완료 · 룸터치 시간 · 좌석 가독성 개선 ===== */
.complete-btn{
  border:1px solid #bfc9c1;background:#f7faf7;color:#2e4237;border-radius:9px;
  padding:7px 8px;font-size:11px;font-weight:950;cursor:pointer;white-space:nowrap;
}
.complete-btn:hover{background:#e9f3eb;border-color:#91aa98}
.reading-row.is-completed{background:#f8fbf8}
.reading-row.sparkle-complete{position:relative;animation:rowSparkle 1.05s ease both}
.reading-row.sparkle-complete td{animation:cellGlow 1.05s ease both}
@keyframes rowSparkle{
  0%{transform:scale(1)} 28%{transform:scale(1.008)} 52%{transform:scale(1)}
}
@keyframes cellGlow{
  0%{box-shadow:inset 0 0 0 rgba(255,220,90,0);background:inherit}
  25%{box-shadow:inset 0 0 18px rgba(255,215,70,.72);background:#fff9d9}
  65%{box-shadow:inset 0 0 9px rgba(255,235,150,.42);background:#fffdf1}
  100%{box-shadow:none}
}
.roomtouch-box{
  border:1px solid #cad8d0;background:#f3f8f5;padding:10px 12px;display:flex;align-items:center;
  gap:8px;flex-wrap:wrap;font-size:12px;font-weight:850;color:#3f5448;
}
.roomtouch-box input[type=time]{width:112px;border:1px solid #bdcbc2;background:#fff;border-radius:8px;padding:7px 8px;font-weight:900}
.roomtouch-label{color:#52655b}

/* 좌석 안의 핵심 정보가 잘리지 않도록 압축 표시함 */
.seat-content{padding:4px 3px;gap:1px;justify-content:flex-start;overflow:hidden}
.seat-name{font-size:17px;margin-top:2px}
.room-tag{margin-top:1px;padding:2px 5px;font-size:9px}
.seat-guest{font-size:8px;line-height:1.15;white-space:normal;overflow:visible;text-overflow:clip}
.seat-meal{font-size:8px;line-height:1.15}
.seat-people{padding:2px 5px;gap:3px}.seat-people small{font-size:7px}.seat-people b{font-size:10px}
.seat .seat-order-wrap{gap:2px;line-height:1}
.seat .order-chip{min-width:auto;min-height:auto;padding:2px 4px;font-size:7.5px;line-height:1.05;border-radius:7px}
.seat-memo-preview,.seat-memo-empty{display:none!important}
.seat-memo-trigger{
  margin-top:2px;border:1px solid #b9c6bd;background:rgba(255,255,255,.88);border-radius:7px;
  padding:2px 5px;font-size:7.5px;font-weight:900;color:#405348;cursor:pointer;line-height:1.15;
}
.seat[title]{cursor:pointer}
.seat:hover,.seat:focus-within{z-index:70;overflow:visible!important;box-shadow:0 8px 22px rgba(25,45,35,.20)}
.seat:hover .seat-content,.seat:focus-within .seat-content{background:#fff;border:1px solid #b8c7be;min-height:100%;height:auto;padding:5px 4px;position:relative;z-index:71}
@media(max-width:700px){
  .seat-name{font-size:14px}.room-tag{font-size:7.5px;padding:1px 4px}.seat-guest,.seat-meal{font-size:6.8px}
  .seat .order-chip{font-size:6.4px;padding:1px 3px}.seat-memo-trigger{font-size:6.5px;padding:1px 3px}
}


/* ===== v4 · 시간즉시정렬 / 연박강조 / 완료요약 / 모바일 ===== */
.table-wrap{
  overflow-x:scroll!important; overflow-y:auto; max-width:100%; -webkit-overflow-scrolling:touch;
  scrollbar-gutter:stable both-edges; overscroll-behavior-x:contain;
}
#breakfastTable{min-width:1360px!important}
.table-wrap::-webkit-scrollbar{height:13px;width:11px}
.table-wrap::-webkit-scrollbar-thumb{background:#9aa79f;border:3px solid #f2f4f1;border-radius:999px}
.table-wrap::-webkit-scrollbar-track{background:#eef1ed}
.menu-stack-cell{padding:3px!important;background:#fbfcfa}
.menu-stack-cell .menu-stack-title{font-size:9px;font-weight:950;color:#405048;margin-bottom:2px}
.menu-stack-cell select{font-size:12px;font-weight:950;border:1px solid #d2d9d4;background:#fff;border-radius:7px;padding:4px 2px}
.reading-row.is-stay td{background:#e8f5e4!important}
.reading-row.is-stay td:first-child{box-shadow:inset 4px 0 0 #79a96f}
.stay-detail-row.is-stay td{background:#f1f9ee!important}
.completed-summary-cell{padding:6px 9px!important;text-align:left!important;background:#f8fbf8!important}
.completed-summary{display:flex;align-items:center;gap:6px;flex-wrap:wrap;min-height:34px}
.completed-summary .sum-room{font-size:13px;font-weight:950;color:#22382d}
.completed-summary .sum-pill{display:inline-flex;align-items:center;min-height:24px;padding:3px 7px;border-radius:999px;background:#fff;border:1px solid #d4ddd7;font-size:10.5px;font-weight:800;color:#4a5b52}
.completed-summary .sum-pill.stay{background:#e3f2df;border-color:#bdd7b7;color:#3d6637}
.completed-summary .sum-menu{font-weight:950;color:#2f4639}
.completed-summary .sum-spacer{flex:1}
.complete-btn.edit-mode{background:#fff5db;border-color:#dec98a;color:#66551f}
.time-burst{animation:timeMoveFlash .7s ease both}
@keyframes timeMoveFlash{0%{box-shadow:inset 0 0 0 999px rgba(219,239,251,.75)}100%{box-shadow:none}}
.seat-content{overflow:visible!important;justify-content:center!important}
.seat-menu-text{font-size:8px!important;line-height:1.15;font-weight:950;color:#2d4236;white-space:normal;word-break:keep-all;text-align:center;max-width:100%;padding:1px 2px}
.seat-time{font-size:8px;font-weight:900;color:#52645a}
.floorplan-shell{overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-gutter:stable}
.floorplan-shell::-webkit-scrollbar{height:10px}.floorplan-shell::-webkit-scrollbar-thumb{background:#9aa79f;border-radius:999px}
@media(max-width:700px){
  .wrap{padding:12px 8px 42px}
  .layout-card{padding:10px;margin-top:10px}
  .layout-help{font-size:10px}
  .floorplan-shell{max-width:100%;padding:4px}
  .floorplan{width:720px!important;min-width:720px!important;min-height:535px!important;aspect-ratio:auto!important}
  #breakfastTable{min-width:1180px!important}
  th{font-size:9.5px;padding:4px 2px}td{font-size:10.5px;padding:4px 2px}
  .completed-summary{gap:4px}.completed-summary .sum-pill{font-size:9px;padding:2px 5px;min-height:21px}
  .completed-summary .sum-room{font-size:11px}.complete-btn{font-size:10px;padding:6px}
  .seat-name{font-size:14px!important}.room-tag{font-size:7.5px!important}.seat-guest,.seat-time,.seat-menu-text{font-size:7px!important}
}


/* ===== v4.2 · 완료행 선명도 + 전체 입력 확정 ===== */
#breakfastTable.focus-mode tbody tr.reading-row.is-completed{opacity:1!important}
.reading-row.is-completed td{
  color:#22352b!important;
  border-color:#b9c9be!important;
  font-weight:800;
}
.reading-row.is-completed:not(.is-stay) td{background:#f3f8f4!important}
.reading-row.is-completed.is-stay td,
.reading-row.is-stay.is-completed td{
  background:#dff1da!important;
  border-color:#a9c7a2!important;
}
.reading-row.is-stay.is-completed td:first-child{box-shadow:inset 5px 0 0 #5f9b55}
.reading-row.is-completed .completed-summary-cell{
  background:#f3f8f4!important;
}
.reading-row.is-stay.is-completed .completed-summary-cell{
  background:#dff1da!important;
}
.reading-row.is-completed .completed-summary .sum-room{color:#183426!important}
.reading-row.is-completed .completed-summary .sum-pill{
  background:#fff!important;border-color:#b9c9be!important;color:#30483a!important;
}
.reading-row.is-stay.is-completed .completed-summary .sum-pill{
  background:#f7fff5!important;border-color:#a9c7a2!important;color:#315c2d!important;
}
.reading-row.is-stay.is-completed .completed-summary .sum-pill.stay{
  background:#cfe8c8!important;border-color:#8fb785!important;color:#285622!important;
}
.finalize-wrap{
  margin:14px 0 18px;
  display:flex;
  justify-content:center;
}
.finalize-btn{
  width:min(720px,100%);
  min-height:52px;
  border:1px solid #9bb3a1;
  border-radius:14px;
  background:linear-gradient(135deg,#eaf5e8 0%,#f8f6dc 100%);
  color:#294235;
  font-size:14px;
  font-weight:950;
  letter-spacing:-.02em;
  cursor:pointer;
  box-shadow:0 8px 20px rgba(45,75,55,.08);
  transition:.18s ease;
}
.finalize-btn:not(:disabled):hover{transform:translateY(-1px);box-shadow:0 10px 24px rgba(45,75,55,.14)}
.finalize-btn:disabled{
  cursor:not-allowed;
  opacity:.55;
  background:#f1f2ef;
  border-color:#d3d7d2;
  color:#8a928c;
  box-shadow:none;
}
.finalize-btn.is-saved{
  background:linear-gradient(135deg,#d6efd2 0%,#fff1b8 100%);
  border-color:#7fa875;
  color:#1f5128;
  animation:finalSaveSparkle .95s ease both;
}
@keyframes finalSaveSparkle{
  0%{transform:scale(1);box-shadow:0 0 0 rgba(255,220,80,0)}
  28%{transform:scale(1.015);box-shadow:0 0 0 8px rgba(255,224,100,.20),0 10px 28px rgba(65,100,70,.18)}
  60%{transform:scale(1);box-shadow:0 0 24px rgba(255,220,80,.45)}
  100%{box-shadow:0 8px 20px rgba(45,75,55,.10)}
}
.finalize-sub{
  display:block;
  margin-top:3px;
  font-size:10px;
  font-weight:750;
  color:inherit;
  opacity:.72;
}
@media(max-width:700px){
  .finalize-wrap{margin:10px 0 14px}
  .finalize-btn{min-height:48px;border-radius:12px;font-size:13px;padding:8px 10px}
}

/* ===== v4.3 · 좌석 가독성 / 좌석→차트 이동 / 전체확정 잠금 ===== */
.seat-content{gap:3px!important;padding:6px 5px!important}
.seat-name{font-size:23px!important;font-weight:1000!important;color:#18261f!important;line-height:1!important}
.room-tag{font-size:14px!important;font-weight:1000!important;padding:4px 8px!important;background:#203c30!important;color:#fff!important;line-height:1.1!important}
.seat-people{display:flex!important;flex-direction:column!important;gap:0!important;padding:3px 7px!important;background:#eef5f0!important;border:1px solid #bfd0c4!important;border-radius:9px!important}
.seat-people small{font-size:8px!important;line-height:1!important;color:#66766d!important}
.seat-people b{font-size:14px!important;line-height:1.1!important;color:#1f382b!important}
.seat-guest{font-size:10px!important;line-height:1.2!important;font-weight:850!important;color:#53665a!important;white-space:normal!important}
.seat-time{font-size:10px!important;font-weight:950!important;color:#445a4d!important}
.seat-menu-text{font-size:11.5px!important;line-height:1.25!important;font-weight:1000!important;color:#1f3f2d!important;white-space:normal!important}
.seat-memo-trigger{font-size:9px!important;padding:3px 7px!important;margin-top:2px!important}
.seat.occupied{cursor:pointer}
.seat.occupied:hover .room-tag,.seat.occupied:focus-within .room-tag{background:#142d23!important}
.reading-row.seat-jump-highlight td{animation:seatJumpGlow 1.5s ease both!important}
@keyframes seatJumpGlow{
  0%{box-shadow:inset 0 0 0 999px rgba(255,230,115,.70)}
  35%{box-shadow:inset 0 0 0 999px rgba(255,244,188,.62)}
  100%{box-shadow:none}
}
.reading-row.final-unused td{filter:blur(3px);opacity:.30!important;pointer-events:none!important;user-select:none!important}
#breakfastTable.final-locked .complete-btn.edit-mode{pointer-events:none!important;opacity:.48!important}
.finalize-wrap{flex-direction:column;align-items:center;gap:8px}
.final-edit-btn{
  display:none;width:min(720px,100%);min-height:40px;border:1px solid #c3c9c5;border-radius:12px;
  background:#fff;color:#59645d;font-size:12px;font-weight:900;cursor:pointer
}
.final-edit-btn.show{display:block}
.final-edit-btn:hover{background:#f4f6f4}
@media(max-width:700px){
  .seat-name{font-size:20px!important}.room-tag{font-size:12px!important;padding:3px 6px!important}
  .seat-people b{font-size:12px!important}.seat-guest{font-size:8.5px!important}.seat-time{font-size:8.5px!important}.seat-menu-text{font-size:9.5px!important}
  .seat-memo-trigger{font-size:8px!important}
  .final-edit-btn{font-size:11px;min-height:38px}
}


/* ===== v4.4 · 완료행 열 유지 / 좌석 핵심정보 확대 ===== */
.reading-row.is-completed td{opacity:1!important;filter:none!important;background:#fff!important;color:#1f2a24!important;border-color:#b7c2ba!important;padding:8px 5px!important;font-weight:800!important}
.reading-row.is-completed.is-stay td{background:#dff2e4!important;border-color:#9bc7a7!important;color:#173c25!important}
.reading-row.is-completed .done-cell{font-size:11.5px!important;line-height:1.25!important;vertical-align:middle!important}
.reading-row.is-completed .done-room{font-size:14px!important;font-weight:1000!important}
.reading-row.is-completed .done-time{font-size:12.5px!important;font-weight:1000!important}
.reading-row.is-completed .done-menu{font-size:12px!important;font-weight:950!important}
.reading-row.is-completed .done-people b{display:block;font-size:13px!important}.reading-row.is-completed .done-people small{display:block;font-size:9px!important;color:#4e6255!important;margin-top:2px}
.reading-row.is-completed .done-stay{background:#cfeacf!important;color:#184d2b!important}
.reading-row.is-completed .done-total{font-size:14px!important;font-weight:1000!important;background:#eef4ef!important}
.seat{overflow:visible!important;transform-origin:center center;transition:transform .28s cubic-bezier(.2,.8,.2,1),box-shadow .28s ease,z-index 0s}
.seat:hover,.seat.expanded{z-index:120!important;transform:scale(1.38);box-shadow:0 18px 45px rgba(22,45,32,.28)!important}
.seat:hover .seat-content,.seat.expanded .seat-content{background:#fffefb!important;border:2px solid #7ea68b!important;border-radius:8px;min-height:100%;height:auto!important;padding:8px 7px!important}
.seat-name{font-size:24px!important;color:#22352b!important;margin-bottom:2px!important}
.room-tag.seat-room-jump{appearance:none;border:0!important;background:#24483a!important;color:#fff!important;border-radius:999px!important;padding:5px 10px!important;font-size:14px!important;font-weight:1000!important;line-height:1!important;cursor:pointer}
.seat-people{display:flex!important;flex-direction:column!important;align-items:center!important;gap:1px!important;padding:4px 8px!important;border-radius:10px!important}
.seat-people small{font-size:8.5px!important}.seat-people b{font-size:15px!important}.seat-people em{font-style:normal;font-size:9px!important;font-weight:900!important;color:#51665a!important;white-space:nowrap}
.seat-menu-label{font-size:8px;font-weight:900;color:#7a887f;line-height:1;margin-top:1px}
.seat-menu-text{font-size:13px!important;line-height:1.18!important;font-weight:1000!important;color:#173b28!important;white-space:normal!important;overflow:visible!important}
.seat-salad{font-size:11px!important;font-weight:900!important;color:#325c43;background:#eef7f0;border:1px solid #c4ddca;border-radius:8px;padding:3px 7px;white-space:nowrap}.seat-salad b{font-size:12px!important}
.seat-memo-trigger{font-size:9px!important;padding:3px 7px!important}
@media(max-width:760px){
  .seat:hover{transform:none;box-shadow:none!important}.seat.expanded{transform:scale(1.55)!important;z-index:180!important;box-shadow:0 20px 50px rgba(22,45,32,.34)!important}
  .seat.expanded .seat-name{font-size:22px!important}.seat.expanded .room-tag.seat-room-jump{font-size:13px!important}.seat.expanded .seat-menu-text{font-size:12px!important}.seat.expanded .seat-salad{font-size:10px!important}
  .reading-row.is-completed .done-cell{font-size:10.5px!important}.reading-row.is-completed .done-room{font-size:12.5px!important}
}


/* ===== v4.5 · 좌석배치 대형화 + 모바일 가독성 + 완료행 터치 상세 ===== */
html,body{touch-action:pan-x pan-y pinch-zoom;-webkit-text-size-adjust:100%}
.layout-card{overflow:visible}
.floorplan-shell{
  max-width:1180px!important;
  overflow:auto!important;
  -webkit-overflow-scrolling:touch;
  scrollbar-gutter:stable both-edges;
  padding:12px!important;
  touch-action:pan-x pan-y pinch-zoom;
}
.floorplan{
  width:100%!important;
  min-width:1080px!important;
  min-height:790px!important;
  aspect-ratio:1.36/1!important;
}
.floorplan-shell::-webkit-scrollbar{height:12px}
.floorplan-shell::-webkit-scrollbar-thumb{background:#93a198;border:3px solid #eef2ef;border-radius:999px}
.floorplan-shell::-webkit-scrollbar-track{background:#eef2ef}

/* 모든 테이블의 핵심 정보가 기본 상태에서도 읽히도록 확대 */
.seat-content{padding:9px 7px!important;gap:5px!important;justify-content:center!important}
.seat-name{font-size:31px!important;font-weight:1000!important;line-height:.95!important;color:#17271f!important}
.room-tag.seat-room-jump{font-size:16px!important;padding:6px 12px!important;min-height:32px!important}
.seat-people{padding:6px 10px!important;border-radius:12px!important;min-width:92px}
.seat-people small{font-size:10px!important}.seat-people b{font-size:18px!important}.seat-people em{font-size:11px!important}
.seat-menu-label{font-size:10px!important;margin-top:2px!important}
.seat-menu-text{font-size:15px!important;line-height:1.25!important;max-width:100%!important}
.seat-salad{font-size:13px!important;padding:5px 9px!important}.seat-salad b{font-size:14px!important}
.seat-memo-trigger{font-size:11px!important;min-height:30px!important;padding:5px 9px!important}

/* PC에서는 호버, 모바일에서는 탭으로 부드럽게 확대 */
@media (hover:hover) and (pointer:fine){
  .seat.occupied:hover{transform:scale(1.24)!important;z-index:180!important}
  .seat.occupied:hover .seat-content{border:2px solid #729c82!important;box-shadow:0 16px 36px rgba(22,45,32,.24)!important}
}
.seat.expanded{transform:scale(1.30)!important;z-index:200!important}
.seat.expanded .seat-content{border:2px solid #729c82!important;box-shadow:0 18px 42px rgba(22,45,32,.28)!important;background:#fffefb!important}

/* 완료된 행은 칸 구분을 유지하고 클릭/터치 시 전체 행을 선명하게 강조 */
.reading-row.is-completed{cursor:pointer;transition:box-shadow .18s ease,transform .18s ease}
.reading-row.is-completed td{transition:background .18s ease,border-color .18s ease,color .18s ease}
.reading-row.is-completed.row-selected td{
  background:#fff7cf!important;
  border-top:2px solid #b79c32!important;
  border-bottom:2px solid #b79c32!important;
  color:#17271f!important;
}
.reading-row.is-completed.is-stay.row-selected td{
  background:#cdeccb!important;
  border-color:#699d68!important;
}

/* 선택한 완료행 아래에 첨부 예시처럼 세로형 핵심카드 표시 */
.completed-detail-row td{padding:10px!important;background:#f7f8f5!important;border:0!important}
.completed-detail-card{
  width:min(340px,94vw);margin:0 auto;padding:12px 14px 14px;
  display:flex;flex-direction:column;align-items:center;gap:7px;
  background:#fffefa;border:1px solid #bfcac2;border-radius:18px;
  box-shadow:0 10px 24px rgba(30,50,38,.10);color:#1d3026;
}
.completed-detail-card.is-stay{background:#eef8eb;border-color:#9cc99b}
.completed-detail-seat{font-size:25px;font-weight:1000;line-height:1;color:#193326}
.completed-detail-room{background:#17392c;color:#fff;border-radius:999px;padding:5px 12px;font-size:17px;font-weight:1000;line-height:1.1}
.completed-detail-people{display:flex;flex-direction:column;align-items:center;padding:7px 12px;border:1px solid #bdd1c2;background:#eef5f0;border-radius:12px;line-height:1.08}
.completed-detail-people small{font-size:10px;font-weight:900;color:#607267}.completed-detail-people b{font-size:19px;font-weight:1000}.completed-detail-people span{font-size:11px;font-weight:900;color:#496354;margin-top:3px}
.completed-detail-label{font-size:10px;color:#7a8980;font-weight:900;line-height:1;margin-top:1px}
.completed-detail-menu{font-size:18px;font-weight:1000;line-height:1.2;text-align:center;color:#163c28;word-break:keep-all}
.completed-detail-salad{font-size:14px;font-weight:950;padding:5px 10px;background:#edf7ef;border:1px solid #bdd9c3;border-radius:10px}
.completed-detail-meta{font-size:11px;font-weight:850;color:#56685e;text-align:center}
.completed-detail-actions{display:flex;gap:7px;justify-content:center;flex-wrap:wrap}
.completed-detail-actions button{border:1px solid #c6d0c9;background:#fff;border-radius:9px;padding:6px 10px;font-size:11px;font-weight:900;cursor:pointer}

/* 표 영역도 손가락 확대/축소 및 좌우 스크롤을 방해하지 않음 */
.table-wrap{touch-action:pan-x pan-y pinch-zoom!important}
#breakfastTable{touch-action:pan-x pan-y pinch-zoom}

@media(max-width:700px){
  .wrap{padding:12px 8px 50px!important}
  .layout-card{padding:12px 8px!important}
  .layout-head{align-items:flex-start!important;flex-direction:column!important;gap:5px!important}
  .layout-help{font-size:11px!important}
  .floorplan-shell{padding:6px!important;margin:0!important;max-width:100%!important}
  .floorplan{min-width:980px!important;min-height:720px!important}
  .seat-name{font-size:28px!important}
  .room-tag.seat-room-jump{font-size:15px!important;padding:6px 11px!important}
  .seat-people b{font-size:17px!important}.seat-people small{font-size:9px!important}.seat-people em{font-size:10px!important}
  .seat-menu-text{font-size:14px!important}.seat-salad{font-size:12px!important}.seat-salad b{font-size:13px!important}
  .seat.expanded{transform:scale(1.38)!important}
  .table-wrap{margin-left:-2px;margin-right:-2px}
  table{min-width:1160px!important}
  .reading-row.is-completed td{font-size:11.5px!important;min-height:42px!important;padding:8px 6px!important}
  .reading-row.is-completed .done-room{font-size:14px!important}
  .reading-row.is-completed .done-time{font-size:13px!important}
  .reading-row.is-completed .done-menu{font-size:12px!important}
  .completed-detail-card{width:min(315px,92vw);padding:11px 12px}
  .completed-detail-menu{font-size:17px}
}

/* ===== 조식 LIST 텍스트 가져오기 ===== */
.sheet-import-bar{display:flex;align-items:center;justify-content:space-between;gap:12px;margin:12px 0 16px;padding:13px 14px;border:1px solid #d7ded9;background:#fbfcfa;border-radius:14px;flex-wrap:wrap}
.sheet-import-copy b{display:block;font-size:13px;color:#34443b}.sheet-import-copy small{display:block;margin-top:3px;color:#7b857f;font-size:10.5px;line-height:1.55}
.sheet-import-btn{appearance:none;border:1px solid #b9cbbf;background:#edf5ef;color:#294437;border-radius:999px;padding:10px 15px;font:850 12px/1.2 inherit;cursor:pointer;min-height:40px}.sheet-import-btn:hover{background:#e5f0e8}
.ocr-modal{position:fixed;inset:0;z-index:45000;background:rgba(28,36,31,.46);display:none;place-items:center;padding:14px;backdrop-filter:blur(5px)}
.ocr-modal.open{display:grid}.ocr-dialog{width:min(900px,100%);max-height:92vh;overflow:auto;background:#fff;border:1px solid #cfd8d2;border-radius:18px;box-shadow:0 24px 70px rgba(18,30,22,.24)}
.ocr-head{position:sticky;top:0;z-index:2;display:flex;justify-content:space-between;gap:14px;align-items:flex-start;padding:17px 18px;background:#f6faf7;border-bottom:1px solid #dce4df}.ocr-head h3{margin:0;font-size:17px;color:#2f4036}.ocr-head p{margin:5px 0 0;font-size:11px;color:#738078;line-height:1.6}.ocr-close{border:1px solid #ccd5cf;background:white;border-radius:50%;width:36px;height:36px;font-size:20px;cursor:pointer}
.ocr-body{padding:18px}.json-paste{width:100%;min-height:290px;resize:vertical;border:1px solid #cfd8d2;border-radius:12px;padding:14px;font:12px/1.65 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;background:#fbfcfb;color:#24352b;outline:none}.json-paste:focus{border-color:#7f9a89;box-shadow:0 0 0 3px rgba(89,125,101,.10)}
.json-help{margin:0 0 12px;padding:11px 12px;border-left:3px solid #95ad9d;background:#f5f9f6;color:#5e6c63;font-size:11px;line-height:1.65}.json-status{min-height:20px;margin:9px 0 0;font-size:11px;font-weight:850;color:#5b6b61}.json-status.error{color:#a13b3b}.json-status.ok{color:#2e7048}
.ocr-actions{display:flex;justify-content:flex-end;gap:8px;margin-top:14px;flex-wrap:wrap}.ocr-secondary,.ocr-apply{min-height:42px;border-radius:10px;padding:9px 14px;font-weight:900;cursor:pointer}.ocr-secondary{border:1px solid #ccd5cf;background:white;color:#56645b}.ocr-apply{border:0;background:#315c43;color:white}
.json-example{margin-top:12px}.json-example summary{cursor:pointer;font-size:11px;font-weight:850;color:#617067}.json-example pre{white-space:pre-wrap;word-break:break-word;background:#f6f7f5;border:1px solid #dde3df;border-radius:10px;padding:11px;font:10.5px/1.55 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;color:#59665e}
.import-highlight td{animation:importGlow 1.4s ease both}@keyframes importGlow{0%{background:#fff0ad!important}55%{background:#fff8d9!important}100%{}}
@media(max-width:760px){.sheet-import-bar{padding:11px}.sheet-import-btn{width:100%;font-size:13px}.ocr-modal{padding:5px}.ocr-dialog{max-height:96vh;border-radius:14px}.ocr-body{padding:12px}.json-paste{min-height:320px;font-size:16px}.ocr-head{padding:14px}}


/* ===== v6 · 행 삭제 / 날짜별 화면 분리 ===== */
.row-action-wrap{display:flex;align-items:center;justify-content:center;gap:5px;flex-wrap:wrap}
.delete-row-btn{appearance:none;border:1px solid #e4c3c3;background:#fff7f7;color:#9b4242;border-radius:8px;padding:6px 8px;font:850 10px/1 inherit;cursor:pointer;white-space:nowrap}
.delete-row-btn:hover{background:#ffeded;border-color:#d99f9f}
.date-switch-burst{animation:dateSwitchBurst .65s ease both}
@keyframes dateSwitchBurst{0%{opacity:.35;transform:translateY(5px)}100%{opacity:1;transform:none}}
@media(max-width:760px){.row-action-wrap{gap:4px}.delete-row-btn,.complete-btn{min-height:34px;padding:7px 8px}}

/* ===== v5 · ORDER TO JSON / 연박 검수 ===== */
.sheet-import-actions{display:flex;align-items:center;justify-content:flex-end;gap:8px;flex-wrap:wrap;margin-left:auto}
.sheet-order-btn{appearance:none;border:1px solid #c6cbc8;background:#fff;color:#59635e;border-radius:999px;padding:10px 14px;font:900 10px/1.2 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;letter-spacing:.08em;cursor:pointer;min-height:40px;white-space:nowrap}
.sheet-order-btn:hover{background:#f3f4f3;border-color:#aeb6b1}
.order-prompt-text{min-height:430px!important;font-size:11px!important;line-height:1.6!important}
.reading-row.is-stay:not(.is-completed) td{background:#fff0f0!important;border-color:#e6b6b6!important}
.reading-row.is-stay:not(.is-completed) td:first-child{box-shadow:inset 5px 0 0 #d94b4b!important}
.reading-row.is-stay:not(.is-completed) select[data-k="status"]{color:#b52222!important;font-weight:950!important}
.stay-detail-row.is-stay td{background:#fff7f7!important;border-color:#ecc5c5!important}
.stay-pending-note{margin-top:8px;padding:9px 11px;border:1px solid #e9bcbc;background:#fff0f0;color:#a62b2b;font-size:11px;font-weight:850;line-height:1.55;border-radius:10px}
.stay-lump-check{display:inline-flex;align-items:center;gap:7px;margin-top:8px;padding:8px 10px;border:1px solid #d9c7b2;background:#fffaf1;border-radius:10px;color:#65513b;font-size:11px;font-weight:850}
.stay-lump-check input{width:auto!important;accent-color:#6d7d70}
.stay-review-modal{position:fixed;inset:0;z-index:49000;display:none;place-items:center;padding:16px;background:rgba(36,20,20,.46);backdrop-filter:blur(6px)}
.stay-review-modal.open{display:grid}
.stay-review-dialog{width:min(480px,100%);background:#fff;border:1px solid #e4b3b3;border-radius:18px;padding:24px;box-shadow:0 24px 80px rgba(80,20,20,.24);text-align:center}
.stay-review-mark{width:42px;height:42px;border-radius:50%;display:grid;place-items:center;margin:0 auto 10px;background:#c93d3d;color:#fff;font-size:24px;font-weight:950}
.stay-review-dialog h3{margin:0;color:#8d2525;font-size:19px}.stay-review-question{margin:13px 0 7px;color:#632727;font-size:15px;font-weight:900;line-height:1.55}.stay-review-sub{margin:0;color:#806464;font-size:11px;line-height:1.65}
.stay-review-actions{display:flex;gap:8px;justify-content:center;margin-top:18px;flex-wrap:wrap}.stay-review-actions button{min-height:42px;padding:9px 14px;border-radius:10px;font-weight:900;cursor:pointer}.stay-review-back{border:1px solid #d8c7c7;background:#fff;color:#725858}.stay-review-ok{border:0;background:#ad3333;color:#fff}
@media(max-width:760px){.sheet-import-actions{width:100%;display:grid;grid-template-columns:minmax(0,1fr) auto}.sheet-import-btn{width:auto!important}.sheet-order-btn{font-size:9px;padding:10px 11px}.order-prompt-text{min-height:52vh!important;font-size:12px!important}.stay-review-dialog{padding:20px 16px}.stay-review-question{font-size:14px}}

</style>

<style id="seatLayoutV7">
/* v7 · 좌석 카드 겹침 방지: 테이블 번호는 카드 바깥 작은 라벨로 분리 */
.seat{overflow:visible!important;position:absolute!important}
.seat .seat-content{
  position:relative!important;width:100%!important;height:100%!important;min-height:0!important;
  padding:10px 8px 8px!important;display:flex!important;flex-direction:column!important;
  justify-content:center!important;align-items:center!important;gap:3px!important;
  overflow:visible!important;background:rgba(237,247,241,.93)!important;
}
.seat .seat-name{
  position:absolute!important;left:5px!important;top:-20px!important;z-index:6!important;
  width:auto!important;min-width:26px!important;height:18px!important;margin:0!important;padding:2px 6px!important;
  border:1px solid #b7c4bc!important;border-radius:6px!important;background:#fff!important;color:#46564d!important;
  font-size:10px!important;font-weight:950!important;line-height:12px!important;letter-spacing:.02em!important;
  text-align:center!important;white-space:nowrap!important;box-shadow:0 2px 5px rgba(30,55,40,.08)!important;
}
.seat .room-tag.seat-room-jump{font-size:15px!important;min-height:30px!important;padding:5px 11px!important;margin:0 0 2px!important}
.seat .seat-people{min-width:90px!important;padding:5px 8px!important;gap:0!important;margin:0!important}
.seat .seat-people small{font-size:9px!important;line-height:1.05!important}
.seat .seat-people b{font-size:17px!important;line-height:1.05!important}
.seat .seat-people em{font-size:9px!important;line-height:1.1!important;white-space:nowrap!important}
.seat .seat-menu-label{display:none!important}
.seat .seat-menu-text{
  display:block!important;width:100%!important;max-width:100%!important;margin:2px 0!important;padding:0 3px!important;
  color:#173b28!important;font-size:13px!important;font-weight:1000!important;line-height:1.25!important;
  text-align:center!important;white-space:normal!important;word-break:keep-all!important;overflow-wrap:anywhere!important;
  overflow:visible!important;
}
.seat .seat-salad{font-size:10.5px!important;padding:3px 6px!important;margin-top:1px!important;line-height:1.15!important}
.seat .seat-salad b{font-size:11px!important}
.seat .seat-memo-trigger{font-size:9px!important;min-height:26px!important;padding:4px 7px!important;margin-top:1px!important}
.seat .seat-guest{font-size:9px!important;line-height:1.2!important;text-align:center!important}
.seat .seat-dup{font-size:10px!important;line-height:1.1!important}
@media (hover:hover) and (pointer:fine){
 .seat.occupied:hover{transform:scale(1.13)!important;z-index:190!important}
}
.seat.expanded{transform:scale(1.18)!important;z-index:210!important}
@media(max-width:760px){
  .floorplan-shell{padding:28px 10px 14px!important}
  .floorplan{min-width:980px!important;min-height:760px!important}
  .seat .seat-name{top:-19px!important;font-size:10px!important}
  .seat .room-tag.seat-room-jump{font-size:15px!important}
  .seat .seat-menu-text{font-size:13px!important;line-height:1.24!important}
  .seat.expanded{transform:scale(1.25)!important}
}
</style>

</head>
<body>
<div class="wrap">
  <div class="topbar">
    <div></div>
    <div class="title">조식 LIST <span class="notice">연박 고객은 조식패키지 구분을 잘 확인하기.</span></div>
    <div class="datebox">
      <label>작성일</label>
      <input id="sheetDate" type="date">
    </div>
  </div>





  <div class="sheet-import-bar">
    <div class="sheet-import-copy">
      <b>📋 조식LIST 텍스트 입력</b>
      <small>정리된 조식LIST 텍스트를 붙여넣으면 객실별 인원·시간·메뉴·연박과 프런트·주방 전달사항을  입력함.</small>
    </div>
    <div class="sheet-import-actions">
      <button type="button" class="sheet-import-btn" id="openSheetTextBtn">📋 조식LIST 텍스트 넣기</button>
      <button type="button" class="sheet-order-btn" id="openOrderPromptBtn">ORDER TO JSON</button>
    </div>
  </div>

  <section class="layout-card">
    <div class="layout-head">
      <div>
        <h2>레스토랑 좌석 배치</h2>
        <div class="layout-help">조식 LIST에서 <b>자리(1~6 / R1~R5)</b>를 선택하면 해당 위치에 <b>객실 호수·인원·시간·메뉴</b>가  표시됨.</div>
      </div>
    </div>
    <div class="floorplan-shell">
      <div class="floorplan">
        <div class="plan-box seat r2" data-seat="R2"></div>
        <div class="plan-box seat r1" data-seat="R1"></div>
        <div class="plan-box seat r3" data-seat="R3"></div>
        <div class="plan-box seat r4" data-seat="R4"></div>
        <div class="plan-box seat r5" data-seat="R5"></div>
        <div class="plan-box fixed-label wc1">화장실</div>
        <div class="plan-box fixed-label wc2">화장실</div>

        <div class="plan-box seat s6" data-seat="6"></div>
        <div class="plan-box seat s4" data-seat="4"></div>
        <div class="plan-box seat s2" data-seat="2"></div>
        <div class="plan-box seat s5" data-seat="5"></div>
        <div class="plan-box seat s3" data-seat="3"></div>
        <div class="plan-box seat s1" data-seat="1"></div>

        <div class="plan-box counter"></div>
        <div class="plan-box fixed-label kitchen-box">주방</div>
        <div class="entrance">출입구</div>
        <div class="entrance-arrow">←</div>
        <div class="plan-box small-box"></div>
      </div>
    </div>
    <div class="plan-legend">
      <span><i class="legend-dot"></i>빈 자리</span>
      <span><i class="legend-dot used"></i>객실 배정</span>
      <span><i class="legend-dot dup"></i>중복 배정</span>
      <span>💬 배정된 좌석을 누르면 객실별 메모 작성</span>
    </div>

    <div class="room-memo-panel" id="roomMemoPanel">
      <div class="room-memo-head">
        <div>
          <b id="roomMemoTitle">객실 메모</b>
          <span id="roomMemoSub">좌석을 선택함.</span>
        </div>
        <button class="room-memo-close" type="button" id="roomMemoClose">닫기</button>
      </div>
      <textarea id="roomMemoText" placeholder="이 객실에 관한 메모를 자유롭게 입력함. 예: 유아의자 필요 / 알레르기 / 8:30 먼저 착석 / 메뉴 재확인"></textarea>
    </div>
  </section>

  <div class="focus-guide"><b>입력 집중 모드</b> · 한 줄을 누르거나 입력칸을 선택하면 그 객실 줄만 선명하게 보이고 다른 줄은 흐리게 표시됨.</div>

  <div class="table-wrap">
    <table id="breakfastTable">
      <colgroup>
        <col style="width:42px"><col style="width:72px"><col style="width:70px"><col style="width:140px">
        <col style="width:135px"><col style="width:205px"><col style="width:105px"><col style="width:80px">
        <col style="width:76px"><col style="width:76px"><col style="width:215px"><col style="width:76px">
        <col style="width:55px"><col style="width:70px">
      </colgroup>
      <thead>
        <tr>
          <th>순번</th>
          <th>호실</th>
          <th>자리</th>
          <th>연박 / 룸터치 / 딜리</th>
          <th>이름</th>
          <th>인원</th>
          <th>결제</th>
          <th>시간</th>
          <th class="menu-head">톳밥</th>
          <th class="menu-head">죽</th>
          <th class="menu-head">어린이용</th>
          <th class="menu-head">스프</th>
          <th>총</th>
          <th>완성</th>
        </tr>
      </thead>
      <tbody id="rows"></tbody>
      <tfoot>
        <tr class="total-row">
          <td colspan="8" style="text-align:right;padding-right:12px">총계</td>
          <td class="sumcell"><small>톳밥</small><br><b id="sumTot">0</b></td>
          <td class="sumcell"><small>죽</small><br><b id="sumPorridge">0</b></td>
          <td class="sumcell"><small>어린이용</small><br><b id="sumChild">0</b></td>
          <td class="sumcell"><small>스프</small><br><b id="sumSoup">0</b></td>
          <td class="grand" id="grandTotal">0</td><td></td>
        </tr>
      </tfoot>
    </table>
  </div>

  <div class="finalize-wrap">
    <button type="button" class="finalize-btn" id="finalizeAllBtn" disabled>
      모두 입력하였습니다
      <span class="finalize-sub" id="finalizeAllSub">작성한 객실의 [완성]을 모두 눌러야 활성화됨.</span>
    </button>
    <button type="button" class="final-edit-btn" id="editFinalizedBtn">전체저장 수정하기 · 빈 줄 블러 해제</button>
  </div>

  <div class="notes">
    <section class="note-section front">
      <div class="note-title">&lt;프런트&gt;</div>
      <div class="memo-toolbar" id="frontNums"></div>
      <div class="memo-toolbar" id="frontEmos"></div>
      <div class="memo-editor">
        <textarea id="frontText" placeholder="① 프런트 전달사항을 입력함."></textarea>
      </div>
      <div class="memo-list" id="frontList"></div>
    </section>

    <section class="note-section kitchen">
      <div class="note-title">&lt;주방&gt;</div>
      <div class="memo-toolbar" id="kitchenNums"></div>
      <div class="memo-toolbar" id="kitchenEmos"></div>
      <div class="memo-editor">
        <textarea id="kitchenText" placeholder="① 주방 전달사항을 입력함."></textarea>
      </div>
      <div class="memo-list" id="kitchenList"></div>
    </section>
  </div>

  <div class="actions">
    <button class="action archive" id="archiveSaveBtn">💾 조식 리딩 저장하기</button>
    <button class="action secondary" id="copyBtn">카톡용 요약 복사</button>
    <button class="action" onclick="window.print()">인쇄 / PDF</button>
    <button class="action excel" id="excelBtn">엑셀로 저장</button>
    <button class="action danger" id="resetBtn">전체 초기화</button>
  </div>
  <div class="save-state" id="saveState">●  저장됨</div>
  <div class="hint">작성 중인 내용은  저장됨. 최종 작성이 끝나면 <b>조식 리딩 저장하기</b>를 눌러 날짜별 기록으로 남김.</div>

  <section class="reading-archive">
    <div class="archive-head">
      <div>
        <h2>날짜별 조식 리딩 저장내용</h2>
        <p>저장한 날짜를 누르면 당시 객실별 조식·연박·메모·프런트/주방 전달사항을 다시 볼 수 있음.</p>
      </div>
      <span class="archive-count" id="archiveCount">0건</span>
    </div>
    <div class="archive-list" id="archiveList"></div>
  </section>
</div>

<div class="archive-modal" id="archiveModal" aria-hidden="true">
  <section class="archive-dialog" role="dialog" aria-modal="true" aria-labelledby="archiveViewTitle">
    <header class="archive-dialog-head">
      <div>
        <h3 id="archiveViewTitle">조식 리딩 저장내용</h3>
        <p id="archiveViewMeta"></p>
      </div>
      <button class="archive-close" id="archiveCloseBtn" type="button" aria-label="닫기">×</button>
    </header>
    <div class="archive-dialog-body">
      <div class="archive-summary" id="archiveSummary"></div>
    </div>
    <footer class="archive-dialog-actions">
      <button type="button" id="archiveModalCloseBtn">닫기</button>
      <button type="button" class="load" id="archiveLoadBtn">이 기록 다시 불러오기</button>
    </footer>
  </section>
</div>
<div class="archive-toast" id="archiveToast">저장됨.</div>

<div class="ocr-modal" id="sheetTextModal" aria-hidden="true">
  <section class="ocr-dialog" role="dialog" aria-modal="true" aria-labelledby="sheetTextTitle">
    <header class="ocr-head"><div><h3 id="sheetTextTitle">📋 조식LIST 텍스트 넣기</h3><p>정리된 <b>TOYU_BREAKFAST_V1 형식</b>의 텍스트를 그대로 붙여넣음. 적용 후 각 행을 확인하고 [완성]을 누름.</p></div><button type="button" class="ocr-close" id="sheetTextClose" aria-label="닫기">×</button></header>
    <div class="ocr-body">
      <div class="json-help">① 조식LIST 내용을 정리된 형식으로 준비함 → ② 아래 입력칸에 결과 전체를 붙여넣음 → ③ [조식차트에 적용]을 누름 → ④ 각 행을 확인 후 [완성]함.</div>
      <textarea class="json-paste" id="sheetTextJson" spellcheck="false" placeholder="여기에 정리된 조식LIST 텍스트를 붙여넣음."></textarea>
      <div class="json-status" id="sheetTextStatus">붙여넣은 뒤 적용함.</div>
      <details class="json-example"><summary>입력 형식 예시 보기</summary><pre>{
  "format": "TOYU_BREAKFAST_V1",
  "date": "2026-08-27",
  "rows": [{
    "room": "305", "name": "이상은", "status": "",
    "adults": 2, "kids": 0, "type": "패키지", "time": "8시",
    "tot": 2, "porridge": 0, "childPorridge": 0, "childRice": 0, "childSoup": 0, "soup": 0
  }]
}</pre></details>
      <div class="ocr-actions"><button type="button" class="ocr-secondary" id="sheetTextCancel">취소</button><button type="button" class="ocr-apply" id="sheetTextApply">조식차트에 적용</button></div>
    </div>
  </section>
</div>

<div class="ocr-modal" id="orderPromptModal" aria-hidden="true">
  <section class="ocr-dialog order-prompt-dialog" role="dialog" aria-modal="true" aria-labelledby="orderPromptTitle">
    <header class="ocr-head"><div><h3 id="orderPromptTitle">ORDER TO JSON</h3><p>Copy the instruction below and use it together with one breakfast LIST capture.</p></div><button type="button" class="ocr-close" id="orderPromptClose" aria-label="닫기">×</button></header>
    <div class="ocr-body">
      <textarea class="json-paste order-prompt-text" id="orderPromptText" readonly spellcheck="false"></textarea>
      <div class="json-status" id="orderPromptStatus">Ready to copy.</div>
      <div class="ocr-actions"><button type="button" class="ocr-secondary" id="orderPromptCancel">닫기</button><button type="button" class="ocr-apply" id="copyOrderPromptBtn">COPY ORDER</button></div>
    </div>
  </section>
</div>

<div class="stay-review-modal" id="stayReviewModal" aria-hidden="true">
  <section class="stay-review-dialog" role="dialog" aria-modal="true" aria-labelledby="stayReviewTitle">
    <div class="stay-review-mark">!</div>
    <h3 id="stayReviewTitle">연박 조식 확인</h3>
    <p class="stay-review-question">연박동안 다음날 조식 메뉴 제대로 골랐는지 검토하였습니까?</p>
    <p class="stay-review-sub">1박당 무료조식 2개 기준. 조식을 몰아사용하지 않는 경우 다음날 메뉴·시간을 다시 확인함.</p>
    <div class="stay-review-actions"><button type="button" class="stay-review-back" id="stayReviewBack">돌아가서 확인</button><button type="button" class="stay-review-ok" id="stayReviewOk">검토 완료 · 완성</button></div>
  </section>
</div>

<script>
const ROOMS=['101','102','201','202','203','204','205','301','302','303','304','305','401'];

const SEATS=['','1','2','3','4','5','6','R1','R2','R3','R4','R5'];
const TYPES=['','패키지','야놀자','여기어때','아고다','부킹닷컴','익스피디아','에어비앤비','전화예약','워크인','현장결제','기타'];
const TIMES=['','8시','8시30분','9시'];
const STATUS=['','연박','룸터치','딜리','연박/룸터치','연박/딜리'];
const EMOJIS=['없음','⚠️','❗','‼️','📌','⭐','🔔','✅','🔥','👀'];
const CIRCLED=['①','②','③','④','⑤','⑥','⑦','⑧','⑨','⑩'];
const KEY='breakfastList_v4';
const LEGACY_KEY='breakfastList_v3';
const ARCHIVE_KEY='breakfastReadingArchives_v1';

const emptyRow=()=>({
  room:'',seat:'',status:'',name:'',people:'',type:'',time:'',roomTouchStart:'',roomTouchEnd:'',completed:false,completedAt:'',
  peopleEnabled:false,peopleGroups:[],
  tot:0,porridge:0,
  childEnabled:false,childMenus:[],
  soup:0,
  roomMemo:'',stayNights:'',stayBreakfasts:[],stayBreakfastLumped:false
});
const init={
  date:new Date().toISOString().slice(0,10),
  rows:Array.from({length:13},emptyRow),
  memo:{front:{active:1,items:{}},kitchen:{active:1,items:{}}},
  finalLocked:false
};
let state=structuredClone(init);

function normalizePeopleGroups(groups){
  if(!Array.isArray(groups)) return [];
  return groups.map(g=>({type:(g&&g.type)||'성인',count:Number(g&&g.count||0)})).filter(g=>g.count>0 || g.type);
}
function normalizeChildMenus(items){
  if(!Array.isArray(items)) return [];
  return items.map(g=>({type:(g&&g.type)||'어죽',count:Number(g&&g.count||0)})).filter(g=>g.count>0 || g.type);
}
function normalizeRow(r){
  const base={...emptyRow(),...(r||{})};
  let peopleGroups=normalizePeopleGroups(base.peopleGroups);
  if(!peopleGroups.length && Number(base.people||0)>0) peopleGroups=[{type:'성인',count:Number(base.people||0)}];
  const legacyChild=[];
  if(Number(base.childPorridge||0)>0) legacyChild.push({type:'어죽',count:Number(base.childPorridge||0)});
  if(Number(base.childEtc||0)>0) legacyChild.push({type:(base.childEtcType||'어밥'),count:Number(base.childEtc||0)});
  let childMenus=normalizeChildMenus(base.childMenus);
  if(!childMenus.length && legacyChild.length) childMenus=legacyChild;
  return {
    ...emptyRow(),
    ...base,
    peopleEnabled:Boolean(base.peopleEnabled || peopleGroups.length),
    peopleGroups,
    childEnabled:Boolean(base.childEnabled || childMenus.length),
    childMenus,
    tot:Number(base.tot||0),
    porridge:Number(base.porridge||0),
    soup:Number(base.soup||0),
    stayBreakfasts:Array.isArray(base.stayBreakfasts)?base.stayBreakfasts:[],
    stayBreakfastLumped:Boolean(base.stayBreakfastLumped)
  };
}

function personTotal(r){ return normalizePeopleGroups(r?.peopleGroups).reduce((a,g)=>a+Number(g.count||0),0); }
function personDetailText(r){
  return normalizePeopleGroups(r?.peopleGroups).filter(g=>Number(g.count||0)>0).map(g=>`${g.type} ${g.count}`).join(' · ');
}
function peopleSummaryText(r){ const total=personTotal(r); return total?`${total}명`:'-'; }
function childMenuItems(r){ return normalizeChildMenus(r?.childMenus).filter(g=>Number(g.count||0)>0).map(g=>`${g.type}${g.count}`); }
function peopleBreakdownText(r){
  const groups=normalizePeopleGroups(r?.peopleGroups);
  let adults=0, kids=0;
  groups.forEach(g=>{
    const n=Number(g.count||0);
    if(String(g.type||'').includes('아이')) kids+=n; else adults+=n;
  });
  if(!adults && !kids) return '';
  return `(어른 ${adults} + 아이 ${kids})`;
}
function adultChildCounts(r){
  const groups=normalizePeopleGroups(r?.peopleGroups);
  let adults=0, kids=0;
  groups.forEach(g=>{
    const n=Number(g.count||0);
    if(String(g.type||'').includes('아이')) kids+=n; else adults+=n;
  });
  return {adults,kids,total:adults+kids};
}

function childTotal(r){ return normalizeChildMenus(r?.childMenus).reduce((a,g)=>a+Number(g.count||0),0); }
function mealItems(r){
  const a=[];
  if(Number(r.tot)) a.push(`톳밥${r.tot}`);
  if(Number(r.porridge)) a.push(`죽${r.porridge}`);
  a.push(...childMenuItems(r));
  if(Number(r.soup)) a.push(`스프${r.soup}`);
  return a;
}
function mealText(r){ return mealItems(r).join(' · '); }

function load(){
  try{
    let raw=localStorage.getItem(KEY);
    if(!raw) raw=localStorage.getItem(LEGACY_KEY);
    const x=raw?JSON.parse(raw):null;
    if(x) state={...structuredClone(init),...x};
    if(!Array.isArray(state.rows)) state.rows=Array.from({length:13},emptyRow);
    while(state.rows.length<13) state.rows.push(emptyRow());
    state.rows=state.rows.slice(0,13).map(normalizeRow);
  }catch(e){}
}
function save(){
  localStorage.setItem(KEY,JSON.stringify(state));
  const s=document.getElementById('saveState');
  s.textContent='● 저장 중...';
  clearTimeout(window.sTimer);
  window.sTimer=setTimeout(()=>s.textContent='●  저장됨',300);
}
function opts(list,val){
  return list.map(x=>`<option value="${x}" ${x===val?'selected':''}>${x||'선택'}</option>`).join('');
}
function qtyOpts(val,max=10,start=0){
  let s='';
  for(let i=start;i<=max;i++) s+=`<option value="${i}" ${Number(val)===i?'selected':''}>${i}</option>`;
  return s;
}
function escAttr(s){return String(s||'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;')}
function safeHtml(v){return String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]))}

function peopleEditorHtml(r,i){
  const enabled=Boolean(r.peopleEnabled || personTotal(r));
  const groups=normalizePeopleGroups(r.peopleGroups);
  const effective=groups.length?groups:[{type:'성인',count:1}];
  return `
    <div class="people-cell-wrap">
      <div class="people-summary ${personTotal(r)?'filled':''}"><small>인원</small><b>${peopleSummaryText(r)}</b></div>
      ${personDetailText(r)?`<div class="inline-summary">${safeHtml(personDetailText(r))}</div>`:''}
      <label class="inline-check"><input type="checkbox" data-person-enabled data-i="${i}" ${enabled?'checked':''}> 인원</label>
      <div class="stack-editor ${enabled?'open':''}">
        ${effective.map((g,gi)=>`
          <div class="editor-row ${effective.length===1?'single':''}">
            <select data-person-type data-i="${i}" data-gi="${gi}">
              <option value="성인" ${g.type==='성인'?'selected':''}>성인</option>
              <option value="아이" ${g.type==='아이'?'selected':''}>아이</option>
            </select>
            <select data-person-count data-i="${i}" data-gi="${gi}">${qtyOpts(g.count,10,1)}</select>
            ${effective.length>1?`<button type="button" class="row-remove" data-remove-person data-i="${i}" data-gi="${gi}">－</button>`:''}
          </div>`).join('')}
        <button type="button" class="add-mini" data-add-person data-i="${i}">＋ 인원 추가</button>
      </div>
    </div>`;
}

function childEditorHtml(r,i){
  const enabled=Boolean(r.childEnabled || childTotal(r));
  const menus=normalizeChildMenus(r.childMenus);
  const effective=menus.length?menus:[{type:'어죽',count:1}];
  return `
    <div class="child-cell-wrap">
      ${childMenuItems(r).length?`<div class="child-pill-row">${childMenuItems(r).map(x=>`<span class="child-pill">${safeHtml(x)}</span>`).join('')}</div>`:`<div class="child-placeholder">어린이용 미선택</div>`}
      <label class="inline-check"><input type="checkbox" data-child-enabled data-i="${i}" ${enabled?'checked':''}> 어린이용</label>
      <div class="stack-editor ${enabled?'open':''}">
        ${effective.map((g,gi)=>`
          <div class="editor-row ${effective.length===1?'single':''}">
            <select data-child-type data-i="${i}" data-gi="${gi}">
              <option value="어죽" ${g.type==='어죽'?'selected':''}>어죽</option>
              <option value="어밥" ${g.type==='어밥'?'selected':''}>어밥</option>
              <option value="어미" ${g.type==='어미'?'selected':''}>어미</option>
            </select>
            <select data-child-count data-i="${i}" data-gi="${gi}">${qtyOpts(g.count,10,1)}</select>
            ${effective.length>1?`<button type="button" class="row-remove" data-remove-child data-i="${i}" data-gi="${gi}">－</button>`:''}
          </div>`).join('')}
        <button type="button" class="add-mini" data-add-child data-i="${i}">＋ 어린이용 추가</button>
      </div>
    </div>`;
}

function compactCompletedHtml(r,i){
  const counts=adultChildCounts(r);
  const childText=childMenuItems(r).join(' · ') || '-';
  const total=Number(r.tot||0)+Number(r.porridge||0)+childTotal(r)+Number(r.soup||0);
  const status=String(r.status||'').trim()||'-';
  const people=counts.total ? `<b>${counts.total}명</b><small>(어른 ${counts.adults} + 아이 ${counts.kids})</small>` : '-';
  return `
    <td class="done-cell done-no">${i+1}</td>
    <td class="done-cell done-room">${r.room?safeHtml(r.room)+'호':'-'}</td>
    <td class="done-cell">${r.seat?safeHtml(r.seat):'-'}</td>
    <td class="done-cell ${status.includes('연박')?'done-stay':''}">${safeHtml(status)}</td>
    <td class="done-cell">${r.name?safeHtml(r.name):'-'}</td>
    <td class="done-cell done-people">${people}</td>
    <td class="done-cell">${r.type?safeHtml(r.type):'-'}</td>
    <td class="done-cell done-time">${r.time?safeHtml(r.time):'-'}</td>
    <td class="done-cell done-menu">${Number(r.tot||0)||'-'}</td>
    <td class="done-cell done-menu">${Number(r.porridge||0)||'-'}</td>
    <td class="done-cell done-menu">${safeHtml(childText)}</td>
    <td class="done-cell done-menu">${Number(r.soup||0)||'-'}</td>
    <td class="done-cell done-total">${total}</td>
    <td class="done-cell"><div class="row-action-wrap"><button type="button" class="complete-btn edit-mode" data-edit-row="${i}">수정</button><button type="button" class="delete-row-btn" data-delete-row="${i}">삭제</button></div></td>`;
}

function rowHasBreakfastInput(r){
  if(!r) return false;
  return !!(
    String(r.room||'').trim() ||
    String(r.name||'').trim() ||
    String(r.seat||'').trim() ||
    String(r.status||'').trim() ||
    personTotal(r) ||
    String(r.type||'').trim() ||
    String(r.time||'').trim() ||
    mealItems(r).length ||
    String(r.roomMemo||'').trim()
  );
}
function updateFinalizeAllButton(){
  const btn=document.getElementById('finalizeAllBtn');
  const editBtn=document.getElementById('editFinalizedBtn');
  const table=document.getElementById('breakfastTable');
  if(!btn) return;
  const used=state.rows.slice(0,13).filter(rowHasBreakfastInput);
  const incomplete=used.filter(r=>!r.completed);
  const ready=used.length>0 && incomplete.length===0;
  table?.classList.toggle('final-locked',Boolean(state.finalLocked));
  editBtn?.classList.toggle('show',Boolean(state.finalLocked));
  if(state.finalLocked){
    btn.disabled=true;
    btn.classList.add('is-saved');
    btn.innerHTML=`✓ 모두 입력 완료 · 저장됨<span class="finalize-sub">작성하지 않은 빈 줄은 블러 처리됨. 수정하려면 아래 버튼을 누름.</span>`;
    return;
  }
  btn.disabled=!ready;
  btn.classList.remove('is-saved');
  btn.innerHTML=ready
    ? `모두 입력하였습니다<span class="finalize-sub" id="finalizeAllSub">${used.length}개 객실 입력 완료 · 누르면 최종 저장함.</span>`
    : `모두 입력하였습니다<span class="finalize-sub" id="finalizeAllSub">${used.length?`미완성 ${incomplete.length}개 · 각 줄의 [완성]을 먼저 누름.`:'작성된 객실이 없음.'}</span>`;
}
function finalizeAllReading(){
  const btn=document.getElementById('finalizeAllBtn');
  const used=state.rows.slice(0,13).filter(rowHasBreakfastInput);
  const incomplete=used.filter(r=>!r.completed);
  if(!used.length || incomplete.length) return;
  state.finalLocked=true;
  save();
  saveReadingArchive();
  renderRows();
  if(btn){
    btn.classList.add('is-saved');
    btn.scrollIntoView({behavior:'smooth',block:'center'});
    setTimeout(()=>btn.classList.remove('is-saved'),1300);
  }
  showArchiveToast('모두 입력 완료 · 저장되었습니다.');
}
function editFinalizedReading(){
  state.finalLocked=false;
  save();
  renderRows();
  document.getElementById('breakfastTable')?.scrollIntoView({behavior:'smooth',block:'start'});
  showArchiveToast('전체저장 수정 모드 · 빈 줄 블러가 해제되었습니다.');
}


let selectedCompletedRow=null;
function closeCompletedDetail(){
  document.querySelectorAll('.completed-detail-row').forEach(x=>x.remove());
  document.querySelectorAll('.reading-row.row-selected').forEach(x=>x.classList.remove('row-selected'));
  selectedCompletedRow=null;
}
function openCompletedDetail(i,tr){
  if(!state.rows[i] || !state.rows[i].completed) return;
  if(selectedCompletedRow===i && document.querySelector('.completed-detail-row')){closeCompletedDetail();return;}
  closeCompletedDetail();
  selectedCompletedRow=i;
  tr.classList.add('row-selected');
  const r=state.rows[i];
  const c=adultChildCounts(r);
  const meals=mealItems(r);
  const menu=meals.length?safeHtml(meals.join(' · ')):'메뉴 미입력';
  const detail=document.createElement('tr');
  detail.className='completed-detail-row';
  detail.dataset.detailRow=String(i);
  detail.innerHTML=`<td colspan="14"><div class="completed-detail-card ${String(r.status||'').includes('연박')?'is-stay':''}">
    <div class="completed-detail-seat">${safeHtml(r.seat||'-')}</div>
    <div class="completed-detail-room">${r.room?safeHtml(r.room)+'호':'호실 미선택'}</div>
    <div class="completed-detail-people"><small>총인원</small><b>${c.total}명</b><span>(어른 ${c.adults} + 아이 ${c.kids})</span></div>
    <div class="completed-detail-label">메뉴</div><div class="completed-detail-menu">${menu}</div>
    <div class="completed-detail-salad">🥗 샐러드 ${c.adults}개</div>
    <div class="completed-detail-meta">${safeHtml(r.time||'시간 미선택')}${r.status?' · '+safeHtml(r.status):''}${r.name?' · '+safeHtml(r.name):''}</div>
    <div class="completed-detail-actions"><button type="button" data-detail-memo="${i}">💬 메모</button><button type="button" data-detail-close>닫기</button></div>
  </div></td>`;
  tr.insertAdjacentElement('afterend',detail);
  detail.querySelector('[data-detail-memo]')?.addEventListener('click',()=>openRoomMemo(i));
  detail.querySelector('[data-detail-close]')?.addEventListener('click',closeCompletedDetail);
  requestAnimationFrame(()=>detail.scrollIntoView({behavior:'smooth',block:'nearest'}));
}



function deleteReadingRow(i){
  if(!Number.isInteger(i)||!state.rows[i]) return;
  const r=state.rows[i];
  const label=String(r.room||'').trim()?`${r.room}호`: `${i+1}번째 줄`;
  if(rowHasBreakfastInput(r) && !confirm(`${label} 조식 내용을 삭제할까요?\n이 줄에 입력한 내용만 초기화됨.`)) return;
  state.rows[i]=normalizeRow(emptyRow());
  state.finalLocked=false;
  selectedMemoRow=null;
  closeCompletedDetail();
  save();
  renderRows();
  showArchiveToast(`${label} 삭제 완료.`);
}

function renderRows(){
  closeCompletedDetail();
  const body=document.getElementById('rows');
  body.innerHTML='';
  state.rows.slice(0,13).forEach((r,i)=>{
    const isStay=String(r.status||'').includes('연박');
    const tr=document.createElement('tr');
    const isUnused=Boolean(state.finalLocked && !rowHasBreakfastInput(r));
    tr.className='reading-row'+(r.completed?' is-completed':'')+(isStay?' is-stay':'')+(isUnused?' final-unused':'');
    tr.dataset.rowIndex=String(i);

    if(r.completed){
      tr.innerHTML=compactCompletedHtml(r,i);
      body.appendChild(tr);
      return;
    }

    tr.innerHTML=`
      <td><b>${i+1}</b></td>
      <td><select class="room-select" data-i="${i}" data-k="room">${opts(['',...ROOMS],r.room)}</select></td>
      <td><select class="seat-select" data-i="${i}" data-k="seat">${opts(SEATS,r.seat)}</select></td>
      <td><select data-i="${i}" data-k="status">${opts(STATUS,r.status)}</select></td>
      <td><input class="name-input" data-i="${i}" data-k="name" value="${escAttr(r.name)}" placeholder="예약자명"></td>
      <td>${peopleEditorHtml(r,i)}</td>
      <td><select class="type-select" data-i="${i}" data-k="type">${opts(TYPES,r.type)}</select></td>
      <td><select class="time-select" data-i="${i}" data-k="time">${opts(TIMES,r.time)}</select></td>
      <td class="menu-stack-cell"><div class="menu-stack-title">톳밥</div><select class="qty" data-i="${i}" data-k="tot">${qtyOpts(r.tot)}</select></td>
      <td class="menu-stack-cell"><div class="menu-stack-title">죽</div><select class="qty" data-i="${i}" data-k="porridge">${qtyOpts(r.porridge)}</select></td>
      <td>${childEditorHtml(r,i)}</td>
      <td class="menu-stack-cell"><div class="menu-stack-title">스프</div><select class="qty" data-i="${i}" data-k="soup">${qtyOpts(r.soup)}</select></td>
      <td class="row-total" id="rt${i}">0</td><td><div class="row-action-wrap"><button type="button" class="complete-btn" data-complete-row="${i}">완성</button><button type="button" class="delete-row-btn" data-delete-row="${i}">삭제</button></div></td>`;
    body.appendChild(tr);

    const hasStay=isStay;
    const hasRoomTouch=String(r.status||'').includes('룸터치');
    if(hasStay || hasRoomTouch){
      const detail=document.createElement('tr');
      detail.className='stay-detail-row'+(hasStay?' is-stay':'');
      detail.dataset.rowIndex=String(i);
      let detailHtml='<div style="display:grid;gap:10px">';
      if(hasRoomTouch){
        detailHtml+=`<div class="roomtouch-box"><span class="roomtouch-label">고객 외출</span><input type="time" step="60" data-roomtouch-start data-i="${i}" value="${escAttr(r.roomTouchStart||'')}"><span>~</span><input type="time" step="60" data-roomtouch-end data-i="${i}" value="${escAttr(r.roomTouchEnd||'')}"><span>까지 · 룸터치 가능</span></div>`;
      }
      if(hasStay){
        const nights=Math.max(1,Math.min(10,Number(r.stayNights)||1));
        let breakfastFields='';
        for(let n=0;n<nights;n++){
          breakfastFields+=`<div class="stay-breakfast-item"><label>${n+1}박 후 다음 아침 조식메뉴</label><textarea data-stay-breakfast data-i="${i}" data-night="${n}" placeholder="예: 성인 톳밥 2 / 아이 어죽 1 / 8시 30분">${safeHtml((r.stayBreakfasts||[])[n]||'')}</textarea></div>`;
        }
        detailHtml+=`<div class="stay-editor"><div class="stay-count-box"><label>연박 박수</label><select data-stay-nights data-i="${i}">${Array.from({length:10},(_,x)=>`<option value="${x+1}" ${nights===x+1?'selected':''}>${x+1}박</option>`).join('')}</select></div><label class="stay-lump-check"><input type="checkbox" data-stay-lumped data-i="${i}" ${r.stayBreakfastLumped?'checked':''}> 조식을 한 번에 몰아사용함</label><div class="stay-breakfast-grid">${breakfastFields}</div><div class="stay-pending-note">연박 미완성 · 1박당 무료조식 2개 기준. 몰아사용하지 않는 경우 다음날 조식 메뉴·시간을 확인한 뒤 [완성]함.</div><div class="stay-help">연박 선택 시 박수만큼 다음 날 아침 조식 내용을 기록함. 메뉴·시간·인원·특이사항을 자유롭게 적을 수 있음.</div></div>`;
      }
      detailHtml+='</div>';
      detail.innerHTML=`<td colspan="14">${detailHtml}</td>`;
      body.appendChild(detail);
    }
  });

  body.querySelectorAll('input[data-k],select[data-k]').forEach(el=>{
    el.addEventListener('input',e=>{
      const i=Number(e.target.dataset.i), k=e.target.dataset.k;
      let v=e.target.value;
      if(['tot','porridge','soup'].includes(k)) v=Number(v||0);
      state.rows[i][k]=v;
      if(k==='time'){
        const movedRow=state.rows[i];
        sortRowsByBreakfastTime();
        save();
        renderRows();
        const ni=state.rows.indexOf(movedRow);
        const moved=document.querySelector(`#rows tr.reading-row[data-row-index="${ni}"]`);
        if(moved){moved.classList.add('time-burst');moved.scrollIntoView({behavior:'smooth',block:'center'});setTimeout(()=>moved.classList.remove('time-burst'),800)}
        return;
      }
      save();calc();renderFloorplan();
      if(k==='status') renderRows();
    });
  });
  body.querySelectorAll('[data-person-enabled]').forEach(el=>el.addEventListener('change',e=>{
    const i=Number(e.target.dataset.i), checked=e.target.checked;
    state.rows[i].peopleEnabled=checked;
    if(checked && !normalizePeopleGroups(state.rows[i].peopleGroups).length) state.rows[i].peopleGroups=[{type:'성인',count:1}];
    if(!checked) state.rows[i].peopleGroups=[];
    save();renderRows();
  }));
  body.querySelectorAll('[data-add-person]').forEach(btn=>btn.addEventListener('click',()=>{
    const i=Number(btn.dataset.i); state.rows[i].peopleEnabled=true;
    state.rows[i].peopleGroups=normalizePeopleGroups(state.rows[i].peopleGroups); state.rows[i].peopleGroups.push({type:'성인',count:1}); save();renderRows();
  }));
  body.querySelectorAll('[data-remove-person]').forEach(btn=>btn.addEventListener('click',()=>{
    const i=Number(btn.dataset.i),gi=Number(btn.dataset.gi); state.rows[i].peopleGroups=normalizePeopleGroups(state.rows[i].peopleGroups); state.rows[i].peopleGroups.splice(gi,1); if(!state.rows[i].peopleGroups.length) state.rows[i].peopleEnabled=false; save();renderRows();
  }));
  body.querySelectorAll('[data-person-type],[data-person-count]').forEach(el=>el.addEventListener('change',e=>{
    const i=Number(e.target.dataset.i),gi=Number(e.target.dataset.gi); state.rows[i].peopleGroups=normalizePeopleGroups(state.rows[i].peopleGroups); const item=state.rows[i].peopleGroups[gi]||{type:'성인',count:1}; if(e.target.hasAttribute('data-person-type')) item.type=e.target.value; if(e.target.hasAttribute('data-person-count')) item.count=Number(e.target.value||0); state.rows[i].peopleGroups[gi]=item; state.rows[i].people=String(personTotal(state.rows[i])||''); save();renderRows();
  }));
  body.querySelectorAll('[data-child-enabled]').forEach(el=>el.addEventListener('change',e=>{
    const i=Number(e.target.dataset.i),checked=e.target.checked; state.rows[i].childEnabled=checked; if(checked&&!normalizeChildMenus(state.rows[i].childMenus).length) state.rows[i].childMenus=[{type:'어죽',count:1}]; if(!checked) state.rows[i].childMenus=[]; save();renderRows();
  }));
  body.querySelectorAll('[data-add-child]').forEach(btn=>btn.addEventListener('click',()=>{
    const i=Number(btn.dataset.i); state.rows[i].childEnabled=true; state.rows[i].childMenus=normalizeChildMenus(state.rows[i].childMenus); state.rows[i].childMenus.push({type:'어죽',count:1}); save();renderRows();
  }));
  body.querySelectorAll('[data-remove-child]').forEach(btn=>btn.addEventListener('click',()=>{
    const i=Number(btn.dataset.i),gi=Number(btn.dataset.gi); state.rows[i].childMenus=normalizeChildMenus(state.rows[i].childMenus); state.rows[i].childMenus.splice(gi,1); if(!state.rows[i].childMenus.length) state.rows[i].childEnabled=false; save();renderRows();
  }));
  body.querySelectorAll('[data-child-type],[data-child-count]').forEach(el=>el.addEventListener('change',e=>{
    const i=Number(e.target.dataset.i),gi=Number(e.target.dataset.gi); state.rows[i].childMenus=normalizeChildMenus(state.rows[i].childMenus); const item=state.rows[i].childMenus[gi]||{type:'어죽',count:1}; if(e.target.hasAttribute('data-child-type')) item.type=e.target.value; if(e.target.hasAttribute('data-child-count')) item.count=Number(e.target.value||0); state.rows[i].childMenus[gi]=item; save();renderRows();
  }));
  body.querySelectorAll('[data-stay-nights]').forEach(el=>el.addEventListener('change',e=>{
    const i=Number(e.target.dataset.i); const nights=Math.max(1,Math.min(10,Number(e.target.value)||1)); state.rows[i].stayNights=nights; if(!Array.isArray(state.rows[i].stayBreakfasts)) state.rows[i].stayBreakfasts=[]; state.rows[i].stayBreakfasts=state.rows[i].stayBreakfasts.slice(0,nights); while(state.rows[i].stayBreakfasts.length<nights) state.rows[i].stayBreakfasts.push(''); save();renderRows();
  }));
  body.querySelectorAll('[data-stay-breakfast]').forEach(el=>el.addEventListener('input',e=>{const i=Number(e.target.dataset.i),n=Number(e.target.dataset.night); if(!Array.isArray(state.rows[i].stayBreakfasts)) state.rows[i].stayBreakfasts=[]; state.rows[i].stayBreakfasts[n]=e.target.value; save();}));
  body.querySelectorAll('[data-stay-lumped]').forEach(el=>el.addEventListener('change',e=>{const i=Number(e.target.dataset.i);state.rows[i].stayBreakfastLumped=Boolean(e.target.checked);save();renderRows();}));
  body.querySelectorAll('[data-roomtouch-start],[data-roomtouch-end]').forEach(el=>el.addEventListener('input',e=>{const i=Number(e.target.dataset.i); if(e.target.hasAttribute('data-roomtouch-start')) state.rows[i].roomTouchStart=e.target.value; else state.rows[i].roomTouchEnd=e.target.value; save();}));
  body.querySelectorAll('[data-complete-row]').forEach(btn=>btn.addEventListener('click',()=>completeReadingRow(Number(btn.dataset.completeRow))));
  body.querySelectorAll('[data-edit-row]').forEach(btn=>btn.addEventListener('click',()=>editReadingRow(Number(btn.dataset.editRow))));
  body.querySelectorAll('[data-delete-row]').forEach(btn=>btn.addEventListener('click',e=>{e.stopPropagation();deleteReadingRow(Number(btn.dataset.deleteRow))}));
  calc();renderFloorplan();
  updateFinalizeAllButton();
}

function breakfastTimeRank(t){
  const order={'8시':0,'8시30분':1,'9시':2};
  return Object.prototype.hasOwnProperty.call(order,t)?order[t]:9;
}
function sortRowsByBreakfastTime(){
  const decorated=state.rows.slice(0,13).map((r,idx)=>({r,idx}));
  decorated.sort((a,b)=>{
    const ta=breakfastTimeRank(a.r.time),tb=breakfastTimeRank(b.r.time);
    if(ta!==tb) return ta-tb;
    if(a.r.completed!==b.r.completed) return a.r.completed?-1:1;
    if(a.r.completed&&b.r.completed){
      const c=String(a.r.completedAt||'').localeCompare(String(b.r.completedAt||''));
      if(c) return c;
    }
    return a.idx-b.idx;
  });
  state.rows=decorated.map(x=>x.r);
}
let pendingStayCompleteIndex=null;
function finishReadingRow(i){
  if(!Number.isInteger(i)||!state.rows[i]) return;
  const target=state.rows[i];
  target.completed=true; target.completedAt=new Date().toISOString();
  sortRowsByBreakfastTime(); save(); renderRows();
  const newIndex=state.rows.indexOf(target);
  const tr=document.querySelector(`#rows tr.reading-row[data-row-index="${newIndex}"]`);
  if(tr){tr.classList.add('sparkle-complete');tr.scrollIntoView({behavior:'smooth',block:'center'});setTimeout(()=>tr.classList.remove('sparkle-complete'),1200)}
}
function openStayReview(i){pendingStayCompleteIndex=i;const m=document.getElementById('stayReviewModal');m?.classList.add('open');m?.setAttribute('aria-hidden','false')}
function closeStayReview(){pendingStayCompleteIndex=null;const m=document.getElementById('stayReviewModal');m?.classList.remove('open');m?.setAttribute('aria-hidden','true')}
function completeReadingRow(i){
  if(!Number.isInteger(i)||!state.rows[i]) return;
  const target=state.rows[i];
  const isStay=String(target.status||'').includes('연박');
  if(isStay && !target.stayBreakfastLumped){openStayReview(i);return;}
  finishReadingRow(i);
}
function editReadingRow(i){
  if(state.finalLocked) return;
  if(!Number.isInteger(i)||!state.rows[i]) return;
  const target=state.rows[i]; target.completed=false; target.completedAt=''; save(); renderRows();
  const ni=state.rows.indexOf(target);
  const tr=document.querySelector(`#rows tr.reading-row[data-row-index="${ni}"]`);
  if(tr){tr.scrollIntoView({behavior:'smooth',block:'center'});setTimeout(()=>tr.querySelector('select,input,button')?.focus({preventScroll:true}),220)}
}

let activeReadingRow=null;
function setActiveReadingRow(i){
  if(!Number.isInteger(i) || i<0 || i>=13) return;
  activeReadingRow=i;
  const table=document.getElementById('breakfastTable');
  table.classList.add('focus-mode');
  table.querySelectorAll('tbody tr[data-row-index]').forEach(tr=>{
    tr.classList.toggle('row-focused',Number(tr.dataset.rowIndex)===i);
  });
}
function restoreReadingFocus(){
  if(activeReadingRow!==null) setActiveReadingRow(activeReadingRow);
}
document.getElementById('rows').addEventListener('focusin',e=>{
  const tr=e.target.closest('tr[data-row-index]');
  if(tr) setActiveReadingRow(Number(tr.dataset.rowIndex));
});
document.getElementById('rows').addEventListener('click',e=>{
  const tr=e.target.closest('tr[data-row-index]');
  if(tr) setActiveReadingRow(Number(tr.dataset.rowIndex));
});
let selectedMemoRow=null;

function jumpToReadingRow(rowIndex){
  if(!Number.isInteger(rowIndex) || !state.rows[rowIndex]) return;
  setActiveReadingRow(rowIndex);
  const tr=document.querySelector(`#rows tr.reading-row[data-row-index="${rowIndex}"]`);
  if(!tr) return;
  tr.scrollIntoView({behavior:'smooth',block:'center',inline:'nearest'});
  tr.classList.add('seat-jump-highlight');
  setTimeout(()=>tr.classList.remove('seat-jump-highlight'),1600);
}

function openRoomMemo(rowIndex){
  if(!Number.isInteger(rowIndex) || !state.rows[rowIndex]) return;
  selectedMemoRow=rowIndex;
  const r=state.rows[rowIndex];
  const panel=document.getElementById('roomMemoPanel');
  const title=document.getElementById('roomMemoTitle');
  const sub=document.getElementById('roomMemoSub');
  const tx=document.getElementById('roomMemoText');
  title.textContent=`${r.room||'호실 미선택'}호 객실 메모`;
  sub.textContent=`좌석 ${r.seat||'-'} · ${r.name||'예약자명 미입력'}`;
  tx.value=r.roomMemo||'';
  panel.classList.add('active');
  setTimeout(()=>tx.focus({preventScroll:true}),30);
  panel.scrollIntoView({behavior:'smooth',block:'nearest'});
}


function renderFloorplan(){
  const grouped={};
  SEATS.filter(Boolean).forEach(x=>grouped[x]=[]);
  state.rows.slice(0,13).forEach((r,i)=>{if(r.seat&&grouped[r.seat])grouped[r.seat].push({...r,rowIndex:i})});
  document.querySelectorAll('.seat[data-seat]').forEach(box=>{
    const seat=box.dataset.seat, list=grouped[seat]||[];
    box.classList.toggle('occupied',list.length===1);
    box.classList.toggle('duplicate',list.length>1);
    box.setAttribute('tabindex',list.length?'0':'-1');
    box.setAttribute('role',list.length?'group':'group');
    box.onclick=null; box.onkeydown=null; box.removeAttribute('title');

    if(!list.length){
      box.innerHTML=`<div class="seat-content"><div class="seat-name">${seat}</div></div>`;
      box.classList.remove('expanded');
    }else if(list.length===1){
      const r=list[0], room=r.room?`${safeHtml(r.room)}호`:'호실 미선택';
      const meals=mealItems(r);
      const menuLine=meals.length?safeHtml(meals.join(' · ')):'메뉴 미입력';
      const counts=adultChildCounts(r);
      const tooltip=[room,counts.total?`총 ${counts.total}명 (어른 ${counts.adults}+아이 ${counts.kids})`:'',menuLine,`샐러드 ${counts.adults}개`,r.roomMemo?`메모: ${r.roomMemo}`:''].filter(Boolean).join(' · ');
      box.setAttribute('title',tooltip);
      box.innerHTML=`<div class="seat-content">
        <div class="seat-name">${seat}</div>
        <button type="button" class="room-tag seat-room-jump" data-jump-row="${r.rowIndex}">${room}</button>
        ${counts.total?`<div class="seat-people"><small>총인원</small><b>${counts.total}명</b><em>(어른 ${counts.adults} + 아이 ${counts.kids})</em></div>`:''}
        <div class="seat-menu-label">메뉴</div><div class="seat-menu-text">${menuLine}</div>
        <div class="seat-salad">🥗 샐러드 <b>${counts.adults}개</b></div>
        <button type="button" class="seat-memo-trigger" data-room-row="${r.rowIndex}">💬 메모</button>
      </div>`;
      box.querySelector('[data-jump-row]')?.addEventListener('click',e=>{e.stopPropagation();jumpToReadingRow(r.rowIndex)});
      box.querySelector('[data-room-row]')?.addEventListener('click',e=>{e.stopPropagation();openRoomMemo(r.rowIndex)});
      box.onclick=(e)=>{
        if(e.target.closest('button')) return;
        document.querySelectorAll('.seat.expanded').forEach(x=>{if(x!==box)x.classList.remove('expanded')});
        box.classList.toggle('expanded');
      };
      box.onkeydown=e=>{if((e.key==='Enter'||e.key===' ')&&!e.target.closest('button')){e.preventDefault();box.classList.toggle('expanded')}};
    }else{
      box.innerHTML=`<div class="seat-content"><div class="seat-name">${seat}</div><div class="seat-dup">⚠ 중복 배정</div>${list.map(r=>{const c=adultChildCounts(r);return `<button type="button" class="room-tag seat-room-jump" data-jump-row="${r.rowIndex}">${safeHtml(r.room||'?')}호</button><div class="seat-guest">총 ${c.total}명 (어른 ${c.adults}+아이 ${c.kids})</div><div class="seat-menu-text">${mealItems(r).length?safeHtml(mealItems(r).join(' · ')):'메뉴 미입력'}</div><div class="seat-salad">🥗 샐러드 <b>${c.adults}개</b></div><button type="button" class="seat-memo-trigger" data-room-row="${r.rowIndex}">💬 메모</button>`}).join('')}</div>`;
      box.querySelectorAll('[data-jump-row]').forEach(btn=>btn.addEventListener('click',e=>{e.stopPropagation();jumpToReadingRow(Number(btn.dataset.jumpRow))}));
      box.querySelectorAll('[data-room-row]').forEach(btn=>btn.addEventListener('click',e=>{e.stopPropagation();openRoomMemo(Number(btn.dataset.roomRow))}));
      box.onclick=(e)=>{if(e.target.closest('button'))return;box.classList.toggle('expanded')};
    }
  });

  if(selectedMemoRow!==null && state.rows[selectedMemoRow]){
    const r=state.rows[selectedMemoRow];
    const title=document.getElementById('roomMemoTitle');
    const sub=document.getElementById('roomMemoSub');
    if(title) title.textContent=`${r.room||'호실 미선택'}호 객실 메모`;
    if(sub) sub.textContent=`좌석 ${r.seat||'-'} · ${r.name||'예약자명 미입력'}`;
  }
}
function calc(){
  let sums={tot:0,porridge:0,child:0,soup:0,grand:0};
  state.rows.slice(0,13).forEach((r,i)=>{
    const child=childTotal(r);
    const total=Number(r.tot||0)+Number(r.porridge||0)+child+Number(r.soup||0);
    const c=document.getElementById('rt'+i); if(c)c.textContent=total;
    sums.tot+=Number(r.tot||0);sums.porridge+=Number(r.porridge||0);sums.child+=child;
    sums.soup+=Number(r.soup||0);sums.grand+=total;
    state.rows[i].people=String(personTotal(r)||'');
  });
  document.getElementById('sumTot').textContent=sums.tot;
  document.getElementById('sumPorridge').textContent=sums.porridge;
  document.getElementById('sumChild').textContent=sums.child;
  document.getElementById('sumSoup').textContent=sums.soup;
  document.getElementById('grandTotal').textContent=sums.grand;
}
function memoItem(dept,n){
  const k=String(n);
  if(!state.memo[dept].items[k]) state.memo[dept].items[k]={text:'',emoji:'없음'};
  return state.memo[dept].items[k];
}
function setupMemo(dept){
  const nums=document.getElementById(dept+'Nums');
  const emos=document.getElementById(dept+'Emos');
  nums.innerHTML='';emos.innerHTML='';
  for(let n=1;n<=10;n++){
    const b=document.createElement('button');b.type='button';b.className='num-btn';b.textContent=n;
    b.onclick=()=>{
      state.memo[dept].active=n;save();renderMemo(dept);
    };nums.appendChild(b);
  }
  EMOJIS.forEach(em=>{
    const b=document.createElement('button');b.type='button';b.className='emo-btn';b.textContent=em==='없음'?'—':em;
    b.onclick=()=>{memoItem(dept,state.memo[dept].active).emoji=em;save();renderMemo(dept)};
    emos.appendChild(b);
  });
  document.getElementById(dept+'Text').addEventListener('input',e=>{
    memoItem(dept,state.memo[dept].active).text=e.target.value;save();renderMemoList(dept);markNums(dept);
  });
}
function markNums(dept){
  [...document.getElementById(dept+'Nums').children].forEach((b,i)=>{
    const it=state.memo[dept].items[String(i+1)];
    b.classList.toggle('active',state.memo[dept].active===i+1);
    b.classList.toggle('has',!!(it&&it.text&&it.text.trim()));
  });
}
function renderMemo(dept){
  const n=state.memo[dept].active,it=memoItem(dept,n);
  const tx=document.getElementById(dept+'Text');tx.value=it.text||'';tx.placeholder=`${CIRCLED[n-1]} ${dept==='front'?'프런트':'주방'} 전달사항 입력.`;
  markNums(dept);
  [...document.getElementById(dept+'Emos').children].forEach((b,i)=>b.classList.toggle('active',EMOJIS[i]===it.emoji));
  renderMemoList(dept);
}
function renderMemoList(dept){
  const box=document.getElementById(dept+'List');
  let lines='';
  for(let n=1;n<=10;n++){
    const it=state.memo[dept].items[String(n)];
    if(it&&it.text&&it.text.trim()){
      lines+=`<div class="memo-line"><span class="memo-circle">${n}</span><span>${it.emoji&&it.emoji!=='없음'?it.emoji+' ':''}${escapeHtml(it.text)}</span></div>`;
    }
  }
  box.innerHTML=lines||'<span class="small">작성된 메모 없음.</span>';
}
function escapeHtml(s){return String(s||'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]))}

document.getElementById('roomMemoText').addEventListener('input',e=>{
  if(selectedMemoRow===null || !state.rows[selectedMemoRow]) return;
  state.rows[selectedMemoRow].roomMemo=e.target.value;
  save();
  renderFloorplan();
});
document.getElementById('roomMemoClose').addEventListener('click',()=>{
  selectedMemoRow=null;
  document.getElementById('roomMemoPanel').classList.remove('active');
});

function blankStateForDate(date){
  const fresh=structuredClone(init);
  fresh.date=date||new Date().toISOString().slice(0,10);
  fresh.rows=Array.from({length:13},()=>normalizeRow(emptyRow()));
  fresh.finalLocked=false;
  return fresh;
}
function switchReadingDate(nextDate){
  nextDate=String(nextDate||'').trim();
  if(!nextDate || nextDate===state.date) return;
  const saved=archiveLoadAll().find(x=>x && x.date===nextDate && x.snapshot);
  if(saved){
    state=JSON.parse(JSON.stringify(saved.snapshot));
    state.date=nextDate;
    if(!Array.isArray(state.rows)) state.rows=[];
    while(state.rows.length<13) state.rows.push(emptyRow());
    state.rows=state.rows.slice(0,13).map(normalizeRow);
  }else{
    state=blankStateForDate(nextDate);
  }
  selectedMemoRow=null; activeReadingRow=null; closeCompletedDetail();
  document.getElementById('breakfastTable')?.classList.remove('focus-mode');
  document.getElementById('roomMemoPanel')?.classList.remove('active');
  document.getElementById('sheetDate').value=nextDate;
  save(); renderRows(); renderMemo('front'); renderMemo('kitchen');
  document.querySelector('.table-wrap')?.classList.add('date-switch-burst');
  setTimeout(()=>document.querySelector('.table-wrap')?.classList.remove('date-switch-burst'),700);
  showArchiveToast(saved ? `${formatReadingTitle(nextDate)} 저장본 불러옴.` : `${formatReadingTitle(nextDate)} 새 화면.`);
}
document.getElementById('sheetDate').addEventListener('change',e=>switchReadingDate(e.target.value));

function kakaoText(){
  const used=state.rows.filter(r=>r.room||r.name||personTotal(r)||mealItems(r).length);
  let out=[`[조식 LIST ${state.date}]`];
  used.forEach(r=>{
    const menus=mealItems(r);
    let line=`${r.room||'-'}호 / 자리 ${r.seat||'-'} / ${r.name||'-'} / 인원 ${peopleSummaryText(r)}${personDetailText(r)?` (${personDetailText(r)})`:''} / ${r.type||'-'} / ${r.time||'-'} / ${menus.join(', ')||'메뉴없음'}`;
    if(String(r.status||'').includes('룸터치') && (r.roomTouchStart||r.roomTouchEnd)){
      line+=` / 룸터치 ${r.roomTouchStart||'?'}~${r.roomTouchEnd||'?'}`;
    }
    if(String(r.status||'').includes('연박')){
      const nights=Number(r.stayNights)||1;
      line+=` / 연박 ${nights}박`;
      (r.stayBreakfasts||[]).slice(0,nights).forEach((x,idx)=>{if(String(x||'').trim()) line+=` / ${idx+1}박후 조식: ${String(x).trim()}`});
    }
    if(String(r.roomMemo||'').trim()) line+=` / 객실메모: ${String(r.roomMemo).trim()}`;
    out.push(line);
  });
  out.push('','<프런트>');
  for(let n=1;n<=10;n++){const it=state.memo.front.items[String(n)];if(it&&it.text&&it.text.trim())out.push(`${n}. ${it.emoji!=='없음'?it.emoji+' ':''}${it.text.trim()}`)}
  out.push('','<주방>');
  for(let n=1;n<=10;n++){const it=state.memo.kitchen.items[String(n)];if(it&&it.text&&it.text.trim())out.push(`${n}. ${it.emoji!=='없음'?it.emoji+' ':''}${it.text.trim()}`)}
  return out.join('\n');
}
document.getElementById('copyBtn').onclick=async()=>{
  try{await navigator.clipboard.writeText(kakaoText());alert('카톡용 요약 복사 완료.')}catch(e){alert('복사 권한 확인 필요.')}
};

document.getElementById('resetBtn').onclick=()=>{
  if(!confirm('조식표와 메모를 모두 초기화할까요?'))return;
  localStorage.removeItem(KEY);state=structuredClone(init);selectedMemoRow=null;document.getElementById('roomMemoPanel').classList.remove('active');document.getElementById('sheetDate').value=state.date;renderRows();renderMemo('front');renderMemo('kitchen');save();
};


function archiveLoadAll(){
  try{
    const x=JSON.parse(localStorage.getItem(ARCHIVE_KEY));
    return Array.isArray(x)?x:[];
  }catch(e){ return []; }
}

function archiveStoreAll(items){
  localStorage.setItem(ARCHIVE_KEY,JSON.stringify(items));
}

function formatReadingTitle(dateStr){
  const m=String(dateStr||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if(!m) return '작성일 미지정 조식리딩';
  return `${Number(m[1])}년${Number(m[2])}월${Number(m[3])}일 조식리딩`;
}

function formatSavedTime(iso){
  try{
    const d=new Date(iso);
    return `${d.getFullYear()}.${String(d.getMonth()+1).padStart(2,'0')}.${String(d.getDate()).padStart(2,'0')} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
  }catch(e){ return ''; }
}

function cloneStateForArchive(){
  return JSON.parse(JSON.stringify(state));
}

let selectedArchiveId=null;
let archiveToastTimer=null;

function showArchiveToast(text){
  const el=document.getElementById('archiveToast');
  el.textContent=text;
  el.classList.add('show');
  clearTimeout(archiveToastTimer);
  archiveToastTimer=setTimeout(()=>el.classList.remove('show'),1600);
}

function saveReadingArchive(){
  const date=state.date || new Date().toISOString().slice(0,10);
  state.date=date;
  document.getElementById('sheetDate').value=date;

  const items=archiveLoadAll();
  const now=new Date().toISOString();
  const existingIndex=items.findIndex(x=>x && x.date===date);
  const item={
    id:existingIndex>=0 ? items[existingIndex].id : `${date}_${Date.now()}`,
    date,
    title:formatReadingTitle(date),
    savedAt:now,
    snapshot:cloneStateForArchive()
  };
  if(existingIndex>=0) items[existingIndex]=item;
  else items.push(item);

  items.sort((a,b)=>String(b.date).localeCompare(String(a.date)) || String(b.savedAt).localeCompare(String(a.savedAt)));
  archiveStoreAll(items);
  renderArchiveList();
  showArchiveToast(`${item.title} 저장 완료`);
}

function renderArchiveList(){
  const list=document.getElementById('archiveList');
  const count=document.getElementById('archiveCount');
  const items=archiveLoadAll();
  count.textContent=`${items.length}건`;
  if(!items.length){
    list.innerHTML='<div class="archive-empty">아직 날짜별로 저장한 조식 리딩 없음.</div>';
    return;
  }
  list.innerHTML=items.map(item=>`
    <button type="button" class="archive-item" data-archive-id="${safeHtml(item.id)}">
      <span>
        <b>${safeHtml(item.title || formatReadingTitle(item.date))}</b>
        <small>최종 저장 ${safeHtml(formatSavedTime(item.savedAt))}</small>
      </span>
      <span class="open-mark">›</span>
    </button>
  `).join('');

  list.querySelectorAll('[data-archive-id]').forEach(btn=>{
    btn.addEventListener('click',()=>openArchive(btn.dataset.archiveId));
  });
}

function memoLinesForArchive(dept){
  const out=[];
  const items=dept?.items||{};
  for(let n=1;n<=10;n++){
    const it=items[String(n)];
    if(it && String(it.text||'').trim()){
      out.push(`${n}. ${it.emoji && it.emoji!=='없음' ? it.emoji+' ' : ''}${String(it.text).trim()}`);
    }
  }
  return out;
}

function openArchive(id){
  const item=archiveLoadAll().find(x=>x.id===id);
  if(!item) return;
  selectedArchiveId=id;
  const snap=item.snapshot||{};
  const rows=Array.isArray(snap.rows)?snap.rows:[];
  const used=rows.map(normalizeRow).filter(r=>r && (r.room||r.name||personTotal(r)||r.time||r.roomMemo||mealText(r)));

  document.getElementById('archiveViewTitle').textContent=item.title || formatReadingTitle(item.date);
  document.getElementById('archiveViewMeta').textContent=`저장 ${formatSavedTime(item.savedAt)} · ${used.length}개 객실 작성`;

  const roomHtml=used.length ? used.map(r=>{
    const menus=mealText(r)||'메뉴 미입력';
    let stay='';
    if(String(r.status||'').includes('연박')){
      const nights=Number(r.stayNights)||1;
      const each=(r.stayBreakfasts||[]).slice(0,nights).map((x,i)=>String(x||'').trim()?`<div><b>${i+1}박 후:</b> ${safeHtml(String(x).trim())}</div>`:'').join('');
      stay=`<div class="archive-stay"><b>연박 ${nights}박</b>${each||'<div>박수별 조식메뉴 미입력</div>'}</div>`;
    }
    return `
      <article class="archive-room">
        <div class="archive-room-head">
          <b>${safeHtml(r.room||'-')}호 · 자리 ${safeHtml(r.seat||'-')}</b>
          <span>${safeHtml(r.status||'-')} · 인원 ${safeHtml(peopleSummaryText(r))} · ${safeHtml(r.time||'-')}</span>
        </div>
        <div class="archive-room-body">
          <div><b>예약자</b> ${safeHtml(r.name||'-')}</div>
          <div class="archive-people"><b>인원 상세</b> ${safeHtml(personDetailText(r)||'-')}</div>
          <div><b>결제/예약</b> ${safeHtml(r.type||'-')}</div>
          <div><b>조식</b> ${safeHtml(menus)}</div>
          ${stay}
          ${String(r.roomMemo||'').trim()?`<div class="memo"><b>객실 메모</b><br>${safeHtml(String(r.roomMemo).trim()).replace(/\n/g,'<br>')}</div>`:''}
        </div>
      </article>`;
  }).join('') : '<div class="archive-empty">이 저장본에는 작성된 객실 없음.</div>';

  const front=memoLinesForArchive(snap.memo?.front);
  const kitchen=memoLinesForArchive(snap.memo?.kitchen);
  const notes=`
    <div class="archive-notes">
      <div class="archive-note-card"><b>&lt;프런트&gt;</b>${front.length?front.map(x=>`<div>${safeHtml(x)}</div>`).join(''):'작성 없음'}</div>
      <div class="archive-note-card"><b>&lt;주방&gt;</b>${kitchen.length?kitchen.map(x=>`<div>${safeHtml(x)}</div>`).join(''):'작성 없음'}</div>
    </div>`;

  document.getElementById('archiveSummary').innerHTML=roomHtml+notes;
  const modal=document.getElementById('archiveModal');
  modal.classList.add('open');
  modal.setAttribute('aria-hidden','false');
}

function closeArchive(){
  const modal=document.getElementById('archiveModal');
  modal.classList.remove('open');
  modal.setAttribute('aria-hidden','true');
  selectedArchiveId=null;
}

function loadSelectedArchive(){
  const item=archiveLoadAll().find(x=>x.id===selectedArchiveId);
  if(!item || !item.snapshot) return;
  if(!confirm(`${item.title} 내용을 현재 조식 리딩 화면으로 불러올까요?\n현재 작성 중인 내용은 해당 저장본으로 바뀝니다.`)) return;
  state=JSON.parse(JSON.stringify(item.snapshot));
  state.finalLocked=false;
  if(!Array.isArray(state.rows)) state.rows=Array.from({length:13},emptyRow);
  while(state.rows.length<13) state.rows.push(emptyRow());
  state.rows=state.rows.slice(0,13).map(normalizeRow);
  selectedMemoRow=null;
  activeReadingRow=null;
  document.getElementById('breakfastTable').classList.remove('focus-mode');
  document.getElementById('roomMemoPanel').classList.remove('active');
  document.getElementById('sheetDate').value=state.date||'';
  renderRows();
  renderMemo('front');
  renderMemo('kitchen');
  save();
  closeArchive();
  window.scrollTo({top:0,behavior:'smooth'});
  showArchiveToast('저장 기록 불러오기 완료.');
}

document.getElementById('archiveSaveBtn').addEventListener('click',saveReadingArchive);
document.getElementById('finalizeAllBtn')?.addEventListener('click',finalizeAllReading);
document.getElementById('editFinalizedBtn')?.addEventListener('click',editFinalizedReading);
document.getElementById('archiveCloseBtn').addEventListener('click',closeArchive);
document.getElementById('archiveModalCloseBtn').addEventListener('click',closeArchive);
document.getElementById('archiveLoadBtn').addEventListener('click',loadSelectedArchive);
document.getElementById('archiveModal').addEventListener('mousedown',e=>{
  if(e.target.id==='archiveModal') closeArchive();
});
document.addEventListener('keydown',e=>{
  if(e.key==='Escape' && document.getElementById('archiveModal').classList.contains('open')) closeArchive();
});


/* ---------- Excel (.xlsx) 저장 ---------- */
function excelDate(){
  return (state.date||'').replaceAll('-','.');
}
function excelRowsAoA(){
  const aoa=[];
  aoa.push(['조식 LIST',`작성일 ${excelDate()}`]);
  aoa.push(['순번','호실','자리','상태','이름','인원','인원상세','결제','시간','톳밥','죽','어린이용','스프','총','룸터치 시작','룸터치 종료','객실메모']);
  state.rows.slice(0,13).forEach((r,i)=>{
    const childText=childMenuItems(r).join(', ');
    const total=Number(r.tot||0)+Number(r.porridge||0)+childTotal(r)+Number(r.soup||0);
    aoa.push([i+1,r.room||'',r.seat||'',r.status||'',r.name||'',peopleSummaryText(r),personDetailText(r),r.type||'',r.time||'',Number(r.tot||0),Number(r.porridge||0),childText,Number(r.soup||0),total,r.roomTouchStart||'',r.roomTouchEnd||'',String(r.roomMemo||'').trim()]);
  });
  aoa.push([]);
  aoa.push(['<프런트>']);
  for(let n=1;n<=10;n++){
    const it=state.memo.front.items[String(n)];
    if(it&&it.text&&it.text.trim()) aoa.push([`${n}. ${it.emoji&&it.emoji!=='없음'?it.emoji+' ':''}${it.text.trim()}`]);
  }
  aoa.push([]);
  aoa.push(['<주방>']);
  for(let n=1;n<=10;n++){
    const it=state.memo.kitchen.items[String(n)];
    if(it&&it.text&&it.text.trim()) aoa.push([`${n}. ${it.emoji&&it.emoji!=='없음'?it.emoji+' ':''}${it.text.trim()}`]);
  }
  return aoa;
}
function makeExcel(){
  const aoa=excelRowsAoA();
  if(typeof XLSX==='undefined'){
    fallbackExcelCsv();
    return;
  }
  const wb=XLSX.utils.book_new();
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols']=[{wch:5},{wch:8},{wch:8},{wch:16},{wch:16},{wch:9},{wch:18},{wch:13},{wch:9},{wch:7},{wch:7},{wch:20},{wch:7},{wch:7},{wch:24}];
  XLSX.utils.book_append_sheet(wb,ws,'조식 LIST');
  const file=`조식_LIST_${(state.date||'').replaceAll('-','')||'작성일'}.xlsx`;
  XLSX.writeFile(wb,file);
}
function fallbackExcelCsv(){
  const aoa=excelRowsAoA();
  const csv='﻿'+aoa.map(row=>row.map(v=>`"${String(v??'').replace(/"/g,'""')}"`).join(',')).join('\n');
  const blob=new Blob([csv],{type:'text/csv;charset=utf-8;'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download=`조식_LIST_${(state.date||'').replaceAll('-','')||'작성일'}.csv`;
  a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);
}
document.getElementById('excelBtn').onclick=makeExcel;



// 완료된 행은 PC 클릭/모바일 터치 시 세로형 핵심 정보카드로 크게 표시함.
document.getElementById('rows')?.addEventListener('click',e=>{
  if(e.target.closest('button,input,select,textarea,label')) return;
  const tr=e.target.closest('tr.reading-row.is-completed');
  if(!tr) return;
  const i=Number(tr.dataset.rowIndex);
  if(Number.isFinite(i)) openCompletedDetail(i,tr);
});


document.getElementById('stayReviewBack')?.addEventListener('click',closeStayReview);
document.getElementById('stayReviewModal')?.addEventListener('mousedown',e=>{if(e.target.id==='stayReviewModal')closeStayReview()});
document.getElementById('stayReviewOk')?.addEventListener('click',()=>{const i=pendingStayCompleteIndex;const m=document.getElementById('stayReviewModal');m?.classList.remove('open');m?.setAttribute('aria-hidden','true');pendingStayCompleteIndex=null;if(Number.isInteger(i))finishReadingRow(i)});

/* ===== 조식LIST 텍스트 / 차트 입력 ===== */
const TEXT_ROOM_LIST=['101','102','201','202','203','204','205','301','302','303','304','305','401'];
function setSheetTextStatus(message,type=''){const el=document.getElementById('sheetTextStatus');if(!el)return;el.textContent=message;el.className='json-status'+(type?' '+type:'')}
function normalizeImportTime(v){const s=String(v||'').trim().replace(/\s+/g,'');if(['8','8시','08:00','8:00','오전8시','아침8시'].includes(s))return '8시';if(['8시30분','8시반','08:30','8:30','오전8시30분','오전8시반'].includes(s))return '8시30분';if(['9','9시','09:00','9:00','오전9시','아침9시'].includes(s))return '9시';return TIMES.includes(v)?v:''}
function normalizePayment(v){const s=String(v||'').trim();if(!s)return '';if(/패키지/.test(s))return '패키지';if(/서비스/.test(s))return '서비스';if(/현장/.test(s))return '현장결제';return TYPES.includes(s)?s:'기타'}
function clampCount(v){return Math.max(0,Math.min(10,Number(v)||0))}
function stripJsonFence(text){let t=String(text||'').trim();t=t.replace(/^```(?:json)?\s*/i,'').replace(/\s*```$/,'').trim();const a=t.indexOf('{'),b=t.lastIndexOf('}');if(a>=0&&b>a)t=t.slice(a,b+1);return t}
function parseBreakfastImport(text){const raw=stripJsonFence(text);if(!raw)throw new Error('붙여넣은 내용이 없음.');const data=JSON.parse(raw);if(!data||!Array.isArray(data.rows))throw new Error('rows 배열을 찾지 못함.');const rows=data.rows.map(x=>({room:String(x.room||'').trim(),name:String(x.name||'').trim(),status:String(x.status||'').trim(),adults:clampCount(x.adults),kids:clampCount(x.kids),type:normalizePayment(x.type),time:normalizeImportTime(x.time),tot:clampCount(x.tot),porridge:clampCount(x.porridge),childPorridge:clampCount(x.childPorridge),childRice:clampCount(x.childRice),childSoup:clampCount(x.childSoup),soup:clampCount(x.soup)})).filter(x=>TEXT_ROOM_LIST.includes(x.room));if(!rows.length)throw new Error('유효한 객실번호를 찾지 못함.');const cleanNotes=v=>Array.isArray(v)?v.map(x=>String(x??'').trim()).filter(Boolean).slice(0,10):[];return {date:String(data.date||'').trim(),rows,frontNotes:cleanNotes(data.frontNotes),kitchenNotes:cleanNotes(data.kitchenNotes)}}
function applyImportedNotes(dept,notes){if(!state.memo||!state.memo[dept])return;state.memo[dept].items={};(notes||[]).slice(0,10).forEach((text,idx)=>{state.memo[dept].items[String(idx+1)]={text:String(text||'').trim(),emoji:'없음'}});state.memo[dept].active=1;}
function applyBreakfastImport(){let parsed;try{parsed=parseBreakfastImport(document.getElementById('sheetTextJson')?.value||'')}catch(e){setSheetTextStatus('적용 실패 · '+(e.message||'JSON 형식을 확인함.'),'error');return}let applied=0;const touched=[];parsed.rows.forEach(c=>{let idx=state.rows.findIndex(r=>String(r.room)===c.room);if(idx<0)idx=state.rows.findIndex(r=>!String(r.room||'').trim());if(idx<0)return;const r=state.rows[idx];r.room=c.room;r.name=c.name;r.status=c.status==='연박'?'연박':c.status;r.type=c.type;r.time=c.time;const groups=[];if(c.adults>0)groups.push({type:'성인',count:c.adults});if(c.kids>0)groups.push({type:'아이',count:c.kids});r.peopleEnabled=groups.length>0;r.peopleGroups=groups;r.people=c.adults+c.kids;r.tot=c.tot;r.porridge=c.porridge;r.soup=c.soup;const cm=[];if(c.childPorridge>0)cm.push({type:'어죽',count:c.childPorridge});if(c.childRice>0)cm.push({type:'어밥',count:c.childRice});if(c.childSoup>0)cm.push({type:'어미',count:c.childSoup});r.childEnabled=cm.length>0;r.childMenus=cm;r.completed=false;r.completedAt='';touched.push(c.room);applied++});if(parsed.date&&/^\d{4}-\d{2}-\d{2}$/.test(parsed.date)){state.date=parsed.date;const d=document.getElementById('sheetDate');if(d)d.value=parsed.date}applyImportedNotes('front',parsed.frontNotes);applyImportedNotes('kitchen',parsed.kitchenNotes);state.finalLocked=false;save();renderRows();renderMemo('front');renderMemo('kitchen');setSheetTextStatus(`${applied}개 객실을 차트에 입력함. 각 행을 확인 후 [완성]을 누름.`,'ok');setTimeout(()=>{touched.forEach(room=>{const i=state.rows.findIndex(r=>r.room===room);document.querySelector(`tr.reading-row[data-row-index="${i}"]`)?.classList.add('import-highlight')});setTimeout(()=>document.querySelectorAll('.import-highlight').forEach(x=>x.classList.remove('import-highlight')),1800)},80);setTimeout(closeSheetText,650)}
const ORDER_TO_JSON_PROMPT='Analyze the attached single TOYU breakfast LIST capture and convert every populated guest row into the exact JSON schema below. Return JSON only. Do not add explanations, markdown commentary, or any text outside the JSON code block. \n \nREAD THESE FIELDS FOR EACH POPULATED ROOM: \n- room number \n- guest name (preserve Korean, English, Chinese, or other original text exactly as shown in the capture) \n- stay status \n- guest count \n- payment type \n- breakfast time \n- Totbap quantity \n- porridge quantity \n- child menu quantities \n- soup quantity \n \nVALID ROOMS ONLY: \n101, 102, 201, 202, 203, 204, 205, 301, 302, 303, 304, 305, 401 \n \nCRITICAL NAME-READING RULES: \n1. Read guest names ONLY from the guest-name cell that is visibly filled in the currently attached image. \n2. Preserve the visible guest name exactly as shown. Do not correct spelling, translate it, normalize it, romanize it, or replace it with a similar name. \n3. NEVER infer, guess, reconstruct, autofill, or reuse a guest name from: \n   - previous images \n   - previous conversation turns \n   - memory \n   - another room \n   - bottom notes \n   - reservation patterns \n   - prior TOYU breakfast LIST captures \n4. If the guest-name cell for a room is blank, unclear, cut off, or unreadable, DO NOT create a row for that room. \n5. A room number, payment value, breakfast time, menu quantity, colored row, or bottom note by itself does NOT make that row a valid guest row. A clearly visible guest name is required. \n6. Before finalizing the JSON, re-check every output name against the currently attached image one by one. Every name in the JSON must be visibly present in that image. \n7. If a name does not visibly exist in the currently attached image, it must not appear anywhere in the rows array. \n \nINTERPRETATION RULES: \n1. A light-green guest row means status = "연박". Otherwise status = "" unless another status is explicitly visible. \n2. Guest count format: \n   - "2" means adults=2, kids=0. \n   - "2+1" means adults=2, kids=1. \n   - "6+4" means adults=6, kids=4. \n3. Payment values should preserve the operational Korean labels when visible, especially "패키지", "서비스", or "현장결제". \n4. Normalize breakfast time to one of: "8시", "8시30분", "9시". \n5. Read each menu quantity from its own column. Do not shift a number into a neighboring menu column. \n6. Menu mapping: \n   - 톳밥 -> tot \n   - 죽 -> porridge \n   - 어죽 -> childPorridge \n   - 어밥 -> childRice \n   - 어미 -> childSoup \n   - 스프 -> soup \n7. If a menu quantity is blank, output 0. \n8. Ignore blank guest rows completely. \n9. Read the date shown in the upper-right area and output YYYY-MM-DD. \n10. Do not guess information that is not visible. \n11. Cross-check each room\'s menu total against the visible guest information when possible, and cross-check the bottom menu totals when they are visible. \n12. If any field other than the guest name is unreadable or not visibly present, do not guess it. Use only information visible in the current capture. \n13. The currently attached image is the ONLY source of truth for guest-row extraction. \n \nBOTTOM NOTES: \nThe capture may contain sections labeled <프런트> and <주방> below the main chart. \nExtract each written note in its original Korean wording and original order. \nDo not mix these notes into guest rows. \nDo not use a name appearing only inside a bottom note as the guest name for a room row. \nPut the notes into frontNotes and kitchenNotes. \nIf a section has no written note, return an empty array. \n \nOUTPUT SCHEMA: \n{ \n  "format": "TOYU_BREAKFAST_V1", \n  "date": "YYYY-MM-DD", \n  "rows": [ \n    { \n      "room": "305", \n      "name": "이상은", \n      "status": "", \n      "adults": 2, \n      "kids": 0, \n      "type": "패키지", \n      "time": "8시", \n      "tot": 2, \n      "porridge": 0, \n      "childPorridge": 0, \n      "childRice": 0, \n      "childSoup": 0, \n      "soup": 0 \n    } \n  ], \n  "frontNotes": ["프런트 메모 1", "프런트 메모 2"], \n  "kitchenNotes": ["주방 메모 1", "주방 메모 2"] \n} \n \nFINAL VALIDATION BEFORE OUTPUT: \n- Confirm that every room in rows is one of the VALID ROOMS. \n- Confirm that every row has a guest name visibly written in the currently attached image. \n- Confirm that every output guest name exactly matches the visible text in that image. \n- Remove any row whose guest-name cell is blank. \n- Remove any name that came from memory, a previous image, a previous conversation, or a bottom note. \n- Never fill a missing name based on what was shown in an earlier TOYU LIST. \n- Output one valid TOYU_BREAKFAST_V1 JSON object only. \n \nIMPORTANT: \nThe instruction is written in English, but all extracted Korean content must remain in Korean exactly as shown in the capture. \nUse ONLY the currently attached image. \nOutput one valid TOYU_BREAKFAST_V1 JSON object only.';
function openSheetText(){document.getElementById('sheetTextModal')?.classList.add('open');document.getElementById('sheetTextModal')?.setAttribute('aria-hidden','false');setSheetTextStatus('붙여넣은 뒤 적용함.');setTimeout(()=>document.getElementById('sheetTextJson')?.focus(),80)}
function closeSheetText(){document.getElementById('sheetTextModal')?.classList.remove('open');document.getElementById('sheetTextModal')?.setAttribute('aria-hidden','true')}
document.getElementById('openSheetTextBtn')?.addEventListener('click',openSheetText);document.getElementById('sheetTextClose')?.addEventListener('click',closeSheetText);document.getElementById('sheetTextCancel')?.addEventListener('click',closeSheetText);document.getElementById('sheetTextModal')?.addEventListener('mousedown',e=>{if(e.target.id==='sheetTextModal')closeSheetText()});document.getElementById('sheetTextApply')?.addEventListener('click',applyBreakfastImport);document.getElementById('sheetTextJson')?.addEventListener('input',()=>setSheetTextStatus('텍스트가 입력됨. [조식차트에 적용]을 누름.'));
function openOrderPrompt(){const m=document.getElementById('orderPromptModal'),t=document.getElementById('orderPromptText');if(t)t.value=ORDER_TO_JSON_PROMPT;m?.classList.add('open');m?.setAttribute('aria-hidden','false');document.getElementById('orderPromptStatus').textContent='Ready to copy.'}
function closeOrderPrompt(){const m=document.getElementById('orderPromptModal');m?.classList.remove('open');m?.setAttribute('aria-hidden','true')}
async function copyOrderPrompt(){const t=document.getElementById('orderPromptText'),s=document.getElementById('orderPromptStatus');if(!t)return;let ok=false;try{await navigator.clipboard.writeText(t.value);ok=true}catch(e){try{t.focus();t.select();ok=document.execCommand('copy')}catch(_){ok=false}}if(s)s.textContent=ok?'Copied.':'Select the text and copy it manually.'}
document.getElementById('openOrderPromptBtn')?.addEventListener('click',openOrderPrompt);document.getElementById('orderPromptClose')?.addEventListener('click',closeOrderPrompt);document.getElementById('orderPromptCancel')?.addEventListener('click',closeOrderPrompt);document.getElementById('copyOrderPromptBtn')?.addEventListener('click',copyOrderPrompt);document.getElementById('orderPromptModal')?.addEventListener('mousedown',e=>{if(e.target.id==='orderPromptModal')closeOrderPrompt()});

load();
document.getElementById('sheetDate').value=state.date;
renderRows();
setupMemo('front');setupMemo('kitchen');
renderMemo('front');renderMemo('kitchen');
renderArchiveList();
</script>
</body>
</html>
