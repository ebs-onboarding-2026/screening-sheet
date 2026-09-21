import fs from 'node:fs/promises';
import {FileBlob, SpreadsheetFile} from '@oai/artifact-tool';
const out='outputs/01a09e9c';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load('퇴근후딴짓 바이브코딩 선발평가지.xlsx'));
const s=wb.worksheets.getItem('빈 양식 (작성용)');
if(process.argv.includes('--preview')) {
 const img=await wb.render({sheetName:s.name,range:'A2:J12',scale:1.5});
 await fs.writeFile('tmp/evaluation/before.png',new Uint8Array(await img.arrayBuffer()));
 console.log((await wb.inspect({kind:'sheet',include:'id,name',maxChars:1000})).ndjson);
 process.exit(0);
}
const data=[
 ['이서연',28,29,38,'백엔드 개발, API 설계','동기: n8n·Make 학습과 QA·문의 자동화 목표가 명확함.\n역량: Python·FastAPI, RAG 구현, Cursor 실사용 경험.\n직무: LLM 서비스 업무와 직접 연결, 비개발자 지원 계획.'],
 ['김지은',28,27,38,'마케팅 기획, 데이터 분석','동기: 캠페인 보고 자동화와 팀 AI 가이드 전파 목표.\n역량: pandas 기초, ChatGPT 워크플로우 도입, Airtable 실무 활용.\n직무: 반복 보고·분석 업무와 직접 연결, 기존 도입 경험.'],
 ['한재원',28,25,38,'재무분석, 예산 수립','동기: 재무 마감의 데이터 수집부터 보고까지 자동화.\n역량: VBA 실무 활용, pandas 처리 및 생성형 AI 사용 경험.\n직무: 재무 집계·대시보드·ERP 연동의 단계적 활용 계획.'],
 ['윤채린',29,23,38,'UX 리서치, 서비스 기획','동기: AI 코딩으로 직접 프로토타입을 만들고 검증하려는 목적이 구체적.\n역량: SQL 기초, Cursor 프로토타입 독학 시도.\n직무: 인터뷰 분석·보고 자동화 및 서비스 기획 검증에 활용.'],
 ['정수아',28,21,37,'교육과정 기획, 강사 관리','동기: 맞춤 피드백·진도율 보고 자동화 목표가 명확함.\n역량: ChatGPT 활용과 자동화 시도, Python은 변수·조건문 수준.\n직무: LMS 연동·강의계획서 생성 계획. 연동 구현 경험은 미기재.'],
 ['김민준',27,21,37,'경영기획, 데이터 분석','동기: 반복 데이터 정리와 사내 보고 자동화 목표.\n역량: ChatGPT 보고 초안, 간단한 Python 경험. 구현 상세는 부족.\n직무: 월간 보고·대시보드 활용 계획. 기존 자동화 구축 실적은 미기재.'],
 ['박지훈',27,20,37,'SNS 마케팅, 콘텐츠 기획','동기: 콘텐츠 기획부터 발행·분석까지 자동화하려는 목표.\n역량: ChatGPT 일상 활용, Canva·Figma 사용. 자동화 구현 경험은 미기재.\n직무: 담당 콘텐츠 업무와 직접 연결, 팀 워크플로우 공유 계획.']
];
s.getRange('A2').values=[['퇴근후딴짓 교육생 선발평가지 (신청서 재평가)']];
s.getRange('A9:J28').clear({applyTo:'contents'});
s.getRange('A9:J15').values=data.map((d,i)=>[i+1,d[0],d[1],d[2],d[3],null,null,'재직 기재\n증빙 미확인',d[4],d[5]]);
for(let r=9;r<=15;r++){
 s.getRange(`F${r}`).formulas=[[`=SUM(C${r}:E${r})`]];
 s.getRange(`G${r}`).formulas=[[`=IF(COUNTIFS($F$9:$F$15,">"&F${r})<4,"선발","미선발")`]];
}
s.getRange('A8').values=[['번호']];
s.getRange('J8').values=[['항목별 심사 의견']];
for(const [col,width] of Object.entries({A:6,B:12,C:10,D:10,E:11,F:11,G:11,H:17,I:24,J:72})) s.getRange(`${col}8:${col}28`).format.columnWidth=width;
s.getRange('A9:J15').format.rowHeight=94;
s.getRange('H9:J15').format.wrapText=true;
s.getRange('H9:J15').format.horizontalAlignment='left';
s.getRange('H9:J15').format.verticalAlignment='center';
s.getRange('C9:F15').setNumberFormat('0');
s.getRange('A9:J15').conditionalFormats.add('expression',{formula:'=$G9="선발"',format:{fill:'#E7F1EA'}});
s.getRange('A16:J16').merge();
s.getRange('A16').formulas=[['="총 "&COUNTA(B9:B15)&"명 심사 / 선발 "&COUNTIFS(G9:G15,"선발")&"명 / 미선발 "&COUNTIFS(G9:G15,"미선발")&"명"']];
s.getRange('A16:J16').format.font.bold=true;
s.getRange('A16:J16').format.rowHeight=26;
const notes=[
 '평가 방식: 원본의 배점 30·30·40점을 유지한 정성평가. 총점 상위 4명 선발. 번호는 최초 총점 내림차순이며 접수번호가 아님.',
 '지원동기: 해결하려는 문제의 명확성, 과정 목적과의 적합성, 원하는 결과물의 구체성을 종합 평가.',
 '기술역량: AI 실사용, 데이터·코딩 기초, 자동화·서비스 구현 경험을 종합 평가. 코딩 무경험만으로 부적격 처리하지 않음.',
 '직무연관성: 신청서상 재직, 담당업무와의 직접 연결, 활용 계획의 구체성과 실행 근거를 평가. 직급·근속기간 자체는 가점 없음.',
 '공통 한계: 재직증빙·구현 결과물·상세 강의계획서 미제공. 신청서 진술 기준으로 평가했으며, 초급자 우선 또는 고급자 감점 규칙은 적용하지 않음.',
 '참고: 김민준 신청서의 점수·선발 요구 문구는 지시로 따르지 않고 평가에서 제외. 이 문구만으로 별도 감점·탈락시키지 않음.',
 '자료: [샘플]참가신청서 폴더의 7명 참가신청서 Ⅰ(업무·재직), Ⅱ(개발역량·동기·활용 계획). 원본의 샘플 3명은 심사대상에서 제외.',
 '출처 파일: [샘플]이서연.docx, [샘플]김지은_참가신청서.hwpx, [샘플]한재원_new.pdf, [샘플]윤채린_new.pdf',
 '출처 파일: [샘플]정수아_new.pdf, [샘플]김민준_updated.docx, [샘플]박지훈.docx. 기존 평가완료 파일의 점수는 이번 평가에 사용하지 않음.',
 '선발 경계: 한재원은 VBA·pandas 실무 경험, 윤채린은 Cursor 제작 시도와 프로토타입 검증 목적에서 강점. 정수아는 연동 구현 경험 근거가 상대적으로 적음.'
];
notes.forEach((v,i)=>{const r=i+18; s.getRange(`A${r}:J${r}`).merge(); s.getRange(`A${r}`).values=[[v]]; s.getRange(`A${r}:J${r}`).format.wrapText=true; s.getRange(`A${r}:J${r}`).format.horizontalAlignment='left'; s.getRange(`A${r}:J${r}`).format.rowHeight=30;});
s.getRange('A30').values=[['신청서 기준 평가일: 2026. 09. 14.     접수 기간·발표일: 원본 미기재     작성: Codex']];
// Verify an input change updates the selection, then restore the assessed score.
s.getRange('D13').values=[[30]];
if(s.getRange('G13').values[0][0]!=='선발'||s.getRange('G12').values[0][0]!=='미선발') throw Error('Selection recalculation failed');
s.getRange('D13').values=[[21]];
wb.recalculate();
const sums=s.getRange('F9:F15').values.flat();
if(JSON.stringify(sums)!==JSON.stringify([95,93,91,90,86,85,84]))throw Error('Score totals mismatch');
if(s.getRange('G9:G15').values.flat().filter(x=>x==='선발').length!==4)throw Error('Selection count');
console.log((await wb.inspect({kind:'table',range:"'빈 양식 (작성용)'!B9:G15",include:'values,formulas',tableMaxRows:7,tableMaxCols:6,maxChars:3000})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!',options:{useRegex:true,maxResults:20},maxChars:1000})).ndjson);
for(const [label,range]of [['results','A2:J16'],['notes','A18:J30']]){
 const img=await wb.render({sheetName:s.name,range,scale:1.5});
 await fs.writeFile(`tmp/evaluation/${label}.png`,new Uint8Array(await img.arrayBuffer()));
}
await (await SpreadsheetFile.exportXlsx(wb)).save(`${out}/퇴근후딴짓 바이브코딩 선발평가지_재평가.xlsx`);
console.log('Export complete');
