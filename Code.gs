function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔖 Quiz')
    .addItem('Generate Quiz', 'generateQuiz')
    .addToUi();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔖 Export')
    .addItem('Export Custom Range as CSV (복사용)', 'exportPromptRangeToCSVCopy')
    .addToUi();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📚 니모닉 스토리북')
    .addItem('🔁 자동 이야기 생성 (200개)', 'generateStoriesFrom200Words')
    .addToUi();
}

function onEdit(e) {
  if (!e) return;
  const sheet    = e.source.getActiveSheet();
  const name     = sheet.getName().toLowerCase();
  const col      = e.range.getColumn();
  const startRow = e.range.getRow();
  const numRows  = e.range.getNumRows();

  if (name === 'master' && col === 1) {
    for (let i = 0; i < numRows; i++) {
      const row = startRow + i;
      if (row >= 2) handleMasterEdit(sheet, row);
    }
  }
  else if (name === 'quiz' && col === 3) {
    for (let i = 0; i < numRows; i++) {
      const row = startRow + i;
      if (row >= 2) handleQuizEdit(sheet, row, e.source);
    }
  }
}


function handleMasterEdit(sheet, row) {
  const word = sheet.getRange(row,1).getValue().toString().trim();
  if (!word) return;

  // B열: 번역
  sheet.getRange(row,2)
       .setFormula(`=GOOGLETRANSLATE(A${row},"en","ko")`);

  // C/G열: 의미·어원·mnemonic + 한 문장 요약
  const marker = '### SUMMARY ###';
  const prompt =
`단어 "${word}"에 대해 한글로 아래 형식을 그대로 작성해주세요.

1. 뜻 / 품사 / 발음 / 자,타동사여부
- 뜻: (한국어로 간단히)
- 품사: (명사, 형용사 등)
- 발음: (IPA 기호)

2. 어원 분석 (etymology)
구성 요소      의미
[접두사]        [뜻]
[어근] (라틴어) [뜻]
[접미사]        [뜻]
→ 위 어근 조합으로 전체 의미를 간단히 설명. 라틴어 어근/접두사/접미사 여부도 표기

3. mnemonic
- 청각/촉각/후각까지 연상하도록 설계 이미지 기억에 남는 비유 (예: "apple → 아담과 이브의 빨간 사과") 
- 경선식 영어 스타일묘사 
- 경선식 영어 스타일묘사 2

4. 관련 단어 (2~3개)
- 같은 어근을 가진 단어 + 뜻 (간결하게)

표현은 '~입니다', '~되었습니다' 사용 금지, 핵심 중심으로.

작성 후 반드시 단독 줄에 다음 마커를 넣고,
그 아래 줄에 "${word}" 의 한국식 영발음을 이용하여 뜻이 잘생각나는 mnemonic 스토리를 작성해주세요:

${marker}`;
  const resp = callOpenAI(prompt).trim();
  const idx  = resp.indexOf(marker);
  let detail = resp, summary = '';
  if (idx !== -1) {
    detail  = resp.substring(0, idx).trim();
    summary = resp.substring(idx + marker.length).trim();
  }
  sheet.getRange(row,3).setValue(detail);   // C열
  sheet.getRange(row,7).setValue(summary);  // G열

 // E열: 대중문화 용례
const fact = callOpenAI(
  `1) key  : ${word}
2) value: "<용례> – <출처>", 한국어, 70자 이내
   └ 용례는 게임·영화·연설·커뮤니티 밈·수학·과학 표기처럼 널리 알려진 사례거나 주변에서 흔히 사용되는 사례, 강력하게 인상적인 느낌을 재현하는게 목표
   └ 단어 의미가 즉시 연상될 만큼 직관적이어야 함
   └ 20대 남성들이 열광하는 주제면 가산점

[예시]
hazardous : "Biohazard – 좀비 공포게임 레지던트이블",
derivative: "d/dx – 교과서 미분 기호가 이 뜻임",
inevitable: "I am inevitable – 어벤저스엔드게임에서 타노스 손튕길 때",
fragile   : "Handle with care – 포장박스 깨지기 쉬움 경고",
unstoppable: "Unstoppable – 롤 연속킬 알림"

위 규칙과 예시 형식 참고해 작성하라.`).trim();

sheet.getRange(row, 5).setValue(fact);

  // H열: 이미지 URL + D열: 이미지 표시
  const imageUrl = fetchImageUrl(word);
  sheet.getRange(row,8).setValue(imageUrl);
  sheet.getRange(row,4).setFormula(`=IMAGE(H${row})`);


// J열: GPT로 어근 생성
var rootPrompt =  `${word}"의 어원(영어 기준)을 접두사, 어근, 접미사 구조로 분석해서 "접두사(뜻) + 어근(뜻) + 접미사(뜻): 전체 의미" 형태로 한문장으로 매우 짧게 적어줘. 최종적으로는 언어의 뜻과 일치되게, 연상에 도움되게 결론짓어 예시: analogous = ana(다시) + log(말하다, 이치) + ous(형용사형): 다시 말하다-> 술먹고 한말을 계속하다-> 유사한`
var root = callOpenAI(rootPrompt);  // gptCompletion → callOpenAI 로 맞춤
sheet.getRange(row, 10).setValue(root); // J열

// K열: GPT로 예문 생성
var sentPrompt = `Write a simple English sentence using the word "${word}", 문장 구조 분석을 / 표시로 나타내고, 밑에는 한국말 해석 그리고. "${word}" 와 - 같은 어근을 가진 단어 + 뜻 (간결하게)4개
예시)
She / will / impose / new / rules.  
주어 / 조동사 / 동사 / 형용사 / 목적어
그녀는 새로운 규칙을 부과할 것이다.

- position: 위치, 자리 (pos = 놓다)  
- deposit: 예금하다, 맡기다 (de- 아래로 + pos)  
- oppose: 반대하다 (op- 맞서 + pos)

` ;
var sentence = callOpenAI(sentPrompt);
sheet.getRange(row, 11).setValue(sentence); // K열
}


function handleQuizEdit(sheet, row, ss) {
  const userStory = sheet.getRange(row,3).getValue().toString().trim();
  const word      = sheet.getRange(row,2).getValue().toString().trim();
  if (!userStory || !word) return;

  const masterSheet = ss.getSheetByName('Master');
  const finder = masterSheet.createTextFinder(word).findNext();
  if (!finder) return;
  const mRow      = finder.getRow();
  const origStory = masterSheet.getRange(mRow,6).getValue().toString().trim();

const evalPrompt =
  `암기 대상 단어: ${word}\n` +
  `원본 암기 스토리: ${origStory}\n` +
  `학생 암기 스토리: ${userStory}\n\n` +
  `[평가 기준]\n` +
  `① 단어 뜻 유추 정확성 (60점)\n` +
  `   · 스토리에 단어의 뜻이 나와있으면 만점\n` +
  `② 핵심 키워드·문구 포함도 (30점)\n` +
  `   · 원본 스토리의 핵심 키워드의 50% 이상이 포함되면 만점\n` +
  `③ 니모닉·상황 묘사 창의성 (10점)\n` +
  `   · 구체적‧생생한 장면이나 연상이 있으면 만점\n\n` +
  `총점 = ① + ② + ③\n\n` +
  `[응답 형식]\n` +
  `- 70점 이상 → ✅ 통과\n` +
  `- 60점 미만 → ⚠️ 부족: <부족한 이유 한 문장>\n` +
  `점수와 한 줄 피드백만 반환`;
const feedback = callOpenAI(evalPrompt).split('\n')[0].trim();
sheet.getRange(row, 4).setValue(feedback);

  // 틀린 경우 Master 시트 I열(9) 오답 횟수 1 증가
if (feedback.startsWith('⚠️')) {
  const countCell = masterSheet.getRange(mRow, 9);
  const prevCount = parseInt(countCell.getValue(), 10) || 0;
  masterSheet.getRange(mRow, 9).setValue(prevCount + 1);
}
}

function generateQuiz() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const ui     = SpreadsheetApp.getUi();
  const master = ss.getSheetByName('Master');
  let quiz     = ss.getSheetByName('Quiz');
  if (!quiz) quiz = ss.insertSheet('Quiz'); else quiz.clear();
  quiz.appendRow(['No','Question','Your Story','Feedback']);

  const start = parseInt(ui.prompt('퀴즈 시작 행','몇 번째 행부터?',ui.ButtonSet.OK).getResponseText(),10) || 2;
  const end   = parseInt(ui.prompt('퀴즈 끝 행','몇 번째 행까지?',ui.ButtonSet.OK).getResponseText(),10) || master.getLastRow();
  const cnt   = parseInt(ui.prompt('문제 수','몇 개 출제할까요?',ui.ButtonSet.OK).getResponseText(),10) || 10;

  const words = master.getRange(start,1,end-start+1).getValues().flat().filter(v=>v);
  for (let i=words.length-1;i>0;i--){
    const j=Math.floor(Math.random()*(i+1)); [words[i],words[j]]=[words[j],words[i]];
  }
  for (let i=0;i<Math.min(cnt,words.length);i++){
    quiz.appendRow([i+1,words[i],'','']);
  }
}

function resetWrongCounts() {
  const master = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master');
  const last   = master.getLastRow();
  master.getRange(2,9,last-1).setValue(0);
}

function callOpenAI(prompt) {
  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions',{method:'post',contentType:'application/json',headers:{Authorization:`Bearer ${OPENAI_API_KEY}`},payload:JSON.stringify({model:'gpt-4.1-mini',messages:[{role:'user',content:prompt}],temperature:0.5})});
  return JSON.parse(res.getContentText()).choices[0].message.content;
}

function fetchImageUrl(word) {
  try{
    const url = `https://api.pexels.com/v1/search?query=${encodeURIComponent(word+' mnemonic')}&per_page=1`;
    const data = JSON.parse(UrlFetchApp.fetch(url,{headers:{Authorization:PEXELS_API_KEY}}).getContentText());
    return data.photos[0]?.src?.original||'https://via.placeholder.com/150';
  }catch(e){return 'https://via.placeholder.com/150';}
}

 // 이야기 펑션
 
function generateStoriesFrom200Words() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const sheetName = ui.prompt("시트 이름을 입력하세요 (예: Master)").getResponseText().trim();
  const startRow = parseInt(ui.prompt("시작 행 번호를 입력하세요 (예: 2)").getResponseText().trim(), 10);
  const totalCount = parseInt(ui.prompt("몇 개 행을 사용할까요? (예: 200)").getResponseText().trim(), 10);

  const sourceSheet = ss.getSheetByName(sheetName);
  if (!sourceSheet) {
    ui.alert(`시트 "${sheetName}"를 찾을 수 없습니다.`);
    return;
  }

  const data = sourceSheet.getRange(startRow, 1, totalCount, 4).getValues(); // A~D열 읽기
  const chunkSize = 10;
  const numChunks = Math.floor(data.length / chunkSize);
  const storySheet = ss.getSheetByName('Story') || ss.insertSheet('Story');

  for (let i = 0; i < numChunks; i++) {
    const chunk = data.slice(i * chunkSize, (i + 1) * chunkSize);
    const prompt = buildPromptFromChunk(chunk);
    const story = callOpenAIForStory(prompt);

    const usedWords = chunk.map(row => row[0]).join(', ');
    const rangeInfo = `A${startRow + i * chunkSize}~A${startRow + (i + 1) * chunkSize - 1}`;
    const nextRow = storySheet.getLastRow() + 1;

    storySheet.getRange(nextRow, 1).setValue(new Date());
    storySheet.getRange(nextRow, 2).setValue(`단어: ${usedWords}`);
    storySheet.getRange(nextRow, 3).setValue(`범위: ${rangeInfo}`);
    storySheet.getRange(nextRow, 4).setValue(story);
  }

  ui.alert(`${numChunks}개의 니모닉 이야기 생성 완료!`);
}
function buildPromptFromChunk(chunk) {
  const entries = chunk.map(([word, meaning, root, pun]) =>
    `- ${word} | ${meaning} | ${root} | ${pun}`
  ).join('\n');

  return `
아래 10개 단어를 모두 사용해서, 개연성 있고 유머러스한 짧은 스토리를 써줘.

- 각 단어는 반드시 이야기 속 등장인물, 사건, 배경 등으로 자연스럽게 등장해야 해.
- 각 단어는 **영어 원형, 어근(간단히), 경선식 말장난/소리 연상**을 한두 문장에 짧고 임팩트 있게 녹여줘.
- 한 단어를 두 번 이상 반복하거나 설명을 늘이지 말고, 강렬하고 간결하게 써줘.
- 스토리는 원인→갈등→해결→교훈의 흐름을 갖추고, 상황이 자연스럽게 연결되게 해줘.
- 전체 이야기 자체가 읽기 쉽고 재미있게, 유머와 캐릭터화된 묘사가 살아 있게 해줘.

단어 목록:
${entries}
`;
}


function callOpenAIForStory(prompt) {
  const url = 'https://api.openai.com/v1/chat/completions';
  const payload = {
    model: 'gpt-4.1-mini',
    messages: [{ role: 'user', content: prompt }],
    temperature: 0.7,
    max_tokens: 2048
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${OPENAI_API_KEY}` },
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  return data.choices[0].message.content.trim();
}
