/**
 * 순회 방문 신청 - 공통 상수와 시트 메뉴
 *
 * 파일 구성
 *   Code.gs       상수, 시트 메뉴
 *   Utils.gs      해시, 토큰, 잠금, 로그, 시트 읽기
 *   Setup.gs      초기 설정, 전화번호 해시 변환
 *   Api.gs        웹앱 API (login / getData / toggle / updateExtra)
 *   Status.gs     신청현황, 전도인별 탭 그리기와 자동 갱신
 *   Sync.gs       배정 체크박스 동기화, 대기 표시
 *   Recommend.gs  자동 추천 배정
 *   TestData.gs   테스트 데이터 넣기 / 지우기
 */

var TZ = 'Asia/Seoul';

var TAB = {
  ROSTER: '명단',
  SCHEDULE: '일정',
  APPLY: '신청',
  CONFIG: '설정',
  STATUS: '신청현황',
  RECO: '추천결과',
  PERSON: '전도인별',
  LOG: '로그'
};

var HEAD = {
  ROSTER: ['이름', '전화번호해시', '성별'],
  SCHEDULE: ['칸ID', '날짜', '요일', '항목', '팀', '시간', '장소', '필요인원', '열림여부', '추천성별'],
  APPLY: ['칸ID', '이름', '신청시각', '상태', '차량제공', '식사방식', '대접가능팀수', '메모'],
  PERSON: ['이름', '성별', '신청한 칸', '배정된 칸', '신청 수', '배정 수'],
  LOG: ['시각', '이름', '동작', '칸ID']
};

var ST = { APPLY: '신청', ASSIGN: '배정', WAIT: '대기', CANCEL: '취소' };

var ITEMS = ['오전봉사', '점심식사', '오후봉사', '저녁식사'];

// 신청현황 / 추천결과 탭의 열 위치 (1부터 시작)
var VIEW = {
  CHECK: 1, NAME: 2, GENDER: 3, CAR: 4, MEAL: 5, TEAMS: 6, MEMO: 7,
  TOTAL_APPLY: 8, TOTAL_ASSIGN: 9, TIME: 10, KEY_SLOT: 11, KEY_NAME: 12,
  WIDTH: 12,
  FIRST_ROW: 3,          // 1행: 갱신 시각, 2행: 열 이름
  HEADER_MARK: '#H'      // 블록 머리줄 표시 (숨김 열)
};

var COLOR = {
  HEAD_NORMAL: '#cfe2f3',  // 항목 색이 정해지지 않은 칸의 머리줄
  HEAD_SHORT: '#f6b26b',   // 신청자가 필요 인원보다 적음 (인원 칸에만 칠함)
  HEAD_FULL: '#93c47d',    // 배정이 필요 인원만큼 참 (인원 칸에만 칠함)
  LABEL: '#eeeeee',
  RECO: '#fbd5b5',         // 추천결과 탭에서 새로 추천된 줄
  DATE_BG: '#3c4043',      // 날짜 구분 줄
  DATE_TEXT: '#ffffff',
  TEXT: '#000000'
};

// 항목별 색: head는 블록 머리줄, tint는 그 아래 신청자 줄
var ITEM_COLOR = {
  '오전봉사': { head: '#9fc5e8', tint: '#eaf3fb' },
  '점심식사': { head: '#ffe599', tint: '#fff9e3' },
  '오후봉사': { head: '#b4a7d6', tint: '#f0ecf8' },
  '저녁식사': { head: '#a2c4c9', tint: '#e8f2f3' }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('순회방문 관리')
    .addItem('초기 설정', 'setup')
    .addItem('전화번호 해시 변환', 'convertPhones')
    .addSeparator()
    .addItem('신청현황 새로고침', 'refreshStatus')
    .addItem('자동 갱신 켜기', 'enableAutoRefresh')
    .addItem('자동 갱신 끄기', 'disableAutoRefresh')
    .addItem('배정 동기화 트리거 설치', 'installSyncTrigger')
    .addSeparator()
    .addItem('자동 추천 배정 (미리보기 생성)', 'recommend')
    .addItem('추천 결과 반영', 'applyRecommendation')
    .addItem('체크 안 된 신청을 모두 대기로 표시', 'markWaiting')
    .addSeparator()
    .addItem('테스트 데이터 넣기', 'insertTestData')
    .addItem('테스트 데이터 지우기', 'removeTestData')
    .addToUi();
}
