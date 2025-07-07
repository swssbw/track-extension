// 파일 input이 변경되었을 때 엑셀 파일을 읽고 처리하는 메인 핸들러
document.getElementById('fileInput').addEventListener('change', handleFileInput);

function handleFileInput(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = async (evt) => {
    const sheet = readExcel(evt.target.result); // 1. 엑셀 파일 → JSON 시트
    const resultData = processData(sheet); // 2. 공휴일 처리 및 초과/차감 계산

    const keys = Object.keys(resultData[0]).filter((k) => k !== '상태'); // '상태'는 제외
    const { firstDate, lastDate, totalOvertime, totalDeduction } = calculateSummary(resultData);

    renderSummary(firstDate, lastDate, totalOvertime, totalDeduction); // 요약 영역 표시

    // "자세히 보기" 버튼 활성화 및 이벤트 바인딩
    document.getElementById('editDetail').style.display = 'inline-block';
    document.getElementById('editDetail').onclick = () => setupEditPopup(resultData, keys, firstDate, lastDate);
  };

  reader.readAsArrayBuffer(file); // 파일 읽기 시작
}

// 엑셀 파일을 JSON 배열로 변환 (헤더 정의 + 시작 행 설정)
function readExcel(arrayBuffer) {
  const data = new Uint8Array(arrayBuffer);
  const workbook = XLSX.read(data, { type: 'array' });
  const sheetName = workbook.SheetNames[0];

  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    range: 5,
    header: ['일자', '업무시작', '업무종료', '총 근무시간', '기본', '연장', '야간', '상태'],
  });
}

// 주간 요약 제외, 공휴일 자동 처리, 초과/차감 계산 포함된 최종 row 데이터 가공
function processData(sheet) {
  return sheet
    .map((row) => {
      const date = row['일자'] || '';
      let status = row['상태'] || '';
      const rawTime = row['총 근무시간'] || '';
      const workMinutes = Math.round(parseFloat(rawTime || 0) * 1440);

      // "주간 근무시간" 행은 제외
      if (date.includes('주간 근무시간')) return {};

      // 주말이면 자동으로 공휴일 처리
      if (date.includes('토') || date.includes('일')) {
        status = '공휴일';
      }

      const tempRow = {
        일자: date,
        '총 근무(분)': workMinutes,
        상태: status,
      };

      const { overtime, deduction } = calculateWorkInfo(tempRow);

      return {
        ...tempRow,
        '초과(분)': overtime,
        '차감(분)': deduction,
      };
    })
    .filter((row) => Object.keys(row).length !== 0); // 빈 객체 제거
}

// 각 행(row)에 대해 초과근무/차감시간을 계산
function calculateWorkInfo(row) {
  const status = row['상태'] || '';
  const workMinutes = row['총 근무(분)'] || 0;

  let overtime = 0;
  let deduction = 0;

  // 공휴일이면 전 근무시간이 초과로 인정, 차감 없음
  if (status === '공휴일') {
    overtime = workMinutes;
  } else {
    // 연차 8시간
    if (status.includes('8.00h')) {
      deduction = 0;
    }
    // 반차인데 4시간보다 덜 일함
    else if (status.includes('4.00h') && workMinutes < 240) {
      deduction = 240 - workMinutes;
    }
    // 연차가 아니고 8시간 미만 근무
    else if (workMinutes < 480) {
      deduction = 480 - workMinutes;
    }

    // 9시간(540분) 초과한 경우 초과근무로 인정
    if (workMinutes > 540) {
      overtime = workMinutes - 540;
    }
  }

  return { overtime, deduction };
}

// 전체 요약: 총 초과시간, 총 차감시간, 첫날/마지막 날짜 추출
function calculateSummary(data) {
  let totalOvertime = 0,
    totalDeduction = 0,
    firstDate = '',
    lastDate = '';

  data.forEach((row, index) => {
    if (index === 0) firstDate = row['일자'];
    if (index === data.length - 1) lastDate = row['일자'];
    totalOvertime += parseInt(row['초과(분)']) || 0;
    totalDeduction += parseInt(row['차감(분)']) || 0;
  });

  return { totalOvertime, totalDeduction, firstDate, lastDate };
}

// 메인 페이지 상단에 대체휴가 요약 정보 출력
function renderSummary(firstDate, lastDate, totalOvertime, totalDeduction) {
  const delta = totalOvertime - totalDeduction;
  const days = Math.floor(delta / 540);
  const hours = Math.floor((((delta % 540) + 540) % 540) / 60);

  document.getElementById('summary').innerHTML = `
    <h3>기간: ${firstDate} ~ ${lastDate}</h3>
    <h3>대체휴가: ${days}일 ${hours}시간</h3>
  `;
}

// 팝업 창 열고 HTML 삽입, 팝업 로드 후 동작 삽입
function setupEditPopup(resultData, keys, firstDate, lastDate) {
  const popup = window.open('', '_blank', 'width=700,height=800');

  popup.document.write(generatePopupHtml(resultData, keys, firstDate, lastDate));
  popup.onload = () => {
    injectPopupScript(popup, resultData); // 팝업에 기능 삽입
  };

  popup.document.close();
}

// 팝업 HTML 마크업 문자열 생성 (테이블 + 체크박스)
function generatePopupHtml(resultData, keys, firstDate, lastDate) {
  const headerHtml = `<tr>${keys.map((k) => `<th>${k}</th>`).join('')}<th>휴일</th></tr>`;
  const rowsHtml = resultData
    .map((row, idx) => {
      const checked = row['상태'] === '공휴일' ? 'checked' : '';
      const cells = keys.map((k) => `<td>${row[k] ?? ''}</td>`).join('');
      return `<tr data-index="${idx}">${cells}
        <td style="text-align:center"><input type="checkbox" ${checked}></td>
      </tr>`;
    })
    .join('');

  return `
    <html>
      <head>
        <title>자세히 보기</title>
        <style>
          body { padding: 16px; font-size: 14px;}
          table { border-collapse: collapse; width: 100%; font-size: 14px; }
          th, td { border: 1px solid #aaa; padding: 4px; text-align: left; }
          input[type="checkbox"] { transform: scale(1.3); }
        </style>
      </head>
      <body>
        <h3>${firstDate} ~ ${lastDate}</h3>
        <div id="summaryArea"></div>
        <table id="editTable">${headerHtml}${rowsHtml}</table>
      </body>
    </html>
  `;
}

// 팝업 내에서 체크박스 변경 감지 → 상태 업데이트 + 실시간 요약 반영
function injectPopupScript(popup, resultData) {
  const table = popup.document.getElementById('editTable');
  const summary = popup.document.getElementById('summaryArea');

  // 실시간 요약 영역 업데이트 함수
  function updateSummary() {
    let totalOvertime = 0;
    let totalDeduction = 0;

    resultData.forEach((row) => {
      const { overtime, deduction } = calculateWorkInfo(row);
      totalOvertime += overtime;
      totalDeduction += deduction;
    });

    const delta = totalOvertime - totalDeduction;
    const days = Math.floor(delta / 540);
    const hours = Math.floor((((delta % 540) + 540) % 540) / 60);

    summary.innerHTML = `
      <p>
        <strong>초과:</strong> ${totalOvertime}분,
        <strong>차감:</strong> ${totalDeduction}분,
        <strong>총 대체휴가:</strong> ${days}일 ${hours}시간
      </p>
    `;
  }

  // 각 체크박스에 이벤트 바인딩
  table.querySelectorAll('input[type="checkbox"]').forEach((checkbox, idx) => {
    checkbox.addEventListener('change', (e) => {
      resultData[idx]['상태'] = e.target.checked ? '공휴일' : '';

      // 행 배경색 업데이트
      const rowEl = table.querySelectorAll('tr')[idx + 1];
      rowEl.style.backgroundColor = e.target.checked ? '#ff334b26' : '';

      updateSummary(); // 요약 재계산
    });
  });

  // 초기 공휴일 표시 행에 배경색 적용
  table.querySelectorAll('tr').forEach((tr, i) => {
    if (i === 0) return;
    const idx = i - 1;
    if (resultData[idx]['상태'] === '공휴일') {
      tr.style.backgroundColor = '#ff334b26';
    }
  });

  updateSummary(); // 최초 요약 렌더링
}
