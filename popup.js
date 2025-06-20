function processData(sheet) {
  return sheet
    .map((row) => {
      const date = row['일자'] || '';
      const status = row['상태'] || '';
      const rawTime = row['총 근무시간'] || '';
      const workMinutes = Math.round(parseFloat(rawTime || 0) * 1440);

      let overtimeMinutes = 0, // 초과시간
        deductionMinutes = 0; //차감시간

      if (date.includes('주간 근무시간')) {
        return {}; // 주간 근무시간 행은 계산하지 않음
      }

      if (status === '공휴일') {
        overtimeMinutes = workMinutes;
      } else {
        // 연차
        if (status.includes('8.00h')) {
          deductionMinutes = 0;
        } else if (status.includes('4.00h')) {
          // 반차이며 4시간이내로 근무한 경우
          if (workMinutes < 240) {
            deductionMinutes = 240 - workMinutes;
          }
        } else if (workMinutes < 480) {
          // 평일이며 연차나 반차가 아닌데 8시간 미만으로 근무한 경우
          deductionMinutes = 480 - workMinutes;
        }

        // 평일이면서 9시간 이상 근무했을때
        if (workMinutes > 540) {
          overtimeMinutes = workMinutes - 540;
        }
      }

      return {
        일자: date,
        '총 근무(분)': workMinutes,
        '초과(분)': overtimeMinutes,
        '차감(분)': deductionMinutes,
        상태: status,
      };
    })
    .filter((row) => Object.keys(row).length !== 0);
}

document.getElementById('fileInput').addEventListener('change', (e) => {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];

    const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      range: 5,
      header: ['일자', '업무시작', '업무종료', '총 근무시간', '기본', '연장', '야간', '상태'],
    });

    const resultData = processData(sheet);
    const keys = Object.keys(resultData[0]);

    let firstDate = '',
      lastDate = '',
      totalOvertime = 0,
      totalDeduction = 0;
    const rowsHtml = [];

    // 테이블 헤더
    const headerHtml = `<tr>${keys.map((k) => `<th>${k}</th>`).join('')}</tr>`;

    resultData.forEach((row, index) => {
      if (index === 0) firstDate = row['일자'];
      if (index === resultData.length - 1) lastDate = row['일자'];

      let trStyle = '';
      if (!!row['상태']) {
        trStyle = 'background-color:#ff334b26';
      }

      const rowHtml = `<tr style=${trStyle}>${keys.map((k) => `<td>${row[k] ?? ''}</td>`).join('')}</tr>`;
      rowsHtml.push(rowHtml);

      const overtime = parseInt(row['초과(분)']) || 0;
      const deduction = parseInt(row['차감(분)']) || 0;

      totalOvertime += overtime;
      totalDeduction += deduction;
    });

    const delta = totalOvertime - totalDeduction;
    const days = Math.floor(delta / 540);
    const hours = Math.floor((((delta % 540) + 540) % 540) / 60);

    // 본 화면에 텍스트만 표시
    document.getElementById('summary').innerHTML = `
      <h3>기간: ${firstDate} ~ ${lastDate}</h3>
      <h3>대체휴가: ${days}일 ${hours}시간</h3>
    `;

    // 버튼 활성화
    const btn = document.getElementById('showDetail');
    btn.style.display = 'inline-block';

    btn.onclick = () => {
      const popup = window.open('', '_blank', 'width=500,height=700');
      popup.document.write(`
        <html>
          <head>
            <title>근무 자세히 보기</title>
            <style>
              body { padding: 16px; font-size: 14px;}
              table { border-collapse: collapse; width: 100%; font-size: 14px; }
              th, td { border: 1px solid #aaa; padding: 4px; text-align: left; }
            </style>
          </head>
          <body>
            <h3>${firstDate} ~ ${lastDate}</h2>
            <p>
              <strong>초과:</strong> ${totalOvertime}분, 
              <strong>차감:</strong> ${totalDeduction}분, 
              <strong>총 대체휴가:</strong> ${days}일 ${hours}시간
            </p>
            <table>${headerHtml}${rowsHtml.join('')}</table>
          </body>
        </html>
      `);
      popup.document.close();
    };
  };

  reader.readAsArrayBuffer(file);
});
