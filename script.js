const headerList = [
  "No.",
  "타이틀",
  "상태",
  "발송수",
  "성공수",
  "실패수",
  "발송일시",
  "작성자",
];

const data = [
  [1, "제목1", "완료", 100, 95, 3, "2024-11-23", "홍길동1"],
  [2, "제목2", "실패", 100, 95, 3, "2024-11-24", "홍길동2"],
  [3, "제목3", "완료", 100, 95, 3, "2024-11-25", "홍길동3"],
  [4, "제목4", "실패", 100, 95, 3, "2024-11-26", "홍길동4"],
  [5, "제목5", "완료", 100, 95, 3, "2024-11-27", "홍길동5"],
];

document
  .getElementById("downloadButton")
  .addEventListener("click", excelDownload);

function excelDownload() {
  // 새로운 Workbook 생성
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sheet 1"); // Excel Sheet파일명

  // 헤더 지정
  const columList = headerList.map((key) => {
    return { header: key, key, width: key.length * 5 };
  });
  worksheet.columns = columList;

  // 헤더 스타일 설정
  worksheet.getRow(1).eachCell((cell) => {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFCCCCCC" },
    };
    cell.font = { bold: true };
    cell.border = {
      bottom: { style: "thin" },
    };
  });

  // 행 추가
  data.forEach((row) => {
    worksheet.addRow(row);
  });

  workbook.xlsx.writeBuffer().then(function (buffer) {
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "제목수정 여기서 하십쇼.xlsx"; // Excel 파일명
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  });
}
