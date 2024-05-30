import useExcelDownload from "./share/hooks/useExcelDownload";
import type { ExcelSheet } from "./share/hooks/useExcelDownload";

function App() {
  const { onClickDownloadExcelFile } = useExcelDownload({
    fileName: "excel_sample",
  });

  const handleClickExcelDownload = () => {
    const excelSheet: ExcelSheet[] = [
      {
        sheetName: "첫 번째 시트",
        headers: ["나이", "이름", "직업"],
        width: [30, 40, 50],
        headerCellStyle: (cell) => {
          cell.font = {
            size: 24,
          };
        },
        data: [
          {
            age: 24,
            name: "존",
            job: "student",
          },
          {
            age: 55,
            name: "시나",
            job: "professor",
          },
          {
            age: 66,
            name: "포이즌",
            job: "정년퇴직",
          },
        ],
      },
      {
        sheetName: "두 번째 시트",
        titleRow: {
          title: "두 번째 시트의 타이틀 입니다.",
          mergeCell: "A1:B1",
        },
        data: [
          {
            label: "기간",
            value: "1994/10/25 ~ 1994/10/25",
          },
          {
            label: "기념일 명칭",
            value: "내 생일",
          },
        ],
      },
    ];

    onClickDownloadExcelFile(excelSheet);
  };
  return (
    <div>
      <button onClick={handleClickExcelDownload}>엑셀 다운로드</button>
    </div>
  );
}

export default App;
