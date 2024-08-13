import ExcelJS from "exceljs";
import { useCallback } from "react";

const DEFAULT_COL_WIDTH_SIZE = 10;
const MAX_COL_WIDTH_SIZE = 50;
const LENGTH_CORRECTION_RATIO = 1.5;

// 기본 헤더 컬럼 스타일 정의
const setDefaultHeaderCellStyle = (cell: ExcelJS.Cell) => {
  cell.font = {
    bold: true,
  };
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: {
      argb: "E0E0E0",
    },
  };
  cell.alignment = {
    horizontal: "center",
  };
};

// 기본 데이터 컬럼 스타일 정의
const setDefaultDataCellStyle = (cell: ExcelJS.Cell) => {
  cell.alignment = {
    horizontal: "center",
  };
  cell.font = {
    size: 10,
  };
  cell.border = {
    top: {
      style: "thin",
    },
    bottom: {
      style: "thin",
    },
    right: {
      style: "thin",
    },
    left: {
      style: "thin",
    },
  };
};

export type ExcelSheet = {
  sheetName: string;
  data: unknown[];
  titleRow?: {
    title: string;
    mergeCell?: string;
    titleCellStyle?: (cell: ExcelJS.Cell) => void;
  };
  headers?: string[];
  width?: number[];
  headerCellStyle?: (cell: ExcelJS.Cell) => void;
  dataCellStyle?: (cell: ExcelJS.Cell) => void;
};

export type CustomParsingFunction = (
  excelSheet: ExcelSheet[]
) => ExcelJS.Workbook;

type UseExcelDownloadProps = {
  fileName: string;
  noDataLabel?: string;
  customParsingFunction?: CustomParsingFunction;
};

type UseExcelDownloadResult = {
  onClickDownloadExcelFile: (excelSheet: ExcelSheet[]) => void;
};

const useExcelDownload = ({
  fileName,
  noDataLabel,
  customParsingFunction,
}: UseExcelDownloadProps): UseExcelDownloadResult => {
  const downloadExcelFile = useCallback(
    async (workbook: ExcelJS.Workbook, fileName: string) => {
      try {
        const fileData = await workbook.xlsx.writeBuffer();
        const blob = new Blob([fileData], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        const url = window.URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = fileName + ".xlsx";
        anchor.click();
        window.URL.revokeObjectURL(url);
      } catch (err) {
        console.error(err);
      }
    },
    []
  );

  const parsingToWorkbook = useCallback(
    (excelSheetList: ExcelSheet[], workbook: ExcelJS.Workbook) => {
      excelSheetList.forEach((excelSheet) => {
        const sheet = workbook.addWorksheet(excelSheet.sheetName);

        // Title Row
        if (excelSheet.titleRow) {
          const { title, titleCellStyle, mergeCell } = excelSheet.titleRow;
          const titleRow = sheet.addRow([title]);

          sheet.mergeCells(mergeCell ?? "A1:E1");
          titleRow.eachCell((cell) => {
            cell.font = {
              size: 15,
              bold: true,
            };
          });

          if (titleCellStyle && titleCellStyle instanceof Function) {
            titleRow.eachCell((cell) => {
              titleCellStyle(cell);
            });
          }
        }

        // Column Header Row
        if (excelSheet.headers) {
          const headerRow = sheet.addRow(excelSheet.headers);
          headerRow.eachCell((cell, colNumber) => {
            if (!excelSheet?.width) {
              const valueLength = cell?.value?.toString().length ?? 0;
              sheet.getColumn(colNumber).width =
                valueLength < DEFAULT_COL_WIDTH_SIZE
                  ? DEFAULT_COL_WIDTH_SIZE
                  : valueLength * LENGTH_CORRECTION_RATIO;
            }

            setDefaultHeaderCellStyle(cell);
            // apply custom header cell style
            if (excelSheet?.headerCellStyle instanceof Function) {
              excelSheet.headerCellStyle(cell);
            }
          });
        }

        // Data Row
        excelSheet.data.forEach((value) => {
          const row: unknown[] = [];
          if (Array.isArray(value)) {
            row.push(value.toString());
          } else if (value instanceof Object) {
            const rawData = value as Record<string, unknown>;
            Object.keys(rawData).forEach((key) => row.push(rawData[key] ?? ""));
          } else {
            row.push(value);
          }

          const appendRow = sheet.addRow(row);
          appendRow.eachCell((cell, colNumber) => {
            setDefaultDataCellStyle(cell);

            // apply custom data cell style
            if (excelSheet?.dataCellStyle instanceof Function) {
              excelSheet.dataCellStyle(cell);
            }

            // apply cell width
            if (excelSheet?.width) {
              sheet.getColumn(colNumber).width =
                excelSheet.width[colNumber - 1];
            } else if (cell.value) {
              const valueLength = cell.value.toString().length;
              const colWidth =
                sheet.getColumn(colNumber).width ?? DEFAULT_COL_WIDTH_SIZE;

              // Set Default Column Width
              const adjustedWidth =
                colWidth > DEFAULT_COL_WIDTH_SIZE
                  ? colWidth
                  : DEFAULT_COL_WIDTH_SIZE;

              // Set Column width
              sheet.getColumn(colNumber).width =
                adjustedWidth > MAX_COL_WIDTH_SIZE
                  ? MAX_COL_WIDTH_SIZE
                  : adjustedWidth < valueLength * LENGTH_CORRECTION_RATIO
                  ? valueLength * LENGTH_CORRECTION_RATIO
                  : adjustedWidth;
            }
          });
        });
      });
    },
    []
  );

  const onClickDownloadExcelFile = useCallback(
    async (excelSheetList: ExcelSheet[]) => {
      const isEmptyData =
        !excelSheetList.length ||
        excelSheetList.every(({ data }) => !data.length);

      if (isEmptyData) return noDataLabel ?? "데이터가 존재하지 않습니다.";

      const workbook = new ExcelJS.Workbook();

      if (customParsingFunction instanceof Function) {
        const workbook = customParsingFunction(excelSheetList);
        return downloadExcelFile(workbook, fileName);
      }

      parsingToWorkbook(excelSheetList, workbook);
      await downloadExcelFile(workbook, fileName);
    },
    [
      customParsingFunction,
      parsingToWorkbook,
      downloadExcelFile,
      noDataLabel,
      fileName,
    ]
  );

  return { onClickDownloadExcelFile };
};

export default useExcelDownload;
