import ExcelJS from "exceljs";
import {
  DEFAULT_COL_WIDTH_SIZE,
  LENGTH_CORRECTION_RATIO,
  MAX_COL_WIDTH_SIZE,
  REGEX_KOREAN,
} from "../constants";
import { CustomParsingFunction, ExcelSheet } from "../types/excel.type";

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

type ParseData = {
  sheet: ExcelJS.Worksheet;
  excelSheet: ExcelSheet;
  columnWidths?: number[];
};

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
  const downloadExcelFile = async (
    workbook: ExcelJS.Workbook,
    fileName: string
  ) => {
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
  };

  const calculateCellLength = (cellText: string): number => {
    let textWidth = 0;

    for (const char of cellText) {
      if (REGEX_KOREAN.test(char)) {
        textWidth += 2;
      } else {
        textWidth += 1;
      }
    }

    return textWidth * LENGTH_CORRECTION_RATIO;
  };

  const adjustColumnWidths = (
    sheet: ExcelJS.Worksheet,
    columnWidths: number[]
  ) => {
    columnWidths.forEach((width, colIndex) => {
      sheet.getColumn(colIndex + 1).width = width;
    });
  };

  const parseTitleRow = (parseData: ParseData) => {
    const { sheet, excelSheet } = parseData;
    const { title, mergeCell, titleCellStyle } = excelSheet.titleRow!;
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
  };

  const parseHeaderRow = (parseData: ParseData) => {
    const { sheet, excelSheet, columnWidths = [] } = parseData;
    const headerRow = sheet.addRow(excelSheet.headers);
    headerRow.eachCell((cell, colNumber) => {
      setDefaultHeaderCellStyle(cell);

      const cellText = cell?.value?.toString() || "";
      const textWidth = calculateCellLength(cellText);

      if (
        !columnWidths[colNumber - 1] ||
        textWidth > columnWidths[colNumber - 1]
      ) {
        if (textWidth < DEFAULT_COL_WIDTH_SIZE) {
          columnWidths[colNumber - 1] = DEFAULT_COL_WIDTH_SIZE;
        } else if (textWidth > MAX_COL_WIDTH_SIZE) {
          columnWidths[colNumber - 1] = MAX_COL_WIDTH_SIZE;
        } else {
          columnWidths[colNumber - 1] = textWidth;
        }
      }

      // apply custom header cell style
      if (excelSheet?.headerCellStyle instanceof Function) {
        excelSheet.headerCellStyle(cell);
      }
    });
  };

  const parseDataRow = (parseData: ParseData) => {
    const { sheet, excelSheet, columnWidths = [] } = parseData;
    excelSheet.data.forEach((value) => {
      const row: unknown[] = [];
      if (Array.isArray(value)) {
        row.push(value.toString());
      } else if (value instanceof Object) {
        const rawData = value as Record<string, unknown>;
        const orderedKeys = excelSheet.headers ?? Object.keys(rawData);
        orderedKeys.forEach((key) => {
          row.push(rawData[key] ?? "");
        });
      } else {
        row.push(value);
      }

      const appendRow = sheet.addRow(row);
      appendRow.eachCell((cell, colNumber) => {
        setDefaultDataCellStyle(cell);

        const cellValue = cell?.value?.toString() || "";
        let valueLength = 0;

        for (const char of cellValue) {
          if (/[\uac00-\ud7af]/.test(char)) {
            valueLength += 2;
          } else {
            valueLength += 1;
          }
        }

        valueLength *= LENGTH_CORRECTION_RATIO;

        if (
          !columnWidths[colNumber - 1] ||
          valueLength > columnWidths[colNumber - 1]
        ) {
          if (valueLength < DEFAULT_COL_WIDTH_SIZE) {
            columnWidths[colNumber - 1] = DEFAULT_COL_WIDTH_SIZE;
          } else if (valueLength > MAX_COL_WIDTH_SIZE) {
            columnWidths[colNumber - 1] = MAX_COL_WIDTH_SIZE;
          } else {
            columnWidths[colNumber - 1] = valueLength;
          }
        }

        // apply custom data cell style
        if (excelSheet?.dataCellStyle instanceof Function) {
          excelSheet.dataCellStyle(cell);
        }
      });
    });
  };

  const parsingToWorkbook = (
    excelSheetList: ExcelSheet[],
    workbook: ExcelJS.Workbook
  ) => {
    excelSheetList.forEach((excelSheet) => {
      const sheet = workbook.addWorksheet(excelSheet.sheetName);
      const columnWidths: number[] = [];

      // Title Row
      if (excelSheet.titleRow) {
        parseTitleRow({ sheet, excelSheet });
      }

      // Column Header Row
      if (excelSheet.headers) {
        parseHeaderRow({ sheet, excelSheet, columnWidths });
      }

      // Data Row
      if (excelSheet.data) {
        parseDataRow({ sheet, excelSheet, columnWidths });
      }

      adjustColumnWidths(sheet, columnWidths);
    });
  };

  const onClickDownloadExcelFile = async (excelSheetList: ExcelSheet[]) => {
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
  };

  return { onClickDownloadExcelFile };
};

export default useExcelDownload;
