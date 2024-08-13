import ExcelJS from "exceljs";

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
