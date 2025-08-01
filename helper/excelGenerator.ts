import { Sheet, utils, WorkBook, WorkSheet, writeFile } from "xlsx-js-style";
import { getDayWithSuffix, globalVariables } from "./helper";

export interface ClaimDay {
  day?: number;
  type?: "claim" | "leave";
  desc?: string;
}

export interface ExcelData {
  month: string;
  days: Array<ClaimDay>;
}

interface CellValue {
  value: string | number;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  hCenter?: "left" | "center" | "right";
  vCenter?: "top" | "center" | "bottom";
  borderBottom?: boolean;
  fullBorder?: boolean;
  borderStyle?: "thin" | "medium" | "thick";
  wrap?: boolean;
  fill?: string;
  type?: "s" | "n";
  formula?: string;
  formattedValue?: string;
}

const setCellValue = ({
  value,
  size = 10,
  bold = false,
  italic = false,
  hCenter = "left",
  vCenter = "center",
  borderBottom = false,
  fullBorder = false,
  borderStyle = "thin",
  wrap = false,
  fill = "",
  type = "s",
  formula = "",
  formattedValue = "",
}: CellValue) => {
  const styling: any = {
    font: { name: "Arial Narrow", sz: size, bold: bold, italic: italic },
    alignment: { vertical: vCenter, horizontal: hCenter, wrapText: wrap },
  };

  // #region Border Section
  if (borderBottom) {
    styling.border = {
      bottom: { style: borderStyle, color: { rgb: "000000" } },
    };
  }

  if (fullBorder) {
    styling.border = {
      top: { style: borderStyle, color: { rgb: "000000" } },
      right: { style: borderStyle, color: { rgb: "000000" } },
      bottom: { style: borderStyle, color: { rgb: "000000" } },
      left: { style: borderStyle, color: { rgb: "000000" } },
    };
  }
  // #endregion Border Section

  // #region Fill section
  if (fill) {
    styling.fill = {
      patternType: "solid",
      fgColor: { rgb: fill },
    };
  }
  // #endregion Fill section

  if (formattedValue) {
    styling["numFmt"] = formattedValue;
  }

  const result = {
    v: value,
    t: type,
    s: styling,
  };

  if (formula) {
    result["f"] = formula;
  }

  return result;
};

const getExcelDetails = (month: string, days: Array<ClaimDay>) => {
  const additionalMergeRow: Array<number> = [];

  // This is the main title of the excel
  const title = [
    setCellValue({
      value: "MEAL REIMBURSEMENT FORM",
      size: 16,
      bold: true,
      hCenter: "center",
    }),
  ];

  // Headers with name and details
  const headers = [
    [
      setCellValue({ value: "Name" }),
      setCellValue({ value: globalVariables.FULLNAME, borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "Supervisor" }),
      setCellValue({ value: "" }),
      setCellValue({ value: globalVariables.SUPERVISOR, borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
    ],
    [
      setCellValue({ value: "Position" }),
      setCellValue({ value: globalVariables.POSITION, borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "" }),
      setCellValue({ value: "" }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
    ],
    [
      setCellValue({ value: "Company" }),
      setCellValue({ value: globalVariables.COMPANY, borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "Month:" }),
      setCellValue({ value: "" }),
      setCellValue({ value: month, borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
    ],
  ];

  // Table header
  const tableHeaders = [
    [
      setCellValue({
        value: "Date",
        fullBorder: true,
        hCenter: "center",
        fill: "c0c0c0",
        bold: true,
      }),
      setCellValue({
        value: "Description",
        fullBorder: true,
        hCenter: "center",
        fill: "c0c0c0",
        bold: true,
      }),
      setCellValue({
        value: "",
        fullBorder: true,
        hCenter: "center",
        fill: "c0c0c0",
        bold: true,
      }),
      setCellValue({
        value: "",
        fullBorder: true,
        hCenter: "center",
        fill: "c0c0c0",
        bold: true,
      }),
      setCellValue({
        value: "Location",
        fullBorder: true,
        hCenter: "center",
        fill: "c0c0c0",
        bold: true,
      }),
      setCellValue({
        value: "",
        fullBorder: true,
        hCenter: "center",
        fill: "c0c0c0",
        bold: true,
      }),
      setCellValue({
        value: "TIME",
        fullBorder: true,
        hCenter: "center",
        fill: "c0c0c0",
        bold: true,
      }),
      setCellValue({
        value: "",
        fullBorder: true,
        hCenter: "center",
        fill: "c0c0c0",
        bold: true,
      }),
      setCellValue({
        value: "",
        fullBorder: true,
        hCenter: "center",
        fill: "c0c0c0",
        bold: true,
      }),
      setCellValue({
        value: "Amount\n(RM)",
        fullBorder: true,
        hCenter: "center",
        wrap: true,
        fill: "c0c0c0",
        bold: true,
      }),
    ],
    [
      setCellValue({ value: "", fullBorder: true, fill: "c0c0c0", bold: true }),
      setCellValue({ value: "", fullBorder: true, fill: "c0c0c0", bold: true }),
      setCellValue({ value: "", fullBorder: true, fill: "c0c0c0", bold: true }),
      setCellValue({ value: "", fullBorder: true, fill: "c0c0c0", bold: true }),
      setCellValue({ value: "", fullBorder: true, fill: "c0c0c0", bold: true }),
      setCellValue({ value: "", fullBorder: true, fill: "c0c0c0", bold: true }),
      setCellValue({
        value: "Start",
        fullBorder: true,
        hCenter: "center",
        fill: "c0c0c0",
        bold: true,
      }),
      setCellValue({
        value: "End",
        fullBorder: true,
        hCenter: "center",
        fill: "c0c0c0",
        bold: true,
      }),
      setCellValue({ value: "", fullBorder: true, fill: "c0c0c0", bold: true }),
      setCellValue({ value: "", fullBorder: true, fill: "c0c0c0", bold: true }),
    ],
  ];

  // Each row of the dates with total amount
  const dateRow: Array<Array<any>> = [];
  for (let i = 1; i <= 31; i++) {
    const canClaim = days.find((d: ClaimDay) => d.day == i);

    if (canClaim) {
      const isLeave = canClaim.type == "leave";
      // if (isLeave) {
      //   additionalMergeRow.push(i);
      // }

      dateRow.push([
        setCellValue({
          value: getDayWithSuffix(i),
          fullBorder: true,
          hCenter: "center",
        }),
        setCellValue({
          value: canClaim.desc || "",
          fullBorder: true,
          // hCenter: canClaim.type == "leave" ? "center" : "left",
        }),
        setCellValue({ value: "", fullBorder: true }),
        setCellValue({ value: "", fullBorder: true }),
        setCellValue({
          value: isLeave ? "" : globalVariables.LOCATION,
          fullBorder: true,
          hCenter: "center",
        }),
        setCellValue({ value: "", fullBorder: true }),
        setCellValue({
          value: isLeave ? "" : globalVariables.START_TIME,
          fullBorder: true,
          hCenter: "right",
        }),
        setCellValue({
          value: isLeave ? "" : globalVariables.END_TIME,
          fullBorder: true,
          hCenter: "center",
        }),
        setCellValue({ value: "", fullBorder: true }),
        setCellValue({
          value: isLeave ? "" : globalVariables.AMOUNT,
          fullBorder: true,
          hCenter: "right",
          type: "n",
          formattedValue: "0.00",
        }),
      ]);
    } else {
      const previous = dateRow.push([
        setCellValue({
          value: getDayWithSuffix(i),
          fullBorder: true,
          hCenter: "center",
        }),
        setCellValue({ value: "", fullBorder: true }),
        setCellValue({ value: "", fullBorder: true }),
        setCellValue({ value: "", fullBorder: true }),
        setCellValue({ value: "", fullBorder: true }),
        setCellValue({ value: "", fullBorder: true }),
        setCellValue({ value: "", fullBorder: true }),
        setCellValue({ value: "", fullBorder: true }),
        setCellValue({ value: "", fullBorder: true }),
        setCellValue({ value: "", fullBorder: true }),
      ]);
    }
  }
  dateRow.push([
    setCellValue({ value: "NOTE :", bold: true, italic: true }),
    setCellValue({ value: "" }),
    setCellValue({ value: "" }),
    setCellValue({ value: "" }),
    setCellValue({ value: "" }),
    setCellValue({ value: "" }),
    setCellValue({ value: "Total Amt Claim" }),
    setCellValue({ value: "" }),
    setCellValue({ value: "" }),
    setCellValue({
      value: 0,
      type: "n",
      formula: "SUM(J10:J40)",
      formattedValue: "0.00",
      bold: true,
      fullBorder: true,
      borderStyle: "medium",
      hCenter: "right",
    }),
  ]);

  // The notes at the bottom of the page
  const bottomNotes = [
    [
      setCellValue({
        value:
          "1)  Form to be submitted to HRD on or before 15th of the following month. Eg. Meal allowance for the month of January is to be submitted via e-claim by 15th of the following month",
        wrap: true,
      }),
    ],
    [
      setCellValue({
        value:
          "2) This is only claimable based on the total actual working days in a month and must be physically in Bangsar South office for a minimum of 4 working hours in a day (between 9:00 am and 6:00 pm) excluding lunch hour.",
        wrap: true,
      }),
    ],
    [
      setCellValue({
        value:
          "3) Not applicable for staff who are on leave including half day leave, on medical leave, on emergency leave, during public holidays, on business trip (outstation / overseas), away for training and those who are working from home.",
        wrap: true,
      }),
    ],
  ];

  // Section for the signatures
  const signs = [
    [
      setCellValue({ value: "Claimed by", bold: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "" }),
      setCellValue({ value: "Verified by", bold: true, hCenter: "right" }),
      setCellValue({ value: "" }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
    ],
    [
      setCellValue({ value: "" }),
      setCellValue({ value: "(Staff Signature)", hCenter: "center" }),
      setCellValue({ value: "" }),
      setCellValue({ value: "" }),
      setCellValue({ value: "" }),
      setCellValue({ value: "" }),
      setCellValue({ value: "Team Lead & HOD", hCenter: "center" }),
    ],
    [],
    [],
    [
      setCellValue({ value: "Checked by", bold: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "" }),
      setCellValue({ value: "Approved by", bold: true, hCenter: "right" }),
      setCellValue({ value: "" }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
      setCellValue({ value: "", borderBottom: true }),
    ],
    [
      setCellValue({ value: "" }),
      setCellValue({ value: "PM / PD", hCenter: "center" }),
      setCellValue({ value: "" }),
      setCellValue({ value: "" }),
      setCellValue({ value: "" }),
      setCellValue({ value: "" }),
      setCellValue({ value: "Head of Operations", hCenter: "center" }),
    ],
  ];

  return {
    merge: additionalMergeRow,
    details: [
      title,
      [],
      ...headers,
      [],
      [],
      ...tableHeaders,
      ...dateRow,
      [],
      ...bottomNotes,
      [],
      [],
      ...signs,
    ],
  };
};

const setExcelOptions = (sheet: Sheet, extraMerge: Array<number>) => {
  // Setting all the merge in the excel
  const mergeCells = [
    { s: { c: 1, r: 30 }, e: { c: 3, r: 30 } },
    { s: { c: 4, r: 30 }, e: { c: 5, r: 30 } },
    { s: { c: 7, r: 30 }, e: { c: 8, r: 30 } },
    { s: { c: 0, r: 44 }, e: { c: 9, r: 44 } },
    { s: { c: 0, r: 47 }, e: { c: 1, r: 47 } },
    { s: { c: 4, r: 47 }, e: { c: 5, r: 47 } },
    { s: { c: 1, r: 48 }, e: { c: 2, r: 48 } },
    { s: { c: 6, r: 48 }, e: { c: 9, r: 48 } },
    { s: { c: 4, r: 51 }, e: { c: 5, r: 51 } },
    { s: { c: 0, r: 51 }, e: { c: 1, r: 51 } },
    { s: { c: 1, r: 52 }, e: { c: 2, r: 52 } },
    { s: { c: 6, r: 52 }, e: { c: 9, r: 52 } },
    { s: { c: 1, r: 39 }, e: { c: 3, r: 39 } },
    { s: { c: 4, r: 39 }, e: { c: 5, r: 39 } },
    { s: { c: 7, r: 39 }, e: { c: 8, r: 39 } },
    { s: { c: 6, r: 40 }, e: { c: 8, r: 40 } },
    { s: { c: 0, r: 42 }, e: { c: 9, r: 42 } },
    { s: { c: 0, r: 43 }, e: { c: 9, r: 43 } },
    { s: { c: 1, r: 37 }, e: { c: 3, r: 37 } },
    { s: { c: 4, r: 37 }, e: { c: 5, r: 37 } },
    { s: { c: 7, r: 37 }, e: { c: 8, r: 37 } },
    { s: { c: 1, r: 38 }, e: { c: 3, r: 38 } },
    { s: { c: 4, r: 38 }, e: { c: 5, r: 38 } },
    { s: { c: 7, r: 38 }, e: { c: 8, r: 38 } },
    { s: { c: 1, r: 35 }, e: { c: 3, r: 35 } },
    { s: { c: 4, r: 35 }, e: { c: 5, r: 35 } },
    { s: { c: 7, r: 35 }, e: { c: 8, r: 35 } },
    { s: { c: 1, r: 36 }, e: { c: 3, r: 36 } },
    { s: { c: 4, r: 36 }, e: { c: 5, r: 36 } },
    { s: { c: 7, r: 36 }, e: { c: 8, r: 36 } },
    { s: { c: 1, r: 33 }, e: { c: 3, r: 33 } },
    { s: { c: 4, r: 33 }, e: { c: 5, r: 33 } },
    { s: { c: 7, r: 33 }, e: { c: 8, r: 33 } },
    { s: { c: 1, r: 34 }, e: { c: 3, r: 34 } },
    { s: { c: 4, r: 34 }, e: { c: 5, r: 34 } },
    { s: { c: 7, r: 34 }, e: { c: 8, r: 34 } },
    { s: { c: 1, r: 31 }, e: { c: 3, r: 31 } },
    { s: { c: 4, r: 31 }, e: { c: 5, r: 31 } },
    { s: { c: 7, r: 31 }, e: { c: 8, r: 31 } },
    { s: { c: 1, r: 32 }, e: { c: 3, r: 32 } },
    { s: { c: 4, r: 32 }, e: { c: 5, r: 32 } },
    { s: { c: 7, r: 32 }, e: { c: 8, r: 32 } },
    { s: { c: 1, r: 29 }, e: { c: 3, r: 29 } },
    { s: { c: 4, r: 29 }, e: { c: 5, r: 29 } },
    { s: { c: 7, r: 29 }, e: { c: 8, r: 29 } },
    { s: { c: 1, r: 27 }, e: { c: 3, r: 27 } },
    { s: { c: 4, r: 27 }, e: { c: 5, r: 27 } },
    { s: { c: 7, r: 27 }, e: { c: 8, r: 27 } },
    { s: { c: 1, r: 28 }, e: { c: 3, r: 28 } },
    { s: { c: 4, r: 28 }, e: { c: 5, r: 28 } },
    { s: { c: 7, r: 28 }, e: { c: 8, r: 28 } },
    { s: { c: 1, r: 25 }, e: { c: 3, r: 25 } },
    { s: { c: 4, r: 25 }, e: { c: 5, r: 25 } },
    { s: { c: 7, r: 25 }, e: { c: 8, r: 25 } },
    { s: { c: 1, r: 26 }, e: { c: 3, r: 26 } },
    { s: { c: 4, r: 26 }, e: { c: 5, r: 26 } },
    { s: { c: 7, r: 26 }, e: { c: 8, r: 26 } },
    { s: { c: 1, r: 23 }, e: { c: 3, r: 23 } },
    { s: { c: 4, r: 23 }, e: { c: 5, r: 23 } },
    { s: { c: 7, r: 23 }, e: { c: 8, r: 23 } },
    { s: { c: 1, r: 24 }, e: { c: 3, r: 24 } },
    { s: { c: 4, r: 24 }, e: { c: 5, r: 24 } },
    { s: { c: 7, r: 24 }, e: { c: 8, r: 24 } },
    { s: { c: 1, r: 21 }, e: { c: 3, r: 21 } },
    { s: { c: 4, r: 21 }, e: { c: 5, r: 21 } },
    { s: { c: 7, r: 21 }, e: { c: 8, r: 21 } },
    { s: { c: 1, r: 22 }, e: { c: 3, r: 22 } },
    { s: { c: 4, r: 22 }, e: { c: 5, r: 22 } },
    { s: { c: 7, r: 22 }, e: { c: 8, r: 22 } },
    { s: { c: 1, r: 19 }, e: { c: 3, r: 19 } },
    { s: { c: 4, r: 19 }, e: { c: 5, r: 19 } },
    { s: { c: 7, r: 19 }, e: { c: 8, r: 19 } },
    { s: { c: 1, r: 20 }, e: { c: 3, r: 20 } },
    { s: { c: 4, r: 20 }, e: { c: 5, r: 20 } },
    { s: { c: 7, r: 20 }, e: { c: 8, r: 20 } },
    { s: { c: 1, r: 17 }, e: { c: 3, r: 17 } },
    { s: { c: 4, r: 17 }, e: { c: 5, r: 17 } },
    { s: { c: 7, r: 17 }, e: { c: 8, r: 17 } },
    { s: { c: 1, r: 18 }, e: { c: 3, r: 18 } },
    { s: { c: 4, r: 18 }, e: { c: 5, r: 18 } },
    { s: { c: 7, r: 18 }, e: { c: 8, r: 18 } },
    { s: { c: 1, r: 15 }, e: { c: 3, r: 15 } },
    { s: { c: 4, r: 15 }, e: { c: 5, r: 15 } },
    { s: { c: 7, r: 15 }, e: { c: 8, r: 15 } },
    { s: { c: 1, r: 16 }, e: { c: 3, r: 16 } },
    { s: { c: 4, r: 16 }, e: { c: 5, r: 16 } },
    { s: { c: 7, r: 16 }, e: { c: 8, r: 16 } },
    { s: { c: 1, r: 13 }, e: { c: 3, r: 13 } },
    { s: { c: 4, r: 13 }, e: { c: 5, r: 13 } },
    { s: { c: 7, r: 13 }, e: { c: 8, r: 13 } },
    { s: { c: 1, r: 14 }, e: { c: 3, r: 14 } },
    { s: { c: 4, r: 14 }, e: { c: 5, r: 14 } },
    { s: { c: 7, r: 14 }, e: { c: 8, r: 14 } },
    { s: { c: 1, r: 11 }, e: { c: 3, r: 11 } },
    { s: { c: 4, r: 11 }, e: { c: 5, r: 11 } },
    { s: { c: 7, r: 11 }, e: { c: 8, r: 11 } },
    { s: { c: 1, r: 12 }, e: { c: 3, r: 12 } },
    { s: { c: 4, r: 12 }, e: { c: 5, r: 12 } },
    { s: { c: 7, r: 12 }, e: { c: 8, r: 12 } },
    { s: { c: 1, r: 9 }, e: { c: 3, r: 9 } },
    { s: { c: 4, r: 9 }, e: { c: 5, r: 9 } },
    { s: { c: 7, r: 9 }, e: { c: 8, r: 9 } },
    { s: { c: 1, r: 10 }, e: { c: 3, r: 10 } },
    { s: { c: 4, r: 10 }, e: { c: 5, r: 10 } },
    { s: { c: 7, r: 10 }, e: { c: 8, r: 10 } },
    { s: { c: 0, r: 7 }, e: { c: 0, r: 8 } },
    { s: { c: 1, r: 7 }, e: { c: 3, r: 8 } },
    { s: { c: 4, r: 7 }, e: { c: 5, r: 8 } },
    { s: { c: 6, r: 7 }, e: { c: 8, r: 7 } },
    { s: { c: 9, r: 7 }, e: { c: 9, r: 8 } },
    { s: { c: 7, r: 8 }, e: { c: 8, r: 8 } },
    { s: { c: 0, r: 0 }, e: { c: 9, r: 0 } },
    { s: { c: 4, r: 2 }, e: { c: 5, r: 2 } },
    { s: { c: 4, r: 3 }, e: { c: 5, r: 3 } },
    { s: { c: 4, r: 4 }, e: { c: 5, r: 4 } },
    { s: { c: 1, r: 2 }, e: { c: 3, r: 2 } },
    { s: { c: 6, r: 2 }, e: { c: 9, r: 2 } },
    { s: { c: 1, r: 3 }, e: { c: 3, r: 3 } },
    { s: { c: 1, r: 4 }, e: { c: 3, r: 4 } },
    { s: { c: 6, r: 4 }, e: { c: 9, r: 4 } },
  ];
  extraMerge.forEach((merge) => {
    const rowNum = merge + 8;
    let toRemoveIndex = mergeCells.findIndex(
      (mc) => mc.s.r == rowNum && mc.e.r == rowNum
    );
    while (toRemoveIndex >= 0) {
      mergeCells.splice(toRemoveIndex, 1);
      toRemoveIndex = mergeCells.findIndex(
        (mc) => mc.s.r == rowNum && mc.e.r == rowNum
      );
    }

    mergeCells.push({ s: { c: 1, r: rowNum }, e: { c: 9, r: rowNum } });
  });
  sheet["!merges"] = mergeCells;

  // Setting each row height
  sheet["!rows"] = [
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 21, hpx: 21 },
    { hpt: 21, hpx: 21 },
    { hpt: 21, hpx: 21 },
    { hpt: 15, hpx: 15 },
    { hpt: 5.1, hpx: 5.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 20.1, hpx: 20.1 },
    { hpt: 30.75, hpx: 30.75 },
    { hpt: 33.6, hpx: 33.6 },
    { hpt: 32.7, hpx: 32.7 },
    { hpt: 12.75, hpx: 12.75 },
    { hpt: 12.75, hpx: 12.75 },
    { hpt: 25.2, hpx: 25.2 },
    { hpt: 12.75, hpx: 12.75 },
    { hpt: 12.75, hpx: 12.75 },
    { hpt: 12.75, hpx: 12.75 },
    { hpt: 25.2, hpx: 25.2 },
    { hpt: 12.75, hpx: 12.75 },
    { hpt: 25.2, hpx: 25.2 },
    { hpt: 12.75, hpx: 12.75 },
  ];

  // Setting the margin
  sheet["!margins"] = {
    left: 0.25,
    right: 0.25,
    top: 0.5,
    bottom: 0.25,
    header: 0,
    footer: 0,
  };

  // Setting the width of each columns
  sheet["!cols"] = [
    { wpx: 68 / 1.511 },
    { wpx: 150 / 1.511 },
    { wpx: 150 / 1.511 },
    { wpx: 163 / 1.511 },
    { wpx: 45 / 1.511 },
    { wpx: 45 / 1.511 },
    { wpx: 86 / 1.511 },
    { wpx: 49 / 1.511 },
    { wpx: 41 / 1.511 },
    { wpx: 95 / 1.511 },
  ];
};

export const generateMealExcel = (data: ExcelData) => {
  const workbook: WorkBook = utils.book_new();

  const { merge, details } = getExcelDetails(data.month, data.days);
  const sheet: WorkSheet = utils.aoa_to_sheet(details);

  setExcelOptions(sheet, merge);

  const month = new Date().getMonth();
  utils.book_append_sheet(workbook, sheet, "Testing Sheet");
  const saveLocation = `${globalVariables.SAVE_LOCATION}\\${
    globalVariables.SHORTNAME
  } - ${globalVariables.EXCEL_FILENAME}.${new Date().getFullYear()}-${
    month - 1 < 0 ? 12 : month
  }.xlsx`;
  writeFile(workbook, saveLocation);
  console.log(`Excel generated successfully. Saved to ${saveLocation}`);
};
