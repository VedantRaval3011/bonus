import * as ExcelJS from "exceljs";
import { Employee, MonthlyData, BonusCalculation } from "./types";

function toNumber(cell: ExcelJS.Cell): number {
  const raw = cell.result ?? cell.value; // prefer calculated result
  if (typeof raw === "number") return raw; // number – return as-is
  if (raw instanceof Date) return +raw; // date – epoch millis → number
  const n = Number(String(raw).trim()); // string or others
  return isFinite(n) ? n : 0; // fallback if NaN
}

export class ExcelProcessor {
  // Parse Staff.xlsx - Extract from multiple monthly sheets
  static async parseStaffFile(
    buffer: ArrayBuffer
  ): Promise<{ employees: Employee[]; summary: any }> {
    try {
      const workbook = new ExcelJS.Workbook();
      const employeeMap = new Map<string, Employee>();

      // Validate buffer
      if (!buffer || buffer.byteLength === 0) {
        console.error("Invalid or empty staff file buffer");
        return { employees: [], summary: null };
      }

      await workbook.xlsx.load(buffer);
      console.log(
        "Staff sheet names:",
        workbook.worksheets.map((ws) => ws.name)
      );

      // Process each monthly sheet (sheets ending with "O")
      workbook.worksheets.forEach((worksheet) => {
        const sheetName = worksheet.name;

        // Filter for monthly staff sheets (they end with "O")
        if (!sheetName.includes("-") || !sheetName.endsWith(" O")) {
          console.log(`Skipping non-staff sheet: ${sheetName}`);
          return;
        }

        console.log(`Processing staff sheet: ${sheetName}`);

        // Process each row to get calculated values from formulas
        let processedCount = 0;

        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          // Skip header rows (1-3)
          if (rowNumber <= 3) return;

          const empId = row.getCell(2).value?.toString().trim(); // Column B: EMP. ID
          const dept = row.getCell(3).value?.toString().trim() || "S"; // Column C: DEPT
          const name = row.getCell(5).value?.toString().trim(); // Column E: EMPLOYEE NAME

          // Use Column O (SALARY1) which is the calculated working salary
          const salaryCell = row.getCell(15); // Column O: SALARY1
          let salary = 0;

          // Handle both calculated values and raw numbers
          if (salaryCell.result !== undefined) {
            salary = toNumber(salaryCell); // Remove const declaration
          } else if (salaryCell.value !== undefined) {
            salary = toNumber(salaryCell); // Remove const declaration
          }

          const doj = ExcelProcessor.parseDate(row.getCell(33).value); // Column AG: DOJ

          console.log(
            `Row ${rowNumber}: empId=${empId}, name=${name}, salary=${salary}`
          );

          // Skip rows with invalid data
          if (
            !empId ||
            empId === "0" ||
            !name ||
            name.toLowerCase() === "total" ||
            empId.toLowerCase() === "total" ||
            salary <= 0
          ) {
            console.log(`❌ Skipping row ${rowNumber}: invalid data`);
            return;
          }

          console.log(
            `✅ Valid staff employee: ID=${empId}, Name=${name}, Dept=${dept}, Salary=${salary}`
          );
          processedCount++;

          if (!employeeMap.has(empId)) {
            employeeMap.set(empId, {
              empId,
              name,
              department: dept,
              doj,
              salary: 0,
              monthlyData: [],
            });
          }

          const employee = employeeMap.get(empId)!;
          if (!employee.monthlyData) {
            employee.monthlyData = [];
          }
          employee.monthlyData.push({
            month: sheetName,
            salary: salary,
          });
        });

        console.log(
          `Processed ${processedCount} staff employees from sheet: ${sheetName}`
        );
      });

      const result = Array.from(employeeMap.values());
      console.log(`✅ Total parsed staff employees: ${result.length}`);
      const staffSummary = ExcelProcessor.calculateStaffSummary(result);
      console.log("Staff Summary:", staffSummary);

      return { employees: result, summary: staffSummary };
    } catch (error) {
      console.error("Error parsing staff file:", error);
      return { employees: [], summary: null };
    }
  }

  static async parseWorkerFile(
    buffer: ArrayBuffer
  ): Promise<{ employees: Employee[]; summary: any }> {
    try {
      const workbook = new ExcelJS.Workbook();
      const employeeMap = new Map<string, Employee>();

      // Validate buffer
      if (!buffer || buffer.byteLength === 0) {
        console.error("Invalid or empty worker file buffer");
        return { employees: [], summary: null };
      }

      await workbook.xlsx.load(buffer);
      console.log(
        "Worker sheet names:",
        workbook.worksheets.map((ws) => ws.name)
      );

      // Process each monthly sheet (sheets ending with "W")
      workbook.worksheets.forEach((worksheet) => {
        const sheetName = worksheet.name;

        // Filter for monthly worker sheets (they end with "W")
        if (!sheetName.includes("-") || !sheetName.endsWith(" W")) {
          console.log(`Skipping non-worker sheet: ${sheetName}`);
          return;
        }

        console.log(`Processing worker sheet: ${sheetName}`);

        let processedCount = 0;

        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          // Skip header rows (1-2)
          if (rowNumber <= 2) return;

          const empId = row.getCell(2).value?.toString().trim(); // Column B: EMP. ID
          const dept = row.getCell(3).value?.toString().trim() || "W"; // Column C: DEPT
          const name = row.getCell(4).value?.toString().trim(); // Column D: EMPLOYEE NAME

          // Use Column I (Salary1) which is the calculated monthly salary
          const salaryCell = row.getCell(9); // Column I: Salary1
          let salary = 0;

          // Handle both calculated values and raw numbers - FIX: Remove const declarations
          if (salaryCell.result !== undefined) {
            salary = toNumber(salaryCell); // ✅ Fixed - assign to outer variable
          } else if (salaryCell.value !== undefined) {
            salary = toNumber(salaryCell); // ✅ This was already correct
          }

          // DOJ is in Column Z (26) based on Excel structure
          const doj = ExcelProcessor.parseDate(row.getCell(26).value); // Column Z: DOJ

          console.log(
            `Row ${rowNumber}: empId=${empId}, name=${name}, salary=${salary}`
          );

          // Skip rows with invalid data
          if (
            !empId ||
            empId === "0" ||
            !name ||
            name.toLowerCase() === "total" ||
            empId.toLowerCase() === "total" ||
            salary <= 0
          ) {
            console.log(`❌ Skipping row ${rowNumber}: invalid data`);
            return;
          }

          console.log(
            `✅ Valid worker employee: ID=${empId}, Name=${name}, Dept=${dept}, Salary=${salary}`
          );
          processedCount++;

          if (!employeeMap.has(empId)) {
            employeeMap.set(empId, {
              empId,
              name,
              department: dept,
              doj,
              salary: 0,
              monthlyData: [],
            });
          }

          const employee = employeeMap.get(empId)!;
          if (!employee.monthlyData) {
            employee.monthlyData = [];
          }
          employee.monthlyData.push({
            month: sheetName,
            salary: salary,
          });
        });

        console.log(
          `Processed ${processedCount} worker employees from sheet: ${sheetName}`
        );
      });

      const result = Array.from(employeeMap.values());
      console.log(`✅ Total parsed worker employees: ${result.length}`);

      // Calculate worker summary
      const workerSummary = ExcelProcessor.calculateWorkerSummary(result);
      console.log("Worker Summary:", workerSummary);

      return { employees: result, summary: workerSummary };
    } catch (error) {
      console.error("Error parsing worker file:", error);
      return { employees: [], summary: null };
    }
  }

  // Parse HR Comparison File
  static async parseHRComparisonFile(
    buffer: ArrayBuffer
  ): Promise<Map<string, { bonus: number; department: string }>> {
    try {
      const workbook = new ExcelJS.Workbook();
      const bonusMap = new Map<string, { bonus: number; department: string }>();

      // Validate buffer
      if (!buffer || buffer.byteLength === 0) {
        console.error("Invalid or empty HR comparison file buffer");
        return bonusMap;
      }

      await workbook.xlsx.load(buffer);
      console.log(
        "HR Comparison sheet names:",
        workbook.worksheets.map((ws) => ws.name)
      );

      // Validate that workbook has worksheets
      if (!workbook.worksheets || workbook.worksheets.length === 0) {
        console.log("No worksheets found in HR comparison file");
        return bonusMap;
      }

      // Parse all department sheets
      workbook.worksheets.forEach((worksheet) => {
        const sheetName = worksheet.name;

        console.log(`Processing HR comparison sheet: ${sheetName}`);

        // Convert worksheet to array format
        const data: any[][] = [];
        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
          const rowData: any[] = [];
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            rowData[colNumber - 1] = cell.value;
          });
          data[rowNumber - 1] = rowData;
        });

        // Find the header row
        let headerRowIndex = -1;
        for (let i = 0; i < Math.min(data.length, 10); i++) {
          const row = data[i];
          if (row && (row[1] === "EMP Code" || row[1] === "EMP. ID")) {
            headerRowIndex = i;
            break;
          }
        }

        if (headerRowIndex === -1) return;

        // Process data rows
        for (let i = headerRowIndex + 1; i < data.length; i++) {
          const row = data[i];
          if (!row || !row[1] || !row[3]) continue; // Skip if no EMP Code or Name

          const empId = row[1]?.toString().trim();
          const department = row[2]?.toString().trim() || sheetName.charAt(0);

          // Find Final RTGS column (varies by sheet structure)
          let finalRTGS = 0;
          // Check common Final RTGS column positions
          const possibleColumns = [18, 19, 20, 21, 22, 23]; // Columns S, T, U, V, W, X
          for (const col of possibleColumns) {
            if (row[col] && parseFloat(row[col]) > 0) {
              finalRTGS = parseFloat(row[col]);
              break;
            }
          }

          if (empId && finalRTGS > 0) {
            bonusMap.set(empId, {
              bonus: finalRTGS,
              department: department,
            });
          }
        }
      });

      console.log(`Parsed ${bonusMap.size} HR bonus records`);
      return bonusMap;
    } catch (error) {
      console.error("Error parsing HR comparison file:", error);
      return new Map();
    }
  }

  // Generate Final Bonus Excel with department-wise separation and formulas
  static async generateFinalBonusExcel(
    calculations: BonusCalculation[],
    staffSummary?: any,
    workerSummary?: any
  ): Promise<ArrayBuffer> {
    try {
      const workbook = new ExcelJS.Workbook();

      // Validate input
      if (!Array.isArray(calculations) || calculations.length === 0) {
        console.error("Invalid or empty calculations array");
        // Create empty workbook with message
        const worksheet = workbook.addWorksheet("Error");
        worksheet.getCell("A1").value =
          "No data available for bonus calculations";
        return await workbook.xlsx.writeBuffer();
      }

      // Group by departments
      const departments = this.groupByDepartments(calculations);
      console.log("Departments found:", Object.keys(departments));

      // Generate separate sheets for each department
      for (const dept of Object.keys(departments)) {
        const deptCalculations = departments[dept];
        const isStaff =
          dept === "S" ||
          dept === "Sci Prec-" ||
          dept === "NRTM" ||
          dept === "Sci Prec Life.-";

        let sheetName = dept === "S" ? "Staff" : dept === "W" ? "Worker" : dept;

        if (isStaff) {
          await this.generateStaffSheet(
            workbook,
            sheetName,
            deptCalculations,
            dept
          );
        } else {
          await this.generateWorkerSheet(
            workbook,
            sheetName,
            deptCalculations,
            dept
          );
        }
      }

      // Generate Summary Sheet
      await this.generateSummarySheet(workbook, departments);

      // Generate Salary Summary Sheet
      if (staffSummary && workerSummary) {
        await this.generateSalarySummarySheet(
          workbook,
          staffSummary,
          workerSummary
        );
      }

      return await workbook.xlsx.writeBuffer();

      return await workbook.xlsx.writeBuffer();
    } catch (error) {
      console.error("Error generating final bonus excel:", error);
      throw error;
    }
  }

  private static groupByDepartments(calculations: BonusCalculation[]): {
    [key: string]: BonusCalculation[];
  } {
    const departments: { [key: string]: BonusCalculation[] } = {};

    // Validate calculations array
    if (!Array.isArray(calculations)) {
      console.error("Calculations is not an array:", typeof calculations);
      return departments;
    }

    calculations.forEach((calc) => {
      // Debug logging
      console.log(
        `Grouping employee ${calc.empId} - Department: "${calc.department}"`
      );

      let dept = calc.department?.trim();

      // Handle empty or undefined department
      if (!dept || dept === "" || dept === "null" || dept === "undefined") {
        // Try to determine department from employee ID pattern
        const empIdNum = parseInt(calc.empId);
        if (!isNaN(empIdNum) && empIdNum <= 500) {
          dept = "S"; // Staff
        } else {
          dept = "W"; // Worker
        }
      }

      if (!departments[dept]) {
        departments[dept] = [];
      }
      departments[dept].push(calc);
    });

    console.log("Final departments:", Object.keys(departments));
    return departments;
  }

  private static async generateStaffSheet(
    workbook: ExcelJS.Workbook,
    sheetName: string,
    calculations: BonusCalculation[],
    dept: string
  ): Promise<void> {
    const worksheet = workbook.addWorksheet(sheetName);

    const title = `DIWALI BONUS LIST FROM NOVEMBER-2024 TO OCTOBER-2025 - INDIANA ${sheetName.toUpperCase()}`;

    const headers = [
      "Sr.No.",
      "EMP Code",
      "Deptt.",
      "EMP. NAME",
      "%",
      "NOV-24",
      "DEC-24",
      "JAN-25",
      "FEB-25",
      "MAR-25",
      "APR-25",
      "MAY-25",
      "JUN-25",
      "JUL-25",
      "AUG-25",
      "SEP-25",
      "OCT-25",
      "GROSS SAL.",
      "GROSS 02",
      "Register",
      "Actual",
      "Un Paid",
      "Final RTGS",
      "Reim.",
    ];

    // Add title
    worksheet.mergeCells("A1:X1");
    const titleCell = worksheet.getCell("A1");
    titleCell.value = title;
    titleCell.font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
    titleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF2E7D32" },
    };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };

    // Add empty row
    worksheet.getRow(2).height = 15;

    // Add headers
    const headerRow = worksheet.getRow(3);
    headers.forEach((header, index) => {
      const cell = headerRow.getCell(index + 1);
      cell.value = header;
      cell.font = { bold: true, size: 12, color: { argb: "FF000000" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFB0BEC5" },
      };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });

    let rowIndex = 4; // Start from row 4

    calculations.forEach((calc, index) => {
      const currentRow = worksheet.getRow(rowIndex + index);
      currentRow.height = 25;

      // Get 12 months of data
      const monthlyValues: (number | null)[] = new Array(12).fill(null);
      calc.monthlyData.forEach((month, idx) => {
        if (idx < 12 && month.salary > 0) {
          monthlyValues[idx] = Math.round(month.salary);
        }
      });

      // Basic data (columns A-Q)
      const dataValues = [
        index + 1, // A: Sr.No.
        calc.empId, // B: EMP Code
        calc.department || dept, // C: Deptt.
        calc.name, // D: EMP. NAME
        calc.bonusPercent, // E: %
        ...monthlyValues, // F-Q: Monthly salaries
      ];

      dataValues.forEach((value, colIndex) => {
        const cell = currentRow.getCell(colIndex + 1);
        cell.value = value;

        // Apply styling based on column
        if (colIndex === 3) {
          // Name column - left aligned
          cell.alignment = { horizontal: "left", vertical: "middle" };
        } else if (colIndex >= 4) {
          // Numeric columns - right aligned
          cell.alignment = { horizontal: "right", vertical: "middle" };
          if (colIndex >= 5 && value !== null) {
            // Monthly salary columns
            cell.numFmt = "#,##0";
          }
        } else {
          cell.alignment = { horizontal: "center", vertical: "middle" };
        }

        cell.border = {
          top: { style: "thin", color: { argb: "FFCCCCCC" } },
          bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
          left: { style: "thin", color: { argb: "FFCCCCCC" } },
          right: { style: "thin", color: { argb: "FFCCCCCC" } },
        };
      });

      // Add formulas for calculated columns
      const rowNum = rowIndex + index;

      // R: GROSS SAL. = ROUND(SUM(F:Q), 0)
      const grossSalCell = currentRow.getCell(18); // Column R
      grossSalCell.value = { formula: `ROUND(SUM(F${rowNum}:Q${rowNum}),0)` };
      grossSalCell.numFmt = "#,##0";
      grossSalCell.alignment = { horizontal: "right", vertical: "middle" };
      grossSalCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // S: GROSS 02 = ROUND(IF(E=8.33, R, IF(E>8.33, R*0.6, "")), 0)
      const gross02Cell = currentRow.getCell(19); // Column S
      gross02Cell.value = {
        formula: `ROUND(IF(E${rowNum}=8.33,R${rowNum},IF(E${rowNum}>8.33,R${rowNum}*0.6,"")),0)`,
      };
      gross02Cell.numFmt = "#,##0";
      gross02Cell.alignment = { horizontal: "right", vertical: "middle" };
      gross02Cell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // T: Register = ROUND(R*E/100, 0)
      const registerCell = currentRow.getCell(20); // Column T
      registerCell.value = { formula: `ROUND(R${rowNum}*E${rowNum}/100,0)` };
      registerCell.numFmt = "#,##0";
      registerCell.alignment = { horizontal: "right", vertical: "middle" };
      registerCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // U: Actual = ROUND(IF(E=8.33, R*E/100, IF(E>8.33, S*E/100, "")), 0)
      const actualCell = currentRow.getCell(21); // Column U
      actualCell.value = {
        formula: `ROUND(IF(E${rowNum}=8.33,R${rowNum}*E${rowNum}/100,IF(E${rowNum}>8.33,S${rowNum}*E${rowNum}/100,"")),0)`,
      };
      actualCell.numFmt = "#,##0";
      actualCell.alignment = { horizontal: "right", vertical: "middle" };
      actualCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // V: Un Paid = 0
      const unPaidCell = currentRow.getCell(22); // Column V
      unPaidCell.value = 0;
      unPaidCell.numFmt = "#,##0";
      unPaidCell.alignment = { horizontal: "right", vertical: "middle" };
      unPaidCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // W: Final RTGS = ROUND(IF(E=8.33, T, IF(E>8.33, T*0.6, "")), 0)
      const finalRTGSCell = currentRow.getCell(23); // Column W
      finalRTGSCell.value = {
        formula: `ROUND(IF(E${rowNum}=8.33,T${rowNum},IF(E${rowNum}>8.33,T${rowNum}*0.6,"")),0)`,
      };
      finalRTGSCell.numFmt = "#,##0";
      finalRTGSCell.alignment = { horizontal: "right", vertical: "middle" };
      finalRTGSCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // X: Reim. = 0
      const reimCell = currentRow.getCell(24); // Column X
      reimCell.value = 0;
      reimCell.numFmt = "#,##0";
      reimCell.alignment = { horizontal: "right", vertical: "middle" };
      reimCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };
    });

    // Add totals row
    const totalsRowIndex = rowIndex + calculations.length;
    const totalsRow = worksheet.getRow(totalsRowIndex);
    totalsRow.height = 35;

    // Add "GRAND TOTAL" label
    const grandTotalCell = totalsRow.getCell(1);
    grandTotalCell.value = "GRAND TOTAL";
    grandTotalCell.font = { bold: true, size: 12, color: { argb: "FF000000" } };
    grandTotalCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFF59D" },
    };
    grandTotalCell.alignment = { horizontal: "center", vertical: "middle" };
    grandTotalCell.border = {
      top: { style: "medium", color: { argb: "FF000000" } },
      bottom: { style: "medium", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };

    // Add total formulas
    const dataStartRow = 4;
    const dataEndRow = rowIndex + calculations.length - 1;

    // Columns with totals: R, S, T, U, V, W, X
    const totalColumns = [18, 19, 20, 21, 22, 23, 24]; // R through X
    totalColumns.forEach((colIndex) => {
      const cell = totalsRow.getCell(colIndex);
      const colLetter = String.fromCharCode(64 + colIndex); // Convert to column letter
      cell.value = {
        formula: `SUM(${colLetter}${dataStartRow}:${colLetter}${dataEndRow})`,
      };
      cell.font = { bold: true, size: 12, color: { argb: "FF000000" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFF59D" },
      };
      cell.alignment = { horizontal: "right", vertical: "middle" };
      cell.numFmt = "#,##0";
      cell.border = {
        top: { style: "medium", color: { argb: "FF000000" } },
        bottom: { style: "medium", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });

    // Set column widths
    const colWidths = [
      60, 80, 80, 200, 60, 90, 90, 90, 90, 90, 90, 90, 90, 90, 90, 90, 90, 110,
      100, 100, 100, 90, 110, 80,
    ];
    colWidths.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width / 7; // ExcelJS uses character units
    });

    // Set row heights
    worksheet.getRow(1).height = 35; // Title row
    worksheet.getRow(2).height = 15; // Empty row
    worksheet.getRow(3).height = 30; // Header row
  }

  private static async generateWorkerSheet(
    workbook: ExcelJS.Workbook,
    sheetName: string,
    calculations: BonusCalculation[],
    dept: string
  ): Promise<void> {
    const worksheet = workbook.addWorksheet(sheetName);

    const title = `DIWALI BONUS LIST FROM NOVEMBER-2024 TO OCTOBER-2025 (2024-2025) INDIANA BOYS ${sheetName.toUpperCase()}`;

    const headers = [
      "Sr.No.",
      "EMP Code",
      "Deptt.",
      "EMP. NAME",
      "%",
      "NOV-24",
      "DEC-24",
      "JAN-25",
      "FEB-25",
      "MAR-25",
      "APR-25",
      "MAY-25",
      "JUN-25",
      "JUL-25",
      "AUG-25",
      "SEP-25",
      "Salary12",
      "Gross",
      "Register",
      "Due VC",
      "Final RTGS",
    ];

    // Add title
    worksheet.mergeCells("A1:U1");
    const titleCell = worksheet.getCell("A1");
    titleCell.value = title;
    titleCell.font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
    titleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF2E7D32" },
    };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };

    // Add empty row
    worksheet.getRow(2).height = 15;

    // Add headers
    const headerRow = worksheet.getRow(3);
    headers.forEach((header, index) => {
      const cell = headerRow.getCell(index + 1);
      cell.value = header;
      cell.font = { bold: true, size: 12, color: { argb: "FF000000" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFB0BEC5" },
      };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });

    let rowIndex = 4; // Start from row 4

    calculations.forEach((calc, index) => {
      const currentRow = worksheet.getRow(rowIndex + index);
      currentRow.height = 25;

      // Get 11 months of data for workers
      const monthlyValues: (number | null)[] = new Array(11).fill(null);
      calc.monthlyData.forEach((month, idx) => {
        if (idx < 11 && month.salary > 0) {
          monthlyValues[idx] = Math.round(month.salary);
        }
      });

      // Basic data (columns A-P)
      const dataValues = [
        index + 1, // A: Sr.No.
        calc.empId, // B: EMP Code
        calc.department || dept, // C: Deptt.
        calc.name, // D: EMP. NAME
        calc.bonusPercent, // E: %
        ...monthlyValues, // F-P: Monthly salaries (11 months)
      ];

      dataValues.forEach((value, colIndex) => {
        const cell = currentRow.getCell(colIndex + 1);
        cell.value = value;

        // Apply styling based on column
        if (colIndex === 3) {
          // Name column - left aligned
          cell.alignment = { horizontal: "left", vertical: "middle" };
        } else if (colIndex >= 4) {
          // Numeric columns - right aligned
          cell.alignment = { horizontal: "right", vertical: "middle" };
          if (colIndex >= 5 && value !== null) {
            // Monthly salary columns
            cell.numFmt = "#,##0";
          }
        } else {
          cell.alignment = { horizontal: "center", vertical: "middle" };
        }

        cell.border = {
          top: { style: "thin", color: { argb: "FFCCCCCC" } },
          bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
          left: { style: "thin", color: { argb: "FFCCCCCC" } },
          right: { style: "thin", color: { argb: "FFCCCCCC" } },
        };
      });

      // Add formulas for calculated columns
      const rowNum = rowIndex + index;

      // Q: Salary12 = ROUND(AVERAGE(F:P), 0)
      const salary12Cell = currentRow.getCell(17); // Column Q
      salary12Cell.value = {
        formula: `ROUND(AVERAGE(F${rowNum}:P${rowNum}),0)`,
      };
      salary12Cell.numFmt = "#,##0";
      salary12Cell.alignment = { horizontal: "right", vertical: "middle" };
      salary12Cell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // R: Gross = ROUND(SUM(F:Q), 0)
      const grossCell = currentRow.getCell(18); // Column R
      grossCell.value = { formula: `ROUND(SUM(F${rowNum}:Q${rowNum}),0)` };
      grossCell.numFmt = "#,##0";
      grossCell.alignment = { horizontal: "right", vertical: "middle" };
      grossCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // S: Register = ROUND(R*8.33%, 0)
      const registerCell = currentRow.getCell(19); // Column S
      registerCell.value = { formula: `ROUND(R${rowNum}*8.33%,0)` };
      registerCell.numFmt = "#,##0";
      registerCell.alignment = { horizontal: "right", vertical: "middle" };
      registerCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // T: Due VC = 0
      const dueVCCell = currentRow.getCell(20); // Column T
      dueVCCell.value = 0;
      dueVCCell.numFmt = "#,##0";
      dueVCCell.alignment = { horizontal: "right", vertical: "middle" };
      dueVCCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // U: Final RTGS = ROUND(S-T, 0)
      const finalRTGSCell = currentRow.getCell(21); // Column U
      finalRTGSCell.value = { formula: `ROUND(S${rowNum}-T${rowNum},0)` };
      finalRTGSCell.numFmt = "#,##0";
      finalRTGSCell.alignment = { horizontal: "right", vertical: "middle" };
      finalRTGSCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };
    });

    // Add totals row
    const totalsRowIndex = rowIndex + calculations.length;
    const totalsRow = worksheet.getRow(totalsRowIndex);
    totalsRow.height = 35;

    // Add "GRAND TOTAL" label
    const grandTotalCell = totalsRow.getCell(1);
    grandTotalCell.value = "GRAND TOTAL";
    grandTotalCell.font = { bold: true, size: 12, color: { argb: "FF000000" } };
    grandTotalCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFF59D" },
    };
    grandTotalCell.alignment = { horizontal: "center", vertical: "middle" };
    grandTotalCell.border = {
      top: { style: "medium", color: { argb: "FF000000" } },
      bottom: { style: "medium", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };

    // Add total formulas for workers
    const dataStartRow = 4;
    const dataEndRow = rowIndex + calculations.length - 1;

    // Columns with totals: R, S, U (18, 19, 21)
    const totalColumns = [18, 19, 21]; // R, S, U
    totalColumns.forEach((colIndex) => {
      const cell = totalsRow.getCell(colIndex);
      const colLetter = String.fromCharCode(64 + colIndex); // Convert to column letter
      cell.value = {
        formula: `SUM(${colLetter}${dataStartRow}:${colLetter}${dataEndRow})`,
      };
      cell.font = { bold: true, size: 12, color: { argb: "FF000000" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFF59D" },
      };
      cell.alignment = { horizontal: "right", vertical: "middle" };
      cell.numFmt = "#,##0";
      cell.border = {
        top: { style: "medium", color: { argb: "FF000000" } },
        bottom: { style: "medium", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });

    // Set column widths for workers
    const colWidths = [
      60, 80, 80, 200, 60, 90, 90, 90, 90, 90, 90, 90, 90, 90, 90, 90, 100, 110,
      100, 90, 110,
    ];
    colWidths.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width / 7; // ExcelJS uses character units
    });

    // Set row heights
    worksheet.getRow(1).height = 35; // Title row
    worksheet.getRow(2).height = 15; // Empty row
    worksheet.getRow(3).height = 30; // Header row
  }

  private static async generateSummarySheet(
    workbook: ExcelJS.Workbook,
    departments: { [key: string]: BonusCalculation[] }
  ): Promise<void> {
    const worksheet = workbook.addWorksheet("Summary");

    const title = "DEPARTMENT-WISE BONUS SUMMARY - DIWALI 2024-25";

    const headers = [
      "Department",
      "Employees",
      "Total Gross Salary",
      "Total Bonus",
      "Average Bonus",
    ];

    // Add title
    worksheet.mergeCells("A1:E1");
    const titleCell = worksheet.getCell("A1");
    titleCell.value = title;
    titleCell.font = { bold: true, size: 18, color: { argb: "FFFFFFFF" } };
    titleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1565C0" },
    };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };

    // Add empty row
    worksheet.getRow(2).height = 15;

    // Add headers
    const headerRow = worksheet.getRow(3);
    headers.forEach((header, index) => {
      const cell = headerRow.getCell(index + 1);
      cell.value = header;
      cell.font = { bold: true, size: 14, color: { argb: "FF000000" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF90CAF9" },
      };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "medium", color: { argb: "FF000000" } },
        bottom: { style: "medium", color: { argb: "FF000000" } },
        left: { style: "medium", color: { argb: "FF000000" } },
        right: { style: "medium", color: { argb: "FF000000" } },
      };
    });

    let grandTotalEmployees = 0;
    let grandTotalSalary = 0;
    let grandTotalBonus = 0;
    let rowIndex = 4;

    Object.keys(departments).forEach((dept, index) => {
      const calculations = departments[dept];
      const totalSalary = calculations.reduce(
        (sum, calc) => sum + calc.totalGrossSalary,
        0
      );
      const totalBonus = calculations.reduce(
        (sum, calc) => sum + calc.finalBonus,
        0
      );
      const avgBonus =
        calculations.length > 0 ? totalBonus / calculations.length : 0;

      grandTotalEmployees += calculations.length;
      grandTotalSalary += totalSalary;
      grandTotalBonus += totalBonus;

      const currentRow = worksheet.getRow(rowIndex + index);
      currentRow.height = 30;

      const rowData = [
        dept,
        calculations.length,
        Math.round(totalSalary * 100) / 100,
        Math.round(totalBonus * 100) / 100,
        Math.round(avgBonus * 100) / 100,
      ];

      rowData.forEach((value, colIndex) => {
        const cell = currentRow.getCell(colIndex + 1);
        cell.value = value;

        if (colIndex === 0) {
          // Department name - left aligned
          cell.alignment = { horizontal: "left", vertical: "middle" };
        } else if (colIndex >= 2) {
          // Numeric columns
          cell.alignment = { horizontal: "right", vertical: "middle" };
          cell.numFmt = "#,##0.00";
        } else {
          // Employee count
          cell.alignment = { horizontal: "center", vertical: "middle" };
        }

        cell.border = {
          top: { style: "thin", color: { argb: "FFBBBBBB" } },
          bottom: { style: "thin", color: { argb: "FFBBBBBB" } },
          left: { style: "thin", color: { argb: "FFBBBBBB" } },
          right: { style: "thin", color: { argb: "FFBBBBBB" } },
        };
      });
    });

    // Add empty row
    const emptyRowIndex = rowIndex + Object.keys(departments).length;
    worksheet.getRow(emptyRowIndex).height = 15;

    // Add grand totals
    const grandTotalRowIndex = emptyRowIndex + 1;
    const grandTotalRow = worksheet.getRow(grandTotalRowIndex);
    grandTotalRow.height = 40;

    const grandTotalData = [
      "GRAND TOTAL",
      grandTotalEmployees,
      Math.round(grandTotalSalary * 100) / 100,
      Math.round(grandTotalBonus * 100) / 100,
      Math.round((grandTotalBonus / grandTotalEmployees) * 100) / 100,
    ];

    grandTotalData.forEach((value, colIndex) => {
      const cell = grandTotalRow.getCell(colIndex + 1);
      cell.value = value;
      cell.font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE65100" },
      };

      if (colIndex === 0) {
        cell.alignment = { horizontal: "center", vertical: "middle" };
      } else if (colIndex >= 2) {
        cell.alignment = { horizontal: "right", vertical: "middle" };
        cell.numFmt = "#,##0.00";
      } else {
        cell.alignment = { horizontal: "center", vertical: "middle" };
      }

      cell.border = {
        top: { style: "double", color: { argb: "FF000000" } },
        bottom: { style: "double", color: { argb: "FF000000" } },
        left: { style: "medium", color: { argb: "FF000000" } },
        right: { style: "medium", color: { argb: "FF000000" } },
      };
    });

    // Set column widths for summary
    const summaryColWidths = [150, 100, 150, 130, 130];
    summaryColWidths.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width / 7;
    });

    // Set row heights
    worksheet.getRow(1).height = 40; // Title row
    worksheet.getRow(2).height = 15; // Empty row
    worksheet.getRow(3).height = 35; // Header row
  }

  static async generateComparisonReport(
    comparisons: any[]
  ): Promise<ArrayBuffer> {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Comparison Report");

    const title = "BONUS COMPARISON REPORT";

    const headers = [
      "EMP. ID",
      "Employee Name",
      "Department",
      "System Bonus",
      "HR Bonus",
      "Difference",
      "Status",
    ];

    // Add title
    worksheet.mergeCells("A1:G1");
    const titleCell = worksheet.getCell("A1");
    titleCell.value = title;
    titleCell.font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
    titleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD32F2F" },
    };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };

    // Add empty row
    worksheet.getRow(2).height = 15;

    // Add headers
    const headerRow = worksheet.getRow(3);
    headers.forEach((header, index) => {
      const cell = headerRow.getCell(index + 1);
      cell.value = header;
      cell.font = { bold: true, size: 12, color: { argb: "FF000000" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFCDD2" },
      };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "medium", color: { argb: "FF000000" } },
        bottom: { style: "medium", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });

    // Group by department
    const groupedComparisons = comparisons.reduce((acc, comp) => {
      const dept = comp.department || "Unknown";
      if (!acc[dept]) acc[dept] = [];
      acc[dept].push(comp);
      return acc;
    }, {} as { [key: string]: any[] });

    let currentRow = 4;

    Object.keys(groupedComparisons).forEach((dept) => {
      // Add department header
      const deptHeaderRow = worksheet.getRow(currentRow);
      deptHeaderRow.getCell(1).value = `DEPARTMENT: ${dept}`;
      deptHeaderRow.getCell(1).font = { bold: true, size: 12 };
      deptHeaderRow.getCell(1).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE0E0E0" },
      };
      worksheet.mergeCells(`A${currentRow}:G${currentRow}`);
      currentRow++;

      groupedComparisons[dept].forEach((comp: any) => {
        const dataRow = worksheet.getRow(currentRow);
        dataRow.height = 25;

        const rowData = [
          comp.empId,
          comp.name,
          comp.department || "N/A",
          comp.systemBonus,
          comp.hrBonus,
          comp.difference,
          comp.status,
        ];

        rowData.forEach((value, colIndex) => {
          const cell = dataRow.getCell(colIndex + 1);
          cell.value = value;
          cell.alignment = { horizontal: "center", vertical: "middle" };
          cell.border = {
            top: { style: "thin", color: { argb: "FFCCCCCC" } },
            bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
            left: { style: "thin", color: { argb: "FFCCCCCC" } },
            right: { style: "thin", color: { argb: "FFCCCCCC" } },
          };
        });
        currentRow++;
      });

      currentRow++; // Empty row between departments
    });

    // Set column widths for comparison report
    const compColWidths = [80, 200, 100, 120, 120, 120, 100];
    compColWidths.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width / 7;
    });

    // Set row heights
    worksheet.getRow(1).height = 35; // Title
    worksheet.getRow(2).height = 15; // Empty
    worksheet.getRow(3).height = 30; // Header

    return await workbook.xlsx.writeBuffer();
  }

  private static parseDate(dateValue: any): Date {
    if (!dateValue) return new Date("2020-01-01");

    if (dateValue instanceof Date) return dateValue;

    // Handle Excel date numbers
    if (typeof dateValue === "number") {
      return new Date((dateValue - 25569) * 86400 * 1000);
    }

    // Handle string dates
    if (typeof dateValue === "string") {
      const dateStr = dateValue.toString().trim();

      // Common date formats
      const formats = [
        /^(\d{2})\.(\d{2})\.(\d{2})$/, // DD.MM.YY
        /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/, // D.M.YYYY
        /^(\d{4})-(\d{2})-(\d{2})$/, // YYYY-MM-DD
        /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, // M/D/YYYY
      ];

      for (const format of formats) {
        const match = dateStr.match(format);
        if (match) {
          if (format === formats[0]) {
            // DD.MM.YY
            const day = parseInt(match[1]);
            const month = parseInt(match[2]) - 1;
            const year = parseInt(match[3]) + 2000;
            return new Date(year, month, day);
          } else if (format === formats[1]) {
            // D.M.YYYY
            const day = parseInt(match[1]);
            const month = parseInt(match[2]) - 1;
            const year = parseInt(match[3]);
            return new Date(year, month, day);
          } else if (format === formats[2]) {
            // YYYY-MM-DD
            return new Date(dateStr);
          } else if (format === formats[3]) {
            // M/D/YYYY
            const month = parseInt(match[1]) - 1;
            const day = parseInt(match[2]);
            const year = parseInt(match[3]);
            return new Date(year, month, day);
          }
        }
      }

      // Try direct parsing as last resort
      const parsedDate = new Date(dateStr);
      if (!isNaN(parsedDate.getTime())) {
        return parsedDate;
      }
    }

    return new Date("2020-01-01");
  }

  // Add these new methods at the end of the ExcelProcessor class

  static calculateStaffSummary(employees: Employee[]) {
    const monthlySummary: {
      [month: string]: {
        totalGrossSalary: number;
        totalSalary1: number;
        count: number;
      };
    } = {};

    let totalGrossSalary = 0;
    let totalSalary1 = 0;

    employees.forEach((emp) => {
      if (emp.monthlyData) {
        emp.monthlyData.forEach((monthData) => {
          if (!monthlySummary[monthData.month]) {
            monthlySummary[monthData.month] = {
              totalGrossSalary: 0,
              totalSalary1: 0,
              count: 0,
            };
          }

          // For staff: salary is from SALARY1 column (Column O)
          const salary1 = monthData.salary || 0;
          const grossSalary = salary1; // Staff uses SALARY1 as gross salary

          monthlySummary[monthData.month].totalSalary1 += salary1;
          monthlySummary[monthData.month].totalGrossSalary += grossSalary;
          monthlySummary[monthData.month].count++;

          totalSalary1 += salary1;
          totalGrossSalary += grossSalary;
        });
      }
    });

    // Sort months chronologically
    const sortedMonths = Object.keys(monthlySummary).sort();
    const sortedMonthlySummary: {
      [month: string]: {
        totalGrossSalary: number;
        totalSalary1: number;
        count: number;
      };
    } = {};

    sortedMonths.forEach((month) => {
      sortedMonthlySummary[month] = monthlySummary[month];
    });

    return {
      type: "STAFF",
      monthlySummary: sortedMonthlySummary,
      overallSummary: {
        totalGrossSalary: Math.round(totalGrossSalary * 100) / 100,
        totalSalary1: Math.round(totalSalary1 * 100) / 100,
        totalEmployees: employees.length,
        avgSalary:
          employees.length > 0
            ? Math.round((totalSalary1 / employees.length) * 100) / 100
            : 0,
      },
    };
  }

  static calculateWorkerSummary(employees: Employee[]) {
    const monthlySummary: {
      [month: string]: {
        totalGrossSalary: number;
        totalSalary1: number;
        count: number;
      };
    } = {};

    let totalGrossSalary = 0;
    let totalSalary1 = 0;

    employees.forEach((emp) => {
      if (emp.monthlyData) {
        emp.monthlyData.forEach((monthData) => {
          if (!monthlySummary[monthData.month]) {
            monthlySummary[monthData.month] = {
              totalGrossSalary: 0,
              totalSalary1: 0,
              count: 0,
            };
          }

          // For workers: salary is from Salary1 column (Column I)
          const salary1 = monthData.salary || 0;
          const grossSalary = salary1; // Workers use Salary1 as gross salary

          monthlySummary[monthData.month].totalSalary1 += salary1;
          monthlySummary[monthData.month].totalGrossSalary += grossSalary;
          monthlySummary[monthData.month].count++;

          totalSalary1 += salary1;
          totalGrossSalary += grossSalary;
        });
      }
    });

    // Sort months chronologically
    const sortedMonths = Object.keys(monthlySummary).sort();
    const sortedMonthlySummary: {
      [month: string]: {
        totalGrossSalary: number;
        totalSalary1: number;
        count: number;
      };
    } = {};

    sortedMonths.forEach((month) => {
      sortedMonthlySummary[month] = monthlySummary[month];
    });

    return {
      type: "WORKER",
      monthlySummary: sortedMonthlySummary,
      overallSummary: {
        totalGrossSalary: Math.round(totalGrossSalary * 100) / 100,
        totalSalary1: Math.round(totalSalary1 * 100) / 100,
        totalEmployees: employees.length,
        avgSalary:
          employees.length > 0
            ? Math.round((totalSalary1 / employees.length) * 100) / 100
            : 0,
      },
    };
  }

  private static async generateSalarySummarySheet(
    workbook: ExcelJS.Workbook,
    staffSummary: any,
    workerSummary: any
  ): Promise<void> {
    const worksheet = workbook.addWorksheet("Salary Summary");

    const title = "MONTHLY SALARY SUMMARY - DIWALI 2024-25";

    // Add title
    worksheet.mergeCells("A1:F1");
    const titleCell = worksheet.getCell("A1");
    titleCell.value = title;
    titleCell.font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
    titleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF2E7D32" },
    };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };

    // Add empty row
    worksheet.getRow(2).height = 15;

    let currentRow = 3;

    // STAFF SUMMARY SECTION
    const staffHeaderRow = worksheet.getRow(currentRow);
    staffHeaderRow.getCell(1).value = "STAFF MONTHLY SUMMARY";
    staffHeaderRow.getCell(1).font = { bold: true, size: 14 };
    staffHeaderRow.getCell(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFB0BEC5" },
    };
    worksheet.mergeCells(`A${currentRow}:F${currentRow}`);
    currentRow++;

    // Staff table headers
    const staffHeaders = [
      "Month",
      "Total Gross Salary",
      "Total Salary1",
      "Employee Count",
      "Average Salary",
      "",
    ];
    const staffHeaderRow2 = worksheet.getRow(currentRow);
    staffHeaders.forEach((header, index) => {
      const cell = staffHeaderRow2.getCell(index + 1);
      cell.value = header;
      cell.font = { bold: true, size: 12 };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE3F2FD" },
      };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
      };
    });
    currentRow++;

    // Staff monthly data
    let staffTotalGross = 0;
    let staffTotalSalary1 = 0;
    let staffTotalEmployees = 0;

    Object.keys(staffSummary.monthlySummary).forEach((month) => {
      const monthData = staffSummary.monthlySummary[month];
      const dataRow = worksheet.getRow(currentRow);

      staffTotalGross += monthData.totalGrossSalary;
      staffTotalSalary1 += monthData.totalSalary1;
      staffTotalEmployees = Math.max(staffTotalEmployees, monthData.count);

      const rowData = [
        month,
        Math.round(monthData.totalGrossSalary),
        Math.round(monthData.totalSalary1),
        monthData.count,
        monthData.count > 0
          ? Math.round(monthData.totalSalary1 / monthData.count)
          : 0,
        "",
      ];

      rowData.forEach((value, colIndex) => {
        const cell = dataRow.getCell(colIndex + 1);
        cell.value = value;
        if (colIndex === 0) {
          // Month column
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else if (colIndex >= 1 && colIndex <= 4) {
          // Numeric columns
          cell.alignment = { horizontal: "right", vertical: "middle" };
          if (colIndex >= 1 && colIndex <= 2) {
            // Currency columns
            cell.numFmt = "₹#,##0";
          }
        }
        cell.border = {
          top: { style: "thin", color: { argb: "FFCCCCCC" } },
          bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
          left: { style: "thin", color: { argb: "FFCCCCCC" } },
          right: { style: "thin", color: { argb: "FFCCCCCC" } },
        };
      });
      currentRow++;
    });

    // Staff total row
    const staffTotalRow = worksheet.getRow(currentRow);
    const staffTotalData = [
      "TOTAL (STAFF)",
      staffTotalGross,
      staffTotalSalary1,
      staffTotalEmployees,
      staffTotalEmployees > 0
        ? Math.round(staffTotalSalary1 / staffTotalEmployees)
        : 0,
      "",
    ];

    staffTotalData.forEach((value, colIndex) => {
      const cell = staffTotalRow.getCell(colIndex + 1);
      cell.value = value;
      cell.font = { bold: true, size: 12 };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFF59D" },
      };
      if (colIndex === 0) {
        cell.alignment = { horizontal: "center", vertical: "middle" };
      } else if (colIndex >= 1 && colIndex <= 4) {
        cell.alignment = { horizontal: "right", vertical: "middle" };
        if (colIndex >= 1 && colIndex <= 2) {
          cell.numFmt = "₹#,##0";
        }
      }
      cell.border = {
        top: { style: "medium", color: { argb: "FF000000" } },
        bottom: { style: "medium", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });
    currentRow += 2; // Empty row

    // WORKER SUMMARY SECTION
    const workerHeaderRow = worksheet.getRow(currentRow);
    workerHeaderRow.getCell(1).value = "WORKER MONTHLY SUMMARY";
    workerHeaderRow.getCell(1).font = { bold: true, size: 14 };
    workerHeaderRow.getCell(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFB0BEC5" },
    };
    worksheet.mergeCells(`A${currentRow}:F${currentRow}`);
    currentRow++;

    // Worker table headers
    const workerHeaders = [
      "Month",
      "Total Gross Salary",
      "Total Salary1",
      "Employee Count",
      "Average Salary",
      "",
    ];
    const workerHeaderRow2 = worksheet.getRow(currentRow);
    workerHeaders.forEach((header, index) => {
      const cell = workerHeaderRow2.getCell(index + 1);
      cell.value = header;
      cell.font = { bold: true, size: 12 };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE3F2FD" },
      };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
      };
    });
    currentRow++;

    // Worker monthly data
    let workerTotalGross = 0;
    let workerTotalSalary1 = 0;
    let workerTotalEmployees = 0;

    Object.keys(workerSummary.monthlySummary).forEach((month) => {
      const monthData = workerSummary.monthlySummary[month];
      const dataRow = worksheet.getRow(currentRow);

      workerTotalGross += monthData.totalGrossSalary;
      workerTotalSalary1 += monthData.totalSalary1;
      workerTotalEmployees = Math.max(workerTotalEmployees, monthData.count);

      const rowData = [
        month,
        Math.round(monthData.totalGrossSalary),
        Math.round(monthData.totalSalary1),
        monthData.count,
        monthData.count > 0
          ? Math.round(monthData.totalSalary1 / monthData.count)
          : 0,
        "",
      ];

      rowData.forEach((value, colIndex) => {
        const cell = dataRow.getCell(colIndex + 1);
        cell.value = value;
        if (colIndex === 0) {
          // Month column
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else if (colIndex >= 1 && colIndex <= 4) {
          // Numeric columns
          cell.alignment = { horizontal: "right", vertical: "middle" };
          if (colIndex >= 1 && colIndex <= 2) {
            // Currency columns
            cell.numFmt = "₹#,##0";
          }
        }
        cell.border = {
          top: { style: "thin", color: { argb: "FFCCCCCC" } },
          bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
          left: { style: "thin", color: { argb: "FFCCCCCC" } },
          right: { style: "thin", color: { argb: "FFCCCCCC" } },
        };
      });
      currentRow++;
    });

    // Worker total row
    const workerTotalRow = worksheet.getRow(currentRow);
    const workerTotalData = [
      "TOTAL (WORKER)",
      workerTotalGross,
      workerTotalSalary1,
      workerTotalEmployees,
      workerTotalEmployees > 0
        ? Math.round(workerTotalSalary1 / workerTotalEmployees)
        : 0,
      "",
    ];

    workerTotalData.forEach((value, colIndex) => {
      const cell = workerTotalRow.getCell(colIndex + 1);
      cell.value = value;
      cell.font = { bold: true, size: 12 };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFF59D" },
      };
      if (colIndex === 0) {
        cell.alignment = { horizontal: "center", vertical: "middle" };
      } else if (colIndex >= 1 && colIndex <= 4) {
        cell.alignment = { horizontal: "right", vertical: "middle" };
        if (colIndex >= 1 && colIndex <= 2) {
          cell.numFmt = "₹#,##0";
        }
      }
      cell.border = {
        top: { style: "medium", color: { argb: "FF000000" } },
        bottom: { style: "medium", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });
    currentRow += 2;

    // GRAND TOTAL ROW
    const grandTotalRow = worksheet.getRow(currentRow);
    const grandTotalData = [
      "GRAND TOTAL",
      staffTotalGross + workerTotalGross,
      staffTotalSalary1 + workerTotalSalary1,
      staffTotalEmployees + workerTotalEmployees,
      staffTotalEmployees + workerTotalEmployees > 0
        ? Math.round(
            (staffTotalSalary1 + workerTotalSalary1) /
              (staffTotalEmployees + workerTotalEmployees)
          )
        : 0,
      "",
    ];

    grandTotalData.forEach((value, colIndex) => {
      const cell = grandTotalRow.getCell(colIndex + 1);
      cell.value = value;
      cell.font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE65100" },
      };
      if (colIndex === 0) {
        cell.alignment = { horizontal: "center", vertical: "middle" };
      } else if (colIndex >= 1 && colIndex <= 4) {
        cell.alignment = { horizontal: "right", vertical: "middle" };
        if (colIndex >= 1 && colIndex <= 2) {
          cell.numFmt = "₹#,##0";
        }
      }
      cell.border = {
        top: { style: "double", color: { argb: "FF000000" } },
        bottom: { style: "double", color: { argb: "FF000000" } },
        left: { style: "medium", color: { argb: "FF000000" } },
        right: { style: "medium", color: { argb: "FF000000" } },
      };
    });

    // Set column widths
    const colWidths = [120, 150, 150, 120, 120, 50];
    colWidths.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width / 7;
    });

    // Set row heights
    worksheet.getRow(1).height = 35; // Title row
    worksheet.getRow(2).height = 15; // Empty row
  }
}
