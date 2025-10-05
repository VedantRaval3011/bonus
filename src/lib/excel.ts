import * as ExcelJS from "exceljs";
import { Employee, MonthlyData, BonusCalculation } from "./types";

function toNumber(cell: ExcelJS.Cell): number {
  const raw = cell.result ?? cell.value;
  if (typeof raw === "number") return raw;
  if (raw instanceof Date) return +raw;
  const n = Number(String(raw).trim());
  return isFinite(n) ? n : 0;
}

export class ExcelProcessor {
  private static shouldIncludeOctober(monthlyData: MonthlyData[]): boolean {
    // Find AUG-25 salary
    const augSalary = monthlyData.find(
      (md) =>
        md.month.toUpperCase().includes("AUG-25") ||
        md.month.toUpperCase() === "AUG-25"
    );

    // If AUG-25 exists and has salary > 0, include October
    // If AUG-25 doesn't exist or is 0, exclude October
    return augSalary ? augSalary.salary > 0 : false;
  }

  // ===== MONTH KEY NORMALIZATION =====
  private static normalizeMonthKey(sheetName: string): string {
    const s = String(sheetName || "")
      .toUpperCase()
      .trim();
    const m =
      s.match(/^([A-Z]+)\s*[-/]\s*(\d{2,4})/) ||
      s.match(/^([A-Z]+)\s+(\d{2,4})/);
    if (!m) return s.slice(0, 6);
    const mon = m[1];
    const yr = m[2];
    const map: Record<string, string> = {
      JAN: "JAN",
      JANUARY: "JAN",
      FEB: "FEB",
      FEBRUARY: "FEB",
      MAR: "MAR",
      MARCH: "MAR",
      APR: "APR",
      APRIL: "APR",
      MAY: "MAY",
      JUN: "JUN",
      JUNE: "JUN",
      JUL: "JUL",
      JULY: "JUL",
      AUG: "AUG",
      AUGUST: "AUG",
      SEP: "SEP",
      SEPT: "SEP",
      SEPTEMBER: "SEP",
      OCT: "OCT",
      OCTOBER: "OCT",
      NOV: "NOV",
      NOVEMBER: "NOV",
      DEC: "DEC",
      DECEMBER: "DEC",
    };
    const m3 = map[mon] || mon.slice(0, 3);
    const yy = yr.length === 4 ? yr.slice(2) : yr;
    return `${m3}-${yy}`;
  }

  // Parse Due Voucher List
  static async parseDueVoucherList(
    buffer: ArrayBuffer
  ): Promise<
    Map<string, { alreadyPaid: number; unpaid: number; dept: string }>
  > {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const dueVoucherMap = new Map<
        string,
        { alreadyPaid: number; unpaid: number; dept: string }
      >();
      const worksheet = workbook.getWorksheet("Sheet1");
      if (!worksheet) return dueVoucherMap;

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber <= 1) return;
        const empCode = row.getCell(2).value?.toString().trim();
        const dept = row.getCell(3).value?.toString().trim();
        const category = row.getCell(4).value?.toString().trim()?.toUpperCase();
        const dueVC = toNumber(row.getCell(6));

        if (!empCode || !category || dueVC === 0) return;

        if (!dueVoucherMap.has(empCode)) {
          dueVoucherMap.set(empCode, {
            alreadyPaid: 0,
            unpaid: 0,
            dept: dept || "",
          });
        }
        const entry = dueVoucherMap.get(empCode)!;
        if (category === "A") entry.alreadyPaid += dueVC;
        else if (category === "U") entry.unpaid += dueVC;
      });

      console.log(`Parsed ${dueVoucherMap.size} Due Voucher entries`);
      return dueVoucherMap;
    } catch (error) {
      console.error("Error parsing Due Voucher file:", error);
      return new Map();
    }
  }

  // Parse Loan Deduction file
  static async parseLoanDeduction(
    buffer: ArrayBuffer
  ): Promise<Map<string, number>> {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const loanMap = new Map<string, number>();

      const worksheet = workbook.worksheets[0];
      if (!worksheet) return loanMap;

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber <= 1) return; // Skip header

        const empCode = row.getCell(2).value?.toString().trim(); // Column B: EMP. ID
        const loanAmount = toNumber(row.getCell(7)); // Column G: DEDUCTION LOAN FOR BONUS

        if (empCode && loanAmount > 0) {
          loanMap.set(empCode, loanAmount);
        }
      });

      console.log(`Parsed ${loanMap.size} loan deduction entries`);
      return loanMap;
    } catch (error) {
      console.error("Error parsing Loan Deduction file:", error);
      return new Map();
    }
  }

  // Parse Actual Percentage Bonus Data
  static async parseActualPercentage(
    buffer: ArrayBuffer
  ): Promise<Map<string, number>> {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const percentageMap = new Map<string, number>();
      const worksheet = workbook.worksheets[0];
      if (!worksheet) return percentageMap;

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber <= 1) return;
        const rawCode = row.getCell(2).value;
        const empCode = String(Number(String(rawCode).trim()));
        const percentage = toNumber(row.getCell(5));
        if (empCode && percentage > 0) percentageMap.set(empCode, percentage);
      });
      return percentageMap;
    } catch (error) {
      console.error("Error parsing Actual Percentage file:", error);
      return new Map();
    }
  }

  private static getFinalDepartment(employee: Employee): string {
    // Get the most recent month's department
    if (!employee.monthlyData || employee.monthlyData.length === 0) {
      return employee.department;
    }

    // Sort by month and get the latest one
    const sorted = [...employee.monthlyData].sort((a, b) => {
      // Simple string comparison works for our month format (e.g., "SEP-25" > "AUG-25")
      return b.month.localeCompare(a.month);
    });

    return sorted[0].department || employee.department;
  }

  // Parse Staff.xlsx - Extract from multiple monthly sheets
  static async parseStaffFile(
    buffer: ArrayBuffer
  ): Promise<{ employees: Employee[]; summary: any }> {
    try {
      const workbook = new ExcelJS.Workbook();
      const employeeMap = new Map<string, Employee>();

      if (!buffer || buffer.byteLength === 0) {
        console.error("Invalid or empty staff file buffer");
        return { employees: [], summary: null };
      }

      await workbook.xlsx.load(buffer);
      console.log(
        "Staff sheet names:",
        workbook.worksheets.map((ws) => ws.name)
      );

      workbook.worksheets.forEach((worksheet) => {
        const sheetName = worksheet.name;
        if (!sheetName.includes("-") || !sheetName.endsWith(" O")) {
          console.log(`Skipping non-staff sheet: ${sheetName}`);
          return;
        }
        console.log(`Processing staff sheet: ${sheetName}`);
        let processedCount = 0;
        let skippedCount = 0;

        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          if (rowNumber <= 3) return;

          const empId = row.getCell(2).value?.toString().trim();
          const dept = row.getCell(3).value?.toString().trim() || "S";
          const name = row.getCell(5).value?.toString().trim();
          const salaryCell = row.getCell(15);
          let salary = 0;
          if (salaryCell.result !== undefined) salary = toNumber(salaryCell);
          else if (salaryCell.value !== undefined)
            salary = toNumber(salaryCell);

          const STAFF_DOJ_COL = 33;
          const dojRaw = row.getCell(STAFF_DOJ_COL).value;

          const dojText = String(dojRaw ?? "")
            .trim()
            .toUpperCase();
          const empIdUpper = (empId || "").toUpperCase();
          const isNValue =
            dojText === "N" || empIdUpper === "N" || empIdUpper.startsWith("N");
          const isNotApplicable =
            dojText === "NA" ||
            dojText === "N.A" ||
            dojText === "N.A." ||
            dojText === "N/A" ||
            dojText === "";

          // If DOJ is N/A (but not "N"), skip this row completely
          if (isNotApplicable && !isNValue) {
            skippedCount++;
            console.log(
              `⏭️ Skipping staff ${empId} - ${name} (DOJ: N/A) in sheet ${sheetName}`
            );
            return;
          }

          // Handle "N" employees/DOJ specially
          let doj: Date | null;
          if (isNValue) {
            doj = new Date("2020-01-01"); // Default date for "N" employees
            console.log(
              `📝 Processing "N" employee ${empId} - ${name} in sheet ${sheetName}`
            );
          } else {
            // Parse DOJ normally
            doj =
              ExcelProcessor.parseDate?.(dojRaw) ??
              (dojRaw instanceof Date ? dojRaw : null);

            // Validate parsed date
            if (!(doj instanceof Date) || isNaN(doj.getTime())) {
              skippedCount++;
              console.log(
                `⏭️ Skipping staff ${empId} - ${name} (Invalid DOJ: ${dojRaw}) in sheet ${sheetName}`
              );
              return;
            }
          }

          // Skip if other critical data is missing

          if (
            !empId ||
            empId === "0" ||
            !name ||
            name.toLowerCase() === "total"
          ) {
            return processedCount;
          }
          if (!isNValue && salary === 0) {
            return processedCount;
          }

          // Add to employee map (only if DOJ is valid and not N/A)
          // Create a unique key for the employee map
          // For "N" employees, use name as part of the key to keep them unique
          const mapKey =
            empId === "N" || empId.toUpperCase() === "N" ? `N_${name}` : empId;

          // Add to employee map (only if DOJ is valid and not N/A)
          if (!employeeMap.has(mapKey)) {
            const finalDept = isNValue ? "N" : dept;

            employeeMap.set(mapKey, {
              empId: empId, // Keep original empId as "N"
              name,
              department: finalDept,
              doj,
              salary: 0,
              monthlyData: [],
            });
          }

          if (isNValue || salary > 0) {
            const employee = employeeMap.get(mapKey)!;
            if (!employee.monthlyData) employee.monthlyData = [];
            const monthKey = ExcelProcessor.normalizeMonthKey(sheetName);
            const monthDept =
              empId === "N" || String(dojRaw).trim().toUpperCase() === "N"
                ? "N"
                : dept;

            employee.monthlyData.push({
              month: monthKey,
              salary, // Can be 0 for N employees
              department: monthDept,
            });
          }

          const employee = employeeMap.get(mapKey)!;
          // Update employee's primary department
          const currentDept =
            empId === "N" || String(dojRaw).trim().toUpperCase() === "N"
              ? "N"
              : dept;
          employee.department = currentDept;
        });

        console.log(
          `Processed ${processedCount} staff employees, skipped ${skippedCount} (N/A DOJ) from sheet: ${sheetName}`
        );
      });

      const result = Array.from(employeeMap.values());
      console.log(`✅ Total valid staff employees: ${result.length}`);
      const staffSummary = ExcelProcessor.calculateStaffSummary(result);
      console.log("Staff Summary:", staffSummary);
      return { employees: result, summary: staffSummary };
    } catch (error) {
      console.error("Error parsing staff file:", error);
      return { employees: [], summary: null };
    }
  }

  // ============= COMPLETE parseWorkerFile METHOD =============
  static async parseWorkerFile(
    buffer: ArrayBuffer
  ): Promise<{ employees: Employee[]; summary: any }> {
    try {
      const workbook = new ExcelJS.Workbook();
      const employeeMap = new Map<string, Employee>();

      if (!buffer || buffer.byteLength === 0) {
        console.error("Invalid or empty worker file buffer");
        return { employees: [], summary: null };
      }

      await workbook.xlsx.load(buffer);
      console.log(
        "Worker sheet names:",
        workbook.worksheets.map((ws) => ws.name)
      );

      workbook.worksheets.forEach((worksheet) => {
        const sheetName = worksheet.name;
        if (!sheetName.includes("-") || !sheetName.endsWith(" W")) {
          console.log(`Skipping non-worker sheet: ${sheetName}`);
          return;
        }
        console.log(`Processing worker sheet: ${sheetName}`);
        let processedCount = 0;
        let skippedCount = 0;

        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          if (rowNumber <= 2) return;

          const empId = row.getCell(2).value?.toString().trim();
          const dept = row.getCell(3).value?.toString().trim() || "W";
          const name = row.getCell(4).value?.toString().trim();
          const salaryCell = row.getCell(9);
          let salary = 0;
          if (salaryCell.result !== undefined) salary = toNumber(salaryCell);
          else if (salaryCell.value !== undefined)
            salary = toNumber(salaryCell);

          const WORKER_DOJ_COL = 26;
          const dojRaw = row.getCell(WORKER_DOJ_COL).value;

          // CHECK FOR N/A DOJ VALUES OR N EMPLOYEE ID
          const dojText = String(dojRaw ?? "")
            .trim()
            .toUpperCase();
          const empIdUpper = (empId || "").toUpperCase();
          const isNValue =
            dojText === "N" || empIdUpper === "N" || empIdUpper.startsWith("N");
          const isNotApplicable =
            dojText === "NA" ||
            dojText === "N.A" ||
            dojText === "N.A." ||
            dojText === "N/A" ||
            dojText === "";

          // If DOJ is N/A (but not "N"), skip this row completely
          if (isNotApplicable && !isNValue) {
            skippedCount++;
            console.log(
              `⏭️ Skipping worker ${empId} - ${name} (DOJ: N/A) in sheet ${sheetName}`
            );
            return;
          }

          // Handle "N" employees/DOJ specially
          let doj: Date | null;
          if (isNValue) {
            doj = new Date("2020-01-01"); // Default date for "N" employees
            console.log(
              `📝 Processing "N" worker ${empId} - ${name} in sheet ${sheetName}`
            );
          } else {
            // Parse DOJ normally
            doj =
              ExcelProcessor.parseDate?.(dojRaw) ??
              (dojRaw instanceof Date ? dojRaw : null);

            // Validate parsed date
            if (!(doj instanceof Date) || isNaN(doj.getTime())) {
              skippedCount++;
              console.log(
                `⏭️ Skipping worker ${empId} - ${name} (Invalid DOJ: ${dojRaw}) in sheet ${sheetName}`
              );
              return;
            }
          }

          // Get CASH SALARY column (column N = 14)
          const cashSalaryCell = row.getCell(14);
          const cashSalary = cashSalaryCell.value?.toString().trim() || "";
          const isCashSalary = cashSalary.length > 0;

          // Skip if other critical data is missing
          if (
            !empId ||
            empId === "0" ||
            !name ||
            name.toLowerCase() === "total" ||
            empId.toLowerCase() === "total"
          ) {
            return;
          }

          // MODIFIED: For non-N employees, skip if salary is 0
          // For N employees, allow them through even with 0 salary
          if (!isNValue && salary <= 0) {
            return;
          }

          processedCount++;

          // Add to employee map (always include N employees)
          const mapKey =
            empId === "N" || empId.toUpperCase() === "N" ? `N_${name}` : empId;

          // Add to employee map (always include N employees)
          if (!employeeMap.has(mapKey)) {
            // Assign department "N" if empId is "N" or DOJ is "N"
            const finalDept = isNValue ? "N" : dept;

            employeeMap.set(mapKey, {
              empId: empId, // Keep original empId as "N"
              name,
              department: finalDept,
              doj,
              salary: 0,
              monthlyData: [],
              isCashSalary: false,
            });
          }

          const employee = employeeMap.get(mapKey)!;

          // Update cash salary flag if any month has cash salary
          if (isCashSalary) {
            employee.isCashSalary = true;
          }

          if (!employee.monthlyData) employee.monthlyData = [];
          const monthKey = ExcelProcessor.normalizeMonthKey(sheetName);

          // MODIFIED: Add monthly data for N employees even with 0 salary
          // For regular employees, only add if salary > 0
          if (isNValue || salary > 0) {
            // Use department "N" for monthly data if applicable
            const monthDept =
              empId === "N" || String(dojRaw).trim().toUpperCase() === "N"
                ? "N"
                : dept;

            employee.monthlyData.push({
              month: monthKey,
              salary, // Can be 0 for N employees
              department: monthDept,
            });
          }

          const currentDept =
            empId === "N" || String(dojRaw).trim().toUpperCase() === "N"
              ? "N"
              : dept;
          employee.department = currentDept;
        });

        console.log(
          `Processed ${processedCount} worker employees, skipped ${skippedCount} (N/A DOJ) from sheet: ${sheetName}`
        );
      });

      const result = Array.from(employeeMap.values());
      console.log(`✅ Total valid worker employees: ${result.length}`);
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
  ): Promise<Map<string, any>> {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const hrMap = new Map();

      ["Staff", "Worker"].forEach((sheetName) => {
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) return;

        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber <= 2) return;
          const empCode = row.getCell(2).value?.toString().trim();
          if (!empCode) return;
          const dept =
            row.getCell(3).value?.toString().trim() || sheetName.charAt(0);
          const salary12 = toNumber(row.getCell(17));
          const gross = toNumber(row.getCell(18));
          const register = toNumber(row.getCell(19));
          const dueVC = toNumber(row.getCell(20));
          const finalRTGS = toNumber(row.getCell(21));

          hrMap.set(empCode, {
            bonus: finalRTGS,
            department: dept,
            grossSal: gross,
            gross02: gross,
            register,
            actual: finalRTGS,
            unpaid: dueVC,
            reim: 0,
          });
        });
      });

      console.log(`✅ Parsed ${hrMap.size} HR entries from comparison file`);
      return hrMap;
    } catch (error) {
      console.error("❌ Error parsing HR comparison file:", error);
      return new Map();
    }
  }

  // Generate Final Bonus Excel
  static async generateFinalBonusExcel(
    calculations: BonusCalculation[],
    staffSummary?: any,
    workerSummary?: any,
    hrMonthlyTotals?: Map<string, number>,
    staffBuffer?: ArrayBuffer, // ADD THIS
    workerBuffer?: ArrayBuffer
  ): Promise<ArrayBuffer> {
    try {
      const workbook = new ExcelJS.Workbook();

      if (!Array.isArray(calculations) || calculations.length === 0) {
        console.error("Invalid or empty calculations array");
        const worksheet = workbook.addWorksheet("Error");
        worksheet.getCell("A1").value =
          "No data available for bonus calculations";
        return await workbook.xlsx.writeBuffer();
      }

      const departments = this.groupByDepartments(calculations);
      console.log("Departments found:", Object.keys(departments));

      for (const dept of Object.keys(departments)) {
        const deptCalculations = departments[dept];
        const isStaff =
          dept === "S" ||
          dept === "N" || // Add "N" department
          dept === "Sci Prec-" ||
          dept === "NRTM" ||
          dept === "Sci Prec Life.-";
        let sheetName =
          dept === "S"
            ? "Staff"
            : dept === "W"
            ? "Worker"
            : dept === "N"
            ? "N (Special Cases)"
            : dept;

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

      await this.generateSummarySheet(
        workbook,
        departments,
        hrMonthlyTotals,
        staffBuffer, // ADD THIS
        workerBuffer // ADD THIS
      );
      if (staffSummary && workerSummary) {
        await this.generateSalarySummarySheet(
          workbook,
          staffSummary,
          workerSummary
        );
      }

      return await workbook.xlsx.writeBuffer();
    } catch (error) {
      console.error("Error generating final bonus excel:", error);
      throw error;
    }
  }

  private static recalculateForDepartment(
    calc: BonusCalculation,
    dept: string
  ): BonusCalculation {
    // Filter monthly data for this department only
    const deptMonthlyData = (calc.monthlyData || []).filter(
      (md) => md.department === dept
    );

    // Recalculate total gross salary for this department
    const deptGrossSalary = deptMonthlyData.reduce(
      (sum, md) => sum + (md.salary || 0),
      0
    );

    return {
      ...calc,
      department: dept,
      monthlyData: deptMonthlyData,
      totalGrossSalary: Math.round(deptGrossSalary),
      // Note: Other fields like register, bonus, etc. would need recalculation
      // based on the department-specific gross salary
    };
  }

  private static groupByDepartments(
    calculations: BonusCalculation[]
  ): Record<string, BonusCalculation[]> {
    const departments: Record<string, BonusCalculation[]> = {};
    calculations.forEach((calc) => {
      if (!calc) return;

      // Special handling for "N" department employees
      if (calc.department === "N" || calc.empId === "N") {
        if (!departments["N"]) departments["N"] = [];
        const deptCalc = this.recalculateForDepartment(calc, "N");
        departments["N"].push(deptCalc);
        return;
      }

      // Handle employees with empty monthlyData
      if (!calc.monthlyData || calc.monthlyData.length === 0) {
        const dept = calc.department || "N";
        if (!departments[dept]) departments[dept] = [];
        const deptCalc = this.recalculateForDepartment(calc, dept);
        departments[dept].push(deptCalc);
        return;
      }

      // For other employees, group by departments found in monthly data
      const employeeDepts = new Set<string>();
      calc.monthlyData.forEach((md) => {
        if (md.department) employeeDepts.add(md.department);
      });

      // If no departments found in monthly data, use the employee's primary department
      if (employeeDepts.size === 0) {
        const dept = calc.department || "N";
        if (!departments[dept]) departments[dept] = [];
        const deptCalc = this.recalculateForDepartment(calc, dept);
        departments[dept].push(deptCalc);
      } else {
        employeeDepts.forEach((dept) => {
          if (!departments[dept]) departments[dept] = [];
          const deptCalc = this.recalculateForDepartment(calc, dept);
          departments[dept].push(deptCalc);
        });
      }
    });
    return departments;
  }

  private static async generateStaffSheet(
    workbook: ExcelJS.Workbook,
    sheetName: string,
    calculations: BonusCalculation[],
    dept: string
  ): Promise<void> {
    const worksheet = workbook.addWorksheet(sheetName);
    const title = `DIWALI BONUS LIST FROM NOVEMBER-2024 TO OCTOBER-2025 (2024-2025) INDIANA BOYS ${sheetName.toUpperCase()}`;
    const headers = [
      "SR. No.",
      "EMP Code",
      "Emp Name",
      "DOJ",
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
      "Gross Salary",
      "Gross 2",
      "Register",
      "Already Paid",
      "Unpaid",
      "Eligible",
      "After V",
      "Percentage",
      "Actual",
      "Reim",
      "Loan",
      "Final RTGS",
    ];

    worksheet.mergeCells("A1:AA1");
    const titleCell = worksheet.getCell("A1");
    titleCell.value = title;
    titleCell.font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
    titleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF2E7D32" },
    };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };

    worksheet.getRow(2).height = 15;

    const headerRow = worksheet.getRow(3);
    headers.forEach((header, index) => {
      const cell = headerRow.getCell(index + 1);
      cell.value = header;
      cell.font = { bold: true, size: 11, color: { argb: "FF000000" } };
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

    calculations.sort((a, b) => {
      const na = Number(a.empId),
        nb = Number(b.empId);
      if (isFinite(na) && isFinite(nb) && na !== nb) return na - nb;
      return (
        String(a.empId).localeCompare(String(b.empId)) ||
        a.name.localeCompare(b.name)
      );
    });

    const dataStartRow = 4;
    let rowIndex = dataStartRow;
    const monthOrder = [
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
    ];

    calculations.forEach((calc, index) => {
      const rowNum = rowIndex + index;
      const currentRow = worksheet.getRow(rowNum);
      currentRow.height = 25;

      const monthToSalary = new Map<string, number>();
      for (const md of calc.monthlyData || []) {
        const key = (md.month || "").slice(0, 6).toUpperCase();
        if (!monthToSalary.has(key) && md.salary > 0)
          monthToSalary.set(key, Math.round(md.salary));
      }

      // Check if October should be included
      const includeOctober = this.shouldIncludeOctober(calc.monthlyData || []);

      const monthlyValues: (number | null)[] = monthOrder.map((m) => {
        const v = monthToSalary.get(m);
        return typeof v === "number" ? Math.round(v) : null;
      });

      // Add October only if AUG-25 salary > 0
      if (includeOctober) {
        const octValue = monthToSalary.get("OCT-25");
        monthlyValues.push(
          octValue !== undefined ? Math.round(octValue) : null
        );
      } else {
        monthlyValues.push(null); // Keep null for October
      }

      const basicData = [index + 1, calc.empId, calc.name, calc.doj];
      basicData.forEach((value, colIndex) => {
        const cell = currentRow.getCell(colIndex + 1);
        cell.value = value;
        if (colIndex === 2)
          cell.alignment = { horizontal: "left", vertical: "middle" };
        else if (colIndex === 3 && value instanceof Date) {
          cell.numFmt = "dd-mm-yyyy";
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "thin", color: { argb: "FFCCCCCC" } },
          bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
          left: { style: "thin", color: { argb: "FFCCCCCC" } },
          right: { style: "thin", color: { argb: "FFCCCCCC" } },
        };
      });

      monthlyValues.forEach((value, idx) => {
        const cell = currentRow.getCell(5 + idx);
        cell.value = value; // Will be null for missing months, which Excel displays as blank
        if (value !== null) {
          cell.numFmt = "#,##0";
        }
        cell.alignment = { horizontal: "right", vertical: "middle" };
        cell.border = {
          top: { style: "thin", color: { argb: "FFCCCCCC" } },
          bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
          left: { style: "thin", color: { argb: "FFCCCCCC" } },
          right: { style: "thin", color: { argb: "FFCCCCCC" } },
        };
      });

      const octCell = currentRow.getCell(16);
      octCell.value = {
        formula: `IF(N${rowNum}=0, "", IF(COUNTBLANK(E${rowNum}:O${rowNum})=11, "", ROUND(AVERAGEIF(E${rowNum}:O${rowNum},">0"),0)))`,
      };

      octCell.numFmt = "#,##0";
      octCell.alignment = { horizontal: "right", vertical: "middle" };
      octCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      const grossSalaryCell = currentRow.getCell(17);
      grossSalaryCell.value = {
        formula: `ROUND(SUM(E${rowNum}:P${rowNum}),0)`,
      };
      grossSalaryCell.numFmt = "#,##0";
      grossSalaryCell.alignment = { horizontal: "right", vertical: "middle" };
      grossSalaryCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      const gross2Cell = currentRow.getCell(18);
      gross2Cell.value = {
        formula: `IF(X${rowNum}=8.33,Q${rowNum},IF(X${rowNum}>8.33,Q${rowNum}*0.6,""))`,
      };
      gross2Cell.numFmt = "#,##0";
      gross2Cell.alignment = { horizontal: "right", vertical: "middle" };
      gross2Cell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      const registerCell = currentRow.getCell(19);
      const percentValue = calc.bonusPercent || 8.33;
      registerCell.value = { formula: `ROUND(R${rowNum}*${percentValue}%,0)` };
      registerCell.numFmt = "#,##0";
      registerCell.alignment = { horizontal: "right", vertical: "middle" };
      registerCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      const alreadyPaidCell = currentRow.getCell(20);
      alreadyPaidCell.value = calc.alreadyPaid || 0;
      alreadyPaidCell.numFmt = "#,##0";
      alreadyPaidCell.alignment = { horizontal: "right", vertical: "middle" };
      alreadyPaidCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      const unpaidCell = currentRow.getCell(21);
      unpaidCell.value = calc.unpaid || 0;
      unpaidCell.numFmt = "#,##0";
      unpaidCell.alignment = { horizontal: "right", vertical: "middle" };
      unpaidCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      const eligibleCell = currentRow.getCell(22);
      eligibleCell.value = "Yes";
      eligibleCell.alignment = { horizontal: "center", vertical: "middle" };
      eligibleCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      const afterVCell = currentRow.getCell(23);
      afterVCell.value = { formula: `S${rowNum}-(T${rowNum}+U${rowNum})` };
      afterVCell.numFmt = "#,##0";
      afterVCell.alignment = { horizontal: "right", vertical: "middle" };
      afterVCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      const percentageCell = currentRow.getCell(24);
      percentageCell.value = calc.bonusPercent || 0;
      percentageCell.numFmt = "0.00";
      percentageCell.alignment = { horizontal: "right", vertical: "middle" };
      percentageCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      const actualCell = currentRow.getCell(25);
      actualCell.value = { formula: `S${rowNum}-(T${rowNum}+U${rowNum})` };
      actualCell.numFmt = "#,##0";
      actualCell.alignment = { horizontal: "right", vertical: "middle" };
      actualCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      const reimCell = currentRow.getCell(26);
      reimCell.value = { formula: `W${rowNum}-Y${rowNum}` };
      reimCell.numFmt = "#,##0";
      reimCell.alignment = { horizontal: "right", vertical: "middle" };
      reimCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // NEW: Loan column (column 27)
      const loanCell = currentRow.getCell(27);
      loanCell.value = calc.loan || 0;
      loanCell.numFmt = "#,##0";
      loanCell.alignment = { horizontal: "right", vertical: "middle" };
      loanCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // UPDATE: Final RTGS (now column 28, was 27)
      const finalRTGSCell = currentRow.getCell(28);
      finalRTGSCell.value = { formula: `Y${rowNum}-AA${rowNum}` }; // Actual - Loan
      finalRTGSCell.numFmt = "#,##0";
      finalRTGSCell.alignment = { horizontal: "right", vertical: "middle" };
      finalRTGSCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };
    });

    const dataEndRow = dataStartRow + calculations.length - 1;
    const monthTotalsRowIndex = dataEndRow + 1;
    const monthTotalsRow = worksheet.getRow(monthTotalsRowIndex);
    monthTotalsRow.height = 28;

    const monthTotalsLabel = worksheet.getCell(`A${monthTotalsRowIndex}`);
    monthTotalsLabel.value = "MONTH TOTALS";
    monthTotalsLabel.font = { bold: true };
    monthTotalsLabel.alignment = { horizontal: "center", vertical: "middle" };
    monthTotalsLabel.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFF59D" },
    };
    monthTotalsLabel.border = {
      top: { style: "medium", color: { argb: "FF000000" } },
      bottom: { style: "medium", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };

    const monthCols = [
      { title: "NOV-24", col: "E" },
      { title: "DEC-24", col: "F" },
      { title: "JAN-25", col: "G" },
      { title: "FEB-25", col: "H" },
      { title: "MAR-25", col: "I" },
      { title: "APR-25", col: "J" },
      { title: "MAY-25", col: "K" },
      { title: "JUN-25", col: "L" },
      { title: "JUL-25", col: "M" },
      { title: "AUG-25", col: "N" },
      { title: "SEP-25", col: "O" },
      { title: "OCT-25", col: "P" },
    ];
    for (const col of monthCols) {
      const cell = worksheet.getCell(`${col.col}${monthTotalsRowIndex}`);
      cell.value = {
        formula: `SUM(${col.col}${dataStartRow}:${col.col}${dataEndRow})`,
      };
      cell.numFmt = "₹#,##0";
      cell.font = { bold: true };
      cell.alignment = { horizontal: "right", vertical: "middle" };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFF59D" },
      };
      cell.border = {
        top: { style: "medium", color: { argb: "FF000000" } },
        bottom: { style: "medium", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    }

    // Add Year Total (NOV-SEP) one row below month totals
    const yearNovSepRowIndex = monthTotalsRowIndex + 1;
    const yearNovSepRow = worksheet.getRow(yearNovSepRowIndex);
    yearNovSepRow.height = 28;

    const yearNovSepLabelCell = worksheet.getCell(`A${yearNovSepRowIndex}`);
    yearNovSepLabelCell.value = "YEAR TOTAL (NOV–SEP)";
    yearNovSepLabelCell.font = { bold: true };
    yearNovSepLabelCell.alignment = {
      horizontal: "center",
      vertical: "middle",
    };
    yearNovSepLabelCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFD54F" },
    };
    yearNovSepLabelCell.border = {
      top: { style: "medium", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };

    const yearNovSepTotalCell = worksheet.getCell(`Q${yearNovSepRowIndex}`);
    yearNovSepTotalCell.value = {
      formula: `SUM(E${monthTotalsRowIndex}:O${monthTotalsRowIndex})`,
    };
    yearNovSepTotalCell.numFmt = "₹#,##0";
    yearNovSepTotalCell.font = { bold: true };
    yearNovSepTotalCell.alignment = { horizontal: "right", vertical: "middle" };
    yearNovSepTotalCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFD54F" },
    };
    yearNovSepTotalCell.border = {
      top: { style: "medium", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };

    // Add Full Year Total (NOV–OCT) below NOV-SEP total
    const fullYearRowIndex = yearNovSepRowIndex + 1;
    const fullYearRow = worksheet.getRow(fullYearRowIndex);
    fullYearRow.height = 28;

    const fullYearLabelCell = worksheet.getCell(`A${fullYearRowIndex}`);
    fullYearLabelCell.value = "YEAR TOTAL (NOV–OCT)";
    fullYearLabelCell.font = { bold: true };
    fullYearLabelCell.alignment = { horizontal: "center", vertical: "middle" };
    fullYearLabelCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFD54F" },
    };
    fullYearLabelCell.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "medium", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };

    const fullYearTotalCell = worksheet.getCell(`Q${fullYearRowIndex}`);
    fullYearTotalCell.value = {
      formula: `SUM(E${monthTotalsRowIndex}:P${monthTotalsRowIndex})`,
    };
    fullYearTotalCell.numFmt = "₹#,##0";
    fullYearTotalCell.font = { bold: true };
    fullYearTotalCell.alignment = { horizontal: "right", vertical: "middle" };
    fullYearTotalCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFD54F" },
    };
    fullYearTotalCell.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "medium", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };

    // Grand totals row comes after the year totals
    const totalsRowIndex = fullYearRowIndex + 1;
    const totalsRow = worksheet.getRow(totalsRowIndex);
    totalsRow.height = 35;

    const grandTotalCell = totalsRow.getCell(1);
    grandTotalCell.value = "GRAND TOTAL :- ";
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

    const totalColumns = [17, 18, 19, 20, 21, 23, 25, 26, 27, 28]; // Added 27, changed last to 28

    const colLetter = (n: number) => {
      const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
      if (n <= 26) return letters[n - 1];
      const first = letters[Math.floor((n - 1) / 26) - 1];
      const second = letters[(n - 1) % 26];
      return `${first}${second}`;
    };

    totalColumns.forEach((ci) => {
      const cell = totalsRow.getCell(ci);
      const letter = colLetter(ci);
      cell.value = {
        formula: `SUM(${letter}${dataStartRow}:${letter}${dataEndRow})`,
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

    const colWidths = [
      50,
      80,
      180,
      90, // SR.No, EMP Code, Name, DOJ
      80,
      80,
      80,
      80,
      80,
      80,
      80,
      80,
      80,
      80,
      80,
      80, // 12 months
      100,
      100,
      100,
      100,
      100,
      80,
      100,
      80,
      100,
      100, // Gross to Reim
      100, // Loan - ADD THIS
      110, // Final RTGS
    ];
    colWidths.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width / 7;
    });

    worksheet.getRow(1).height = 35;
    worksheet.getRow(2).height = 15;
    worksheet.getRow(3).height = 30;
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
      "SR. No.",
      "EMP Code",
      "Emp Name",
      "DOJ",
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
      "Gross Salary",
      "Register",
      "Already Paid",
      "Unpaid",
      "Eligible",
      "After V",
      "Percentage",
      "Actual",
      "Reim",
      "Loan",
      "Final RTGS",
    ];

    // Title row
    worksheet.mergeCells("A1:AB1");
    const titleCell = worksheet.getCell("A1");
    titleCell.value = title;
    titleCell.font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
    titleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF2E7D32" },
    };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };

    worksheet.getRow(2).height = 15;

    // Header row
    const headerRow = worksheet.getRow(3);
    headers.forEach((header, index) => {
      const cell = headerRow.getCell(index + 1);
      cell.value = header;
      cell.font = { bold: true, size: 11, color: { argb: "FF000000" } };
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

    // Sort calculations
    calculations.sort((a, b) => {
      const na = Number(a.empId),
        nb = Number(b.empId);
      if (isFinite(na) && isFinite(nb) && na !== nb) return na - nb;
      return (
        String(a.empId).localeCompare(String(b.empId)) ||
        a.name.localeCompare(b.name)
      );
    });

    const dataStartRow = 4;
    let rowIndex = dataStartRow;
    const monthOrder = [
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
    ];

    calculations.forEach((calc, index) => {
      const rowNum = rowIndex + index;
      const currentRow = worksheet.getRow(rowNum);
      currentRow.height = 25;

      const monthToSalary = new Map<string, number>();
      for (const md of calc.monthlyData || []) {
        const key = (md.month || "").slice(0, 6).toUpperCase();
        if (!monthToSalary.has(key) && md.salary > 0)
          monthToSalary.set(key, Math.round(md.salary));
      }

      const monthlyValues: (number | null)[] = monthOrder.map((m) => {
        const v = monthToSalary.get(m);
        return typeof v === "number" ? Math.round(v) : null;
      });

      // Basic data (SR. No., EMP Code, Name, DOJ)
      const basicData = [index + 1, calc.empId, calc.name, calc.doj];
      basicData.forEach((value, colIndex) => {
        const cell = currentRow.getCell(colIndex + 1);
        cell.value = value;
        if (colIndex === 2)
          cell.alignment = { horizontal: "left", vertical: "middle" };
        else if (colIndex === 3 && value instanceof Date) {
          cell.numFmt = "dd-mm-yyyy";
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "thin", color: { argb: "FFCCCCCC" } },
          bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
          left: { style: "thin", color: { argb: "FFCCCCCC" } },
          right: { style: "thin", color: { argb: "FFCCCCCC" } },
        };
      });

      // Monthly values (NOV-24 to SEP-25)
      monthlyValues.forEach((value, idx) => {
        const cell = currentRow.getCell(5 + idx);
        cell.value = value;
        if (value !== null) {
          cell.numFmt = "#,##0";
        }
        cell.alignment = { horizontal: "right", vertical: "middle" };

        cell.border = {
          top: { style: "thin", color: { argb: "FFCCCCCC" } },
          bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
          left: { style: "thin", color: { argb: "FFCCCCCC" } },
          right: { style: "thin", color: { argb: "FFCCCCCC" } },
        };
      });

      // Salary12 (Column P, index 16) = AVERAGE(E to O)
      const salary12Cell = currentRow.getCell(16);
      salary12Cell.value = {
        formula: `IF(N${rowNum}=0, "", IF(COUNTBLANK(E${rowNum}:O${rowNum})=11, "", ROUND(AVERAGEIF(E${rowNum}:O${rowNum},">0"),0)))`,
      };
      salary12Cell.numFmt = "#,##0";
      salary12Cell.alignment = { horizontal: "right", vertical: "middle" };
      salary12Cell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // Gross Salary (Column Q, index 17) = SUM(E to P)
      const grossSalaryCell = currentRow.getCell(17);
      grossSalaryCell.value = {
        formula: `ROUND(SUM(E${rowNum}:P${rowNum}),0)`,
      };
      grossSalaryCell.numFmt = "#,##0";
      grossSalaryCell.alignment = { horizontal: "right", vertical: "middle" };
      grossSalaryCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // Register (Column R, index 18) = Gross Salary * 8.33%
      // Register (Column R, index 18) = Gross Salary * 8.33% OR 0 if cash salary
      const registerCell = currentRow.getCell(18);
      if (calc.isCashSalary) {
        // For cash salary employees, register is 0
        registerCell.value = 0;
      } else {
        // For regular employees, calculate register
        registerCell.value = { formula: `ROUND(Q${rowNum}*8.33%,0)` };
      }
      registerCell.numFmt = "#,##0";
      registerCell.alignment = { horizontal: "right", vertical: "middle" };
      registerCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // Already Paid (Column S, index 19)
      const alreadyPaidCell = currentRow.getCell(19);
      alreadyPaidCell.value = calc.alreadyPaid || 0;
      alreadyPaidCell.numFmt = "#,##0";
      alreadyPaidCell.alignment = { horizontal: "right", vertical: "middle" };
      alreadyPaidCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // Unpaid (Column T, index 20)
      const unpaidCell = currentRow.getCell(20);
      unpaidCell.value = calc.unpaid || 0;
      unpaidCell.numFmt = "#,##0";
      unpaidCell.alignment = { horizontal: "right", vertical: "middle" };
      unpaidCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // Eligible (Column U, index 21)
      const eligibleCell = currentRow.getCell(21);
      eligibleCell.value = calc.isEligible ? "Yes" : "No";
      eligibleCell.alignment = { horizontal: "center", vertical: "middle" };
      eligibleCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // After V (Column V, index 22) = Register - (Already Paid + Unpaid)
      const afterVCell = currentRow.getCell(22);
      afterVCell.value = { formula: `R${rowNum}-(S${rowNum}+T${rowNum})` };
      afterVCell.numFmt = "#,##0";
      afterVCell.alignment = { horizontal: "right", vertical: "middle" };
      afterVCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // Percentage (Column W, index 23)
      const percentageCell = currentRow.getCell(23);
      percentageCell.value = calc.bonusPercent || 0;
      percentageCell.numFmt = "0.00";
      percentageCell.alignment = { horizontal: "right", vertical: "middle" };
      percentageCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // Actual (Column X, index 24)
      const actualCell = currentRow.getCell(24);
      actualCell.value = { formula: `IF(U${rowNum}="Yes",V${rowNum},0)` };
      actualCell.numFmt = "#,##0";
      actualCell.alignment = { horizontal: "right", vertical: "middle" };
      actualCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // Reim (Column Y, index 25) = After V - Actual
      const reimCell = currentRow.getCell(25);
      reimCell.value = { formula: `V${rowNum}-X${rowNum}` };
      reimCell.numFmt = "#,##0";
      reimCell.alignment = { horizontal: "right", vertical: "middle" };
      reimCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // Loan (Column Z, index 26)
      const loanCell = currentRow.getCell(26);
      loanCell.value = calc.loan || 0;
      loanCell.numFmt = "#,##0";
      loanCell.alignment = { horizontal: "right", vertical: "middle" };
      loanCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };

      // Final RTGS (Column AA, index 27) = Actual - Loan
      const finalRTGSCell = currentRow.getCell(27);
      finalRTGSCell.value = { formula: `X${rowNum}-Z${rowNum}` };
      finalRTGSCell.numFmt = "#,##0";
      finalRTGSCell.alignment = { horizontal: "right", vertical: "middle" };
      finalRTGSCell.border = {
        top: { style: "thin", color: { argb: "FFCCCCCC" } },
        bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
        left: { style: "thin", color: { argb: "FFCCCCCC" } },
        right: { style: "thin", color: { argb: "FFCCCCCC" } },
      };
    });

    // Grand totals row
    const dataEndRow = dataStartRow + calculations.length - 1;
    const totalsRowIndex = dataEndRow + 1;
    const totalsRow = worksheet.getRow(totalsRowIndex);
    totalsRow.height = 35;

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

    // After the grand totals row code, add month totals and year totals

    // Month Totals Row
    const monthTotalsRowIndex = totalsRowIndex + 1;
    const monthTotalsRow = worksheet.getRow(monthTotalsRowIndex);
    monthTotalsRow.height = 28;

    const monthTotalsLabel = worksheet.getCell(`A${monthTotalsRowIndex}`);
    monthTotalsLabel.value = "MONTH TOTALS";
    monthTotalsLabel.font = { bold: true };
    monthTotalsLabel.alignment = { horizontal: "center", vertical: "middle" };
    monthTotalsLabel.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFF59D" },
    };
    monthTotalsLabel.border = {
      top: { style: "medium", color: { argb: "FF000000" } },
      bottom: { style: "medium", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };

    // Month columns (E to O = NOV-24 to SEP-25)
    const monthCols = [
      { title: "NOV-24", col: "E" },
      { title: "DEC-24", col: "F" },
      { title: "JAN-25", col: "G" },
      { title: "FEB-25", col: "H" },
      { title: "MAR-25", col: "I" },
      { title: "APR-25", col: "J" },
      { title: "MAY-25", col: "K" },
      { title: "JUN-25", col: "L" },
      { title: "JUL-25", col: "M" },
      { title: "AUG-25", col: "N" },
      { title: "SEP-25", col: "O" },
    ];

    for (const col of monthCols) {
      const cell = worksheet.getCell(`${col.col}${monthTotalsRowIndex}`);
      cell.value = {
        formula: `SUM(${col.col}${dataStartRow}:${col.col}${dataEndRow})`,
      };
      cell.numFmt = "₹#,##0";
      cell.font = { bold: true };
      cell.alignment = { horizontal: "right", vertical: "middle" };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFF59D" },
      };
      cell.border = {
        top: { style: "medium", color: { argb: "FF000000" } },
        bottom: { style: "medium", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    }

    // Year Total (NOV-SEP) - one row below month totals
    const yearNovSepRowIndex = monthTotalsRowIndex + 1;
    const yearNovSepRow = worksheet.getRow(yearNovSepRowIndex);
    yearNovSepRow.height = 28;

    const yearNovSepLabelCell = worksheet.getCell(`A${yearNovSepRowIndex}`);
    yearNovSepLabelCell.value = "YEAR TOTAL (NOV–SEP)";
    yearNovSepLabelCell.font = { bold: true };
    yearNovSepLabelCell.alignment = {
      horizontal: "center",
      vertical: "middle",
    };
    yearNovSepLabelCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFD54F" },
    };
    yearNovSepLabelCell.border = {
      top: { style: "medium", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };

    const yearNovSepTotalCell = worksheet.getCell(`Q${yearNovSepRowIndex}`);
    yearNovSepTotalCell.value = {
      formula: `SUM(E${monthTotalsRowIndex}:O${monthTotalsRowIndex})`,
    };
    yearNovSepTotalCell.numFmt = "₹#,##0";
    yearNovSepTotalCell.font = { bold: true };
    yearNovSepTotalCell.alignment = { horizontal: "right", vertical: "middle" };
    yearNovSepTotalCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFD54F" },
    };
    yearNovSepTotalCell.border = {
      top: { style: "medium", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };

    // Full Year Total (NOV–SEP + Salary12) - below NOV-SEP total
    const fullYearRowIndex = yearNovSepRowIndex + 1;
    const fullYearRow = worksheet.getRow(fullYearRowIndex);
    fullYearRow.height = 28;

    const fullYearLabelCell = worksheet.getCell(`A${fullYearRowIndex}`);
    fullYearLabelCell.value = "YEAR TOTAL (NOV–SEP + Salary12)";
    fullYearLabelCell.font = { bold: true };
    fullYearLabelCell.alignment = { horizontal: "center", vertical: "middle" };
    fullYearLabelCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFD54F" },
    };
    fullYearLabelCell.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "medium", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };

    const fullYearTotalCell = worksheet.getCell(`Q${fullYearRowIndex}`);
    fullYearTotalCell.value = {
      formula: `SUM(E${monthTotalsRowIndex}:P${monthTotalsRowIndex})`,
    };
    fullYearTotalCell.numFmt = "₹#,##0";
    fullYearTotalCell.font = { bold: true };
    fullYearTotalCell.alignment = { horizontal: "right", vertical: "middle" };
    fullYearTotalCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFD54F" },
    };
    fullYearTotalCell.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "medium", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };

    const totalColumns = [17, 18, 19, 20, 22, 24, 25, 26, 27]; // Q, R, S, T, V, X, Y, Z, AA

    const colLetter = (n: number) => {
      const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
      if (n <= 26) return letters[n - 1];
      const first = letters[Math.floor((n - 1) / 26) - 1];
      const second = letters[(n - 1) % 26];
      return `${first}${second}`;
    };

    totalColumns.forEach((ci) => {
      const cell = totalsRow.getCell(ci);
      const letter = colLetter(ci);
      cell.value = {
        formula: `SUM(${letter}${dataStartRow}:${letter}${dataEndRow})`,
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

    // Column widths
    const colWidths = [
      50,
      80,
      180,
      90, // SR.No, EMP Code, Name, DOJ
      80,
      80,
      80,
      80,
      80,
      80,
      80,
      80,
      80,
      80,
      80, // 11 months
      100,
      100,
      100,
      100,
      100,
      80,
      100,
      80,
      100,
      100,
      100,
      110, // Salary12 to Final RTGS
    ];
    colWidths.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width / 7;
    });

    worksheet.getRow(1).height = 35;
    worksheet.getRow(2).height = 15;
    worksheet.getRow(3).height = 30;
  }

  static async computeHRMonthlyTotals(
    hrBuffer: ArrayBuffer
  ): Promise<Map<string, number>> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(hrBuffer);

    const names = [
      "JAN",
      "FEB",
      "MAR",
      "APR",
      "MAY",
      "JUN",
      "JUL",
      "AUG",
      "SEP",
      "OCT",
      "NOV",
      "DEC",
    ];
    const toMonthKey = (val: any): string | null => {
      if (!val) return null;
      if (val instanceof Date) {
        const m = val.getMonth();
        const y = (val.getFullYear() % 100).toString().padStart(2, "0");
        return `${names[m]}-${y}`;
      }
      const s = String(val).trim();

      // Match patterns like "2024-11-01" or "Nov-24" or "NOV-24"
      const m1 = s.match(/\b(20\d{2})-(\d{2})-(\d{2})/);
      if (m1) {
        const mm = parseInt(m1[2], 10) - 1;
        const yy = (parseInt(m1[1], 10) % 100).toString().padStart(2, "0");
        return `${names[mm]}-${yy}`;
      }
      const m2 = s.match(
        /^(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[- ]?(\d{2,4})$/i
      );
      if (m2) {
        const mon = m2[1].toUpperCase();
        const yr = m2[2];
        const y2 = (yr.length === 4 ? yr.slice(2) : yr).padStart(2, "0");
        return `${mon}-${y2}`;
      }
      return null;
    };

    const totals = new Map<string, number>();

    for (const ws of workbook.worksheets) {
      const wsName = (ws.name || "").toUpperCase();
      console.log(`\n📊 Processing sheet: ${ws.name}`);

      // Skip loan deduction sheet
      if (wsName.includes("LOAN")) {
        console.log(`  ⏭️  Skipping loan sheet`);
        continue;
      }

      // Find header row (contains "Sr.No." in column A)
      let headerRowIndex = -1;
      for (let r = 1; r <= Math.min(ws.rowCount, 15); r++) {
        const cellA = String(ws.getRow(r).getCell(1).value || "").toUpperCase();
        if (
          cellA.includes("SR") &&
          (cellA.includes("NO") || cellA.includes("."))
        ) {
          headerRowIndex = r;
          console.log(`  ✓ Found header at row ${r}`);
          break;
        }
      }

      if (headerRowIndex === -1) {
        console.log(`  ❌ No header row found`);
        continue;
      }

      const header = ws.getRow(headerRowIndex);
      const monthCols: Array<{ col: number; key: string }> = [];

      // Scan all columns for month headers
      console.log(`  🔍 Scanning for month columns...`);
      for (let c = 1; c <= Math.min(ws.columnCount, 30); c++) {
        const cellValue = header.getCell(c).value;
        const key = toMonthKey(cellValue);

        if (key) {
          monthCols.push({ col: c, key });
          console.log(`     Column ${c}: ${key} = "${cellValue}"`);
        }
      }

      if (monthCols.length === 0) {
        console.log(`  ❌ No month columns found`);
        continue;
      }

      console.log(`  ✓ Found ${monthCols.length} month columns`);

      // Find "Grand Total" row - check multiple columns
      let grandTotalRowIndex = -1;
      console.log(`  🔍 Searching for Grand Total row...`);

      for (let r = headerRowIndex + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);

        // Check columns A through E for "Grand Total"
        for (let c = 1; c <= 5; c++) {
          const cellValue = String(row.getCell(c).value || "")
            .trim()
            .toUpperCase();
          if (cellValue.includes("GRAND") && cellValue.includes("TOTAL")) {
            grandTotalRowIndex = r;
            console.log(
              `  ✓ Found "Grand Total" at row ${r}, column ${c}: "${
                row.getCell(c).value
              }"`
            );
            break;
          }
        }

        if (grandTotalRowIndex !== -1) break;
      }

      if (grandTotalRowIndex === -1) {
        console.warn(`  ⚠️  No Grand Total row found in sheet ${ws.name}`);
        continue;
      }

      // Extract totals from Grand Total row
      const grandTotalRow = ws.getRow(grandTotalRowIndex);
      console.log(`  📊 Extracting totals from row ${grandTotalRowIndex}:`);

      for (const mc of monthCols) {
        const cell = grandTotalRow.getCell(mc.col);
        const raw = (cell as any).result ?? cell.value;
        let n = 0;

        if (typeof raw === "number") {
          n = raw;
        } else {
          const parsed = Number(String(raw ?? "").replace(/[^0-9.-]/g, ""));
          n = isFinite(parsed) ? parsed : 0;
        }

        if (n > 0) {
          const existing = totals.get(mc.key) || 0;
          totals.set(mc.key, existing + n);
          console.log(
            `     ${mc.key}: ${n.toLocaleString()} (cumulative: ${(
              existing + n
            ).toLocaleString()})`
          );
        } else {
          console.log(`     ${mc.key}: 0 or empty`);
        }
      }
    }

    console.log(`\n✅ Final HR Monthly Totals:`);
    totals.forEach((value, key) => {
      console.log(`   ${key}: ₹${value.toLocaleString()}`);
    });

    return totals;
  }

  static async computeOurMonthlyTotals(
    staffBuffer: ArrayBuffer,
    workerBuffer: ArrayBuffer
  ): Promise<Map<string, number>> {
    const totals = new Map<string, number>();

    console.log("\n📊 Computing OUR monthly totals...");

    // Process Staff file (Column R = 18)
    const staffWb = new ExcelJS.Workbook();
    await staffWb.xlsx.load(staffBuffer);

    for (const ws of staffWb.worksheets) {
      if (!ws.name.includes("-") || !ws.name.endsWith(" O")) continue;

      const monthKey = this.normalizeMonthKey(ws.name);
      console.log(`\n  Processing Staff sheet: ${ws.name} -> ${monthKey}`);

      let grandTotalFound = false;
      let staffGross = 0;

      // Look for the Grand Total row more comprehensively
      for (
        let rowNum = ws.rowCount;
        rowNum > Math.max(1, ws.rowCount - 20);
        rowNum--
      ) {
        const row = ws.getRow(rowNum);

        // Check first few columns for "Grand Total" or just "Total"
        for (let col = 1; col <= 5; col++) {
          const cellValue = String(row.getCell(col).value || "")
            .trim()
            .toUpperCase();

          if (
            (cellValue.includes("GRAND") && cellValue.includes("TOTAL")) ||
            (cellValue === "TOTAL" && !grandTotalFound)
          ) {
            // Get the value from Column R (18)
            const cell = row.getCell(18);
            const value = (cell as any).result ?? cell.value;
            staffGross = toNumber(cell);

            if (staffGross > 0) {
              grandTotalFound = true;
              console.log(
                `    ✓ Found Staff Grand Total at row ${rowNum}: ₹${staffGross.toLocaleString()}`
              );
              totals.set(`STAFF_${monthKey}`, staffGross);
              break;
            }
          }
        }
        if (grandTotalFound) break;
      }

      if (!grandTotalFound) {
        console.log(`    ⚠️ No Grand Total found in Staff sheet ${ws.name}`);
      }
    }

    // Process Worker file (Column I = 9)
    const workerWb = new ExcelJS.Workbook();
    await workerWb.xlsx.load(workerBuffer);

    for (const ws of workerWb.worksheets) {
      if (!ws.name.includes("-") || !ws.name.endsWith(" W")) continue;

      const monthKey = this.normalizeMonthKey(ws.name);
      console.log(`\n  Processing Worker sheet: ${ws.name} -> ${monthKey}`);

      let grandTotalFound = false;
      let workerGross = 0;

      // Look for the Grand Total row
      for (
        let rowNum = ws.rowCount;
        rowNum > Math.max(1, ws.rowCount - 20);
        rowNum--
      ) {
        const row = ws.getRow(rowNum);

        // Check first few columns for "Grand Total" or just "Total"
        for (let col = 1; col <= 5; col++) {
          const cellValue = String(row.getCell(col).value || "")
            .trim()
            .toUpperCase();

          if (
            (cellValue.includes("GRAND") && cellValue.includes("TOTAL")) ||
            (cellValue === "TOTAL" && !grandTotalFound) ||
            cellValue.includes("TOTAL")
          ) {
            // Get the value from Column I (9)
            const cell = row.getCell(9);
            const value = (cell as any).result ?? cell.value;
            workerGross = toNumber(cell);

            if (workerGross > 0) {
              grandTotalFound = true;
              console.log(
                `    ✓ Found Worker Grand Total at row ${rowNum}: ₹${workerGross.toLocaleString()}`
              );

              // Add to existing staff total
              const existing = totals.get(`STAFF_${monthKey}`) || 0;
              const combined = existing + workerGross;
              totals.set(monthKey, combined); // Final combined total
              console.log(
                `    📊 Combined total for ${monthKey}: ₹${combined.toLocaleString()}`
              );
              break;
            }
          }
        }
        if (grandTotalFound) break;
      }

      if (!grandTotalFound) {
        console.log(`    ⚠️ No Grand Total found in Worker sheet ${ws.name}`);
        // Still set the combined total even if worker total is 0
        const existing = totals.get(`STAFF_${monthKey}`) || 0;
        if (existing > 0) {
          totals.set(monthKey, existing);
        }
      }
    }

    console.log("\n✅ Final OUR Monthly Totals:");
    totals.forEach((value, key) => {
      if (!key.startsWith("STAFF_")) {
        console.log(`   ${key}: ₹${value.toLocaleString()}`);
      }
    });

    return totals;
  }

  private static toMonthKey(value: string): string | null {
    if (!value) return null;
    const s = String(value).toUpperCase().trim();
    const m =
      s.match(/^([A-Z]+)\s*[-/]\s*(\d{2,4})/) ||
      s.match(/^([A-Z]+)\s+(\d{2,4})/);
    if (!m) return null;

    const monthMap: Record<string, string> = {
      JAN: "JAN",
      JANUARY: "JAN",
      FEB: "FEB",
      FEBRUARY: "FEB",
      MAR: "MAR",
      MARCH: "MAR",
      APR: "APR",
      APRIL: "APR",
      MAY: "MAY",
      JUN: "JUN",
      JUNE: "JUN",
      JUL: "JUL",
      JULY: "JUL",
      AUG: "AUG",
      AUGUST: "AUG",
      SEP: "SEP",
      SEPT: "SEP",
      SEPTEMBER: "SEP",
      OCT: "OCT",
      OCTOBER: "OCT",
      NOV: "NOV",
      NOVEMBER: "NOV",
      DEC: "DEC",
      DECEMBER: "DEC",
    };

    const mon = monthMap[m[1]] || m[1].slice(0, 3);
    const yr = m[2].length === 4 ? m[2].slice(2) : m[2];
    return `${mon}-${yr}`;
  }

  static async computeHRGrossTotals(buffer: ArrayBuffer): Promise<{
    hrMonthTotals: Record<string, number>;
    hrYearTotals: Record<string, number>;
    hrGrandTotal: number;
  }> {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      const hrMonthTotals: Record<string, number> = {};
      const hrYearTotals: Record<string, number> = {};
      let hrGrandTotal = 0;

      ["Staff", "Worker"].forEach((sheetName) => {
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) return;

        // Find header row (assuming row 3 based on your structure)
        const headerRow = worksheet.getRow(3);
        const monthColumns: Array<{ col: number; month: string }> = [];
        let grossSalaryCol = 0;

        // Scan header to find month columns and Gross Salary column
        headerRow.eachCell((cell, colNumber) => {
          const cellValue = String(cell.value || "").trim();
          const monthKey = this.toMonthKey(cellValue);

          if (monthKey) {
            monthColumns.push({ col: colNumber, month: monthKey });
          }

          // Find Gross Salary column (or similar)
          if (
            cellValue.toLowerCase().includes("gross") &&
            (cellValue.toLowerCase().includes("salary") ||
              cellValue.toLowerCase() === "gross")
          ) {
            grossSalaryCol = colNumber;
          }
        });

        console.log(
          `${sheetName}: Found ${monthColumns.length} month columns, Gross Salary at col ${grossSalaryCol}`
        );

        // Sum each month column
        monthColumns.forEach(({ col, month }) => {
          let monthSum = 0;
          worksheet.eachRow((row, rowNumber) => {
            if (rowNumber <= 3) return; // Skip header rows
            const value = toNumber(row.getCell(col));
            if (value > 0) monthSum += value;
          });

          hrMonthTotals[month] = (hrMonthTotals[month] || 0) + monthSum;

          // Add to year total
          const year = month.split("-")[1];
          hrYearTotals[`20${year}`] =
            (hrYearTotals[`20${year}`] || 0) + monthSum;
        });

        // Sum Gross Salary column for grand total
        if (grossSalaryCol > 0) {
          worksheet.eachRow((row, rowNumber) => {
            if (rowNumber <= 3) return;
            const value = toNumber(row.getCell(grossSalaryCol));
            if (value > 0) hrGrandTotal += value;
          });
        }
      });

      console.log(
        `✅ HR Totals - Months: ${
          Object.keys(hrMonthTotals).length
        }, Grand: ${hrGrandTotal}`
      );
      return { hrMonthTotals, hrYearTotals, hrGrandTotal };
    } catch (error) {
      console.error("❌ Error computing HR gross totals:", error);
      return { hrMonthTotals: {}, hrYearTotals: {}, hrGrandTotal: 0 };
    }
  }

  private static async generateSummarySheet(
    workbook: ExcelJS.Workbook,
    departments: { [key: string]: BonusCalculation[] },
    hrMonthlyTotals?: Map<string, number>,
    staffBuffer?: ArrayBuffer,
    workerBuffer?: ArrayBuffer
  ): Promise<void> {
    const worksheet = workbook.addWorksheet("Summary");
    const title = "DEPARTMENT-WISE BONUS SUMMARY - DIWALI 2024-25";

    // First section: Department-wise summary
    const headers = [
      "Department",
      "Employees",
      "Total Gross Salary",
      "Total Bonus",
      "Average Bonus",
    ];

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

    worksheet.getRow(2).height = 15;

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
        if (colIndex === 0)
          cell.alignment = { horizontal: "left", vertical: "middle" };
        else if (colIndex >= 2) {
          cell.alignment = { horizontal: "right", vertical: "middle" };
          cell.numFmt = "#,##0.00";
        } else cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "thin", color: { argb: "FFBBBBBB" } },
          bottom: { style: "thin", color: { argb: "FFBBBBBB" } },
          left: { style: "thin", color: { argb: "FFBBBBBB" } },
          right: { style: "thin", color: { argb: "FFBBBBBB" } },
        };
      });
    });

    const emptyRowIndex = rowIndex + Object.keys(departments).length;
    worksheet.getRow(emptyRowIndex).height = 15;

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
      if (colIndex === 0)
        cell.alignment = { horizontal: "center", vertical: "middle" };
      else if (colIndex >= 2) {
        cell.alignment = { horizontal: "right", vertical: "middle" };
        cell.numFmt = "#,##0.00";
      } else cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "double", color: { argb: "FF000000" } },
        bottom: { style: "double", color: { argb: "FF000000" } },
        left: { style: "medium", color: { argb: "FF000000" } },
        right: { style: "medium", color: { argb: "FF000000" } },
      };
    });

    // NEW SECTION: Monthly Comparison
    let currentRow = grandTotalRowIndex + 3;

    // Monthly Comparison Title
    worksheet.mergeCells(`A${currentRow}:D${currentRow}`);
    const monthTitleCell = worksheet.getCell(`A${currentRow}`);
    monthTitleCell.value =
      "MONTHLY GROSS SALARY COMPARISON (OUR SHEET vs HR SHEET)";
    monthTitleCell.font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
    monthTitleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1565C0" },
    };
    monthTitleCell.alignment = { horizontal: "center", vertical: "middle" };
    worksheet.getRow(currentRow).height = 35;
    currentRow++;

    worksheet.getRow(currentRow).height = 15;
    currentRow++;

    // Monthly Comparison Headers
    const monthHeaders = [
      "Month",
      "Our Sheet Total",
      "HR Sheet Total",
      "Difference",
    ];
    const monthHeaderRow = worksheet.getRow(currentRow);
    monthHeaders.forEach((header, index) => {
      const cell = monthHeaderRow.getCell(index + 1);
      cell.value = header;
      cell.font = { bold: true, size: 12, color: { argb: "FF000000" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF90CAF9" },
      };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "medium", color: { argb: "FF000000" } },
        bottom: { style: "medium", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });
    currentRow++;

    // Calculate monthly totals from our calculations
    const monthlyTotals =
      staffBuffer && workerBuffer
        ? await this.computeOurMonthlyTotals(staffBuffer, workerBuffer)
        : new Map<string, number>();

    const monthOrder = [
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
    ];

    let ourYearTotal = 0;
    let hrYearTotal = 0;

    // Monthly data rows
    monthOrder.forEach((month) => {
      const ourTotal = Math.round(monthlyTotals.get(month) || 0);
      const hrTotal = Math.round(hrMonthlyTotals?.get(month) || 0);
      const difference = ourTotal - hrTotal;
      ourYearTotal += ourTotal;
      hrYearTotal += hrTotal;

      const dataRow = worksheet.getRow(currentRow);
      dataRow.height = 25;

      const rowData = [month, ourTotal, hrTotal, difference];
      rowData.forEach((value, colIndex) => {
        const cell = dataRow.getCell(colIndex + 1);
        cell.value = value;

        if (colIndex === 0) {
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else {
          cell.alignment = { horizontal: "right", vertical: "middle" };
          cell.numFmt = "#,##0";
        }

        // Highlight difference in red if > 1
        if (colIndex === 3 && Math.abs(difference) > 1) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF0000" },
          };
          cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
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

    // Yearwise Total Row
    currentRow++;
    const yearTotalRow = worksheet.getRow(currentRow);
    yearTotalRow.height = 30;

    const yearDifference = ourYearTotal - hrYearTotal;
    const yearData = [
      "YEAR TOTAL (NOV-OCT)",
      ourYearTotal,
      hrYearTotal,
      yearDifference,
    ];

    yearData.forEach((value, colIndex) => {
      const cell = yearTotalRow.getCell(colIndex + 1);
      cell.value = value;
      cell.font = { bold: true, size: 12 };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFD54F" },
      };

      if (colIndex === 0) {
        cell.alignment = { horizontal: "center", vertical: "middle" };
      } else {
        cell.alignment = { horizontal: "right", vertical: "middle" };
        cell.numFmt = "#,##0";
      }

      // Highlight year difference in red if > 1
      if (colIndex === 3 && Math.abs(yearDifference) > 1) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF0000" },
        };
        cell.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
      }

      cell.border = {
        top: { style: "medium", color: { argb: "FF000000" } },
        bottom: { style: "medium", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });

    // Grand Gross Total Comparison
    currentRow += 3;
    worksheet.mergeCells(`A${currentRow}:D${currentRow}`);
    const grandGrossTitleCell = worksheet.getCell(`A${currentRow}`);
    grandGrossTitleCell.value = "GRAND GROSS SALARY COMPARISON";
    grandGrossTitleCell.font = {
      bold: true,
      size: 16,
      color: { argb: "FFFFFFFF" },
    };
    grandGrossTitleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE65100" },
    };
    grandGrossTitleCell.alignment = {
      horizontal: "center",
      vertical: "middle",
    };
    worksheet.getRow(currentRow).height = 35;
    currentRow++;

    worksheet.getRow(currentRow).height = 15;
    currentRow++;

    const grandGrossHeaders = [
      "Description",
      "Our Sheet",
      "HR Sheet",
      "Difference",
    ];
    const grandGrossHeaderRow = worksheet.getRow(currentRow);
    grandGrossHeaders.forEach((header, index) => {
      const cell = grandGrossHeaderRow.getCell(index + 1);
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
    currentRow++;

    const ourGrandTotal = Math.round(grandTotalSalary);
    const hrGrandTotal = hrYearTotal;
    const grandDifference = ourGrandTotal - hrGrandTotal;

    const grandGrossRow = worksheet.getRow(currentRow);
    grandGrossRow.height = 35;

    const grandGrossData = [
      "Grand Gross Total",
      ourGrandTotal,
      hrGrandTotal,
      grandDifference,
    ];
    grandGrossData.forEach((value, colIndex) => {
      const cell = grandGrossRow.getCell(colIndex + 1);
      cell.value = value;
      cell.font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE65100" },
      };

      if (colIndex === 0) {
        cell.alignment = { horizontal: "left", vertical: "middle" };
      } else {
        cell.alignment = { horizontal: "right", vertical: "middle" };
        cell.numFmt = "#,##0";
      }

      // Highlight grand difference in red if > 1
      if (colIndex === 3 && Math.abs(grandDifference) > 1) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF0000" },
        };
      }

      cell.border = {
        top: { style: "double", color: { argb: "FF000000" } },
        bottom: { style: "double", color: { argb: "FF000000" } },
        left: { style: "medium", color: { argb: "FF000000" } },
        right: { style: "medium", color: { argb: "FF000000" } },
      };
    });

    // Set column widths
    const summaryColWidths = [150, 150, 150, 150];
    summaryColWidths.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width / 7;
    });

    worksheet.getRow(1).height = 40;
    worksheet.getRow(2).height = 15;
    worksheet.getRow(3).height = 35;
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

    worksheet.getRow(2).height = 15;

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

    const groupedComparisons = comparisons.reduce((acc, comp) => {
      const dept = comp.department || "Unknown";
      if (!acc[dept]) acc[dept] = [];
      acc[dept].push(comp);
      return acc;
    }, {} as { [key: string]: any[] });

    let currentRow = 4;

    Object.keys(groupedComparisons).forEach((dept) => {
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
      currentRow++;
    });

    const compColWidths = [80, 200, 100, 120, 120, 120, 100];
    compColWidths.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width / 7;
    });

    worksheet.getRow(1).height = 35;
    worksheet.getRow(2).height = 15;
    worksheet.getRow(3).height = 30;

    return await workbook.xlsx.writeBuffer();
  }

  private static parseDate(dateValue: any): Date {
    // Handle null, undefined, or empty strings as invalid, return default date
    const s = String(dateValue ?? "")
      .trim()
      .toUpperCase();
    if (!dateValue || s === "N" || s === "NA" || s === "N.A" || s === "N/A") {
      console.warn(
        `⚠️ Could not parse date: ${dateValue}, using default 2020-01-01`
      );
      return new Date("2020-01-01");
    }

    // If already a Date object, return it
    if (dateValue instanceof Date) return dateValue;

    // Handle Excel serial number
    const n = Number(String(dateValue).trim());
    if (typeof dateValue === "number" && isFinite(n) && n > 0) {
      const ms = Math.round((n - 25569) * 86400 * 1000);
      const parsedDate = new Date(ms);
      if (!isNaN(parsedDate.getTime())) return parsedDate;
    }

    // Handle string dates (ISO, dd-mm-yyyy, mm-dd-yyyy, etc.) using native Date constructor
    const parsedDate = new Date(String(dateValue));
    if (!isNaN(parsedDate.getTime())) return parsedDate;

    // If parsing fails, return default date with warning
    console.warn(
      `⚠️ Could not parse date: ${dateValue}, using default 2020-01-01`
    );
    return new Date("2020-01-01");
  }

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
          const salary1 = monthData.salary || 0;
          const grossSalary = salary1;
          monthlySummary[monthData.month].totalSalary1 += salary1;
          monthlySummary[monthData.month].totalGrossSalary += grossSalary;
          monthlySummary[monthData.month].count++;
          totalSalary1 += salary1;
          totalGrossSalary += grossSalary;
        });
      }
    });

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
          const salary1 = monthData.salary || 0;
          const grossSalary = salary1;
          monthlySummary[monthData.month].totalSalary1 += salary1;
          monthlySummary[monthData.month].totalGrossSalary += grossSalary;
          monthlySummary[monthData.month].count++;
          totalSalary1 += salary1;
          totalGrossSalary += grossSalary;
        });
      }
    });

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
    worksheet.getRow(2).height = 15;

    let currentRow = 3;

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

    let staffTotalGross = 0,
      staffTotalSalary1 = 0,
      staffTotalEmployees = 0;
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
        if (colIndex === 0)
          cell.alignment = { horizontal: "center", vertical: "middle" };
        else if (colIndex >= 1 && colIndex <= 4) {
          cell.alignment = { horizontal: "right", vertical: "middle" };
          if (colIndex >= 1 && colIndex <= 2) cell.numFmt = "₹#,##0";
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
      if (colIndex === 0)
        cell.alignment = { horizontal: "center", vertical: "middle" };
      else if (colIndex >= 1 && colIndex <= 4) {
        cell.alignment = { horizontal: "right", vertical: "middle" };
        if (colIndex >= 1 && colIndex <= 2) cell.numFmt = "₹#,##0";
      }
      cell.border = {
        top: { style: "medium", color: { argb: "FF000000" } },
        bottom: { style: "medium", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });
    currentRow += 2;

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

    let workerTotalGross = 0,
      workerTotalSalary1 = 0,
      workerTotalEmployees = 0;
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
        if (colIndex === 0)
          cell.alignment = { horizontal: "center", vertical: "middle" };
        else if (colIndex >= 1 && colIndex <= 4) {
          cell.alignment = { horizontal: "right", vertical: "middle" };
          if (colIndex >= 1 && colIndex <= 2) cell.numFmt = "₹#,##0";
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
      if (colIndex === 0)
        cell.alignment = { horizontal: "center", vertical: "middle" };
      else if (colIndex >= 1 && colIndex <= 4) {
        cell.alignment = { horizontal: "right", vertical: "middle" };
        if (colIndex >= 1 && colIndex <= 2) cell.numFmt = "₹#,##0";
      }
      cell.border = {
        top: { style: "medium", color: { argb: "FF000000" } },
        bottom: { style: "medium", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });
    currentRow += 2;

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
      if (colIndex === 0)
        cell.alignment = { horizontal: "center", vertical: "middle" };
      else if (colIndex >= 1 && colIndex <= 4) {
        cell.alignment = { horizontal: "right", vertical: "middle" };
        if (colIndex >= 1 && colIndex <= 2) cell.numFmt = "₹#,##0";
      }
      cell.border = {
        top: { style: "double", color: { argb: "FF000000" } },
        bottom: { style: "double", color: { argb: "FF000000" } },
        left: { style: "medium", color: { argb: "FF000000" } },
        right: { style: "medium", color: { argb: "FF000000" } },
      };
    });

    const colWidths = [120, 150, 150, 120, 120, 50];
    colWidths.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width / 7;
    });

    worksheet.getRow(1).height = 35;
    worksheet.getRow(2).height = 15;
  }
}
