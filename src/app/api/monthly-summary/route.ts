
import { NextRequest, NextResponse } from "next/server";
import * as XLSX from "xlsx";

interface MonthlyData {
  month: string;
  totalEmployees: number;
  totalGrossSalary: number;
  totalSalary1: number;
  totalBasicSalary: number;
  totalHRA: number;
  totalDA: number;
  totalPF: number;
  totalESIC: number;
  totalOvertimePayment: number;
  totalWorkingDays: number;
  totalLeaveDays: number;
  totalLateDays: number;
  avgGrossSalary: number;
  avgSalary1: number;
}

interface CombinedSummary {
  staff: MonthlyData[];
  worker: MonthlyData[];
  combined: MonthlyData[];
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const staffFile = formData.get("staffTulsi") as File;
    const workerFile = formData.get("workerTulsi") as File;

    if (!staffFile) {
      return NextResponse.json(
        { error: "Staff file is required" },
        { status: 400 }
      );
    }

    // Process Staff File
    const staffBuffer = Buffer.from(await staffFile.arrayBuffer());
    const staffWorkbook = XLSX.read(staffBuffer, { type: "buffer" });
    const staffMonthlyData: MonthlyData[] = [];

    // Process Staff sheets (ending with " O")
    staffWorkbook.SheetNames.filter(name => name.endsWith(" O")).forEach((sheetName) => {
      try {
        const worksheet = staffWorkbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: 2 });

        if (!data || data.length < 4) return;

        const headers = data[2] as string[];
        const rows = data.slice(3);

        const getColumnIndex = (columnName: string) => {
          return headers.findIndex(
            (header) =>
              header &&
              header.toString().toLowerCase().includes(columnName.toLowerCase())
          );
        };

        const empNameCol = getColumnIndex("employee name");
        const grossSalaryCol = getColumnIndex("gross salary");
        const salary1Col = getColumnIndex("salary1");
        const basicCol = getColumnIndex("basic");
        const hraCol = getColumnIndex("hra");
        const daCol = getColumnIndex("d.a.");
        const pfCol = getColumnIndex("pf-12%");
        const esicCol = getColumnIndex("esic");
        const otCol = getColumnIndex("net ot");
        const wdCol = getColumnIndex("wd");
        const ldCol = getColumnIndex("ld");
        const lateCol = getColumnIndex("actual late");

        const validRows = rows.filter(
          (row: unknown): row is unknown[] =>
            Array.isArray(row) &&
            row[empNameCol] &&
            typeof row[empNameCol] === "string" &&
            row[empNameCol].toString().trim() !== ""
        );

        if (validRows.length === 0) return;

        const monthDisplay = sheetName.replace(" O", "");

        const summary: MonthlyData = {
          month: monthDisplay,
          totalEmployees: validRows.length,
          totalGrossSalary: validRows.reduce(
            (sum: number, row: unknown[]) =>
              sum + (parseFloat(String(row[grossSalaryCol])) || 0),
            0
          ),
          totalSalary1: validRows.reduce(
            (sum: number, row: unknown[]) =>
              sum + (parseFloat(String(row[salary1Col])) || 0),
            0
          ),
          totalBasicSalary: validRows.reduce(
            (sum: number, row: unknown[]) =>
              sum + (parseFloat(String(row[basicCol])) || 0),
            0
          ),
          totalHRA: validRows.reduce(
            (sum: number, row: unknown[]) =>
              sum + (parseFloat(String(row[hraCol])) || 0),
            0
          ),
          totalDA: validRows.reduce(
            (sum: number, row: unknown[]) =>
              sum + (parseFloat(String(row[daCol])) || 0),
            0
          ),
          totalPF: validRows.reduce(
            (sum: number, row: unknown[]) =>
              sum + (parseFloat(String(row[pfCol])) || 0),
            0
          ),
          totalESIC: validRows.reduce(
            (sum: number, row: unknown[]) =>
              sum + (parseFloat(String(row[esicCol])) || 0),
            0
          ),
          totalOvertimePayment: validRows.reduce(
            (sum: number, row: unknown[]) =>
              sum + (parseFloat(String(row[otCol])) || 0),
            0
          ),
          totalWorkingDays: validRows.reduce(
            (sum: number, row: unknown[]) =>
              sum + (parseFloat(String(row[wdCol])) || 0),
            0
          ),
          totalLeaveDays: validRows.reduce(
            (sum: number, row: unknown[]) =>
              sum + (parseFloat(String(row[ldCol])) || 0),
            0
          ),
          totalLateDays: validRows.reduce(
            (sum: number, row: unknown[]) =>
              sum + (parseFloat(String(row[lateCol])) || 0),
            0
          ),
          avgGrossSalary: 0,
          avgSalary1: 0,
        };

        if (summary.totalEmployees > 0) {
          summary.avgGrossSalary = summary.totalGrossSalary / summary.totalEmployees;
          summary.avgSalary1 = summary.totalSalary1 / summary.totalEmployees;
        }

        staffMonthlyData.push(summary);
      } catch (error) {
        console.log(`Error processing staff sheet ${sheetName}:`, error);
      }
    });

    // Process Worker File (if provided)
    const workerMonthlyData: MonthlyData[] = [];

    if (workerFile) {
      try {
        const workerBuffer = Buffer.from(await workerFile.arrayBuffer());
        const workerWorkbook = XLSX.read(workerBuffer, { type: "buffer" });

        // Process Worker sheets (ending with " W")
        workerWorkbook.SheetNames.filter(name => name.endsWith(" W")).forEach((sheetName) => {
          try {
            const worksheet = workerWorkbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: 2 });

            if (!data || data.length < 4) return;

            // Worker file has different structure - employee name is usually at index 3
            // Salary columns are at different positions
            const rows = data.slice(1); // Skip header row

            const validRows = rows.filter(
              (row: unknown): row is unknown[] =>
                Array.isArray(row) &&
                row[3] && // Name is at index 3
                typeof row[3] === "string" &&
                row[3].toString().trim() !== "" &&
                row[3] !== "EMPLOYEE NAME" // Skip header
            );

            if (validRows.length === 0) return;

            const monthDisplay = sheetName.replace(" W", "");

            // Worker salary structure analysis based on the data we saw:
            // Index 8: appears to be gross salary
            // Index 12: appears to be net salary
            // Index 4: working days
            // Index 5: leave days

            const summary: MonthlyData = {
              month: monthDisplay,
              totalEmployees: validRows.length,
              totalGrossSalary: validRows.reduce(
                (sum: number, row: unknown[]) =>
                  sum + (parseFloat(String(row[8])) || 0), // Gross salary at index 8
                0
              ),
              totalSalary1: validRows.reduce(
                (sum: number, row: unknown[]) =>
                  sum + (parseFloat(String(row[12])) || 0), // Net salary at index 12
                0
              ),
              totalBasicSalary: validRows.reduce(
                (sum: number, row: unknown[]) =>
                  sum + (parseFloat(String(row[7])) || 0), // Basic at index 7
                0
              ),
              totalHRA: 0, // Not clearly identifiable in worker data
              totalDA: 0,  // Not clearly identifiable in worker data
              totalPF: validRows.reduce(
                (sum: number, row: unknown[]) =>
                  sum + (parseFloat(String(row[10])) || 0), // PF at index 10
                0
              ),
              totalESIC: 0, // Not clearly identifiable
              totalOvertimePayment: validRows.reduce(
                (sum: number, row: unknown[]) =>
                  sum + (parseFloat(String(row[11])) || 0), // OT at index 11
                0
              ),
              totalWorkingDays: validRows.reduce(
                (sum: number, row: unknown[]) =>
                  sum + (parseFloat(String(row[4])) || 0), // WD at index 4
                0
              ),
              totalLeaveDays: validRows.reduce(
                (sum: number, row: unknown[]) =>
                  sum + (parseFloat(String(row[5])) || 0), // LD at index 5
                0
              ),
              totalLateDays: 0, // Not clearly identifiable
              avgGrossSalary: 0,
              avgSalary1: 0,
            };

            if (summary.totalEmployees > 0) {
              summary.avgGrossSalary = summary.totalGrossSalary / summary.totalEmployees;
              summary.avgSalary1 = summary.totalSalary1 / summary.totalEmployees;
            }

            workerMonthlyData.push(summary);
          } catch (error) {
            console.log(`Error processing worker sheet ${sheetName}:`, error);
          }
        });
      } catch (error) {
        console.log("Error processing worker file:", error);
      }
    }

    // Sort both arrays by month order
    const monthOrder = [
      "NOV-24", "DEC-24", "JAN-25", "FEB-25", "MAR-25", "APR-25",
      "MAY-25", "JUN-25", "JULY-25", "AUG-25", "SEP-25", "OCT-25"
    ];

    const sortByMonth = (a: MonthlyData, b: MonthlyData) => {
      const aIndex = monthOrder.indexOf(a.month);
      const bIndex = monthOrder.indexOf(b.month);
      return (aIndex === -1 ? 999 : aIndex) - (bIndex === -1 ? 999 : bIndex);
    };

    staffMonthlyData.sort(sortByMonth);
    workerMonthlyData.sort(sortByMonth);

    // Create combined data
    const combinedMonthlyData: MonthlyData[] = [];

    monthOrder.forEach(month => {
      const staffMonth = staffMonthlyData.find(s => s.month === month);
      const workerMonth = workerMonthlyData.find(w => w.month === month);

      if (staffMonth || workerMonth) {
        const combinedMonth: MonthlyData = {
          month,
          totalEmployees: (staffMonth?.totalEmployees || 0) + (workerMonth?.totalEmployees || 0),
          totalGrossSalary: (staffMonth?.totalGrossSalary || 0) + (workerMonth?.totalGrossSalary || 0),
          totalSalary1: (staffMonth?.totalSalary1 || 0) + (workerMonth?.totalSalary1 || 0),
          totalBasicSalary: (staffMonth?.totalBasicSalary || 0) + (workerMonth?.totalBasicSalary || 0),
          totalHRA: (staffMonth?.totalHRA || 0) + (workerMonth?.totalHRA || 0),
          totalDA: (staffMonth?.totalDA || 0) + (workerMonth?.totalDA || 0),
          totalPF: (staffMonth?.totalPF || 0) + (workerMonth?.totalPF || 0),
          totalESIC: (staffMonth?.totalESIC || 0) + (workerMonth?.totalESIC || 0),
          totalOvertimePayment: (staffMonth?.totalOvertimePayment || 0) + (workerMonth?.totalOvertimePayment || 0),
          totalWorkingDays: (staffMonth?.totalWorkingDays || 0) + (workerMonth?.totalWorkingDays || 0),
          totalLeaveDays: (staffMonth?.totalLeaveDays || 0) + (workerMonth?.totalLeaveDays || 0),
          totalLateDays: (staffMonth?.totalLateDays || 0) + (workerMonth?.totalLateDays || 0),
          avgGrossSalary: 0,
          avgSalary1: 0,
        };

        if (combinedMonth.totalEmployees > 0) {
          combinedMonth.avgGrossSalary = combinedMonth.totalGrossSalary / combinedMonth.totalEmployees;
          combinedMonth.avgSalary1 = combinedMonth.totalSalary1 / combinedMonth.totalEmployees;
        }

        combinedMonthlyData.push(combinedMonth);
      }
    });

    // Calculate summaries for each category
    // Replace the createSummary function with this fixed version
const createSummary = (data: MonthlyData[]) => {
  // Check if data array is empty
  if (!data || data.length === 0) {
    return {
      totalMonths: 0,
      totalGrossSalaryAllMonths: 0,
      totalSalary1AllMonths: 0,
      averageEmployeesPerMonth: 0,
      totalOvertimePayment: 0,
      totalWorkingDays: 0,
      highestSalaryMonth: 'N/A',
      lowestSalaryMonth: 'N/A',
    };
  }

  return {
    totalMonths: data.length,
    totalGrossSalaryAllMonths: data.reduce((sum, month) => sum + month.totalGrossSalary, 0),
    totalSalary1AllMonths: data.reduce((sum, month) => sum + month.totalSalary1, 0),
    averageEmployeesPerMonth: data.reduce((sum, month) => sum + month.totalEmployees, 0) / data.length,
    totalOvertimePayment: data.reduce((sum, month) => sum + month.totalOvertimePayment, 0),
    totalWorkingDays: data.reduce((sum, month) => sum + month.totalWorkingDays, 0),
    highestSalaryMonth: data.reduce((max, month) => 
      month.totalGrossSalary > max.totalGrossSalary ? month : max
    ).month,
    lowestSalaryMonth: data.reduce((min, month) => 
      month.totalGrossSalary < min.totalGrossSalary ? month : min
    ).month,
  };
};


    const staffSummary = createSummary(staffMonthlyData);
    const workerSummary = workerMonthlyData.length > 0 ? createSummary(workerMonthlyData) : null;
    const combinedSummary = createSummary(combinedMonthlyData);

    // Prepare chart data
    const createChartData = (data: MonthlyData[]) => ({
      months: data.map(m => m.month),
      totalGrossSalary: data.map(m => Math.round(m.totalGrossSalary)),
      totalSalary1: data.map(m => Math.round(m.totalSalary1)),
      totalEmployees: data.map(m => m.totalEmployees),
      overtimePayment: data.map(m => Math.round(m.totalOvertimePayment)),
      workingDays: data.map(m => Math.round(m.totalWorkingDays)),
      attendance: {
        totalWorkingDays: data.map(m => Math.round(m.totalWorkingDays)),
        totalLeaveDays: data.map(m => Math.round(m.totalLeaveDays)),
        totalLateDays: data.map(m => Math.round(m.totalLateDays))
      }
    });

    return NextResponse.json({
      success: true,
      staff: {
        monthlyData: staffMonthlyData,
        summary: staffSummary,
        chartData: createChartData(staffMonthlyData)
      },
      worker: workerMonthlyData.length > 0 ? {
        monthlyData: workerMonthlyData,
        summary: workerSummary,
        chartData: createChartData(workerMonthlyData)
      } : null,
      combined: {
        monthlyData: combinedMonthlyData,
        summary: combinedSummary,
        chartData: createChartData(combinedMonthlyData)
      }
    });

  } catch (error) {
    console.error("Error processing monthly summary:", error);
    return NextResponse.json(
      {
        error: "Failed to process monthly summary",
        details: error instanceof Error ? error.message : String(error),
      },
      { status: 500 }
    );
  }
}
