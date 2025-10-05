import { Employee, BonusCalculation, MonthlyData } from "./types";

export class BonusCalculator {
  private static isValidDOJ(d: any): boolean {
    return d instanceof Date && !isNaN(d.getTime());
  }

  static calculateBonus(
    employees: Employee[],
    dueVoucherMap?: Map<
      string,
      { alreadyPaid: number; unpaid: number; dept: string }
    >,
    loanMap?: Map<string, number>,
    percentageMap?: Map<string, number>
  ): BonusCalculation[] {
    const safeEmployees = employees.filter((e) => {
      const d = e.doj instanceof Date ? e.doj : new Date(e.doj as any);
      return d instanceof Date && !isNaN(d.getTime());
    }); // <— ignore DOJ = N/invalid
    return safeEmployees.map((emp) => {
      const monthlyData = emp.monthlyData || [];
      const isCashSalary = emp.isCashSalary || false;

      // Extract monthly salaries (NOV-24 to SEP-25)
      const monthlySalaries = this.extractMonthlySalaries(monthlyData);

      // Calculate OCT-25 (estimated)
      const octSalary = this.calculateOctSalary(monthlySalaries);

      // Calculate Gross Salary (sum of all months including OCT-25)
      const totalGrossSalary = [...monthlySalaries, octSalary]
        .filter((val): val is number => val !== null)
        .reduce((sum, val) => sum + val, 0);

      // Calculate service months
      const serviceMonths = this.calculateServiceMonths(emp.doj);

      // Get bonus percentage
      const bonusPercent = this.calculateBonusPercentage(
  emp.empId,
  emp.doj,
  percentageMap,
  emp.department
);

      // Calculate Gross 2
      const gross2 = this.calculateGross2(totalGrossSalary, bonusPercent);

      // Calculate Register - ZERO if cash salary
      const register = isCashSalary ? 0 : Math.round((gross2 * 8.33) / 100);

      // Get Already Paid and Unpaid - ZERO if cash salary
      const dueVoucherData = dueVoucherMap?.get(emp.empId);
      const alreadyPaid = isCashSalary ? 0 : dueVoucherData?.alreadyPaid || 0;
      const unpaid = isCashSalary ? 0 : dueVoucherData?.unpaid || 0;

      // Determine eligibility
      const isEligible = this.isEligible(emp.department, serviceMonths);

      // Calculate After V
      const afterV = register - (alreadyPaid + unpaid);

      // Calculate Actual
      const actual = isEligible ? afterV : 0;

      // Calculate Reim
      const reim = afterV - actual;

      const loanAmount = loanMap?.get(emp.empId) || 0;

      // Calculate Final RTGS
      const finalRTGS = register - (alreadyPaid + unpaid);

      // Calculate backward-compatible fields
      const calculatedBonus = register;
      const finalBonus = finalRTGS;
      const serviceMultiplier = serviceMonths >= 12 ? 1.0 : serviceMonths / 12;

      return {
        empId: emp.empId,
        name: emp.name,
        department: emp.department,
        doj: emp.doj,
        serviceMonths,

        // New fields
        monthlySalaries,
        octSalary,
        totalGrossSalary: Math.round(totalGrossSalary),
        gross2: Math.round(gross2),
        register: Math.round(register),
        alreadyPaid: Math.round(alreadyPaid),
        unpaid: Math.round(unpaid),
        isEligible,
        afterV: Math.round(afterV),
        bonusPercent,
        actual: Math.round(actual),
        reim: Math.round(reim),
        loan: Math.round(loanAmount),
        finalRTGS: Math.round(finalRTGS),
        isCashSalary, // Add this field

        // Backward compatibility
        calculatedBonus: Math.round(calculatedBonus),
        finalBonus: Math.round(finalBonus),
        serviceMultiplier,
        monthlyData,
      };
    });
  }

  /**
   * Calculate bonus percentage based on years of service
   * Priority: 1) Custom percentage from file, 2) Years of service calculation
   */
  private static calculateBonusPercentage(
  empId: string,
  doj: Date,
  percentageMap?: Map<string, number>,
  department?: string
): number {
  // For Workers (except specific exceptions), always use 8.33%
  const isWorker = department === "W" || department === "M" || 
                   department === "A" || department === "C" || 
                   department?.startsWith("Sci Prec") === false;
  
  // Special exception employees who get custom percentages even as workers
  const customExceptions = ["143", "914"];
  
  if (isWorker && !customExceptions.includes(empId)) {
    console.log(`[${empId}] Worker - using standard 8.33%`);
    return 8.33;
  }

  // For Staff OR special exception employees, check custom percentage file first
  if (percentageMap && percentageMap.has(empId)) {
    const customPercent = percentageMap.get(empId)!;
    console.log(`[${empId}] Using custom percentage: ${customPercent}`);
    return customPercent;
  }

  // Calculate years of service for Staff (or if no custom percentage found)
  const now = new Date();
  const yearsOfService =
    (now.getTime() - doj.getTime()) / (1000 * 60 * 60 * 24 * 365.25);

  console.log(
    `[${empId}], DOJ: ${doj}, Years of Service: ${yearsOfService}`
  );

  // Apply percentage based on years of service (for Staff)
  let percentage: number;
  if (yearsOfService < 1) {
    percentage = 10;
    console.log(`[${empId}] < 1 year: 10%`);
  } else if (yearsOfService >= 1 && yearsOfService < 2) {
    percentage = 12;
    console.log(`[${empId}] 1-2 years: 12%`);
  } else {
    percentage = 8.33;
    console.log(`[${empId}] >= 2 years: 8.33%`);
  }

  return percentage;
}

  private static extractMonthlySalaries(
    monthlyData: MonthlyData[]
  ): (number | null)[] {
    const monthMapping = {
      NOV: 0,
      DEC: 1,
      JAN: 2,
      FEB: 3,
      MAR: 4,
      APR: 5,
      MAY: 6,
      JUN: 7,
      JUL: 8,
      AUG: 9,
      SEP: 10,
    };

    const result: (number | null)[] = new Array(11).fill(null);

    monthlyData.forEach((data) => {
      const monthName = data.month.split("-")[0].toUpperCase().substring(0, 3);
      const index = monthMapping[monthName as keyof typeof monthMapping];

      if (index !== undefined && data.salary > 0) {
        result[index] = data.salary;
      }
    });

    return result;
  }

  private static calculateOctSalary(
    monthlySalaries: (number | null)[]
  ): number {
    // Check if AUG-25 (index 9) >= 1
    const augSalary = monthlySalaries[9];

    if (!augSalary || augSalary < 1) {
      return 0;
    }

    // Calculate average from DEC-24 (index 1) to AUG-25 (index 9)
    // Only include months where salary > 1
    const validSalaries = monthlySalaries
      .slice(1, 10) // DEC-24 to AUG-25
      .filter((s) => s !== null && s > 1) as number[];

    if (validSalaries.length === 0) return 0;

    const average =
      validSalaries.reduce((sum, val) => sum + val, 0) / validSalaries.length;
    return Math.round(average);
  }

  private static calculateGross2(
    totalGrossSalary: number,
    bonusPercent: number
  ): number {
    // If percentage is 8.33%, Gross 2 = Gross Salary
    if (bonusPercent === 8.33) {
      return totalGrossSalary;
    }
    // If percentage > 8.33%, Gross 2 = Gross Salary * 0.6
    if (bonusPercent > 8.33) {
      return totalGrossSalary * 0.6;
    }
    // Otherwise (< 8.33%), return empty/0
    return 0;
  }

  private static isEligible(
    department: string,
    serviceMonths: number
  ): boolean {

    
    // Staff (S, Sci Prec-, NRTM, Sci Prec Life-) = always eligible
    if (
      department === "S" ||
      department.startsWith("Sci Prec") ||
      department === "NRTM"
    ) {
      return true;
    }

    // Worker (W, M) = eligible if DOJ >= 6 months
    if (department === "W" || department === "M") {
      return serviceMonths >= 6;
    }

    return true; // Default to eligible
  }

  private static calculateServiceMonths(doj: Date): number {
    const now = new Date();
    const months =
      (now.getFullYear() - doj.getFullYear()) * 12 +
      (now.getMonth() - doj.getMonth());
    return Math.max(0, months);
  }
}
