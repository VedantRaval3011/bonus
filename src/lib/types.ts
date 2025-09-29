export interface Employee {
  empId: string;
  name: string;
  department: string;
  doj: Date;
  salary: number;
  basicSalary?: number;
  grossSalary?: number;
  monthlyData?: MonthlyData[];
  totalGrossSalary?: number;
  serviceMonths?: number;
  isEligible?: boolean;
  bonusAmount?: number;
  bonusPercent?: number;
  serviceMultiplier?: number;
}

export interface MonthlyData {
  month: string;
  salary: number;
}

export interface BonusCalculation {
  empId: string;
  name: string;
  department?: string; // Added department field
  doj: Date;
  monthlyData: MonthlyData[];
  totalGrossSalary: number;
  serviceMonths: number;
  bonusPercent: number;
  calculatedBonus: number;
  finalBonus: number;
  isEligible: boolean;
  serviceMultiplier: number;
}

export interface ComparisonResult {
  empId: string;
  name: string;
  department: string; // Added department field
  systemBonus: number;
  hrBonus: number;
  difference: number;
  status: 'MATCH' | 'MISMATCH';
}

export interface UploadedFiles {
  staff?: File;
  worker?: File;
  hrComparison?: File;
}

export interface DashboardData {
  staffData: Employee[];
  workerData: Employee[];
  bonusCalculations: BonusCalculation[];
  comparisonResults?: ComparisonResult[];
  summary: {
    totalEmployees: number;
    eligibleEmployees: number;
    ineligibleEmployees: number;
    totalBonusAmount: number;
    departmentBreakdown: {[key: string]: {
      employees: number;
      totalBonus: number;
    }};
  };
}


