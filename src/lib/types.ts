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
   isCashSalary?: boolean;
}


export interface MonthlyData {
  month: string;
  salary: number;
   department?: string;
}

export interface BonusCalculation {
  empId: string;
  name: string;
  department: string;
  doj: Date;
  serviceMonths: number;
  
  // Monthly salary breakdown (NOV-24 to SEP-25)
  monthlySalaries: (number | null)[];
  octSalary: number;
  
  // Gross calculations
  totalGrossSalary: number;
  gross2: number;
  register: number;
  
  // Due Voucher data
  alreadyPaid: number;
  unpaid: number;
  
  // Eligibility
  isEligible: boolean;
  
  // Final calculations
  afterV: number;
  bonusPercent: number;
  actual: number;
  reim: number;
   loan: number;
  finalRTGS: number;
  isCashSalary?: boolean;
  
  // Backward compatibility
  calculatedBonus: number;
  finalBonus: number;
  serviceMultiplier: number;
  monthlyData: MonthlyData[];
}

export interface ComparisonResult {
  empId: string;
  name: string;
  department: string;
  systemBonus: number;
  hrBonus: number;
  difference: number;
  status: 'MATCH' | 'MISMATCH';
}

export interface UploadedFiles {
  staffTulsi?: File;
  workerTulsi?: File;
  dueVoucher?: File;
  bonusSummary?: File;
  actualPercentage?: File;
  monthWise?: File;
  loanDeduction?: File;
  hrComparison?: File;
}

export interface DashboardData {
  staffData?: Employee[];
  workerData?: Employee[];
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
  salarySummaries?: {
    staff: any;
    worker: any;
  };
}
