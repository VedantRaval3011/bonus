import { Employee, BonusCalculation } from './types';
import { calculateServiceMonths } from '@/utils/date';

export class BonusCalculator {
  private static readonly BONUS_PERCENT = 8.33;
  private static readonly MIN_SERVICE_MONTHS = 6;
  
  static calculateBonus(employees: Employee[]): BonusCalculation[] {
    return employees.map(emp => this.calculateEmployeeBonus(emp));
  }
  
 private static calculateEmployeeBonus(employee: Employee): BonusCalculation {
  let doj: Date;
  if (employee.doj instanceof Date) {
    doj = employee.doj;
  } else if (typeof employee.doj === 'string') {
    doj = new Date(employee.doj);
  } else {
    doj = new Date('2020-01-01');
  }
  
  const serviceMonths = calculateServiceMonths(doj, new Date());
  const isEligible = serviceMonths >= this.MIN_SERVICE_MONTHS;
  
  // Dynamic bonus percentage based on service period
  let bonusPercent: number;
  if (serviceMonths >= 6 && serviceMonths < 12) {
    bonusPercent = 10.0;  // 6-12 months: 10%
  } else if (serviceMonths >= 12 && serviceMonths < 24) {
    bonusPercent = 12.0;  // 12-24 months: 12%
  } else if (serviceMonths >= 24) {
    bonusPercent = 8.33;  // 24+ months: 8.33%
  } else {
    bonusPercent = 0;     // Less than 6 months: 0%
  }
  
  const totalGrossSalary = employee.monthlyData?.reduce((sum, month) => sum + (month.salary || 0), 0) || 0;
  
  let calculatedBonus = 0;
  if (isEligible && totalGrossSalary > 0) {
    const baseBonus = (totalGrossSalary * bonusPercent) / 100;
    const serviceMultiplier = this.getServiceMultiplier(serviceMonths);
    calculatedBonus = baseBonus * serviceMultiplier;
  }
  
  return {
    empId: employee.empId,
    name: employee.name,
    department: employee.department,
    doj: doj,
    monthlyData: employee.monthlyData || [],
    totalGrossSalary,
    serviceMonths,
    bonusPercent: bonusPercent, // Now variable based on service
    calculatedBonus: Math.round(calculatedBonus * 100) / 100,
    finalBonus: Math.round(calculatedBonus * 100) / 100,
    isEligible,
    serviceMultiplier: this.getServiceMultiplier(serviceMonths)
  };
}

  
  private static getServiceMultiplier(serviceMonths: number): number {
    const years = serviceMonths / 12;
    
    if (years < 1) return 0.5;
    if (years < 3) return 0.75;
    if (years < 5) return 1.0;
    if (years < 10) return 1.25;
    return 1.5;
  }
}