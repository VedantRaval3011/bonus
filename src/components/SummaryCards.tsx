"use client";

import {
  Users,
  CheckCircle,
  XCircle,
  IndianRupee,
  TrendingUp,
  Calendar,
  BarChart3,
  DollarSign,
} from "lucide-react";

interface SummaryCardsProps {
  summary: {
    totalEmployees: number;
    eligibleEmployees: number;
    ineligibleEmployees: number;
    totalBonusAmount: number;
    monthlySummary?: {
      totalMonths: number;
      totalGrossSalaryAllMonths: number;
      totalSalary1AllMonths: number;
      averageEmployeesPerMonth: number;
      highestSalaryMonth: string;
      lowestSalaryMonth: string;
      totalOvertimePayment: number;
      totalWorkingDays: number;
    };
  };
}

export default function SummaryCards({ summary }: SummaryCardsProps) {
  const formatCurrency = (amount: number) => {
    return new Intl.NumberFormat("en-IN", {
      style: "currency",
      currency: "INR",
      minimumFractionDigits: 0,
    }).format(amount);
  };

  const formatLakhsCrores = (amount: number) => {
    if (amount >= 10000000) {
      return `₹${(amount / 10000000).toFixed(2)} Cr`;
    } else if (amount >= 100000) {
      return `₹${(amount / 100000).toFixed(2)} L`;
    }
    return formatCurrency(amount);
  };

  // Calculate difference between gross and net
  const totalDifference = summary.monthlySummary 
    ? summary.monthlySummary.totalGrossSalaryAllMonths - summary.monthlySummary.totalSalary1AllMonths 
    : 0;

  return (
    <>
      {/* Existing Bonus Summary Cards */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
        <div className="bg-white p-3 rounded shadow-sm">
          <div className="flex items-center">
            <Users className="h-6 w-6 text-blue-500 mr-2" />
            <div>
              <p className="text-xs text-gray-600">Total Employees</p>
              <p className="text-lg font-bold">{summary.totalEmployees}</p>
            </div>
          </div>
        </div>

        <div className="bg-white p-3 rounded shadow-sm">
          <div className="flex items-center">
            <CheckCircle className="h-6 w-6 text-green-500 mr-2" />
            <div>
              <p className="text-xs text-gray-600">Eligible</p>
              <p className="text-lg font-bold">{summary.eligibleEmployees}</p>
            </div>
          </div>
        </div>

        <div className="bg-white p-3 rounded shadow-sm">
          <div className="flex items-center">
            <XCircle className="h-6 w-6 text-red-500 mr-2" />
            <div>
              <p className="text-xs text-gray-600">Ineligible</p>
              <p className="text-lg font-bold">{summary.ineligibleEmployees}</p>
            </div>
          </div>
        </div>

        <div className="bg-white p-3 rounded shadow-sm">
          <div className="flex items-center">
            <IndianRupee className="h-6 w-6 text-green-600 mr-2" />
            <div>
              <p className="text-xs text-gray-600">Total Bonus</p>
              <p className="text-sm font-bold">
                {formatLakhsCrores(summary.totalBonusAmount)}
              </p>
            </div>
          </div>
        </div>
      </div>

      {/* Enhanced Monthly Summary Section with Prominent Totals */}
      {summary.monthlySummary && (
        <div className="mb-6">
          <h3 className="text-xl font-bold mb-4 flex items-center text-gray-800">
            <Calendar className="h-6 w-6 mr-3 text-indigo-600" />
            Salary Analysis ({summary.monthlySummary.totalMonths} Months Overview)
          </h3>

          {/* Main Salary Totals - Prominent Display */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-6">
            {/* Total Gross Salary */}
            <div className="bg-gradient-to-br from-blue-500 to-blue-600 p-6 rounded-xl shadow-lg text-white">
              <div className="flex items-center justify-between mb-4">
                <DollarSign className="h-8 w-8 text-blue-100" />
                <span className="text-blue-100 text-sm font-medium uppercase tracking-wider">
                  Total Gross
                </span>
              </div>
              <p className="text-3xl font-bold mb-2">
                {formatLakhsCrores(summary.monthlySummary.totalGrossSalaryAllMonths)}
              </p>
              <p className="text-blue-100 text-sm">
                Across {summary.monthlySummary.totalMonths} months
              </p>
              <div className="mt-3 pt-3 border-t border-blue-400">
                <p className="text-xs text-blue-100">
                  Monthly Average: {formatLakhsCrores(
                    summary.monthlySummary.totalGrossSalaryAllMonths / summary.monthlySummary.totalMonths
                  )}
                </p>
              </div>
            </div>

            {/* Total Salary1 (Net) */}
            <div className="bg-gradient-to-br from-green-500 to-green-600 p-6 rounded-xl shadow-lg text-white">
              <div className="flex items-center justify-between mb-4">
                <IndianRupee className="h-8 w-8 text-green-100" />
                <span className="text-green-100 text-sm font-medium uppercase tracking-wider">
                  Total Salary1
                </span>
              </div>
              <p className="text-3xl font-bold mb-2">
                {formatLakhsCrores(summary.monthlySummary.totalSalary1AllMonths)}
              </p>
              <p className="text-green-100 text-sm">
                Net payments made
              </p>
              <div className="mt-3 pt-3 border-t border-green-400">
                <p className="text-xs text-green-100">
                  Monthly Average: {formatLakhsCrores(
                    summary.monthlySummary.totalSalary1AllMonths / summary.monthlySummary.totalMonths
                  )}
                </p>
              </div>
            </div>

            {/* Deductions/Difference */}
            <div className="bg-gradient-to-br from-orange-500 to-orange-600 p-6 rounded-xl shadow-lg text-white">
              <div className="flex items-center justify-between mb-4">
                <TrendingUp className="h-8 w-8 text-orange-100 rotate-180" />
                <span className="text-orange-100 text-sm font-medium uppercase tracking-wider">
                  Total Deductions
                </span>
              </div>
              <p className="text-3xl font-bold mb-2">
                {formatLakhsCrores(totalDifference)}
              </p>
              <p className="text-orange-100 text-sm">
                Gross - Net difference
              </p>
              <div className="mt-3 pt-3 border-t border-orange-400">
                <p className="text-xs text-orange-100">
                  Deduction %: {((totalDifference / summary.monthlySummary.totalGrossSalaryAllMonths) * 100).toFixed(1)}%
                </p>
              </div>
            </div>
          </div>

          {/* Additional Metrics */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
            <div className="bg-gradient-to-br from-purple-50 to-purple-100 p-4 rounded-lg shadow-sm">
              <div className="flex items-center justify-between mb-2">
                <Users className="h-5 w-5 text-purple-600" />
                <span className="text-xs text-purple-600 font-medium">
                  AVG EMPLOYEES
                </span>
              </div>
              <p className="text-lg font-bold text-purple-800">
                {Math.round(summary.monthlySummary.averageEmployeesPerMonth)}
              </p>
              <p className="text-xs text-purple-600 mt-1">Per Month</p>
            </div>

            <div className="bg-gradient-to-br from-teal-50 to-teal-100 p-4 rounded-lg shadow-sm">
              <div className="flex items-center justify-between mb-2">
                <TrendingUp className="h-5 w-5 text-teal-600" />
                <span className="text-xs text-teal-600 font-medium">
                  PEAK MONTH
                </span>
              </div>
              <p className="text-sm font-bold text-teal-800">
                {summary.monthlySummary.highestSalaryMonth}
              </p>
              <p className="text-xs text-teal-600 mt-1">Highest Salary</p>
            </div>

            <div className="bg-gradient-to-br from-indigo-50 to-indigo-100 p-4 rounded-lg shadow-sm">
              <div className="flex items-center justify-between mb-2">
                <Calendar className="h-5 w-5 text-indigo-600" />
                <span className="text-xs text-indigo-600 font-medium">
                  WORKING DAYS
                </span>
              </div>
              <p className="text-lg font-bold text-indigo-800">
                {summary.monthlySummary.totalWorkingDays.toLocaleString()}
              </p>
              <p className="text-xs text-indigo-600 mt-1">Total Days</p>
            </div>

            <div className="bg-gradient-to-br from-gray-50 to-gray-100 p-4 rounded-lg shadow-sm">
              <div className="flex items-center justify-between mb-2">
                <BarChart3 className="h-5 w-5 text-gray-600" />
                <span className="text-xs text-gray-600 font-medium">
                  OVERTIME
                </span>
              </div>
              <p className="text-lg font-bold text-gray-800">
                {formatLakhsCrores(summary.monthlySummary.totalOvertimePayment)}
              </p>
              <p className="text-xs text-gray-600 mt-1">Total OT Payment</p>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
