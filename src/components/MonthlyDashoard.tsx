
'use client';

import { useState } from 'react';
import { BarChart3, LineChart, PieChart, Download, Calendar, TrendingUp } from 'lucide-react';

interface MonthlyDashboardProps {
  monthlySummaryData: {
    months: string[];
    totalGrossSalary: number[];
    totalSalary1: number[];
    totalEmployees: number[];
    overtimePayment: number[];
    workingDays: number[];
    attendance: {
      totalWorkingDays: number[];
      totalLeaveDays: number[];
      totalLateDays: number[];
    };
  };
}

export default function MonthlyDashboard({ monthlySummaryData }: MonthlyDashboardProps) {
  const [activeChart, setActiveChart] = useState<'salary' | 'employees' | 'overtime' | 'attendance'>('salary');

  const formatCurrency = (amount: number) => {
    if (amount >= 10000000) {
      return `₹${(amount / 10000000).toFixed(1)}Cr`;
    } else if (amount >= 100000) {
      return `₹${(amount / 100000).toFixed(1)}L`;
    }
    return `₹${amount.toLocaleString()}`;
  };

  const chartButtons = [
    { key: 'salary', label: 'Salary Trends', icon: TrendingUp },
    { key: 'employees', label: 'Employee Count', icon: BarChart3 },
    { key: 'overtime', label: 'Overtime Analysis', icon: LineChart },
    { key: 'attendance', label: 'Attendance Pattern', icon: PieChart }
  ];

  const downloadMonthlySummary = () => {
    // Create CSV data
    const headers = ['Month', 'Total Employees', 'Gross Salary (₹)', 'Salary1 (₹)', 'OT Payment (₹)', 'Working Days'];
    const csvData = [headers];

    monthlySummaryData.months.forEach((month, index) => {
      csvData.push([
        month,
        monthlySummaryData.totalEmployees[index].toString(),
        monthlySummaryData.totalGrossSalary[index].toString(),
        monthlySummaryData.totalSalary1[index].toString(),
        monthlySummaryData.overtimePayment[index].toString(),
        monthlySummaryData.workingDays[index].toString()
      ]);
    });

    const csvContent = csvData.map(row => row.join(',')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Monthly-Summary-${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
  };

  const renderSalaryChart = () => {
    const maxSalary = Math.max(...monthlySummaryData.totalGrossSalary);
    const maxSalary1 = Math.max(...monthlySummaryData.totalSalary1);
    const maxValue = Math.max(maxSalary, maxSalary1);

    return (
      <div className="space-y-4">
        <div className="flex justify-between items-center">
          <h4 className="font-semibold text-gray-800">Monthly Salary Trends</h4>
          <div className="flex space-x-4 text-sm">
            <div className="flex items-center">
              <div className="w-3 h-3 bg-blue-500 rounded mr-2"></div>
              <span>Gross Salary</span>
            </div>
            <div className="flex items-center">
              <div className="w-3 h-3 bg-green-500 rounded mr-2"></div>
              <span>Salary1</span>
            </div>
          </div>
        </div>

        <div className="space-y-3">
          {monthlySummaryData.months.map((month, index) => (
            <div key={month} className="space-y-2">
              <div className="flex justify-between text-sm font-medium">
                <span>{month}</span>
                <span className="text-gray-600">
                  {formatCurrency(monthlySummaryData.totalGrossSalary[index])} / {formatCurrency(monthlySummaryData.totalSalary1[index])}
                </span>
              </div>
              <div className="relative">
                <div className="flex space-x-1 h-8">
                  <div 
                    className="bg-blue-500 rounded-sm flex items-center justify-center text-xs text-white font-medium"
                    style={{ width: `${(monthlySummaryData.totalGrossSalary[index] / maxValue) * 100}%` }}
                  >
                    {(monthlySummaryData.totalGrossSalary[index] / maxValue) * 100 > 15 && formatCurrency(monthlySummaryData.totalGrossSalary[index])}
                  </div>
                  <div 
                    className="bg-green-500 rounded-sm flex items-center justify-center text-xs text-white font-medium"
                    style={{ width: `${(monthlySummaryData.totalSalary1[index] / maxValue) * 100}%` }}
                  >
                    {(monthlySummaryData.totalSalary1[index] / maxValue) * 100 > 15 && formatCurrency(monthlySummaryData.totalSalary1[index])}
                  </div>
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  };

  const renderEmployeeChart = () => {
    const maxEmployees = Math.max(...monthlySummaryData.totalEmployees);
    const minEmployees = Math.min(...monthlySummaryData.totalEmployees);

    return (
      <div className="space-y-4">
        <div className="flex justify-between items-center">
          <h4 className="font-semibold text-gray-800">Employee Count by Month</h4>
          <div className="text-sm text-gray-600">
            Range: {minEmployees} - {maxEmployees} employees
          </div>
        </div>

        <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
          {monthlySummaryData.months.map((month, index) => {
            const count = monthlySummaryData.totalEmployees[index];
            const heightPercentage = ((count - minEmployees) / (maxEmployees - minEmployees)) * 100 || 50;

            return (
              <div key={month} className="text-center">
                <div className="mb-2 relative h-32 flex items-end justify-center">
                  <div 
                    className="bg-purple-500 rounded-t-lg w-12 flex items-end justify-center text-white text-xs font-bold pb-1"
                    style={{ height: `${Math.max(heightPercentage, 20)}%` }}
                  >
                    {count}
                  </div>
                </div>
                <div className="text-xs font-medium text-gray-700">{month}</div>
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  const renderOvertimeChart = () => {
    const maxOT = Math.max(...monthlySummaryData.overtimePayment);

    return (
      <div className="space-y-4">
        <div className="flex justify-between items-center">
          <h4 className="font-semibold text-gray-800">Overtime Payment Trends</h4>
          <div className="text-sm text-gray-600">
            Total: {formatCurrency(monthlySummaryData.overtimePayment.reduce((a, b) => a + b, 0))}
          </div>
        </div>

        <div className="space-y-2">
          {monthlySummaryData.months.map((month, index) => {
            const amount = monthlySummaryData.overtimePayment[index];
            const widthPercentage = (amount / maxOT) * 100;

            return (
              <div key={month} className="flex items-center space-x-3">
                <div className="w-16 text-sm font-medium text-gray-700">{month}</div>
                <div className="flex-1 relative">
                  <div className="w-full bg-gray-200 rounded-full h-6">
                    <div 
                      className="bg-gradient-to-r from-orange-400 to-orange-600 h-6 rounded-full flex items-center justify-center text-white text-xs font-medium"
                      style={{ width: `${Math.max(widthPercentage, 5)}%` }}
                    >
                      {widthPercentage > 20 && formatCurrency(amount)}
                    </div>
                  </div>
                </div>
                <div className="w-20 text-sm text-right text-gray-600">
                  {formatCurrency(amount)}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  const renderAttendanceChart = () => {
    const totalWorkingDays = monthlySummaryData.attendance.totalWorkingDays.reduce((a, b) => a + b, 0);
    const totalLeaveDays = monthlySummaryData.attendance.totalLeaveDays.reduce((a, b) => a + b, 0);
    const totalLateDays = monthlySummaryData.attendance.totalLateDays.reduce((a, b) => a + b, 0);
    const totalDays = totalWorkingDays + totalLeaveDays;

    return (
      <div className="space-y-4">
        <h4 className="font-semibold text-gray-800">Attendance Overview (All Months)</h4>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {/* Overall Stats */}
          <div className="space-y-3">
            <h5 className="font-medium text-gray-700">Overall Statistics</h5>
            <div className="space-y-2">
              <div className="flex justify-between">
                <span className="text-sm text-gray-600">Total Working Days</span>
                <span className="font-medium">{totalWorkingDays.toLocaleString()}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-sm text-gray-600">Total Leave Days</span>
                <span className="font-medium">{totalLeaveDays.toLocaleString()}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-sm text-gray-600">Total Late Days</span>
                <span className="font-medium">{totalLateDays.toLocaleString()}</span>
              </div>
              <div className="flex justify-between border-t pt-2">
                <span className="text-sm font-medium">Attendance Rate</span>
                <span className="font-bold text-green-600">
                  {((totalWorkingDays / totalDays) * 100).toFixed(1)}%
                </span>
              </div>
            </div>
          </div>

          {/* Visual representation */}
          <div className="space-y-3">
            <h5 className="font-medium text-gray-700">Visual Breakdown</h5>
            <div className="space-y-2">
              <div className="flex items-center space-x-2">
                <div className="w-4 h-4 bg-green-500 rounded"></div>
                <span className="text-sm">Working Days</span>
                <span className="text-xs text-gray-500">
                  ({((totalWorkingDays / totalDays) * 100).toFixed(1)}%)
                </span>
              </div>
              <div className="flex items-center space-x-2">
                <div className="w-4 h-4 bg-red-500 rounded"></div>
                <span className="text-sm">Leave Days</span>
                <span className="text-xs text-gray-500">
                  ({((totalLeaveDays / totalDays) * 100).toFixed(1)}%)
                </span>
              </div>
              <div className="flex items-center space-x-2">
                <div className="w-4 h-4 bg-yellow-500 rounded"></div>
                <span className="text-sm">Late Days</span>
                <span className="text-xs text-gray-500">
                  ({((totalLateDays / totalWorkingDays) * 100).toFixed(1)}% of working days)
                </span>
              </div>
            </div>

            <div className="w-full bg-gray-200 rounded-full h-8 flex overflow-hidden">
              <div 
                className="bg-green-500 flex items-center justify-center text-white text-xs font-medium"
                style={{ width: `${(totalWorkingDays / totalDays) * 100}%` }}
              >
                Work
              </div>
              <div 
                className="bg-red-500 flex items-center justify-center text-white text-xs font-medium"
                style={{ width: `${(totalLeaveDays / totalDays) * 100}%` }}
              >
                Leave
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderActiveChart = () => {
    switch (activeChart) {
      case 'salary': return renderSalaryChart();
      case 'employees': return renderEmployeeChart();
      case 'overtime': return renderOvertimeChart();
      case 'attendance': return renderAttendanceChart();
      default: return renderSalaryChart();
    }
  };

  return (
    <div className="mb-6">
      <div className="bg-white rounded-lg shadow-sm border">
        {/* Header */}
        <div className="px-6 py-4 border-b border-gray-200">
          <div className="flex justify-between items-center">
            <div className="flex items-center space-x-2">
              <Calendar className="h-5 w-5 text-blue-600" />
              <h3 className="text-lg font-semibold text-gray-800">Monthly Analysis Dashboard</h3>
            </div>
            <button
              onClick={downloadMonthlySummary}
              className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm flex items-center space-x-2"
            >
              <Download className="h-4 w-4" />
              <span>Export Data</span>
            </button>
          </div>
        </div>

        {/* Chart Navigation */}
        <div className="px-6 py-3 border-b border-gray-100 bg-gray-50">
          <div className="flex space-x-1 overflow-x-auto">
            {chartButtons.map(({ key, label, icon: Icon }) => (
              <button
                key={key}
                onClick={() => setActiveChart(key as any)}
                className={`flex items-center space-x-2 px-4 py-2 rounded-lg text-sm font-medium transition-colors whitespace-nowrap ${
                  activeChart === key
                    ? 'bg-blue-600 text-white'
                    : 'bg-white text-gray-600 hover:bg-gray-100 border border-gray-200'
                }`}
              >
                <Icon className="h-4 w-4" />
                <span>{label}</span>
              </button>
            ))}
          </div>
        </div>

        {/* Chart Content */}
        <div className="px-6 py-6">
          {renderActiveChart()}
        </div>
      </div>
    </div>
  );
}
