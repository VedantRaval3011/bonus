"use client";

import { useState } from "react";
import { Download, RefreshCw, LogOut, Building, Search } from "lucide-react";
import FileUpload from "./FileUpload";
import SummaryCards from "@/components/SummaryCards";
import DataTable from "./DataTable";
import ComparisonView from "./ComparisonView";
import MonthlyDashboard from "./MonthlyDashoard";

interface DashboardProps {
  onLogout: () => void;
}

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

interface DashboardData {
  hrFileBase64?: string | null;
  staffFileBase64?: string; // ADD THIS
  workerFileBase64?: string; // ADD THIS
  bonusCalculations?: any[];
  comparisonResults?: any;
  summary?: any;
  staff?: {
    monthlyData: MonthlyData[];
    summary: any;
    chartData: {
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
  };
  worker?: {
    monthlyData: MonthlyData[];
    summary: any;
    chartData: {
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
  };
  combined?: {
    monthlyData: MonthlyData[];
    summary: any;
    chartData: {
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
  };
}

export default function Dashboard({ onLogout }: DashboardProps) {
  const [dashboardData, setDashboardData] = useState<DashboardData | null>(
    null
  );
  const [isProcessing, setIsProcessing] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const [departmentFilter, setDepartmentFilter] = useState("all");
  const [showMonthlyDashboard, setShowMonthlyDashboard] = useState(false);
  const [showSummaryCards, setShowSummaryCards] = useState(false);
  const [showMonthlyBreakdown, setShowMonthlyBreakdown] = useState(false);

  const handleFilesUploaded = async (files: any) => {
    setIsProcessing(true);

    try {
      const formData = new FormData();
      // ✅ FIXED: Use correct field names
      if (files.staffTulsi) formData.append("staffTulsi", files.staffTulsi);
      if (files.workerTulsi) formData.append("workerTulsi", files.workerTulsi);
      if (files.dueVoucher) formData.append("dueVoucher", files.dueVoucher);
      if (files.loanDeduction)
        formData.append("loanDeduction", files.loanDeduction);
      if (files.actualPercentage)
        formData.append("actualPercentage", files.actualPercentage);
      if (files.hrComparison)
        formData.append("hrComparison", files.hrComparison);

      const response = await fetch("/api/process", {
        method: "POST",
        body: formData,
      });

      if (response.ok) {
        const data = await response.json();

        // Fetch monthly summary data from the separate endpoint
        const monthlySummaryResponse = await fetch("/api/monthly-summary", {
          method: "POST",
          body: formData, // Pass the same formData to process monthly data
        });

        if (monthlySummaryResponse.ok) {
          const monthlySummary = await monthlySummaryResponse.json();

          // Merge the monthly summary data with the existing data
          setDashboardData({
            ...data,
            ...monthlySummary,
          });
        } else {
          const error = await monthlySummaryResponse.json();
          throw new Error(
            `Error fetching monthly summary: ${error.error || "Unknown error"}`
          );
        }
      } else {
        const error = await response.json();
        throw new Error(
          `Error processing files: ${error.details || error.error}`
        );
      }
    } catch (error) {
      console.error("Error:", error);
      alert(error instanceof Error ? error.message : "Error processing files");
    } finally {
      setIsProcessing(false);
    }
  };

  const downloadFinalBonus = async () => {
    try {
      const response = await fetch("/api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          calculations: dashboardData?.bonusCalculations,
          hrFileBase64: dashboardData?.hrFileBase64,
          staffFileBase64: dashboardData?.staffFileBase64, // ADD THIS
          workerFileBase64: dashboardData?.workerFileBase64, // ADD THIS
        }),
      });

      if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `Final-Bonus-Data-${
          new Date().toISOString().split("T")[0]
        }.xlsx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
      }
    } catch (error) {
      console.error("Error downloading file:", error);
    }
  };

  const downloadComparisonReport = async () => {
    if (!dashboardData?.comparisonResults) return;

    try {
      const response = await fetch("/api/compare", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ comparisons: dashboardData.comparisonResults }),
      });

      if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `Comparison-Report-${
          new Date().toISOString().split("T")[0]
        }.xlsx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
      }
    } catch (error) {
      console.error("Error downloading comparison report:", error);
    }
  };

  const formatCurrency = (value: number) => {
    return new Intl.NumberFormat("en-IN", {
      style: "currency",
      currency: "INR",
      minimumFractionDigits: 2,
    }).format(value);
  };

  const formatDate = (value: Date | string) => {
    const date = typeof value === "string" ? new Date(value) : value;
    return date.toLocaleDateString("en-IN", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
    });
  };

  const getDepartmentColor = (department: string) => {
    const colors = {
      S: "bg-blue-100 text-blue-800",
      W: "bg-red-100 text-red-800",
      M: "bg-green-100 text-green-800",
      "Sci Prec-": "bg-purple-100 text-purple-800",
      NRTM: "bg-orange-100 text-orange-800",
      "Sci Prec Life-": "bg-yellow-100 text-yellow-800",
    };
    return (
      colors[department as keyof typeof colors] || "bg-gray-100 text-gray-800"
    );
  };

  const filteredData =
    dashboardData?.bonusCalculations?.filter((employee: any) => {
      const matchesSearch =
        employee.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
        employee.empId.toString().includes(searchTerm);
      const matchesDepartment =
        departmentFilter === "all" || employee.department === departmentFilter;
      return matchesSearch && matchesDepartment;
    }) || [];

  return (
    <div className="container mx-auto p-4 max-w-7xl">
      <div className="flex justify-between items-center mb-6">
        <h1 className="text-2xl font-bold text-gray-900">
          Employee Bonus System
        </h1>
        <button
          onClick={onLogout}
          className="bg-red-600 hover:bg-red-700 text-white px-4 py-2 rounded-lg flex items-center text-sm font-medium transition-colors"
        >
          <LogOut className="h-5 w-5 mr-2" />
          Logout
        </button>
      </div>

      {!dashboardData && (
        <div className="mb-6">
          <FileUpload onFilesUploaded={handleFilesUploaded} />
          {isProcessing && (
            <div className="mt-4 flex items-center justify-center p-4 bg-blue-50 rounded-lg">
              <RefreshCw className="h-5 w-5 animate-spin mr-2 text-blue-600" />
              <span className="text-blue-800 text-sm font-medium">
                Processing...
              </span>
            </div>
          )}
        </div>
      )}

      {dashboardData && (
        <>
          {showSummaryCards && <SummaryCards summary={dashboardData.summary} />}

          {showMonthlyDashboard && dashboardData.combined?.chartData && (
            <MonthlyDashboard
              monthlySummaryData={dashboardData.combined.chartData}
            />
          )}

          {dashboardData.summary?.departmentBreakdown && (
            <div className="mb-6">
              <h3 className="text-lg font-semibold mb-3 flex items-center">
                <Building className="h-5 w-5 mr-2" />
                Departments
              </h3>
              <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4">
                {Object.entries(dashboardData.summary.departmentBreakdown).map(
                  ([dept, data]: [string, any]) => (
                    <div
                      key={dept}
                      className={`p-3 rounded-lg ${getDepartmentColor(
                        dept
                      )} text-center`}
                    >
                      <div className="font-semibold text-sm">{dept}</div>
                      <div className="text-xs">{data.employees} emp</div>
                      <div className="text-xs">
                        {formatCurrency(data.totalBonus)}
                      </div>
                    </div>
                  )
                )}
              </div>
            </div>
          )}

          <div className="mb-6 bg-white p-4 rounded-lg shadow-sm">
            <div className="flex flex-wrap gap-4 items-center">
              <div className="flex items-center space-x-2 flex-1 min-w-0">
                <Search className="h-5 w-5 text-gray-400" />
                <input
                  type="text"
                  placeholder="Search..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="flex-1 min-w-0 px-3 py-2 border rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
                <select
                  value={departmentFilter}
                  onChange={(e) => setDepartmentFilter(e.target.value)}
                  className="px-3 py-2 border rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="all">All Depts</option>
                  {dashboardData.summary?.departmentBreakdown &&
                    Object.keys(dashboardData.summary.departmentBreakdown).map(
                      (dept) => (
                        <option key={dept} value={dept}>
                          {dept}
                        </option>
                      )
                    )}
                </select>
              </div>

              <div className="flex gap-2 flex-wrap">
                <button
                  onClick={downloadFinalBonus}
                  className="bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-lg text-sm flex items-center space-x-2 transition-colors"
                >
                  <Download className="h-4 w-4" />
                  <span>Bonus Excel</span>
                </button>

                {dashboardData.comparisonResults && (
                  <button
                    onClick={downloadComparisonReport}
                    className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm flex items-center space-x-2 transition-colors"
                  >
                    <Download className="h-4 w-4" />
                    <span>Comparison</span>
                  </button>
                )}

                <button
                  onClick={() => setDashboardData(null)}
                  className="bg-gray-600 hover:bg-gray-700 text-white px-4 py-2 rounded-lg text-sm flex items-center space-x-2 transition-colors"
                >
                  <RefreshCw className="h-4 w-4" />
                  <span>New Files</span>
                </button>
              </div>
            </div>
            <div className="text-xs text-gray-500 mt-2">
              Showing {filteredData.length} of{" "}
              {dashboardData.bonusCalculations?.length || 0} employees
            </div>
          </div>

          <DataTable
            title="Employee Bonus Calculations"
            data={filteredData}
            columns={[
              { key: "empId", label: "ID" },
              { key: "name", label: "Name" },
              {
                key: "department",
                label: "Dept",
                render: (dept) => (
                  <span
                    className={`px-2 py-1 rounded-lg text-xs ${getDepartmentColor(
                      dept
                    )}`}
                  >
                    {dept}
                  </span>
                ),
              },
              { key: "doj", label: "DOJ", format: formatDate },
              { key: "serviceMonths", label: "Service (M)" },
              {
                key: "totalGrossSalary",
                label: "Gross Salary",
                format: formatCurrency,
              },
              { key: "bonusPercent", label: "Bonus %" },
              {
                key: "isEligible",
                label: "Status",
                format: (v) => (v ? "✅" : "❌"),
              },
            ]}
            searchTerm={searchTerm}
            onSearchChange={setSearchTerm}
            searchFields={["name", "empId", "department"]}
          />

          {dashboardData.comparisonResults && (
            <ComparisonView comparisons={dashboardData.comparisonResults} />
          )}
        </>
      )}
    </div>
  );
}
