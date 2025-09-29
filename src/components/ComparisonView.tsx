'use client';

interface ComparisonData {
  empId: string;
  name: string;
  department: string;
  hrData: {
    grossSal: number;
    gross02: number;
    register: number;
    actual: number;
    unPaid: number;
    finalRTGS: number;
    reim: number;
  };
  systemData: {
    grossSal: number;
    gross02: number;
    register: number;
    actual: number;
    unPaid: number;
    finalRTGS: number;
    reim: number;
  };
  differences: {
    grossSalDiff: number;
    gross02Diff: number;
    registerDiff: number;
    actualDiff: number;
    unPaidDiff: number;
    finalRTGSDiff: number;
    reimDiff: number;
  };
  status: 'MATCH' | 'MISMATCH' | 'MISSING';
  source: 'both' | 'hr_only' | 'system_only';
}

interface ComparisonViewProps {
  comparisons: ComparisonData[];
}

export default function ComparisonView({ comparisons }: ComparisonViewProps) {
  if (!comparisons || comparisons.length === 0) {
    return (
      <div className="bg-white p-4 rounded shadow-sm mt-4">
        <h3 className="text-lg font-semibold mb-2">Comparison Results</h3>
        <p className="text-gray-500 text-sm">No comparison data</p>
      </div>
    );
  }

  const matched = comparisons.filter(c => c.status === 'MATCH').length;
  const mismatched = comparisons.filter(c => c.status === 'MISMATCH').length;
  const missing = comparisons.filter(c => c.status === 'MISSING').length;

  const formatCurrency = (amount: number) => {
    return new Intl.NumberFormat('en-IN', {
      style: 'currency',
      currency: 'INR'
    }).format(amount || 0);
  };

  const getStatusColor = (status: string) => {
    switch (status) {
      case 'MATCH': return 'bg-green-100 text-green-800';
      case 'MISMATCH': return 'bg-red-100 text-red-800';
      case 'MISSING': return 'bg-yellow-100 text-yellow-800';
      default: return 'bg-gray-100 text-gray-800';
    }
  };

  return (
    <div className="bg-white rounded shadow-sm mt-4 overflow-hidden">
      <div className="px-4 py-2 border-b bg-gray-50">
        <h3 className="text-lg font-semibold">HR vs System Detailed Comparison</h3>
      </div>
      
      <div className="p-4 border-b">
        <div className="grid grid-cols-3 gap-4">
          <div className="bg-green-50 p-3 rounded">
            <div className="text-green-800 font-semibold">Matches: {matched}</div>
          </div>
          <div className="bg-red-50 p-3 rounded">
            <div className="text-red-800 font-semibold">Mismatches: {mismatched}</div>
          </div>
          <div className="bg-yellow-50 p-3 rounded">
            <div className="text-yellow-800 font-semibold">Missing: {missing}</div>
          </div>
        </div>
      </div>

      <div className="overflow-x-auto">
        <table className="min-w-full text-xs">
          <thead className="bg-gray-50">
            <tr>
              <th rowSpan={2} className="px-2 py-3 text-left border-r">Employee</th>
              <th colSpan={7} className="px-2 py-1 text-center border-r bg-blue-50">HR Data</th>
              <th colSpan={7} className="px-2 py-1 text-center border-r bg-green-50">System Data</th>
              <th colSpan={7} className="px-2 py-1 text-center bg-red-50">Differences</th>
              <th rowSpan={2} className="px-2 py-3 text-center">Status</th>
            </tr>
            <tr>
              {/* HR Headers */}
              <th className="px-1 py-2 text-center text-xs bg-blue-50">GROSS SAL</th>
              <th className="px-1 py-2 text-center text-xs bg-blue-50">GROSS 02</th>
              <th className="px-1 py-2 text-center text-xs bg-blue-50">Register</th>
              <th className="px-1 py-2 text-center text-xs bg-blue-50">Actual</th>
              <th className="px-1 py-2 text-center text-xs bg-blue-50">Un Paid</th>
              <th className="px-1 py-2 text-center text-xs bg-blue-50">Final RTGS</th>
              <th className="px-1 py-2 text-center text-xs bg-blue-50 border-r">Reim</th>
              
              {/* System Headers */}
              <th className="px-1 py-2 text-center text-xs bg-green-50">GROSS SAL</th>
              <th className="px-1 py-2 text-center text-xs bg-green-50">GROSS 02</th>
              <th className="px-1 py-2 text-center text-xs bg-green-50">Register</th>
              <th className="px-1 py-2 text-center text-xs bg-green-50">Actual</th>
              <th className="px-1 py-2 text-center text-xs bg-green-50">Un Paid</th>
              <th className="px-1 py-2 text-center text-xs bg-green-50">Final RTGS</th>
              <th className="px-1 py-2 text-center text-xs bg-green-50 border-r">Reim</th>
              
              {/* Difference Headers */}
              <th className="px-1 py-2 text-center text-xs bg-red-50">GROSS SAL</th>
              <th className="px-1 py-2 text-center text-xs bg-red-50">GROSS 02</th>
              <th className="px-1 py-2 text-center text-xs bg-red-50">Register</th>
              <th className="px-1 py-2 text-center text-xs bg-red-50">Actual</th>
              <th className="px-1 py-2 text-center text-xs bg-red-50">Un Paid</th>
              <th className="px-1 py-2 text-center text-xs bg-red-50">Final RTGS</th>
              <th className="px-1 py-2 text-center text-xs bg-red-50">Reim</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-200">
            {comparisons.map((comp, index) => (
              <tr key={index} className="hover:bg-gray-50">
                <td className="px-2 py-2 border-r">
                  <div className="font-medium text-sm">{comp.name}</div>
                  <div className="text-xs text-gray-500">{comp.empId}</div>
                  <div className="text-xs text-gray-400">{comp.department}</div>
                </td>
                
                {/* HR Data */}
                <td className="px-1 py-2 text-right text-xs bg-blue-25">{formatCurrency(comp.hrData.grossSal)}</td>
                <td className="px-1 py-2 text-right text-xs bg-blue-25">{formatCurrency(comp.hrData.gross02)}</td>
                <td className="px-1 py-2 text-right text-xs bg-blue-25">{formatCurrency(comp.hrData.register)}</td>
                <td className="px-1 py-2 text-right text-xs bg-blue-25">{formatCurrency(comp.hrData.actual)}</td>
                <td className="px-1 py-2 text-right text-xs bg-blue-25">{formatCurrency(comp.hrData.unPaid)}</td>
                <td className="px-1 py-2 text-right text-xs bg-blue-25 font-semibold">{formatCurrency(comp.hrData.finalRTGS)}</td>
                <td className="px-1 py-2 text-right text-xs bg-blue-25 border-r">{formatCurrency(comp.hrData.reim)}</td>
                
                {/* System Data */}
                <td className="px-1 py-2 text-right text-xs bg-green-25">{formatCurrency(comp.systemData.grossSal)}</td>
                <td className="px-1 py-2 text-right text-xs bg-green-25">{formatCurrency(comp.systemData.gross02)}</td>
                <td className="px-1 py-2 text-right text-xs bg-green-25">{formatCurrency(comp.systemData.register)}</td>
                <td className="px-1 py-2 text-right text-xs bg-green-25">{formatCurrency(comp.systemData.actual)}</td>
                <td className="px-1 py-2 text-right text-xs bg-green-25">{formatCurrency(comp.systemData.unPaid)}</td>
                <td className="px-1 py-2 text-right text-xs bg-green-25 font-semibold">{formatCurrency(comp.systemData.finalRTGS)}</td>
                <td className="px-1 py-2 text-right text-xs bg-green-25 border-r">{formatCurrency(comp.systemData.reim)}</td>
                
                {/* Differences */}
                <td className="px-1 py-2 text-right text-xs">
                  <span className={comp.differences.grossSalDiff > 0 ? 'text-red-600' : 'text-green-600'}>
                    {formatCurrency(comp.differences.grossSalDiff)}
                  </span>
                </td>
                <td className="px-1 py-2 text-right text-xs">
                  <span className={comp.differences.gross02Diff > 0 ? 'text-red-600' : 'text-green-600'}>
                    {formatCurrency(comp.differences.gross02Diff)}
                  </span>
                </td>
                <td className="px-1 py-2 text-right text-xs">
                  <span className={comp.differences.registerDiff > 0 ? 'text-red-600' : 'text-green-600'}>
                    {formatCurrency(comp.differences.registerDiff)}
                  </span>
                </td>
                <td className="px-1 py-2 text-right text-xs">
                  <span className={comp.differences.actualDiff > 0 ? 'text-red-600' : 'text-green-600'}>
                    {formatCurrency(comp.differences.actualDiff)}
                  </span>
                </td>
                <td className="px-1 py-2 text-right text-xs">
                  <span className={comp.differences.unPaidDiff > 0 ? 'text-red-600' : 'text-green-600'}>
                    {formatCurrency(comp.differences.unPaidDiff)}
                  </span>
                </td>
                <td className="px-1 py-2 text-right text-xs font-semibold">
                  <span className={comp.differences.finalRTGSDiff > 0 ? 'text-red-600' : 'text-green-600'}>
                    {formatCurrency(comp.differences.finalRTGSDiff)}
                  </span>
                </td>
                <td className="px-1 py-2 text-right text-xs">
                  <span className={comp.differences.reimDiff > 0 ? 'text-red-600' : 'text-green-600'}>
                    {formatCurrency(comp.differences.reimDiff)}
                  </span>
                </td>
                
                <td className="px-2 py-2 text-center">
                  <span className={`px-2 py-1 rounded text-xs font-medium ${getStatusColor(comp.status)}`}>
                    {comp.status}
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      
      <div className="p-4 bg-gray-50 text-sm text-gray-600">
        <p><strong>Legend:</strong> Blue columns show HR data, Green columns show System data, Red differences indicate mismatches. Final RTGS is the primary comparison field.</p>
      </div>
    </div>
  );
}
