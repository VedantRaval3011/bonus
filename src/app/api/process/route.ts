import { NextRequest, NextResponse } from 'next/server';
import { ExcelProcessor } from '@/lib/excel';
import { BonusCalculator } from '@/lib/bonus-calculator';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const staffFile = formData.get('staffTulsi') as File;
    const workerFile = formData.get('workerTulsi') as File;
    const dueVoucherFile = formData.get('dueVoucher') as File | null;
    const loanDeductionFile = formData.get('loanDeduction') as File | null;
    const actualPercentageFile = formData.get('actualPercentage') as File | null;
    const hrComparisonFile = formData.get('hrComparison') as File | null;

    if (!staffFile || !workerFile) {
      return NextResponse.json({ error: 'Missing required files' }, { status: 400 });
    }

    console.log('📂 Processing files...');
    
    // Parse main files
    const staffBuffer = await staffFile.arrayBuffer();
    const workerBuffer = await workerFile.arrayBuffer();
    
    const staffResult = await ExcelProcessor.parseStaffFile(staffBuffer);
    const workerResult = await ExcelProcessor.parseWorkerFile(workerBuffer);

    const staffData = staffResult.employees;
    const staffSummary = staffResult.summary;
    const workerData = workerResult.employees;
    const workerSummary = workerResult.summary;

    // Validate arrays
    if (!Array.isArray(staffData) || !Array.isArray(workerData)) {
      console.error('Data is not an array');
      return NextResponse.json({ 
        error: 'Data processing failed', 
        details: 'Employee data is not iterable' 
      }, { status: 500 });
    }

    console.log(`✅ Parsed ${staffData.length} staff and ${workerData.length} workers`);

    // Parse optional files
    let dueVoucherMap: Map<string, { alreadyPaid: number, unpaid: number, dept: string }> | undefined;
    let loanMap: Map<string, number> | undefined;
    let percentageMap: Map<string, number> | undefined;
    let hrFileBuffer: ArrayBuffer | undefined; // Declare here

    if (dueVoucherFile) {
      const buffer = await dueVoucherFile.arrayBuffer();
      dueVoucherMap = await ExcelProcessor.parseDueVoucherList(buffer);
    }

    if (loanDeductionFile) {
      const buffer = await loanDeductionFile.arrayBuffer();
      loanMap = await ExcelProcessor.parseLoanDeduction(buffer);
    }

    if (actualPercentageFile) {
      const buffer = await actualPercentageFile.arrayBuffer();
      percentageMap = await ExcelProcessor.parseActualPercentage(buffer);
    }

    // Store HR file buffer for comparison
    if (hrComparisonFile) {
      hrFileBuffer = await hrComparisonFile.arrayBuffer();
    }

    // Calculate bonus with all data
    const allEmployees = [...staffData, ...workerData];
    const bonusCalculations = BonusCalculator.calculateBonus(
      allEmployees,
      dueVoucherMap,
      loanMap,
      percentageMap
    );

    console.log(`✅ Calculated bonus for ${bonusCalculations.length} employees`);

    // Create department breakdown
    const departmentBreakdown = bonusCalculations.reduce((acc, calc) => {
      const dept = calc.department || 'Unknown';
      if (!acc[dept]) {
        acc[dept] = { employees: 0, totalBonus: 0 };
      }
      acc[dept].employees += 1;
      acc[dept].totalBonus += calc.isEligible ? calc.finalRTGS : 0;
      return acc;
    }, {} as {[key: string]: {employees: number, totalBonus: number}});

    const summary = {
      totalEmployees: allEmployees.length,
      eligibleEmployees: bonusCalculations.filter(c => c.isEligible).length,
      ineligibleEmployees: bonusCalculations.filter(c => !c.isEligible).length,
      totalBonusAmount: bonusCalculations.reduce((sum, c) => sum + (c.isEligible ? c.finalRTGS : 0), 0),
      departmentBreakdown
    };

    // Compare with HR file if provided
    let comparisonResults = null;
    if (hrFileBuffer) {
      console.log('📊 Processing HR comparison file...');
      const hrBonusMap = await ExcelProcessor.parseHRComparisonFile(hrFileBuffer);
      
      comparisonResults = bonusCalculations.map(calc => {
        const hrData = hrBonusMap.get(calc.empId);
        
        // Create comparison structure matching ComparisonView expectations
        const comparison = {
          empId: calc.empId,
          name: calc.name,
          department: calc.department,
          
          // HR Data (from HR file or default to 0)
          hrData: {
            grossSal: hrData?.grossSal || 0,
            gross02: hrData?.gross02 || 0,
            register: hrData?.register || 0,
            actual: hrData?.actual || 0,
            unPaid: hrData?.unpaid || 0,
            finalRTGS: hrData?.bonus || 0,
            reim: hrData?.reim || 0
          },
          
          // System Data (from our calculations)
          systemData: {
            grossSal: calc.totalGrossSalary,
            gross02: calc.gross2,
            register: calc.register,
            actual: calc.actual,
            unPaid: calc.unpaid,
            finalRTGS: calc.finalRTGS,
            reim: calc.reim
          },
          
          // Calculate differences
          differences: {
            grossSalDiff: calc.totalGrossSalary - (hrData?.grossSal || 0),
            gross02Diff: calc.gross2 - (hrData?.gross02 || 0),
            registerDiff: calc.register - (hrData?.register || 0),
            actualDiff: calc.actual - (hrData?.actual || 0),
            unPaidDiff: calc.unpaid - (hrData?.unpaid || 0),
            finalRTGSDiff: calc.finalRTGS - (hrData?.bonus || 0),
            reimDiff: calc.reim - (hrData?.reim || 0)
          },
          
          // Determine status based on all differences (within ±1 Rs margin)
          status: (() => {
            if (!hrData) return 'MISSING' as const;
            
            const allDiffs = [
              Math.abs(calc.totalGrossSalary - (hrData.grossSal || 0)),
              Math.abs(calc.gross2 - (hrData.gross02 || 0)),
              Math.abs(calc.register - (hrData.register || 0)),
              Math.abs(calc.actual - (hrData.actual || 0)),
              Math.abs(calc.unpaid - (hrData.unpaid || 0)),
              Math.abs(calc.finalRTGS - (hrData.bonus || 0)),
              Math.abs(calc.reim - (hrData.reim || 0))
            ];
            
            // Check if all differences are within ±1 Rs margin
            const isMatch = allDiffs.every(diff => diff <= 1);
            return isMatch ? 'MATCH' as const : 'MISMATCH' as const;
          })(),
          
          source: hrData ? 'both' as const : 'system_only' as const
        };
        
        return comparison;
      });
      
      const matches = comparisonResults.filter(c => c.status === 'MATCH').length;
      const mismatches = comparisonResults.filter(c => c.status === 'MISMATCH').length;
      const missing = comparisonResults.filter(c => c.status === 'MISSING').length;
      
      console.log(`✅ Comparison: ${matches} matches, ${mismatches} mismatches, ${missing} missing`);
    }

    // In your first route (document 6), at the end:
const responseData = {
  bonusCalculations: bonusCalculations.map(calc => ({
    ...calc,
    doj: calc.doj.toISOString()
  })),
  comparisonResults,
  summary,
  salarySummaries: {
    staff: staffSummary,
    worker: workerSummary
  },
  hasHRFile: !!hrFileBuffer,
  // NEW: Convert HR file to base64 to pass through JSON
  hrFileBase64: hrFileBuffer ? Buffer.from(hrFileBuffer).toString('base64') : null,
  staffFileBase64: staffBuffer ? Buffer.from(staffBuffer).toString('base64') : null,
  workerFileBase64: workerBuffer ? Buffer.from(workerBuffer).toString('base64') : null
};

    return NextResponse.json(responseData);
    
  } catch (error) {
    console.error('❌ Processing error:', error);
    return NextResponse.json({ 
      error: 'Processing failed', 
      details: error instanceof Error ? error.message : 'Unknown error' 
    }, { status: 500 });
  }
}
