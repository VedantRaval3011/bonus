import { NextRequest, NextResponse } from 'next/server';
import { ExcelProcessor } from '@/lib/excel';
import { BonusCalculator } from '@/lib/bonus-calculator';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const staffFile = formData.get('staff') as File;
    const workerFile = formData.get('worker') as File;
    const hrComparisonFile = formData.get('hrComparison') as File | null;

    if (!staffFile || !workerFile) {
      return NextResponse.json({ error: 'Missing required files' }, { status: 400 });
    }

    console.log('Processing files...');
    
    const staffBuffer = await staffFile.arrayBuffer();
    const workerBuffer = await workerFile.arrayBuffer();
    
    const staffResult = await ExcelProcessor.parseStaffFile(staffBuffer);
const workerResult = await ExcelProcessor.parseWorkerFile(workerBuffer);

// Extract employees and summaries
const staffData = staffResult.employees;
const staffSummary = staffResult.summary;
const workerData = workerResult.employees;
const workerSummary = workerResult.summary;

// ✅ Validate that data is arrays before using them
if (!Array.isArray(staffData)) {
  console.error('Staff data is not an array:', typeof staffData);
  return NextResponse.json({ 
    error: 'Staff data processing failed', 
    details: 'Staff data is not iterable' 
  }, { status: 500 });
}

if (!Array.isArray(workerData)) {
  console.error('Worker data is not an array:', typeof workerData);
  return NextResponse.json({ 
    error: 'Worker data processing failed', 
    details: 'Worker data is not iterable' 
  }, { status: 500 });
}

console.log(`Parsed ${staffData.length} staff and ${workerData.length} workers`);
console.log('Staff Summary:', staffSummary);
console.log('Worker Summary:', workerSummary);

    
    const allEmployees = [...staffData, ...workerData];
    const bonusCalculations = BonusCalculator.calculateBonus(allEmployees);
    
    // Create department breakdown
    const departmentBreakdown = bonusCalculations.reduce((acc, calc) => {
      const dept = calc.department || 'Unknown';
      if (!acc[dept]) {
        acc[dept] = { employees: 0, totalBonus: 0 };
      }
      acc[dept].employees += 1;
      acc[dept].totalBonus += calc.isEligible ? calc.finalBonus : 0;
      return acc;
    }, {} as {[key: string]: {employees: number, totalBonus: number}});
    
    const summary = {
      totalEmployees: allEmployees.length,
      eligibleEmployees: bonusCalculations.filter(c => c.isEligible).length,
      ineligibleEmployees: bonusCalculations.filter(c => !c.isEligible).length,
      totalBonusAmount: bonusCalculations.reduce((sum, c) => sum + (c.isEligible ? c.finalBonus : 0), 0),
      departmentBreakdown
    };

    let comparisonResults = null;
    
    if (hrComparisonFile) {
      console.log('Processing HR comparison file...');
      const hrBuffer = await hrComparisonFile.arrayBuffer();
      // ✅ ADD AWAIT HERE TOO
      const hrBonusMap = await ExcelProcessor.parseHRComparisonFile(hrBuffer);
      
      comparisonResults = bonusCalculations.map(calc => {
        const hrData = hrBonusMap.get(calc.empId);
        const hrBonus = hrData?.bonus || 0;
        const department = hrData?.department || calc.department || 'Unknown';
        const difference = calc.finalBonus - hrBonus;
        
        return {
          empId: calc.empId,
          name: calc.name,
          department,
          systemBonus: calc.finalBonus,
          hrBonus,
          difference,
          status: Math.abs(difference) < 1 ? 'MATCH' : 'MISMATCH'
        };
      });
      
      console.log(`Found ${comparisonResults.filter(c => c.status === 'MATCH').length} matches`);
    }

    const responseData = {
  staffData: staffData.map(emp => ({
    ...emp,
    doj: emp.doj.toISOString()
  })),
  workerData: workerData.map(emp => ({
    ...emp,
    doj: emp.doj.toISOString()
  })),
  bonusCalculations: bonusCalculations.map(calc => ({
    ...calc,
    doj: calc.doj.toISOString()
  })),
  comparisonResults,
  summary,
  // Add salary summaries
  salarySummaries: {
    staff: staffSummary,
    worker: workerSummary
  }
};


    return NextResponse.json(responseData);
    
  } catch (error) {
    console.error('Processing error:', error);
    return NextResponse.json({ 
      error: 'Processing failed', 
      details: error instanceof Error ? error.message : 'Unknown error' 
    }, { status: 500 });
  }
}
