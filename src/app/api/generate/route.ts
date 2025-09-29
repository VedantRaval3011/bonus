import { NextRequest, NextResponse } from 'next/server';
import { ExcelProcessor } from '@/lib/excel';

export async function POST(request: NextRequest) {
  try {
    const { calculations } = await request.json();
    
    if (!calculations || !Array.isArray(calculations)) {
      return NextResponse.json({ error: 'Invalid calculations data' }, { status: 400 });
    }

    // Deserialize dates from string back to Date objects
    const processedCalculations = calculations.map(calc => ({
      ...calc,
      doj: new Date(calc.doj) // Convert string back to Date object
    }));

    // This now returns Uint8Array
    const excelBuffer = await ExcelProcessor.generateFinalBonusExcel(processedCalculations);
    
    return new NextResponse(excelBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="Final-Bonus-Data.xlsx"'
      }
    });
    
  } catch (error) {
    console.error('Generate error:', error);
    return NextResponse.json({ 
      error: 'Generation failed',
      details: error instanceof Error ? error.message : 'Unknown error'
    }, { status: 500 });
  }
}
