import { NextRequest, NextResponse } from 'next/server';
import { ExcelProcessor } from '@/lib/excel';

export async function POST(request: NextRequest) {
  try {
    const { comparisons } = await request.json();
    
    if (!comparisons || !Array.isArray(comparisons)) {
      return NextResponse.json({ error: 'Invalid comparison data' }, { status: 400 });
    }

    // This now returns Uint8Array
    const excelBuffer = await ExcelProcessor.generateComparisonReport(comparisons);
    
    return new NextResponse(excelBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="Comparison-Report.xlsx"'
      }
    });
    
  } catch (error) {
    console.error('Comparison error:', error);
    return NextResponse.json({ 
      error: 'Comparison failed',
      details: error instanceof Error ? error.message : 'Unknown error'
    }, { status: 500 });
  }
}
