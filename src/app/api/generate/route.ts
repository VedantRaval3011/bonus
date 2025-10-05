import { NextRequest, NextResponse } from 'next/server';
import { ExcelProcessor } from '@/lib/excel';

export const runtime = 'nodejs';

export async function POST(request: NextRequest) {
  try {
    const contentType = request.headers.get('content-type') || '';

    let calculations: any[] = [];
    let hrBuffer: ArrayBuffer | undefined;
    let staffBuffer: ArrayBuffer | undefined;
let workerBuffer: ArrayBuffer | undefined;

    // In the JSON handling section (around line 18):
if (contentType.includes('application/json')) {
  const body = await request.json();
  calculations = Array.isArray(body.calculations) ? body.calculations : [];

  // Existing HR buffer code...
  if (body.hrFileBase64) {
    console.log('📥 Decoding HR file from base64...');
    const buffer = Buffer.from(body.hrFileBase64, 'base64');
    hrBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
    console.log(`✅ HR file decoded: ${hrBuffer.byteLength} bytes`);
  }

  // ADD STAFF BUFFER DECODING:
  if (body.staffFileBase64) {
    console.log('📥 Decoding Staff file from base64...');
    const buffer = Buffer.from(body.staffFileBase64, 'base64');
    staffBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
    console.log(`✅ Staff file decoded: ${staffBuffer.byteLength} bytes`);
  }

  // ADD WORKER BUFFER DECODING:
  if (body.workerFileBase64) {
    console.log('📥 Decoding Worker file from base64...');
    const buffer = Buffer.from(body.workerFileBase64, 'base64');
    workerBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
    console.log(`✅ Worker file decoded: ${workerBuffer.byteLength} bytes`);
  }
      
      // If HR file URL is provided, fetch it
      if (body.hrFileUrl) {
        console.log('📥 Fetching HR file from URL:', body.hrFileUrl);
        const resp = await fetch(body.hrFileUrl);
        if (!resp.ok) {
          throw new Error(`Failed to fetch HR file: ${resp.status} ${resp.statusText}`);
        }
        hrBuffer = await resp.arrayBuffer();
        console.log(`✅ HR file fetched: ${hrBuffer.byteLength} bytes`);
      }
    } else if (contentType.includes('multipart/form-data')) {
      const form = await request.formData();
      
      // Get calculations from form data
      const calcsField = form.get('calculations') as string | null;
      if (calcsField) {
        try {
          calculations = JSON.parse(calcsField);
        } catch (parseError) {
          console.error('Failed to parse calculations JSON:', parseError);
          return NextResponse.json(
            { error: 'Invalid calculations JSON format' },
            { status: 400 }
          );
        }
      }
      
      // Get HR file from form data
      const hrFile = form.get('hrFile') as File | null;
      if (hrFile && hrFile.size > 0) {
        console.log(`📥 Processing HR file: ${hrFile.name} (${hrFile.size} bytes)`);
        hrBuffer = await hrFile.arrayBuffer();
        console.log(`✅ HR file loaded: ${hrBuffer.byteLength} bytes`);
      }
    } else {
      return NextResponse.json(
        { error: 'Unsupported content type. Use application/json or multipart/form-data' },
        { status: 415 }
      );
    }

    // Validate calculations data
    if (!Array.isArray(calculations) || calculations.length === 0) {
      return NextResponse.json(
        { error: 'Invalid or empty calculations data' },
        { status: 400 }
      );
    }

    console.log(`📊 Processing ${calculations.length} employee calculations`);

    // Process calculations and convert dates
    const processedCalculations = calculations.map((c) => ({
      ...c,
      doj: new Date(c.doj),
    }));

    // Compute HR monthly totals if HR file was provided
    let hrMonthlyTotals: Map<string, number> | undefined;
    if (hrBuffer) {
      console.log('🔍 Computing HR monthly totals...');
      try {
        hrMonthlyTotals = await ExcelProcessor.computeHRMonthlyTotals(hrBuffer);
        console.log(`✅ HR monthly totals computed: ${hrMonthlyTotals.size} months`);
        
        // Log the totals for debugging
        const totalsArray = Array.from(hrMonthlyTotals.entries());
        console.log('HR Monthly Totals:', totalsArray.map(([k, v]) => `${k}: ₹${v.toLocaleString()}`).join(', '));
      } catch (hrError) {
        console.error('⚠️ Error computing HR totals:', hrError);
        // Continue without HR totals rather than failing completely
        hrMonthlyTotals = undefined;
      }
    } else {
      console.log('ℹ️ No HR file provided, skipping HR comparison');
    }

    // Generate final Excel with HR comparison
    console.log('📝 Generating final bonus Excel file...');
    const excelBuffer = await ExcelProcessor.generateFinalBonusExcel(
      processedCalculations,
      undefined, // staffSummary (optional)
      undefined, // workerSummary (optional)
      hrMonthlyTotals,
      staffBuffer,   // ADD THIS - your original staff file buffer
      workerBuffer   // ADD THIS - your original worker file buffer
    );

    console.log(`✅ Excel generated successfully: ${excelBuffer.byteLength} bytes`);

    // Generate filename with timestamp
    const timestamp = new Date().toISOString().split('T')[0];
    const filename = `Final-Bonus-Data-${timestamp}.xlsx`;

    return new NextResponse(excelBuffer as any, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${filename}"`,
        'Cache-Control': 'no-cache, no-store, must-revalidate',
      },
    });
  } catch (error) {
    console.error('❌ Generate error:', error);
    
    // Provide detailed error information
    const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
    const errorStack = error instanceof Error ? error.stack : undefined;
    
    console.error('Error details:', {
      message: errorMessage,
      stack: errorStack,
    });

    return NextResponse.json(
      {
        error: 'Generation failed',
        details: errorMessage,
        timestamp: new Date().toISOString(),
      },
      { status: 500 }
    );
  }
}