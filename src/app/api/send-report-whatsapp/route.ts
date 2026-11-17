import { NextRequest, NextResponse } from 'next/server';

// Allow requests from any domain (localhost, matga.com, etc.)
const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-Requested-With',
};

export async function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: corsHeaders });
}

// The user only ever supplies a phone number — the WhatsApp template and the
// attached document's file name are fixed on purpose.
const WHATSAPP_TEMPLATE_NAME = 'document';
const REPORT_FILE_NAME = 'Analysis_Report.pptx';

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { phoneNumber, pptBase64 } = body;

    // --- Validation ---
    if (!phoneNumber || typeof phoneNumber !== 'string') {
      return NextResponse.json(
        { success: false, message: 'Missing or invalid phone number' },
        { status: 400, headers: corsHeaders }
      );
    }

    const cleanedNumber = phoneNumber.replace(/[^0-9]/g, '');
    if (cleanedNumber.length < 10) {
      return NextResponse.json(
        { success: false, message: 'Please enter a valid phone number with country code (e.g. 919876543210)' },
        { status: 400, headers: corsHeaders }
      );
    }

    if (!pptBase64 || typeof pptBase64 !== 'string') {
      return NextResponse.json(
        { success: false, message: 'Missing pptBase64 — PPT generation may have failed' },
        { status: 400, headers: corsHeaders }
      );
    }

    // --- Env var check ---
    const apiBaseUrl = process.env.WHATSAPP_API_BASE_URL;
    const apiKey = process.env.WHATSAPP_API_KEY;

    if (!apiBaseUrl || !apiKey) {
      console.error('[send-report-whatsapp] WHATSAPP_API_BASE_URL or WHATSAPP_API_KEY is not set in environment variables.');
      return NextResponse.json(
        { success: false, message: 'WhatsApp service is not configured on this server. Please contact support.' },
        { status: 503, headers: corsHeaders }
      );
    }

    // --- Build multipart form-data & send to the WhatsApp integration service ---
    const pptBuffer = Buffer.from(pptBase64, 'base64');
    const fileBlob = new Blob([pptBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    });

    const form = new FormData();
    form.append('to', cleanedNumber);
    form.append('language_code', 'en');
    form.append('file', fileBlob, REPORT_FILE_NAME);

    const response = await fetch(
      `${apiBaseUrl.replace(/\/$/, '')}/templates/${WHATSAPP_TEMPLATE_NAME}/send-document`,
      {
        method: 'POST',
        headers: { 'X-API-Key': apiKey },
        body: form,
      }
    );

    const rawText = await response.text();
    let result: any = {};
    try { result = rawText ? JSON.parse(rawText) : {}; } catch { result = { message: rawText }; }

    if (!response.ok) {
      const errMsg = result?.detail
        ? (Array.isArray(result.detail) ? result.detail.map((d: any) => d.msg || JSON.stringify(d)).join(', ') : String(result.detail))
        : (result?.message || `HTTP ${response.status}`);
      throw new Error(errMsg);
    }

    return NextResponse.json(
      { success: true, message: 'Report sent via WhatsApp successfully!' },
      { headers: corsHeaders }
    );

  } catch (error) {
    console.error('[send-report-whatsapp] Error:', error);
    const rawMsg = error instanceof Error ? error.message : String(error);

    let displayMsg = rawMsg;
    if (rawMsg.includes('ECONNREFUSED') || rawMsg.includes('ETIMEDOUT') || rawMsg.includes('ENOTFOUND')) {
      displayMsg = 'Cannot connect to the WhatsApp integration service. Please try again later.';
    } else if (rawMsg.toLowerCase().includes('phone number') || rawMsg.toLowerCase().includes('recipient')) {
      displayMsg = 'The phone number is invalid or not reachable on WhatsApp.';
    }

    return NextResponse.json(
      { success: false, message: displayMsg, error: rawMsg },
      { status: 500, headers: corsHeaders }
    );
  }
}
