import { NextRequest, NextResponse } from 'next/server';

const FILTER_DATA_BASE = process.env.FILTER_DATA_BASE_URL || 'https://filter-data-isc-35486280762.us-central1.run.app';
const API_KEY = process.env.FILTER_DATA_API_KEY || 'ytsyxvtywfqvg@1209567TEAHBjhbwxuywbPo()*&%$Rssx';

// Server-side proxy so the browser calls same-origin /api/filter-data/* instead of
// the Cloud Run service directly — avoids CORS preflight failures on the custom
// x-api-key header, which the Cloud Run backend does not allow cross-origin.
export async function GET(req: NextRequest, { params }: { params: Promise<{ path: string[] }> }) {
  const { path } = await params;
  const { search } = new URL(req.url);

  try {
    const res = await fetch(`${FILTER_DATA_BASE}/${path.join('/')}${search}`, {
      headers: { 'x-api-key': API_KEY },
    });
    const body = await res.text();
    return new NextResponse(body, {
      status: res.status,
      headers: { 'Content-Type': res.headers.get('content-type') || 'application/json' },
    });
  } catch (err) {
    return NextResponse.json({ error: 'Failed to reach filter data service' }, { status: 502 });
  }
}
