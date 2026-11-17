import { NextRequest, NextResponse } from 'next/server';

/**
 * POST /api/translate
 * Body: { texts: string[], targetLang: string, sourceLang?: string }
 * Returns: { translations: string[] }
 *
 * Uses MyMemory's free translation API (one legitimate request per text —
 * it doesn't support batching either). We previously called Google's
 * unofficial translate_a/single endpoint here, but Google started returning
 * 429 "your network may be sending automated queries" for this server's
 * outbound IP, which silently broke every translation (the caller falls
 * back to the original English text on any non-2xx response). MyMemory is
 * an actual public API meant for this, not a reverse-engineered endpoint,
 * so it isn't subject to that same abuse-detection blocking.
 */

const CONCURRENCY = 6; // parallel requests per batch
const MYMEMORY_WARNING_PREFIX = 'MYMEMORY WARNING';

async function translateOne(
  text: string,
  targetLang: string,
  sourceLang: string,
): Promise<string> {
  if (!text || !text.trim()) return text;

  const url =
    `https://api.mymemory.translated.net/get` +
    `?q=${encodeURIComponent(text)}&langpair=${encodeURIComponent(sourceLang)}|${encodeURIComponent(targetLang)}`;

  try {
    const res = await fetch(url, {
      headers: { Accept: 'application/json' },
    });

    if (!res.ok) return text;

    const data = await res.json();
    const translated: string | undefined = data?.responseData?.translatedText;

    if (!translated || translated.toUpperCase().startsWith(MYMEMORY_WARNING_PREFIX)) {
      return text;
    }

    return translated;
  } catch {
    return text;
  }
}

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { texts, targetLang, sourceLang = 'en' } = body as {
      texts: string[];
      targetLang: string;
      sourceLang?: string;
    };

    if (!texts || !Array.isArray(texts) || texts.length === 0) {
      return NextResponse.json({ translations: [] });
    }

    if (sourceLang === targetLang) {
      return NextResponse.json({ translations: texts });
    }

    // Translate all texts with limited concurrency
    const translations: string[] = new Array(texts.length);

    for (let i = 0; i < texts.length; i += CONCURRENCY) {
      const chunk = texts.slice(i, i + CONCURRENCY);
      const results = await Promise.allSettled(
        chunk.map((t) => translateOne(t, targetLang, sourceLang)),
      );
      results.forEach((result, j) => {
        translations[i + j] =
          result.status === 'fulfilled' ? result.value : texts[i + j];
      });
    }

    return NextResponse.json({ translations });
  } catch {
    try {
      const fallback = await req.json().catch(() => ({ texts: [] as string[] }));
      return NextResponse.json({ translations: fallback.texts ?? [] });
    } catch {
      return NextResponse.json({ translations: [] });
    }
  }
}
