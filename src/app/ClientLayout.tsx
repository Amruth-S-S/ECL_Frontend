'use client';

import { usePathname, useRouter } from 'next/navigation';
import Sidebar from './components/Sidebar';
import { useEffect, useRef } from 'react';
import { I18nextProvider } from 'react-i18next';
import i18n from './i18n';
import { LanguageProvider } from './context/LanguageContext';

export default function ClientLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  const pathname = usePathname();
  const router = useRouter();
  const isLoginPage = pathname === '/' || pathname === '/login' || pathname === '/Login';
  const isConsultantPage = pathname === '/consultant' || pathname === '/Consultant';
  const isCXOPage = pathname === '/cxo' || pathname === '/CXO';

  useEffect(() => {
    // console.log('Current pathname:', pathname);
  }, [pathname, isConsultantPage, isCXOPage]);

  // ─── Neutralize the browser's Back / Forward / Reload buttons ───────────
  // A page can't actually disable or hide the browser's own toolbar buttons —
  // that's not exposed to any website, by design. This is the closest
  // practical equivalent, applied globally (login page and every page after
  // login): pressing Back/Forward snaps straight back to wherever the user
  // currently is instead of navigating away, and reload/close asks for
  // confirmation first instead of just leaving.
  //
  // Why this needs to be more than "remember the url and replace() on
  // popstate": the browser has already changed window.location to the
  // target page before ANY popstate listener runs — ours and Next's own
  // internal one both see the new URL at the same time. If we react
  // immediately, Next's router transition (which is what actually swaps the
  // rendered page) can finish after us and win the race, so the page still
  // visibly changes. Two things fix that:
  //   1. Only treat a URL as "confirmed" 300ms after we've settled on it —
  //      a pathname change we're about to reverse never gets confirmed, so
  //      the ref can't be corrupted by the very navigation we're undoing.
  //   2. Wait two animation frames before correcting — by then Next has
  //      definitely finished its own transition, so our replace() is
  //      guaranteed to be the last word. There's still an unavoidable brief
  //      flash of the target page first; no page can prevent that outright
  //      (this is also why "it doesn't move" on some sites is usually just
  //      that tab having no history to go back to, not a special trick).
  const confirmedUrlRef = useRef(typeof window !== 'undefined' ? window.location.href : '');
  const confirmTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  useEffect(() => {
    if (typeof window === 'undefined') return;
    if (confirmTimerRef.current) clearTimeout(confirmTimerRef.current);
    confirmTimerRef.current = setTimeout(() => {
      confirmedUrlRef.current = window.location.href;
    }, 300);
    return () => {
      if (confirmTimerRef.current) clearTimeout(confirmTimerRef.current);
    };
  }, [pathname]);

  useEffect(() => {
    if (typeof window === 'undefined') return;

    const onPopState = () => {
      const target = confirmedUrlRef.current;
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          if (window.location.href !== target) {
            router.replace(target.replace(window.location.origin, ''));
          }
        });
      });
    };
    window.addEventListener('popstate', onPopState);

    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      e.preventDefault();
      e.returnValue = '';
    };
    window.addEventListener('beforeunload', onBeforeUnload);

    return () => {
      window.removeEventListener('popstate', onPopState);
      window.removeEventListener('beforeunload', onBeforeUnload);
    };
  }, [router]);

  return (
    <I18nextProvider i18n={i18n}>
      <LanguageProvider>
        <div className="flex h-screen">
          {!isLoginPage && !isConsultantPage && !isCXOPage && (
            <Sidebar clientUserId={''} />
          )}
          <div className="w-full overflow-auto">{children}</div>
        </div>
      </LanguageProvider>
    </I18nextProvider>
  );
}
