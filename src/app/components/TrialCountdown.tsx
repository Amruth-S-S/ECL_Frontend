"use client";

import { useEffect, useState } from "react";
import { Clock, X } from "lucide-react";

/**
 * TrialCountdown
 *
 * Shows how many days are left on the logged-in user's trial — a small
 * always-visible badge in the header, plus a reminder popup the first time
 * the app is opened each day (so it's impossible to miss, without nagging on
 * every single click/navigation).
 *
 * Reads `trial_end_date` from GET /client-users/{userId} (the same field
 * visible in the Client Users API) for whichever user is currently logged
 * in. Renders nothing for non-trial accounts or once no trial data is found.
 */
export default function TrialCountdown() {
  const [daysLeft, setDaysLeft] = useState<number | null>(null);
  const [trialEndDate, setTrialEndDate] = useState<string | null>(null);
  const [showPopup, setShowPopup] = useState(false);

  useEffect(() => {
    const API_BASE_URL = process.env.NEXT_PUBLIC_API_BASE_URL || "";
    const API_KEY = process.env.NEXT_PUBLIC_API_KEY || "";

    let userId: string | null = null;
    try {
      const raw = localStorage.getItem("currentUserData");
      if (raw) userId = String(JSON.parse(raw)?.userId || "") || null;
    } catch { /* ignore */ }
    if (!userId) return;

    fetch(`${API_BASE_URL}/client-users/${userId}`, {
      headers: { Accept: "application/json", "X-API-Key": API_KEY },
    })
      .then((res) => (res.ok ? res.json() : null))
      .then((data) => {
        if (!data?.trial_end_date) return;
        // Only relevant while the account is actually on a trial —
        // once upgraded, subscription won't be "Trial" any more.
        if (data.subscription && data.subscription !== "Trial") return;

        const end = new Date(data.trial_end_date);
        if (isNaN(end.getTime())) return;

        const msPerDay = 1000 * 60 * 60 * 24;
        const days = Math.ceil((end.getTime() - Date.now()) / msPerDay);

        setDaysLeft(days);
        setTrialEndDate(data.trial_end_date);

        // Popup once per calendar day per browser, so it's seen without
        // showing up again on every tab switch / re-render.
        const today = new Date().toDateString();
        const lastShown = localStorage.getItem(`trialPopupShownDate_${userId}`);
        if (lastShown !== today) {
          setShowPopup(true);
          localStorage.setItem(`trialPopupShownDate_${userId}`, today);
        }
      })
      .catch(() => { /* silently skip — trial badge is a nice-to-have */ });
  }, []);

  if (daysLeft === null) return null;

  const expired = daysLeft <= 0;
  const urgent = !expired && daysLeft <= 3;
  const warning = !expired && daysLeft > 3 && daysLeft <= 7;

  const badgeClasses = expired || urgent
    ? "bg-red-50 text-red-700 border-red-200"
    : warning
      ? "bg-amber-50 text-amber-700 border-amber-200"
      : "bg-blue-50 text-blue-700 border-blue-200";

  const formattedEndDate = trialEndDate
    ? new Date(trialEndDate).toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" })
    : "";

  return (
    <>
      {/* Persistent header badge — kept large/prominent at all times, with a
          gentle zoom-in/zoom-out pulse so it keeps catching the eye instead
          of blending into the rest of the header. */}
      <div
        className={`trial-badge-pulse flex items-center gap-1.5 sm:gap-2 px-2.5 py-1 sm:px-4 sm:py-2 rounded-full text-xs sm:text-sm font-bold border-2 shadow-sm flex-shrink-0 whitespace-nowrap ${badgeClasses}`}
        title={expired ? `Trial ended on ${formattedEndDate}` : `Trial ends on ${formattedEndDate}`}
      >
        <Clock className="w-3.5 h-3.5 sm:w-[18px] sm:h-[18px] flex-shrink-0" />
        {/* Shorter text on small screens so this doesn't crowd out the rest
            of the header on a phone. */}
        <span className="sm:hidden">{expired ? "Expired" : `${daysLeft}d left`}</span>
        <span className="hidden sm:inline">{expired ? "Trial expired" : `${daysLeft} day${daysLeft === 1 ? "" : "s"} left in trial`}</span>
      </div>

      <style jsx>{`
        .trial-badge-pulse {
          animation: trialZoomPulse 1.8s ease-in-out infinite;
          transform-origin: center;
        }
        @keyframes trialZoomPulse {
          0%, 100% { transform: scale(1); }
          50% { transform: scale(1.08); }
        }
        @media (prefers-reduced-motion: reduce) {
          .trial-badge-pulse { animation: none; }
        }
      `}</style>

      {/* Once-a-day reminder popup */}
      {showPopup && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[9999] p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-sm w-full overflow-hidden">
            <div className={`flex items-center justify-between px-5 py-3.5 border-b ${
              expired || urgent ? "bg-gradient-to-r from-red-50 to-rose-50 border-red-100"
                : warning ? "bg-gradient-to-r from-amber-50 to-yellow-50 border-amber-100"
                : "bg-gradient-to-r from-blue-50 to-indigo-50 border-blue-100"
            }`}>
              <h3 className={`text-base font-bold flex items-center gap-1.5 ${
                expired || urgent ? "text-red-700" : warning ? "text-amber-700" : "text-blue-700"
              }`}>
                <Clock className="w-4 h-4 flex-shrink-0" />
                {expired ? "Trial Expired" : "Trial Reminder"}
              </h3>
              <button
                onClick={() => setShowPopup(false)}
                className="p-1 text-gray-400 hover:text-gray-600 hover:bg-white rounded-full transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="p-5">
              <p className="text-sm text-gray-700">
                {expired
                  ? `Your trial ended on ${formattedEndDate}. Please upgrade to keep using OneVega.`
                  : <>Your trial has <strong>{daysLeft} day{daysLeft === 1 ? "" : "s"}</strong> left, ending on <strong>{formattedEndDate}</strong>.</>}
              </p>
              <button
                onClick={() => setShowPopup(false)}
                className="mt-4 w-full py-2 text-sm bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 transition-colors"
              >
                Got it
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
