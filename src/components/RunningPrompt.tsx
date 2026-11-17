'use client';

import React, { useEffect, useState } from 'react';

type Props = {
  /** Show this card while your prompt request is in progress */
  onStop?: () => void;
  title?: string;
  /** One entry per step chip. `status` is the sentence shown under the title. */
  steps?: { chip: string; status: string }[];
  /** How long each step stays active before moving to the next (ms) */
  stepMs?: number;
};

const DEFAULT_STEPS = [
  { chip: 'Read data', status: 'Reading your data…' },
  { chip: 'Find answer', status: 'Finding the answer…' },
  { chip: 'Build charts', status: 'Building the charts…' },
];

/**
 * "Running your prompt" card: three floating circles, a status line,
 * step chips and a progress bar, with an optional Stop button.
 *
 * Mount it while the request is in flight and unmount it when the data arrives:
 *
 *   {loading && <RunningPrompt onStop={handleStop} />}
 *
 * The steps advance on a timer and wait on the last one until the card is
 * unmounted, so the bar never claims to be finished before your API answers.
 */
export default function RunningPrompt({
  onStop,
  title = 'Running your prompt',
  steps = DEFAULT_STEPS,
  stepMs = 1100,
}: Props) {
  const [step, setStep] = useState(0);
  const last = steps.length - 1;

  useEffect(() => {
    if (step >= last) return;
    const t = setTimeout(() => setStep(s => s + 1), stepMs);
    return () => clearTimeout(t);
  }, [step, last, stepMs]);

  // 8% to 92%: the last 8% is only reached when the card goes away
  const pct = Math.round(8 + ((step + 0.5) / steps.length) * 84);

  return (
    <div className="rp-card" role="status" aria-live="polite">
      <div className="rp-stage" aria-hidden="true"><i /><i /><i /></div>
      <h2 className="rp-title">{title}</h2>
      <p className="rp-status">{steps[step].status}</p>

      <ol className="rp-steps">
        {steps.map((s, i) => (
          <li key={s.chip} className={i < step ? 'done' : i === step ? 'active' : ''}>
            <span className="rp-d" />
            {s.chip}
          </li>
        ))}
      </ol>

      <div className="rp-bar" aria-hidden="true"><i style={{ width: `${pct}%` }} /></div>

      {onStop && (
        <button type="button" onClick={onStop} className="rp-stop">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><rect x="6" y="6" width="12" height="12" rx="2" /></svg>
          Stop
        </button>
      )}

      <style>{`
        .rp-card { border: 1px solid #DCE5F5; border-radius: 20px; background: #fff; padding: 44px 20px 36px;
                   display: flex; flex-direction: column; align-items: center; text-align: center; gap: 12px; }
        .rp-title { margin: 0; font-size: 22px; font-weight: 700; color: #14203D; }
        .rp-status { margin: 0; min-height: 1.5em; color: #5B6A8C; }

        /* three circles */
        .rp-stage { position: relative; width: 130px; height: 130px; margin-bottom: 6px; }
        .rp-stage i { position: absolute; border-radius: 50%; mix-blend-mode: multiply; will-change: transform; }
        .rp-stage i:nth-child(1) { width: 56%; height: 56%; left: 22%; top: 0; background: #4DD6F7; animation: rp-a 2.4s ease-in-out infinite; }
        .rp-stage i:nth-child(2) { width: 62%; height: 62%; left: 0; bottom: 0; background: #069BCB; animation: rp-b 2.4s ease-in-out infinite; }
        .rp-stage i:nth-child(3) { width: 40%; height: 40%; right: 0; top: 40%; background: #1FBDB8; animation: rp-c 2.4s ease-in-out infinite; }
        @keyframes rp-a { 0%,100% { transform: translate(0,0) scale(1); } 33% { transform: translate(18%,14%) scale(.86); } 66% { transform: translate(-14%,22%) scale(1.08); } }
        @keyframes rp-b { 0%,100% { transform: translate(0,0) scale(1); } 33% { transform: translate(14%,-16%) scale(1.1); } 66% { transform: translate(22%,4%) scale(.9); } }
        @keyframes rp-c { 0%,100% { transform: translate(0,0) scale(1); } 33% { transform: translate(-40%,-22%) scale(1.25); } 66% { transform: translate(-22%,30%) scale(.85); } }

        /* step chips */
        .rp-steps { list-style: none; margin: 10px 0 4px; padding: 0; display: flex; gap: 8px; flex-wrap: wrap; justify-content: center; }
        .rp-steps li { display: inline-flex; align-items: center; gap: 8px; padding: 7px 14px; border-radius: 99px;
                       border: 1px solid #DCE5F5; background: #F5F8FD; color: #5B6A8C; font-weight: 600; font-size: 13px; transition: all .25s; }
        .rp-d { width: 10px; height: 10px; border-radius: 50%; border: 2px solid #B7CBF5; }
        .rp-steps li.active { color: #2563EB; border-color: #2563EB; background: #E8EFFD; }
        .rp-steps li.active .rp-d { border-color: #2563EB; border-top-color: transparent; animation: rp-spin .8s linear infinite; }
        .rp-steps li.done { color: #14804A; border-color: rgba(20,128,74,.4); }
        .rp-steps li.done .rp-d { background: #14804A; border-color: #14804A; }
        @keyframes rp-spin { to { transform: rotate(360deg); } }

        /* progress bar */
        .rp-bar { width: min(340px, 80%); height: 6px; border-radius: 99px; background: #E8EFFD; overflow: hidden; margin-top: 6px; }
        .rp-bar i { display: block; height: 100%; border-radius: 99px; background: linear-gradient(90deg, #2563EB, #14A59B); transition: width .5s ease; }

        .rp-stop { margin-top: 10px; display: inline-flex; align-items: center; gap: 8px; height: 42px; padding: 0 18px;
                   border-radius: 11px; border: 1px solid rgba(209,59,78,.35); background: #FDECEF; color: #D13B4E; font-weight: 700; cursor: pointer; }
        .rp-stop:hover { background: #D13B4E; color: #fff; }

        @media (prefers-reduced-motion: reduce) { .rp-stage i { animation-duration: 8s; } }
      `}</style>
    </div>
  );
}
