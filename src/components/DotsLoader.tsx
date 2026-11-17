"use client";

interface DotsLoaderProps {
  /** Overall size in px — scales the dots proportionally. */
  size?: number;
  /** Text under the dots. Pass "" to show dots only. */
  label?: string;
  /** Render as a fixed, full-screen overlay instead of inline. */
  fullscreen?: boolean;
  className?: string;
}

export default function DotsLoader({
  size = 64,
  label = "Thinking...",
  fullscreen = false,
  className = "",
}: DotsLoaderProps) {
  const dotSize = Math.max(6, Math.round(size / 6));
  const gap = Math.max(4, Math.round(size / 8));

  const content = (
    <div className={`flex flex-col items-center justify-center gap-3 ${className}`}>
      <div className="flex items-end" style={{ gap }}>
        {[0, 1, 2].map((i) => (
          <span
            key={i}
            className="rounded-full bg-blue-600 animate-bounce"
            style={{
              width: dotSize,
              height: dotSize,
              animationDelay: `${i * 0.15}s`,
              animationDuration: "0.9s",
            }}
          />
        ))}
      </div>
      {label && <p className="text-base font-medium text-gray-600">{label}</p>}
    </div>
  );

  if (fullscreen) {
    return (
      <div className="fixed inset-0 z-[200] flex items-center justify-center bg-black/30 backdrop-blur-sm">
        <div className="bg-white rounded-2xl shadow-2xl px-10 py-8">
          {content}
        </div>
      </div>
    );
  }

  return content;
}
