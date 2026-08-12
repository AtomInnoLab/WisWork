import React from 'react'

export function WisWorkAppMark({
  size = 28,
  className,
}: {
  readonly size?: number
  readonly className?: string
}): React.JSX.Element {
  return (
    <svg
      className={className}
      width={size}
      height={size}
      viewBox="0 0 64 64"
      role="img"
      aria-label="WisWork"
    >
      <defs>
        <radialGradient id="wiswork-app-bg" cx="50%" cy="42%" r="72%">
          <stop offset="0" stopColor="#281a82" />
          <stop offset="1" stopColor="#12094f" />
        </radialGradient>
        <linearGradient id="wiswork-app-blue" x1="10" y1="12" x2="36" y2="53">
          <stop stopColor="#44b7ff" />
          <stop offset="1" stopColor="#087cff" />
        </linearGradient>
        <linearGradient id="wiswork-app-violet" x1="56" y1="12" x2="36" y2="53">
          <stop stopColor="#ad82ff" />
          <stop offset="1" stopColor="#7a49df" />
        </linearGradient>
      </defs>
      <rect width="64" height="64" rx="14" fill="url(#wiswork-app-bg)" />
      <path d="M9 16h12l12 34H21L9 16Z" fill="url(#wiswork-app-blue)" />
      <path d="m27 37 7-16 6 16-7 13-6-13Z" fill="#0876db" />
      <path d="M43 16h12L43 50H32l11-34Z" fill="url(#wiswork-app-violet)" />
      <path d="m34 21 11 29H34L27 37l7-16Z" fill="#a87aff" />
    </svg>
  )
}
