import type { ReactNode } from 'react'

export interface IconProps {
  size?: number
}

function Svg({ size = 16, children }: IconProps & { children: ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.2"
      strokeLinecap="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

export function IconSend(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.2 8 13.8 2.6 11 13.4 7.6 9.6z" strokeLinejoin="round" />
      <path d="M7.6 9.6 13.8 2.6" />
    </Svg>
  )
}

export function IconStop(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4" y="4" width="8" height="8" rx="1.5" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/** return/enter arrow (↵) for the icon-only send button */
export function IconEnter(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M13 3.5v4a2.5 2.5 0 0 1-2.5 2.5H3.5" />
      <path d="M6.5 7 3.5 10l3 3" />
    </Svg>
  )
}

/** Thin paperclip used by every agent composer. */
export function IconPaperclip(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M5.25 8.75 9.9 4.1a2.1 2.1 0 0 1 3 3l-5.8 5.8a3.5 3.5 0 0 1-5-5l5.65-5.65" />
      <path d="m5.2 8.8 4.9-4.9" />
    </Svg>
  )
}

export function IconSidebarCollapse(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="3.76" width="10.01" height="8.47" rx="0.77" />
      <path d="M 9.93 3.76 v 8.47" />
      <path d="M 4.54 8 h 3.39 M 6.61 6.38 8.23 8 l -1.62 1.62" strokeWidth="1" />
    </Svg>
  )
}
