import type { VariableMeasure } from '../types/spss';

interface MeasureIconProps {
  measure: VariableMeasure;
  className?: string;
}

export function MeasureIcon({ measure, className = 'w-4 h-4' }: MeasureIconProps) {
  if (measure === 'scale') {
    return (
      <svg
        viewBox="0 0 16 16"
        fill="none"
        className={`${className} text-amber-600 shrink-0`}
        aria-hidden="true"
      >
        <path
          d="M2.5 11.5L11.5 2.5L13.5 4.5L4.5 13.5L2.5 11.5Z"
          stroke="currentColor"
          strokeWidth="1.4"
          strokeLinejoin="round"
        />
        <path
          d="M6 8L7.2 9.2M8 6L9.2 7.2M10 4L11.2 5.2"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinecap="round"
        />
      </svg>
    );
  }

  if (measure === 'ordinal') {
    return (
      <svg
        viewBox="0 0 16 16"
        fill="none"
        className={`${className} text-sky-600 shrink-0`}
        aria-hidden="true"
      >
        <rect x="2.5" y="10" width="2.8" height="3.5" rx="0.5" fill="currentColor" />
        <rect x="6.6" y="6.5" width="2.8" height="7" rx="0.5" fill="currentColor" />
        <rect x="10.7" y="2.5" width="2.8" height="11" rx="0.5" fill="currentColor" />
      </svg>
    );
  }

  return (
    <svg
      viewBox="0 0 16 16"
      fill="none"
      className={`${className} text-blue-700 shrink-0`}
      aria-hidden="true"
    >
      <circle cx="5" cy="10.5" r="2.7" stroke="currentColor" strokeWidth="1.4" />
      <circle cx="11" cy="10.5" r="2.7" stroke="currentColor" strokeWidth="1.4" />
      <circle cx="8" cy="5.2" r="2.7" stroke="currentColor" strokeWidth="1.4" />
    </svg>
  );
}
