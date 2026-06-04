// SVG 图标,移植自旧版 src/icons.jsx。单线条、1.6 stroke-width、currentColor。
import type { ReactNode, SVGProps } from 'react';

interface IconProps extends Omit<SVGProps<SVGSVGElement>, 'd'> {
  size?: number;
}

const Icon = ({ d, size = 16, fill, ...rest }: IconProps & { d: ReactNode }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill={fill || 'none'}
    stroke="currentColor"
    strokeWidth="1.6"
    strokeLinecap="round"
    strokeLinejoin="round"
    {...rest}
  >
    {d}
  </svg>
);

export const IconUpload = (p: IconProps) => (
  <Icon {...p} d={<>
    <path d="M12 16V4" />
    <path d="M7 9l5-5 5 5" />
    <path d="M5 16v3a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-3" />
  </>} />
);

export const IconFile = (p: IconProps) => (
  <Icon {...p} d={<>
    <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z" />
    <path d="M14 3v5h5" />
  </>} />
);

export const IconClose = (p: IconProps) => (
  <Icon {...p} d={<><path d="M6 6l12 12M18 6L6 18" /></>} />
);

export const IconCheck = (p: IconProps) => (
  <Icon {...p} d={<><path d="M5 12.5l4.5 4.5L19 7.5" /></>} />
);

export const IconArrowRight = (p: IconProps) => (
  <Icon {...p} d={<><path d="M5 12h14M13 5l7 7-7 7" /></>} />
);

export const IconDownload = (p: IconProps) => (
  <Icon {...p} d={<>
    <path d="M12 4v12" />
    <path d="M7 11l5 5 5-5" />
    <path d="M5 20h14" />
  </>} />
);

export const IconSpark = (p: IconProps) => (
  <Icon {...p} d={<>
    <path d="M12 3v4" />
    <path d="M12 17v4" />
    <path d="M3 12h4" />
    <path d="M17 12h4" />
    <path d="M5.6 5.6l2.8 2.8" />
    <path d="M15.6 15.6l2.8 2.8" />
    <path d="M5.6 18.4l2.8-2.8" />
    <path d="M15.6 8.4l2.8-2.8" />
  </>} />
);

export const IconWarn = (p: IconProps) => (
  <Icon {...p} d={<>
    <path d="M12 3l10 18H2L12 3z" />
    <path d="M12 10v5" />
    <circle cx="12" cy="18" r="0.6" fill="currentColor" />
  </>} />
);

export const IconAlert = (p: IconProps) => (
  <Icon {...p} d={<>
    <circle cx="12" cy="12" r="9" />
    <path d="M12 7v6" />
    <circle cx="12" cy="16.2" r="0.6" fill="currentColor" />
  </>} />
);

export const IconChart = (p: IconProps) => (
  <Icon {...p} d={<>
    <path d="M4 19V5" />
    <path d="M4 19h16" />
    <path d="M7 15l3-4 3 2 4-6" />
  </>} />
);

export const IconLayers = (p: IconProps) => (
  <Icon {...p} d={<>
    <path d="M12 3l9 5-9 5-9-5 9-5z" />
    <path d="M3 13l9 5 9-5" />
    <path d="M3 17l9 5 9-5" />
  </>} />
);

export const IconGrid = (p: IconProps) => (
  <Icon {...p} d={<>
    <rect x="3" y="3" width="7" height="7" rx="1" />
    <rect x="14" y="3" width="7" height="7" rx="1" />
    <rect x="3" y="14" width="7" height="7" rx="1" />
    <rect x="14" y="14" width="7" height="7" rx="1" />
  </>} />
);

export const IconRefresh = (p: IconProps) => (
  <Icon {...p} d={<>
    <path d="M3 12a9 9 0 0 1 15.5-6.3L21 8" />
    <path d="M21 3v5h-5" />
    <path d="M21 12a9 9 0 0 1-15.5 6.3L3 16" />
    <path d="M3 21v-5h5" />
  </>} />
);

export const IconLogo = ({ size = 20 }: { size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none">
    <rect x="2" y="2" width="20" height="20" rx="5" fill="#1A1A1A" />
    <path d="M6 16l3-5 3 2.5 5-7" stroke="#E74C3C" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" fill="none" />
    <circle cx="17" cy="6.5" r="1.4" fill="#E74C3C" />
  </svg>
);
