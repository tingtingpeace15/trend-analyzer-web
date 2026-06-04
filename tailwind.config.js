/** @type {import('tailwindcss').Config} */
// 色卡 / 字体 / 动画照搬旧版「趋势分析工具 -最终app版本/tailwind.config.js」,保持 UI 风格一致
export default {
  content: ['./index.html', './src/**/*.{tsx,ts}'],
  theme: {
    extend: {
      colors: {
        bg: '#FAFAFA',
        card: '#FFFFFF',
        line: '#EEEEEE',
        ink: '#1A1A1A',
        ink2: '#666666',
        ink3: '#999999',
        brand: '#E74C3C',
        'brand-soft': '#FDECEA',
        'brand-deep': '#B83321',
        ok: '#10B981',
        warn: '#F59E0B',
        'warn-soft': '#FEF3C7',
        term: '#0F0F0F',
        'term-text': '#E5E5E5',
        'term-dim': '#7A7A7A',
        'term-line': '#1F1F1F',
      },
      fontFamily: {
        sans: ['Inter', 'Noto Sans SC', 'system-ui', 'sans-serif'],
        mono: ['JetBrains Mono', 'SF Mono', 'ui-monospace', 'monospace'],
      },
      boxShadow: {
        soft: '0 1px 3px rgba(0,0,0,0.04)',
        softer: '0 1px 2px rgba(0,0,0,0.03)',
        btn: '0 1px 3px rgba(231,60,46,0.18)',
      },
      borderRadius: { card: '10px', btn: '8px' },
      keyframes: {
        shake: {
          '0%,100%': { transform: 'translateX(0)' },
          '20%': { transform: 'translateX(-6px)' },
          '40%': { transform: 'translateX(6px)' },
          '60%': { transform: 'translateX(-4px)' },
          '80%': { transform: 'translateX(4px)' },
        },
        fadeup: {
          from: { opacity: '0', transform: 'translateY(12px)' },
          to: { opacity: '1', transform: 'translateY(0)' },
        },
        logline: {
          from: { opacity: '0', transform: 'translateY(4px)' },
          to: { opacity: '1', transform: 'translateY(0)' },
        },
        pulseRing: {
          '0%': { boxShadow: '0 0 0 0 rgba(231,60,46,0.45)' },
          '70%': { boxShadow: '0 0 0 8px rgba(231,60,46,0)' },
          '100%': { boxShadow: '0 0 0 0 rgba(231,60,46,0)' },
        },
        caret: {
          '0%,49%': { opacity: '1' },
          '50%,100%': { opacity: '0' },
        },
      },
      animation: {
        shake: 'shake 360ms ease-in-out',
        fadeup: 'fadeup 150ms ease-out both',
        logline: 'logline 180ms ease-out both',
        pulsering: 'pulseRing 1.6s ease-out infinite',
        caret: 'caret 1s steps(1) infinite',
      },
    },
  },
  plugins: [],
};
