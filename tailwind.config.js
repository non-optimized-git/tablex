/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        muji: {
          beige: '#D5C8B5',
          cream: '#F5F0E8',
          charcoal: '#3D3D3D',
          gray: '#6B6B6B',
          lightGray: '#E8E4DC',
          warmWhite: '#FAF8F5',
        },
        primary: {
          DEFAULT: '#1E40AF',
          hover: '#1E3A8A',
          light: '#DBEAFE',
        },
        accent: '#D97706',
        success: '#059669',
        destructive: '#DC2626',
        muted: '#F1F5F9',
      },
      fontFamily: {
        sans: ['Inter', '-apple-system', 'BlinkMacSystemFont', 'Segoe UI', 'Hiragino Sans', 'sans-serif'],
      },
      boxShadow: {
        'soft': '0 2px 8px rgba(0, 0, 0, 0.06)',
        'hover': '0 4px 16px rgba(0, 0, 0, 0.1)',
      },
      transitionDuration: {
        '200': '200ms',
      },
    },
  },
  plugins: [],
}