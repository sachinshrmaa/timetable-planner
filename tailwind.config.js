/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      fontFamily: {
        sans: ['Inter', 'sans-serif'],
        display: ['Outfit', 'sans-serif'],
      },
      colors: {
        slate: {
          950: '#030712',
        },
        indigo: {
          500: '#6366f1',
          600: '#4f46e5',
        },
        violet: {
          500: '#8b5cf6',
          600: '#7c3aed',
        }
      },
      boxShadow: {
        'glow-indigo': '0 0 20px rgba(99, 102, 241, 0.15)',
        'glow-violet': '0 0 20px rgba(139, 92, 246, 0.2)',
      }
    },
  },
  plugins: [],
}
