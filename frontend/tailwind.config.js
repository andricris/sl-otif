/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{js,jsx}"],
  theme: {
    extend: {
      colors: {
        paper: "#E9ECE6",
        "paper-2": "#DFE3D9",
        sheet: "#F6F8F4",
        ink: "#191C18",
        "ink-2": "#363B35",
        muted: "#565C53",
        rule: "#CFD4C9",
        "rule-2": "#AEB4A6",
        accent: "#A33A2A",
        good: "#2E6B4A",
        warn: "#8A6516",
      },
      fontFamily: {
        display: ['"Newsreader"', "Georgia", "serif"],
        sans: ['"IBM Plex Sans"', '"Helvetica Neue"', "sans-serif"],
        mono: ['"IBM Plex Mono"', "ui-monospace", "monospace"],
      },
      keyframes: {
        rise: {
          "0%": { opacity: "0", transform: "translateY(6px)" },
          "100%": { opacity: "1", transform: "none" },
        },
      },
      animation: {
        rise: "rise .35s cubic-bezier(.2,.7,.2,1) both",
      },
    },
  },
  plugins: [],
};
