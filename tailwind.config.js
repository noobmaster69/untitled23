export default {
  darkMode: 'class',
  content: ['./index.html', './src/**/*.{js,jsx}'],
  theme: {
    extend: {
      fontFamily: {
        display: [
          '"OpenAI Sans"',
          '"Circle"',
          '"Trebuchet MS"',
          '"Lucida Sans Unicode"',
          '"Lucida Grande"',
          'Verdana',
          'sans-serif',
        ],
        mono: [
          '"Atkinson Hyperlegible Mono"',
          '"SFMono-Regular"',
          'Menlo',
          'Consolas',
          'Liberation Mono',
          'monospace',
        ],
      },
    },
  },
  plugins: [],
}
