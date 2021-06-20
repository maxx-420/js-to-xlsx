// vite config

import legacy from '@vitejs/plugin-legacy'

export default {
  plugins: [
    legacy({
      targets: ['ie >= 11'],
      additionalLegacyPolyfills: ['regenerator-runtime/runtime']
    })
  ]
}