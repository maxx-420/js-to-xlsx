// vite config

import legacy from '@vitejs/plugin-legacy'

export default {
  plugins: [
    legacy({
      targets: ['defaults','IE >= 10'],
      additionalLegacyPolyfills: ['regenerator-runtime/runtime']
    })
  ]
}