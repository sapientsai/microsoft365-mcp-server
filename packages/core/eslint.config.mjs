import baseConfig from "ts-builds/eslint-functype"

export default [
  ...baseConfig,
  {
    // Calibration — @sapientsai/ms-graph-core is the IO kernel that MANUFACTURES the Either abstraction
    // its consumers use: it does the actual fetch/try-catch (with AbortController + timeout + finally,
    // which Try doesn't model cleanly), sequential nextLink pagination, and a mutable TTL upload-ticket
    // store (native Map). Throw/nullable/native-Map/sequential-loop are the irreducible imperative
    // implementation here — the functional surface is the Either it returns. These nudges are scoped off
    // for this package; prefer-either/prefer-fold stay ON.
    files: ["src/**/*.ts"],
    rules: {
      "functype/prefer-option": "off",
      "functype/prefer-try": "off",
      "functype/prefer-functype-map": "off",
      "functype/no-imperative-loops": "off",
    },
  },
]
