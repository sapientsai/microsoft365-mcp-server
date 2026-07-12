import baseConfig from "ts-builds/eslint-functype"

export default [
  ...baseConfig,
  {
    // Calibration — packages/microsoft365 is the adapter / imperative shell over the functional core
    // (@sapientsai/ms-graph-core): it does IO, external-SDK glue (Azure Identity, jsonwebtoken), MCP
    // framework wiring, and native Set/Map interop (filterTools returns a native Set consumed via
    // .has()). Nullable / throw / native-collection is the honest contract at this layer, so these
    // functype nudges are scoped off for this package. The Either discipline (prefer-either /
    // prefer-fold) stays ON — that's the FP invariant that actually matters here.
    files: ["src/**/*.ts"],
    rules: {
      "functype/prefer-option": "off",
      "functype/prefer-try": "off",
      "functype/prefer-functype-set": "off",
      "functype/prefer-functype-map": "off",
    },
  },
  {
    // planner-tools exceptions:
    // - prefer-fold fires on plain `T | undefined` ternaries that are NOT Options — a rule false
    //   positive (fixed upstream in eslint-plugin-functype; remove this once that version lands).
    // - prefer-map fires on the buildPatch accumulator, where a forEach that both collects "skipped"
    //   keys and builds the checklist/reference maps reads clearer than a chained map/reduce.
    files: ["src/tools/planner-tools.ts"],
    rules: {
      "functype/prefer-fold": "off",
      "functype/prefer-map": "off",
    },
  },
]
