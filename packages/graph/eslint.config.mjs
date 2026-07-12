import baseConfig from "ts-builds/eslint-functype"

export default [
  ...baseConfig,
  {
    // Calibration — packages/graph (microsoft-mcp-server) is the app-only adapter shell over the same
    // functional core (@sapientsai/ms-graph-core), parallel to packages/microsoft365. IO / nullable /
    // native-collection interop is the honest contract at this layer.
    files: ["src/**/*.ts"],
    rules: {
      "functype/prefer-option": "off",
      "functype/prefer-try": "off",
      "functype/prefer-functype-set": "off",
      "functype/prefer-functype-map": "off",
    },
  },
  {
    // The MCP tool execute()/route handlers throw at the somamcp/FastMCP framework boundary (its error
    // contract); internally they use core's Either and fold-to-throw at the edge (see graph-passthrough's
    // `unwrap`). So prefer-either is off for these boundary handlers, not the internal logic. prefer-fold
    // here only fires on raw nullable checks (if-guards, conditional spreads, `v == null ? "" : String(v)`)
    // that aren't functype Options — heuristic noise, not real folds.
    files: ["src/tools/**/*.ts", "src/extract/**/*.ts"],
    rules: {
      "functype/prefer-either": "off",
      "functype/prefer-fold": "off",
    },
  },
]
