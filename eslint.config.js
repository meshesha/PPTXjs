// @ts-check
import tseslint from "typescript-eslint";

export default tseslint.config(
  { ignores: ["dist/**", "dist-demo/**", "coverage/**", "legacy/**", "test/output/**", "test/baseline/**", "scripts/port/**"] },
  ...tseslint.configs.recommended,
  {
    rules: {
      "@typescript-eslint/no-explicit-any": "off",
      "@typescript-eslint/no-unused-vars": ["warn", { argsIgnorePattern: "^_", varsIgnorePattern: "^_" }],
      "no-control-regex": "off",
    },
  },
  {
    // Rendering code ported 1:1 from PPTXjs v1 (see src/render/*). It keeps the
    // original style (var, loose equality) so diffs against v1 stay readable.
    files: ["src/render/{shape,text,table,fill,color,border,background,nodes,geometry}.ts"],
    linterOptions: { reportUnusedDisableDirectives: "off" },
    rules: {
      "no-var": "off",
      "no-redeclare": "off",
      "no-fallthrough": "off",
      "no-unsafe-finally": "off",
      "no-empty": "off",
      "prefer-const": "off",
      "no-useless-escape": "off",
      "no-case-declarations": "off",
      "no-unreachable": "off",
      "@typescript-eslint/no-unused-vars": "off",
      "@typescript-eslint/no-unused-expressions": "off",
    },
  },
);
