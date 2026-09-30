import tseslint from "typescript-eslint";
import reactHooks from "eslint-plugin-react-hooks";
import reactRefresh from "eslint-plugin-react-refresh";

export default tseslint.config(
  {
    ignores: [
      "**/dist/**",
      "**/node_modules/**",
      "**/coverage/**",
      "apps/docs/.vitepress/cache/**",
      // Node scripts/config without tsconfig coverage (linted by their own
      // package scopes / run as tests, not by the root type-aware config).
      "apps/docs/eslint.config.mjs",
      "apps/docs/scripts/**",
      "eslint.config.mjs",
      "packages/excel-exporter/scripts/**",
      "packages/excel-exporter/tsup.config.ts",
      "packages/excel-exporter/vitest.config.ts",
      "packages/excel-preview/scripts/**",
      "packages/progress-overlay/tsup.config.ts",
      "packages/progress-overlay/vitest.config.ts",
      "packages/excel-preview/tsup.config.ts",
      "packages/excel-preview/vitest.config.ts",
      "packages/xlsx-core/scripts/**",
      "packages/xlsx-core/tsup.config.ts",
      "packages/xlsx-core/vitest.config.ts",
      "scripts/**",
    ],
  },
  ...tseslint.configs.recommendedTypeChecked,
  {
    languageOptions: {
      parserOptions: {
        projectService: true,
        tsconfigRootDir: import.meta.dirname,
      },
    },
  },
  {
    rules: {
      "@typescript-eslint/no-unused-vars": [
        "error",
        {
          argsIgnorePattern: "^_",
          varsIgnorePattern: "^_",
          destructuredArrayIgnorePattern: "^_",
        },
      ],
      "@typescript-eslint/consistent-type-imports": "error",
      "@typescript-eslint/no-floating-promises": "error",
    },
  },
  {
    files: ["**/*.{ts,tsx}"],
    ...reactHooks.configs.flat["recommended-latest"],
  },
  {
    files: ["**/*.{ts,tsx}"],
    ...reactRefresh.configs.vite,
  },
  {
    // Test files: relax typed rules that are noisy in test context
    files: ["**/__tests__/**/*.ts"],
    rules: {
      "@typescript-eslint/no-non-null-assertion": "off",
      "@typescript-eslint/no-explicit-any": "off",
      "@typescript-eslint/no-unsafe-call": "off",
      "@typescript-eslint/no-unsafe-member-access": "off",
      "@typescript-eslint/no-unsafe-assignment": "off",
      "@typescript-eslint/no-unsafe-return": "off",
      "@typescript-eslint/no-base-to-string": "off",
    },
  },
);
