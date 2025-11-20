const html = require("eslint-plugin-html");

module.exports = [
  {
    files: ["**/*.gs", "**/*.js"],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "module",
      parserOptions: {
        ecmaFeatures: {
          impliedStrict: true
        }
      },
      globals: {
        // Google Apps Script server-side globals
        HtmlService: "readonly",
        SpreadsheetApp: "readonly",
        DriveApp: "readonly",
        Logger: "readonly",
        Utilities: "readonly",
        Session: "readonly",
        PropertiesService: "readonly",
        ScriptApp: "readonly",
        ContentService: "readonly",
        UrlFetchApp: "readonly",
        // Client-side globals
        google: "readonly",
        document: "readonly",
        window: "readonly",
        console: "readonly"
      }
    },
    rules: {
      "max-len": ["error", {"code": 120, "ignoreUrls": true, "ignoreStrings": true, "ignoreTemplateLiterals": true}],
      "no-unused-vars": ["error", {"argsIgnorePattern": "^_"}],
      "quotes": ["error", "double"],
      "indent": ["error", 2],
      "comma-dangle": ["error", "never"],
      "semi": ["error", "always"],
      "no-trailing-spaces": "error",
      "eol-last": ["error", "always"],
      "no-multiple-empty-lines": ["error", {"max": 2}]
    }
  },
  {
    files: ["**/*.html"],
    plugins: {
      html: html
    },
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "script",
      globals: {
        // Google Apps Script client-side globals
        google: "readonly",
        document: "readonly",
        window: "readonly",
        console: "readonly",
        // Additional browser globals
        alert: "readonly",
        FileReader: "readonly",
        FormData: "readonly",
        Blob: "readonly"
      }
    },
    rules: {
      "max-len": ["error", {"code": 120, "ignoreUrls": true, "ignoreStrings": true, "ignoreTemplateLiterals": true}],
      "no-unused-vars": ["error", {"argsIgnorePattern": "^_", "varsIgnorePattern": "^_"}],
      "quotes": ["error", "double"],
      "indent": ["error", 2],
      "comma-dangle": ["error", "never"],
      "semi": ["error", "always"],
      "no-trailing-spaces": "error",
      "eol-last": ["error", "always"],
      "no-multiple-empty-lines": ["error", {"max": 2}]
    }
  }
];
