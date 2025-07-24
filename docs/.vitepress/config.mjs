import { defineConfig } from 'vitepress'

export default defineConfig({
  title: 'abap2xlsx Documentation',
  description: 'Documentation for abap2xlsx, a library for generating and manipulating Excel files in ABAP.',
  base: '/docs/',
  themeConfig: {
    logo: '/docs/public/logo.png',
    editLink: {
      pattern: 'https://github.com/kartefact/abap2xlsx-docs/tree/main/docs/:path',
      text: 'Edit this page on GitHub'
    },
    search: {
      provider: 'local'
    },
    nav: [
      { text: 'Home', link: '/' },
      { text: 'Getting Started', link: '/getting-started/getting-started-for-beginners' },
      { text: 'Guide', link: '/guide/basic-usage' },
      { text: 'API', link: '/api/zcl-excel' },
      { text: 'Examples', link: '/examples/basic-report' },
      { text: 'Contributing', link: '/contributing/coding-guidelines' },
    ],
    sidebar: [
      {
        text: 'Getting Started',
        collapsed: false,
        items: [
          { text: 'Getting Started for Beginners', link: '/getting-started/getting-started-for-beginners' },
          { text: 'Installation', link: '/getting-started/installation' },
          { text: 'Quick Start', link: '/getting-started/quick-start' },
          { text: 'System Requirements', link: '/getting-started/system-requirements' },
          { text: 'FAQ', link: '/getting-started/faq' },
        ],
      },
      {
        text: 'Guide',
        collapsed: false,
        items: [
          { text: 'Basic Usage', link: '/guide/basic-usage' },
          { text: 'ALV Integration', link: '/guide/alv-integration' },
          { text: 'Charts', link: '/guide/charts' },
          { text: 'Data Conversion', link: '/guide/data-conversion' },
          { text: 'Formatting', link: '/guide/formatting' },
          { text: 'Formulas', link: '/guide/formulas' },
          { text: 'Images', link: '/guide/images' },
          { text: 'Performance', link: '/guide/performance' },
          { text: 'Reading Excel', link: '/guide/reading-excel' },
          { text: 'Worksheets', link: '/guide/worksheets' },
        ],
      },
      {
        text: 'Advanced',
        collapsed: false,
        items: [
          { text: 'Automation', link: '/advanced/automation' },
          { text: 'Conditional Formatting', link: '/advanced/conditional-formatting' },
          { text: 'Custom Styles', link: '/advanced/custom-styles' },
          { text: 'Data Validation', link: '/advanced/data-validation' },
          { text: 'Macros', link: '/advanced/macros' },
          { text: 'Password Protection', link: '/advanced/password-protection' },
          { text: 'Pivot Tables', link: '/advanced/pivot-tables' },
          { text: 'Templates', link: '/advanced/templates' },
        ],
      },
      {
        text: 'API Reference',
        collapsed: false,
        items: [
          { text: 'ZCL_EXCEL', link: '/api/zcl-excel' },
          { text: 'ZCL_EXCEL_READER', link: '/api/zcl-excel-reader' },
          { text: 'ZCL_EXCEL_STYLE', link: '/api/zcl-excel-style' },
          { text: 'ZCL_EXCEL_WORKSHEET', link: '/api/zcl-excel-worksheet' },
          { text: 'ZCL_EXCEL_WRITER', link: '/api/zcl-excel-writer' },
          { text: 'Error Handling', link: '/api/error-handling' },
        ],
      },
      {
        text: 'Examples',
        collapsed: false,
        items: [
          { text: 'Basic Report', link: '/examples/basic-report' },
          { text: 'Batch Processing', link: '/examples/batch-processing' },
          { text: 'Dashboard', link: '/examples/dashboard' },
          { text: 'Financial Report', link: '/examples/financial-report' },
          { text: 'Integration Patterns', link: '/examples/integration-patterns' },
        ],
      },
      {
        text: 'Contributing',
        collapsed: false,
        items: [
          { text: 'Coding Guidelines', link: '/contributing/coding-guidelines' },
          { text: 'Development Setup', link: '/contributing/development-setup' },
          { text: 'Documentation', link: '/contributing/documentation' },
          { text: 'Publishing a New Release', link: '/contributing/publishing-a-new-release' },
          { text: 'Testing', link: '/contributing/testing' },
        ],
      },
      {
        text: 'Troubleshooting',
        collapsed: false,
        items: [
          { text: 'Common Issues', link: '/troubleshooting/common-issues' },
          { text: 'Debugging', link: '/troubleshooting/debugging' },
          { text: 'Performance Issues', link: '/troubleshooting/performance-issues' },
          { text: 'SAP Notes', link: '/troubleshooting/sap-notes' },
        ],
      },
      {
        text: 'Migration',
        collapsed: false,
        items: [
          { text: 'Breaking Changes', link: '/migration/breaking-changes' },
          { text: 'From SAPLink', link: '/migration/from-saplink' },
          { text: 'Version History', link: '/migration/version-history' },
        ],
      },
      {
        text: 'Legacy Documentation',
        collapsed: false,
        items: [
          { text: 'abap2xlsx Calendar Gallery', link: '/legacy-docs/abap2xlsx-Calender-Gallery' },
          { text: 'abapGit Installation', link: '/legacy-docs/abapGit-installation' },
          { text: 'Getting ABAP2XLSX to Work on a 620 System', link: '/legacy-docs/Getting-ABAP2XLSX-to-work-on-a-620-System' },
          { text: 'SAPLink Installation', link: '/legacy-docs/SAPLink-installation' },
        ],
      },
    ],
    socialLinks: [
      { icon: 'github', link: 'https://github.com/kartefact/abap2xlsx-docs' },
    ],
    footer: {
      message: `
      <a href="/LICENSE">License</a> |
      <a href="/docs/resources/contact">Contact</a>`,
      copyright: `Copyright © 2010-${new Date().getFullYear()} abap2xlsx Contributors`,
    },
  },
})