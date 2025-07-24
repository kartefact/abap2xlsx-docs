import { defineConfig } from 'vitepress'

export default defineConfig({
  title: 'abap2xlsx Documentation',
  description: 'Comprehensive documentation for the abap2xlsx library',
  srcDir: './',
  themeConfig: {
    nav: [
      { text: 'Home', link: '/' },
      { text: 'Getting Started', link: '/getting-started/installation' },
      { text: 'Guide', link: '/guide/basic-usage' },
      { text: 'API Reference', link: '/api/zcl-excel' },
      { text: 'FAQ', link: '/faq' },
      { text: 'Contributing', link: '/contributing/development-setup' }
    ],
    sidebar: {
      '/': [
        {
          text: 'Introduction',
          items: [
            { text: 'FAQ', link: '/faq' }
          ]
        }
      ],
      '/getting-started/': [
        {
          text: 'Getting Started',
          items: [
            { text: 'Installation', link: '/getting-started/installation' },
            { text: 'Quick Start', link: '/getting-started/quick-start' },
            { text: 'System Requirements', link: '/getting-started/system-requirements' },
            { text: 'Getting Started for Beginners', link: '/getting-started/getting-started-for-beginners' }
          ]
        }
      ],
      '/guide/': [
        {
          text: 'User Guide',
          items: [
            { text: 'Basic Usage', link: '/guide/basic-usage' },
            { text: 'Reading Excel Files', link: '/guide/reading-excel' },
            { text: 'Working with Worksheets', link: '/guide/worksheets' },
            { text: 'Formatting', link: '/guide/formatting' },
            { text: 'Formulas', link: '/guide/formulas' },
            { text: 'Charts', link: '/guide/charts' },
            { text: 'Images', link: '/guide/images' },
            { text: 'Data Conversion', link: '/guide/data-conversion' },
            { text: 'ALV Integration', link: '/guide/alv-integration' },
            { text: 'Performance', link: '/guide/performance' }
          ]
        }
      ],
      '/advanced/': [
        {
          text: 'Advanced Topics',
          items: [
            { text: 'Custom Styles', link: '/advanced/custom-styles' },
            { text: 'Conditional Formatting', link: '/advanced/conditional-formatting' },
            { text: 'Pivot Tables', link: '/advanced/pivot-tables' },
            { text: 'Data Validation', link: '/advanced/data-validation' },
            { text: 'Password Protection', link: '/advanced/password-protection' },
            { text: 'Templates', link: '/advanced/templates' },
            { text: 'Automation', link: '/advanced/automation' },
            { text: 'Macros', link: '/advanced/macros' }
          ]
        }
      ],
      '/api/': [
        {
          text: 'API Reference',
          items: [
            { text: 'ZCL Excel', link: '/api/zcl-excel' },
            { text: 'ZCL Excel Worksheet', link: '/api/zcl-excel-worksheet' },
            { text: 'ZCL Excel Writer', link: '/api/zcl-excel-writer' },
            { text: 'ZCL Excel Reader', link: '/api/zcl-excel-reader' },
            { text: 'ZCL Excel Style', link: '/api/zcl-excel-style' },
            { text: 'Error Handling', link: '/api/error-handling' }
          ]
        }
      ],
      '/examples/': [
        {
          text: 'Examples',
          items: [
            { text: 'Basic Report', link: '/examples/basic-report' },
            { text: 'Financial Report', link: '/examples/financial-report' },
            { text: 'Dashboard', link: '/examples/dashboard' },
            { text: 'Batch Processing', link: '/examples/batch-processing' },
            { text: 'Integration Patterns', link: '/examples/integration-patterns' }
          ]
        }
      ],
      '/migration/': [
        {
          text: 'Migration',
          items: [
            { text: 'From SAPLink', link: '/migration/from-saplink' },
            { text: 'Version History', link: '/migration/version-history' },
            { text: 'Breaking Changes', link: '/migration/breaking-changes' }
          ]
        }
      ],
      '/troubleshooting/': [
        {
          text: 'Troubleshooting',
          items: [
            { text: 'Common Issues', link: '/troubleshooting/common-issues' },
            { text: 'Performance Issues', link: '/troubleshooting/performance-issues' },
            { text: 'Debugging', link: '/troubleshooting/debugging' },
            { text: 'SAP Notes', link: '/troubleshooting/sap-notes' }
          ]
        }
      ],
      '/contributing/': [
        {
          text: 'Contributing',
          items: [
            { text: 'Development Setup', link: '/contributing/development-setup' },
            { text: 'Testing', link: '/contributing/testing' },
            { text: 'Documentation', link: '/contributing/documentation' },
            { text: 'Coding Guidelines', link: '/contributing/coding-guidelines' },
            { text: 'Publishing a New Release', link: '/contributing/publishing-a-new-release' }
          ]
        }
      ],
      '/legacy-docs/': [
        {
          text: 'Legacy Documentation',
          items: [
            { text: 'ABAP2XLSX Calendar Gallery', link: '/legacy-docs/abap2xlsx-Calender-Gallery' },
            { text: 'abapGit Installation', link: '/legacy-docs/abapGit-installation' },
            { text: 'Getting ABAP2XLSX to Work on a 620 System', link: '/legacy-docs/Getting-ABAP2XLSX-to-work-on-a-620-System' },
            { text: 'SAPLink Installation', link: '/legacy-docs/SAPLink-installation' }
          ]
        }
      ]
    },
    socialLinks: [
      { icon: 'github', link: 'https://github.com/abap2xlsx/abap2xlsx' }
    ],
    footer: {
      message: 'Released under the Apache 2.0 License.',
      copyright: 'Copyright © 2025 abap2xlsx contributors'
    }
  }
})