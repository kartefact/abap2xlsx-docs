---
layout: home

sidebar: true

hero:
  name: "abap2xlsx"
  text: "Excel Generation for SAP ABAP"
  tagline: "The powerful ABAP library for reading and generating Excel spreadsheets directly from SAP systems"
  image:
    src: /logo.png
    alt: abap2xlsx
  actions:
    - theme: brand
      text: Get Started
      link: /getting-started/installation
    - theme: alt
      text: View on GitHub
      link: https://github.com/abap2xlsx/abap2xlsx

features:
  - icon: 📊
    title: Complete Excel Support
    details: Generate Excel 2007+ files (.xlsx) with full support for worksheets, styling, formulas, charts, and images. Read existing files with full preservation of formatting.
  - icon: 🚀
    title: High Performance
    details: Optimized writers for large datasets (100,000+ rows), memory-efficient processing, and specialized huge file handling for enterprise workloads.
  - icon: 🎨
    title: Rich Formatting
    details: Advanced styling, conditional formatting, data validation, pivot tables, custom themes, and professional report layouts.
  - icon: 📋
    title: ALV Integration
    details: Seamless conversion from ALV grids and SALV tables to Excel with automatic formatting, table styles, and filter dropdowns.
  - icon: 🔄
    title: Template System
    details: Fill Excel templates with ABAP data, create standardized reports, and maintain consistent corporate formatting across documents.
  - icon: 🛡️
    title: Enterprise Ready
    details: Password protection, macro support (.xlsm), comprehensive error handling, and production-grade reliability for business systems.
---

## What is abap2xlsx?

abap2xlsx is a comprehensive ABAP library for creating and reading Excel spreadsheets directly from SAP systems. It provides a complete object-oriented interface for generating Excel files with advanced formatting, charts, images, and seamless business logic integration.

## 🔧 Quick Examples

### Basic Excel Creation

```abap
DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_writer    TYPE REF TO zcl_excel_writer_2007.

CREATE OBJECT lo_excel.
lo_worksheet = lo_excel->get_active_worksheet( ).
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Hello World!' ).

CREATE OBJECT lo_writer.
DATA(lv_xstring) = lo_writer->write_file( lo_excel ).
```

### Advanced Table Binding

```abap
lo_worksheet->bind_table( 
  ip_table = lt_data
  ip_table_settings = VALUE #( 
    table_style = zcl_excel_table=>builtinstyle_medium2
    show_row_stripes = abap_true
    show_filter_dropdown = abap_true
  )
).
```

### Conditional Formatting

```abap
DATA(lo_conditional) = NEW zcl_excel_style_cond( ).
lo_conditional->set_range( 'A1:A100' )
              ->set_rule_type( zcl_excel_style_cond=>c_rule_cellis )
              ->set_operator( zcl_excel_style_cond=>c_operator_greaterthan )
              ->set_formula( '1000' ).
lo_worksheet->add_style_conditional( lo_conditional ).
```

## 🎯 Core Features Overview

### Excel Generation Capabilities

- **File Formats**: .xlsx, .xlsm (macro-enabled), .csv
- **Large Files**: Optimized writers for massive datasets
- **Formatting**: Fonts, colors, borders, number formats, themes
- **Graphics**: Charts, images, drawings

### Data Integration Features

- **Table Binding**: Direct internal table to Excel conversion
- **ALV Integration**: Convert ALV grids to Excel format
- **Template Processing**: Fill Excel templates with ABAP data
- **Data Validation**: Dropdown lists, input restrictions

### Advanced Excel Features

- **Conditional Formatting**: Cell highlighting based on values
- **Autofilters**: Enable Excel filtering capabilities
- **Named Ranges**: Define and reference cell ranges
- **Hyperlinks**: Internal worksheet and external links
- **Comments**: Cell annotations and notes
- **Formulas**: Excel formula support

## 🏗️ Architecture Overview

### Core Classes

| Class | Purpose |
|-------|---------|
| `zcl_excel` | Main workbook container and entry point |
| `zcl_excel_worksheet` | Individual worksheet management |
| `zcl_excel_writer_2007` | Standard XLSX file writer |
| `zcl_excel_writer_huge_file` | Optimized writer for large datasets |
| `zcl_excel_reader_2007` | Excel file reader and parser |
| `zcl_excel_converter` | Data conversion utilities |

### Installation Components

After installation, verify these key objects exist:

- **Main Classes**: `ZCL_EXCEL`, `ZCL_EXCEL_WORKSHEET`, `ZCL_EXCEL_WRITER_2007`
- **Supporting Classes**: Style management, drawing support, converters
- **Demo Programs**: `ZDEMO_EXCEL*` reports for testing and examples

## 📋 System Requirements

| Requirement | Details |
|-------------|---------|
| **SAP System** | Minimum SAP_ABA 731 (may work on older versions) |
| **Developer Access** | Package creation and transport rights |
| **Installation Tool** | abapGit (recommended) or SAPLink (legacy) |
| **Network Access** | HTTPS connectivity for online installation |

## 🚀 Installation

Get started with abap2xlsx using abapGit:

1. **Install abapGit** in your SAP system
2. **Clone repository**: `https://github.com/abap2xlsx/abap2xlsx.git`
3. **Create package**: `$abap2xlsx` or `ZABAP2XLSX`
4. **Activate objects** and start coding!

### Installation Methods

- **[abapGit Installation](/getting-started/installation)** - Recommended modern approach (SAP_ABA 731+)
- **[SAPLink Migration](/migration/from-saplink)** - For legacy systems
- **[System Requirements](/getting-started/system-requirements)** - Check compatibility before installation

## 🆘 Support & Community

### Getting Help

- **[SAP Community](https://community.sap.com/t5/forums/searchpage/tab/message?q=abap2xlsx)** - General questions and discussions
- **[Slack Channel](https://sapmentors.slack.com/archives/CGG0UHDMG)** - Real-time chat in SAP Mentors & Friends
- **[GitHub Issues](https://github.com/abap2xlsx/abap2xlsx/issues)** - Bug reports and feature requests

### Before Asking for Help

1. **Search existing discussions** on SAP Community and GitHub Issues
2. **Check the [troubleshooting guides](/troubleshooting/common-issues)** for common solutions
3. **Try demo programs** to verify your installation works
4. **Provide system details** when reporting issues (SAP version, installation method)

## 🔍 Troubleshooting Quick Reference

| Issue | Solution |
|-------|---------|
| **Objects won't activate** | Follow exact activation order for SAPLink installations |
| **HTTPS connection fails** | Use offline ZIP installation method |
| **Package naming conflicts** | Use unique names, avoid underscore patterns |
| **Version verification** | Check `ZCL_EXCEL=>VERSION` attribute or demo file properties |

## 🤝 Contributing

We welcome contributions! Whether it's bug fixes, new features, or documentation improvements, your help makes abap2xlsx better for everyone.

- **[Contributing Guidelines](/contributing/development-setup)** - Complete contribution process
- **[Coding Standards](/contributing/documentation)** - Development guidelines and naming conventions

## 📊 Project Information

- **License**: Apache License 2.0 - see [LICENSE](https://github.com/abap2xlsx/abap2xlsx/blob/master/LICENSE) for details
- **Repository**: [github.com/abap2xlsx/abap2xlsx](https://github.com/abap2xlsx/abap2xlsx)
- **Demo Repository**: [github.com/abap2xlsx/demos](https://github.com/abap2xlsx/demos)
- **Release Cycle**: Regular updates every 1-2 months <cite>docs/contributing/publishing-a-new-release.md:3</cite>
- **Versioning**: Semantic versioning (Major.Minor.Patch)

## 📖 Additional Resources

- **[Original Blog Series](http://scn.sap.com/community/abap/blog/2010/07/12/abap2xlsx--generate-your-professional-excel-spreadsheet-from-abap)** - Introduction and tutorials
- **[GitHub Releases](https://github.com/abap2xlsx/abap2xlsx/releases)** - Download specific versions
- **[Demo Programs](https://github.com/abap2xlsx/demos)** - Example implementations and use cases

---

<div class="tip custom-block" style="padding-top: 8px">

Ready to start? Check out the [Quick Start Guide](/getting-started/quick-start) to create your first Excel file in minutes!

</div>

---

*This documentation is maintained by the abap2xlsx community. Found an issue? [Report it](https://github.com/abap2xlsx/abap2xlsx/issues) or [contribute a fix](/contributing/development-setup)!*
