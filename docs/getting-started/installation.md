# Installation Guide

Comprehensive guide for installing abap2xlsx in your SAP system.

## Installation Methods

### Modern Systems (SAP NetWeaver 7.02+)

#### Using abapGit (Recommended)

abapGit provides the most streamlined installation and update experience:

1. **Install abapGit** in your system if not already available
2. **Clone the repository**:
   - Repository URL: `https://github.com/abap2xlsx/abap2xlsx.git`
   - Package: `$ABAP2XLSX` (or your preferred package)
3. **Pull the latest version** using abapGit
4. **Activate all objects** in the package

```abap
" Verify installation by checking version
DATA(lv_version) = zcl_excel=>version.
WRITE: / 'abap2xlsx version:', lv_version.
```

#### Manual Installation via Download

If abapGit is not available:

1. Download the latest release from GitHub
2. Import objects manually using SE80 or ADT
3. Ensure all dependencies are resolved

### Legacy Systems (SAP NetWeaver < 7.02)

#### Using SAPLink

For older systems, use the SAPLink nugget files:

1. **Prerequisites**:
   - [SAPLink](http://www.saplink.org) installed in your system
   - SAPLink Plugins (DDic, Interface) installed

2. **Download nugget file**:
   - Get the latest `.nugg` file from the [build folder](https://github.com/abap2xlsx/abap2xlsx/tree/master/build)
   - Save locally on your system

3. **Import process**:
   - Execute report `ZSAPLINK`
   - Select "Import Nugget"
   - Locate your `.nugg` file
   - Check "overwrite originals" if updating existing installation
   - **Important**: If you get "Interface methods are not implemented" error, import the nugget twice (known SAPLink issue)

## Required SAP Notes

Implement these SAP notes for optimal functionality:

### Essential Notes

- **Note 1151257**: Converting document content
- **Note 1151258**: Error when sending Excel attachments  
- **Note 1385713**: SUBMIT parameter of type STRING

### Verification

After implementing notes, test with demo programs:

```abap
" Test basic functionality
REPORT ztest_abap2xlsx_install.

DATA: lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_writer TYPE REF TO zif_excel_writer.

START-OF-SELECTION.
  TRY.
      CREATE OBJECT lo_excel.
      lo_worksheet = lo_excel->add_new_worksheet( ).
      lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Installation Test' ).
      
      CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
      DATA(lv_file) = lo_writer->write_file( lo_excel ).
      
      WRITE: / 'Installation successful - Excel file generated'.
      
    CATCH zcx_excel INTO DATA(lx_excel).
      WRITE: / 'Installation error:', lx_excel->get_text( ).
  ENDTRY.
```

## Post-Installation Verification

### Run Demo Checker

Execute `ZDEMO_EXCEL_CHECKER` to verify all components:

1. All tests should show green checkmarks
2. Any red indicators require attention
3. Review failed tests for missing dependencies

### Font Configuration (Optional)

For optimal Excel compatibility, upload Calibri font:

1. Go to transaction `SM73`
2. Upload Calibri TTF files (regular, bold, italic, bold-italic)
3. Use exact description "Calibri"

## System Requirements

### Minimum Requirements

- SAP NetWeaver 7.31 (recommended)
- May work on older versions with limitations
- Unicode system recommended

### Memory Considerations

- Standard installation: ~50MB
- Large file processing may require additional memory
- Consider system resources for concurrent users

## Troubleshooting Installation

### Common Issues

#### Objects Not Activating

- Check for missing dependencies
- Verify SAP notes are implemented
- Review syntax errors in newer ABAP versions

#### Demo Programs Not Compiling

- Implement required SAP notes
- Check for missing function modules
- Verify class `CL_BCS_CONVERT` availability

#### Version Detection Issues

Version tracking available from 7.0.1+:

- Check Excel file properties for version info
- Verify `VERSION` attribute in `ZCL_EXCEL` class

### Getting Help

- Check [troubleshooting guide](/troubleshooting/common-issues)
- Search [GitHub issues](https://github.com/abap2xlsx/abap2xlsx/issues)
- Post questions on [SAP Community](https://community.sap.com/) with "ABAP2XLSX" tag

## Next Steps

After successful installation:

1. Review the [Quick Start Guide](/getting-started/quick-start)
2. Check [System Requirements](/getting-started/system-requirements) for compatibility
3. Explore [Basic Usage Examples](/guide/basic-usage)
