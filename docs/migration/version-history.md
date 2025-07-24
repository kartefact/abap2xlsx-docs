# Version History and Compatibility

Complete version history and compatibility information for abap2xlsx.

## Current Version

**Latest Release: 7.16.0**

Version tracking was introduced in version 7.0.1. You can check your installed version by examining the VERSION attribute in class `ZCL_EXCEL` or by generating a demo report and checking the Excel file properties.

## Version Compatibility Matrix

| abap2xlsx Version | SAP NetWeaver | Installation Method | Notes |
|-------------------|---------------|-------------------|-------|
| 7.16.0 | 7.31+ | abapGit | Current release |
| 7.15.0 | 7.31+ | abapGit | Previous release |
| 7.0.1+ | 7.02+ | abapGit/SAPLink | Version tracking introduced |
| < 7.0.1 | 6.20+ | SAPLink | Legacy versions |

## Release History

### 7.16.0 (Latest)

- Enhanced performance for large datasets
- Improved chart generation capabilities
- Bug fixes and stability improvements

### 7.15.0

- Added conditional formatting support
- Enhanced template filling functionality
- Performance optimizations

### 7.0.1

- **Important**: First version with version tracking
- Introduced VERSION attribute in ZCL_EXCEL
- Improved error handling

## Installation Methods by Version

### Modern Approach (7.02+)

Use abapGit for the best development experience:

- Full Git integration
- Easy updates and version management
- Better conflict resolution

### Legacy Approach (< 7.02)

SAPLink nugget files are provided for older systems:

- Download from `/build/` folder
- Import using ZSAPLINK transaction
- Manual version management required

## Upgrade Considerations

### From Legacy Versions (< 7.0.1)

- No automatic version detection available
- Manual verification of functionality required
- Consider full regression testing

### Between Modern Versions (7.0.1+)

- Version information preserved in Excel files
- Backward compatibility maintained
- Incremental testing sufficient

## System Requirements

### Minimum Requirements

- SAP NetWeaver 7.31 (recommended)
- May work on older versions with limitations

### Required SAP Notes

For optimal functionality, implement these SAP notes:

- Note 1151257: Converting document content
- Note 1151258: Error when sending Excel attachments  
- Note 1385713: SUBMIT parameter of type STRING

## Checking Compatibility

### Version Detection

```abap
" Check installed version
DATA(lv_version) = zcl_excel=>version.
IF lv_version IS INITIAL.
  WRITE: / 'Version < 7.0.1 (no version tracking)'.
ELSE.
  WRITE: / 'Version:', lv_version.
ENDIF.
```

### Functionality Testing

Run `ZDEMO_EXCEL_CHECKER` to verify all features work correctly in your environment.
