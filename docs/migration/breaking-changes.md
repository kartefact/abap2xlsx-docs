# Breaking Changes

This document tracks breaking changes across abap2xlsx versions to help with migration planning.

## Version 7.x Series

The 7.x series maintains API compatibility as much as possible. Breaking changes are avoided to prevent disruption to existing implementations.

### Version 7.16.0

- No breaking changes
- All existing APIs remain functional

### Version 7.15.0

- No breaking changes
- Enhanced functionality while maintaining backward compatibility

## Migration Guidelines

### General Principles

#### abap2xlsx follows semantic versioning:

- Major version (7.x): May include breaking changes
- Minor version (x.16.x): New features, no breaking changes  
- Patch version (x.x.0): Bug fixes only

### Checking Your Version

You can verify your installed version:

```abap
DATA(lv_version) = zcl_excel=>version.
WRITE: / 'abap2xlsx version:', lv_version.
```

Or check the VERSION attribute in class `ZCL_EXCEL`.

### Compatibility Testing

Before upgrading:

1. Run `ZDEMO_EXCEL_CHECKER` with your current version
2. Note any failing tests
3. After upgrade, run the checker again
4. Compare results to identify any issues

## Future Breaking Changes

### Planned Deprecations

Currently no major breaking changes are planned for the 7.x series.

### Migration Support

If breaking changes become necessary:

- Advance notice will be provided
- Migration guides will be created
- Deprecated features will be marked clearly
- Transition period will be provided where possible

## Getting Help

If you encounter issues after upgrading:

1. Check this breaking changes document
2. Review the FAQ for common issues
3. Search existing GitHub issues
4. Create a new issue with version details and error information
