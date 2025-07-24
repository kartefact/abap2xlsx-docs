# Migrating from SAPLink Installation

Guide for migrating existing abap2xlsx installations from SAPLink to abapGit.

## Overview

SAPLink installation is obsolete and should only be used on systems lower than SAP_ABA 702 <cite>docs/SAPLink-installation.md:1-2</cite>. This guide helps you migrate to the modern abapGit installation method for better maintainability and updates.

## Migration Strategy

### Assessment Phase

Before migrating, assess your current installation:

```abap
" Check current version
DATA(lv_version) = zcl_excel=>version.
MESSAGE |Current version: { lv_version }| TYPE 'I'.

" Identify custom modifications
" Review any custom changes to abap2xlsx classes
" Document dependencies on specific versions
```

### Migration Options

#### Option 1: Clean Installation (Recommended)

1. **Backup Current System**
   - Export any custom reports using abap2xlsx
   - Document current configuration and customizations

2. **Remove SAPLink Installation**
   - Delete existing abap2xlsx objects from `$TMP` package
   - Clean up any transport requests if objects were moved

3. **Install via abapGit**
   - Follow standard [abapGit installation process](/getting-started/installation.md)
   - Use package `ZABAP2XLSX` for production systems

#### Option 2: Side-by-Side Migration

1. **Install abapGit Version**
   - Install in separate package (e.g., `ZABAP2XLSX_NEW`)
   - Test functionality with existing reports

2. **Gradual Migration**
   - Update reports one by one to use new package
   - Validate functionality after each migration

3. **Cleanup Legacy Installation**
   - Remove old SAPLink objects after successful migration

## Key Differences

### Package Structure

| Aspect | SAPLink | abapGit |
|--------|---------|---------|
| **Default Package** | `$TMP` | `$abap2xlsx` or `ZABAP2XLSX` |
| **Object Status** | Inactive after import | Active after pull |
| **Updates** | Manual nugget import | Git pull operation |
| **Version Control** | None | Full Git history |

### Activation Sequence

SAPLink required specific activation order <cite>docs/SAPLink-installation.md:18-28</cite>, while abapGit handles dependencies automatically.

## Migration Checklist

### Pre-Migration
- [ ] Document current abap2xlsx version
- [ ] Identify custom modifications
- [ ] Backup existing reports and programs
- [ ] Test critical functionality

### During Migration
- [ ] Install abapGit if not present
- [ ] Create new package for abap2xlsx
- [ ] Clone repository from GitHub
- [ ] Verify all objects are active
- [ ] Run demo programs to validate installation

### Post-Migration
- [ ] Update existing reports to use new package
- [ ] Test all custom functionality
- [ ] Remove old SAPLink objects
- [ ] Update documentation and procedures

## Common Migration Issues

### Object Conflicts
```abap
" Handle naming conflicts between old and new installations
" Ensure no duplicate class definitions exist
```

### Custom Modifications
- Document any custom changes to core classes
- Consider creating separate enhancement classes instead of modifying core objects
- Use abapGit's ignore functionality for local customizations

### Transport Dependencies
- Update transport requests to reference new package
- Ensure downstream systems receive updated objects

## Validation Steps

After migration, validate your installation:

```abap
" Test basic functionality
DATA: lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet.

CREATE OBJECT lo_excel.
lo_worksheet = lo_excel->add_new_worksheet( ).
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Migration Test' ).

" Generate test file
DATA: lo_writer TYPE REF TO zcl_excel_writer_2007.
CREATE OBJECT lo_writer.
DATA(lv_file) = lo_writer->write_file( lo_excel ).
```

## Benefits of Migration

### Immediate Benefits
- **Automatic Updates**: Easy updates via Git pull
- **Version Control**: Full change history and rollback capability
- **Better Support**: Active community support on GitHub

### Long-term Benefits
- **Modern Tooling**: Integration with modern development workflows
- **Collaboration**: Easy sharing and contribution of improvements
- **Maintenance**: Simplified maintenance and troubleshooting

## Support and Resources

- **Installation Guide**: [Getting Started](/getting-started/installation.md)
- **GitHub Repository**: https://github.com/abap2xlsx/abap2xlsx
- **Community Support**: [SAP Community](https://community.sap.com/t5/forums/searchpage/tab/message?q=abap2xlsx)

## Troubleshooting

### Common Issues
- **Inactive Objects**: Use SE80 to activate remaining objects
- **Missing Dependencies**: Ensure all required SAP notes are implemented <cite>docs/FAQ.md:13-23</cite>
- **Version Conflicts**: Remove old objects completely before installing new version

For system-specific issues on older SAP versions, see [SAP 620 System Guide](docs/Getting-ABAP2XLSX-to-work-on-a-620-System.md).
