# Frequently Asked Questions

Common questions and solutions for abap2xlsx users. Can't find what you're looking for? Check our [community support channels](/troubleshooting/common-issues).

## Installation & Setup

### How do I verify my installation is working?

After installation, run demo report `ZDEMO_EXCEL1` to create a simple Excel file. If it executes successfully, your installation is complete! 🎉

### Which installation method should I use?

| Method | Best For | Requirements |
|--------|----------|--------------|
| **abapGit Online** | Modern systems with internet access | SAP_ABA 731+, HTTPS connectivity |
| **abapGit Offline** | Restricted networks | SAP_ABA 731+, ZIP file download |
| **SAPLink** | Legacy systems only | SAP_ABA < 702 |

For detailed instructions, see our [Installation Guide](/getting-started/installation).

### Demo programs don't appear after installation

Demo programs are in a [separate repository](https://github.com/abap2xlsx/demos). Install them separately after the main library:

1. Create package: `$ABAP2XLSXDEMOS` or `ZABAP2XLSXDEMOS`
2. Clone: `https://github.com/abap2xlsx/demos`

## Version & Compatibility

### How do I check which version is installed?

**Method 1: Check class constant**

```abap
DATA(lv_version) = zcl_excel=>version.
MESSAGE |Current version: { lv_version }| TYPE 'I'.
```

**Method 2: Generate demo report**

- Run any `ZDEMO_EXCEL*` report
- Check the generated XLSX file properties
- Version appears in the description field

### What SAP versions are supported?

- **Minimum tested**: SAP_ABA 731
- **Older versions**: May work but need community testing
- **Legacy systems**: Use SAPLink for systems < SAP_ABA 702

See our [System Requirements](/getting-started/system-requirements) for detailed compatibility information.

### How often are new versions released?

We follow [Semantic Versioning](https://semver.org/) with regular releases every 1-2 months <cite>docs/contributing/publishing-a-new-release.md:3</cite>.

## Common Issues

### "Interface method are not implemented" error after import

**Solution**: Import the SAPLink nugget twice. This is a known SAPLink issue.

### Objects won't activate after SAPLink installation

**Solution**: Follow the exact activation order described in our [SAPLink Migration Guide](/migration/from-saplink).

### Demo reports don't compile - `CL_BCS_CONVERT` not available

**Solution**: Implement required SAP OSS Notes:

- [Note 1151257 - Converting document content](https://service.sap.com/sap/support/notes/1151257)
- [Note 1151258 - Error when sending Excel attachments](https://service.sap.com/sap/support/notes/1151258)

See our [SAP Notes Guide](/troubleshooting/sap-notes) for complete details.

### `SUBMIT` with `STRING` parameter causes DB036 error

**Solution**: Implement SAP OSS Note:

- [Note 1385713 - SUBMIT: Allowing parameter of type STRING](https://service.sap.com/sap/support/notes/1385713)

### HTTPS connection fails during abapGit installation

**Solutions**:

1. **Use offline method**: Download ZIP and import via abapGit offline
2. **Certificate issues**: Import certificates in transaction `STRUST`
3. **Proxy settings**: Configure your system's internet connection

See our [Installation Troubleshooting](/troubleshooting/common-issues) for more solutions.

### Package naming conflicts

**Avoid these patterns**:

- ❌ `Z_ABAP2XLSX_DEMO` (underscore patterns cause issues)
- ❌ `ZABAP2XLSX_DEMOS` (conflicts with internal naming)

**Use instead**:

- ✅ `$ABAP2XLSX` (local package)
- ✅ `ZABAP2XLSX` (transportable package)
- ✅ `$ABAP2XLSXDEMOS` (demo package)

## Performance & Usage

### How do I download XLSX files in background jobs?

Run report `ZDEMO_EXCEL25` for background download examples. See our [Examples section](/examples/batch-processing) for detailed implementations.

### Which writer should I use for large datasets?

| Writer | Best For | Performance |
|--------|----------|-------------|
| `zcl_excel_writer_2007` | Standard files (< 100k rows) | Good |
| `zcl_excel_writer_huge_file` | Large datasets (> 100k rows) | Optimized |
| `zcl_excel_writer_csv` | Simple data export | Fastest |

See our [Performance Guide](/guide/performance) for detailed optimization strategies.

### How do I optimize Excel file generation?

**Tips for better performance**:

1. Use `zcl_excel_writer_huge_file` for large datasets
2. Minimize conditional formatting rules
3. Avoid complex formulas in large ranges
4. Use table binding instead of individual cell setting

For comprehensive performance guidance, see our [Performance Optimization](/guide/performance) documentation.

### Font issues with Calibri and auto-width calculation

**Solution**: Upload Calibri font files via transaction `SM73`:

- Upload all four variants (regular, bold, italic, bold+italic)
- Use exact description "Calibri"
- Ensures accurate width calculations

## Development & Contributing

### How do I report bugs or request features?

1. **Search first**: Check [existing issues](https://github.com/abap2xlsx/abap2xlsx/issues)
2. **Bug reports**: Use GitHub Issues with system details
3. **General questions**: Use [SAP Community](https://community.sap.com/t5/forums/searchpage/tab/message?q=abap2xlsx)
4. **Real-time help**: Join [Slack #abap2xlsx](https://sapmentors.slack.com/archives/CGG0UHDMG)

### How do I contribute code changes?

See our [Contributing Guidelines](/contributing/development-setup) for:

- Development setup
- Coding standards
- Pull request process
- Review procedures

### Where can I find code examples?

**Demo Programs**: [abap2xlsx/demos](https://github.com/abap2xlsx/demos)

- 50+ working examples
- Common use cases
- Best practices

**Documentation Examples**: Check our [Examples section](/examples/basic-report) for comprehensive tutorials.

## Community Support

### Support Channels

- **💬 [SAP Community](https://community.sap.com/t5/forums/searchpage/tab/message?q=abap2xlsx)** - General questions <cite>CONTRIBUTING.md:19-21</cite>
- **💬 [Slack #abap2xlsx](https://sapmentors.slack.com/archives/CGG0UHDMG)** - Real-time chat
- **🐛 [GitHub Issues](https://github.com/abap2xlsx/abap2xlsx/issues)** - Bug reports

### Before Asking for Help

1. ✅ Search existing discussions and documentation
2. ✅ Try the demo programs to verify your setup
3. ✅ Include system details (SAP version, installation method)
4. ✅ Provide error messages and steps to reproduce

---

*This FAQ is maintained by the abap2xlsx community. Missing something? [Contribute an improvement](/contributing/development-setup)!*
