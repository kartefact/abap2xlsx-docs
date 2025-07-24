# Development Environment Setup

Guide for setting up a development environment for contributing to abap2xlsx.

## Prerequisites

### SAP System Requirements

- SAP NetWeaver 7.31 or higher
- ABAP Development Tools (ADT) recommended
- SAPLink (for systems below 7.02)

### Development Tools

- SAP GUI or SAP Business Client
- ABAP Development Tools (Eclipse-based)
- Git client for version control

## Installation Methods

### Modern Systems (7.02+)

Use abapGit for the most streamlined development experience:

1. Install abapGit in your system
2. Clone the repository: `https://github.com/abap2xlsx/abap2xlsx.git`
3. Pull the latest changes

### Legacy Systems (< 7.02)

Use SAPLink for older systems:

1. Download the latest nugget file from the build folder
2. Import using transaction ZSAPLINK
3. Check "overwrite originals" if updating existing installation

## Development Workflow

### Setting Up Your Fork

```bash
# Fork the repository on GitHub
# Clone your fork locally
git clone https://github.com/yourusername/abap2xlsx.git
cd abap2xlsx

# Add upstream remote
git remote add upstream https://github.com/abap2xlsx/abap2xlsx.git
```

### Creating Feature Branches

```bash
# Create and switch to feature branch
git checkout -b feature/your-feature-name

# Make your changes in SAP system
# Export changes using abapGit or SAPLink

# Commit changes
git add .
git commit -m "Add: your feature description"

# Push to your fork
git push origin feature/your-feature-name
```

## Code Organization

The library follows this structure:

- `/src/` - Main library classes
- `/docs/` - Documentation files
- `/build/` - Build artifacts and nugget files

## Testing Your Changes

Before submitting changes:

1. Run `ZDEMO_EXCEL_CHECKER` to verify all tests pass
2. Test your specific functionality with relevant demo programs
3. Ensure backward compatibility

## Submitting Changes

1. Create a pull request from your feature branch
2. Provide clear description of changes
3. Reference any related issues
4. Wait for review and address feedback
