# Documentation Contribution Guide

This guide helps contributors understand how to contribute to the abap2xlsx documentation.

## Documentation Structure

Our documentation is built with VitePress and follows these conventions:

### File Organization

- `/getting-started/` - New user onboarding
- `/guide/` - Core functionality guides  
- `/advanced/` - Advanced features and techniques
- `/api/` - Class and method reference
- `/examples/` - Practical code examples
- `/troubleshooting/` - Common issues and solutions

### Writing Guidelines

- Use clear, concise language
- Include practical ABAP code examples
- Reference actual class names like `ZCL_EXCEL` and `ZCL_EXCEL_WORKSHEET`
- Test all code examples before publishing

## Local Development

```bash
# Install dependencies
npm install

# Start dev server
npm run docs:dev

# Build for production
npm run docs:build
```

## Contributing Process

1. Fork the repository
2. Create a feature branch for your documentation changes
3. Write or update markdown files
4. Test locally with VitePress
5. Submit a pull request with clear description of changes

## Style Guide

- Use sentence case for headings
- Include code examples for all concepts
- Link to related documentation sections
- Keep paragraphs focused and concise