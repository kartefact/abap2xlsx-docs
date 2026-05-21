# Security Policy

The abap2xlsx project takes security seriously. This page describes how to report a vulnerability and what to expect after you do.

> **Note:** `SECURITY.md` was added to the abap2xlsx source repository in February 2025 (PR [#1289](https://github.com/abap2xlsx/abap2xlsx/pull/1289)).

## Supported Versions

Security fixes are applied to the **latest version** on the `main` branch. The project follows a rolling-release model via abapGit — always stay on the latest `main`.

## Reporting a Vulnerability

**Do not open a public GitHub issue.** Use one of these private channels:

1. **GitHub Private Security Advisory** (preferred) — [Submit here](https://github.com/abap2xlsx/abap2xlsx/security/advisories/new).
2. **Direct email to a maintainer** — See [CODEOWNERS](https://github.com/abap2xlsx/abap2xlsx/blob/main/.github/CODEOWNERS) or [CONTRIBUTING.md](https://github.com/abap2xlsx/abap2xlsx/blob/main/CONTRIBUTING.md).

## What to Include

- Description of the vulnerability and potential impact
- abap2xlsx version (commit SHA or abapGit pull date)
- SAP release and ABAP kernel version
- Step-by-step reproduction instructions
- Any proof-of-concept code or test data

## Response Timeline

| Stage | Target |
|---|---|
| Initial acknowledgement | Within 5 business days |
| Triage and severity assessment | Within 10 business days |
| Fix development | Depends on severity; critical issues prioritised |
| Public disclosure | After fix is merged to `main` |

## Scope

In scope:
- Maliciously crafted `.xlsx` files causing dumps or data exposure in `zcl_excel_reader_2007`
- Formula cell injection risks (`=CMD()` style payloads)
- Path traversal or unsafe file handling in download helpers

Out of scope:
- Vulnerabilities in your own code calling abap2xlsx
- SAP kernel issues unrelated to the library
- Issues requiring attacker SAP developer access

## Related Resources

- [SECURITY.md in source repo](https://github.com/abap2xlsx/abap2xlsx/blob/main/SECURITY.md)
- [GitHub Private Security Advisories docs](https://docs.github.com/en/code-security/security-advisories/guidance-on-reporting-and-writing/privately-reporting-a-security-vulnerability)
- **[Coding Guidelines](/contributing/coding-guidelines)**
- **[Testing](/contributing/testing)**
