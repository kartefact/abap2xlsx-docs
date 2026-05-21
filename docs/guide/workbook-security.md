# Workbook Security and Encryption

abap2xlsx supports two levels of protection:

1. **Worksheet protection** — locks cells/objects on a specific sheet (see
   [Worksheets](/guide/worksheets#worksheet-protection)).
2. **Workbook encryption** — password-protects the entire `.xlsx` file using AES-256,
   requiring a password to open it in Excel.

This page covers workbook-level encryption via `zcl_excel_security`.

## How It Works

The OOXML encryption standard wraps the entire `.xlsx` ZIP package inside an OLE Compound
Document (`CFBF`) container and encrypts it with AES-256-CBC using a key derived from the
password via PBKDF2-SHA512. This is the same mechanism Excel uses for "Encrypt with
Password" (File → Info → Protect Workbook → Encrypt with Password).

## Encrypting a Workbook

```abap
DATA: lo_writer   TYPE REF TO zcl_excel_writer_2007,
      lo_security TYPE REF TO zcl_excel_security.

" 1. Generate the xlsx binary first
CREATE OBJECT lo_writer.
DATA(lv_xlsx) = lo_writer->write_file( lo_excel ).

" 2. Encrypt it with a password
CREATE OBJECT lo_security.
DATA(lv_encrypted) = lo_security->encrypt(
  iv_xstring  = lv_xlsx
  iv_password = 'MyS3cur3P@ssword'
).

" 3. lv_encrypted is now a CFBF-wrapped, AES-256-encrypted blob
"    Save or send lv_encrypted — not lv_xlsx
```

The `encrypt` method returns an `xstring`. The result is an OLE Compound Document binary
(recognisable by the magic bytes `D0 CF 11 E0`) that Excel and LibreOffice can open.

## Reading an Encrypted Workbook

To **read** an encrypted workbook you must decrypt it first:

```abap
DATA: lo_security  TYPE REF TO zcl_excel_security,
      lo_reader    TYPE REF TO zcl_excel_reader_2007.

" 1. Decrypt — provide the same password used during encryption
CREATE OBJECT lo_security.
DATA(lv_plain_xlsx) = lo_security->decrypt(
  iv_xstring  = lv_encrypted_blob
  iv_password = 'MyS3cur3P@ssword'
).

" 2. Read the plain xlsx as normal
CREATE OBJECT lo_reader.
DATA(lo_excel) = lo_reader->load( lv_plain_xlsx ).
```

## Password Requirements

- Passwords are passed as `TYPE string` (Unicode).
- The `zexcel_aes_password` data element defines the domain — its length is 255 characters
  maximum.
- There is no minimum length enforced by the API, but Excel recommends at least 8 characters
  for meaningful security.
- The same password is used for both encryption and decryption; abap2xlsx does not store
  passwords anywhere.

## Limitations

- `zcl_excel_security` uses SAP's `CL_SEC_SXML_WRITER` and AES-related kernel functions;
  these are available on SAP Basis 7.50+.
- The class is in the main `src/` package and is **cloud-compatible**.
- Worksheet-level protection (cell locking, sheet password) is separate from workbook
  encryption — you can apply both independently.
- abap2xlsx cannot currently **modify** an encrypted workbook in-place; the workflow is
  always: decrypt → modify → re-encrypt.

## Next Steps

- **[Worksheets](/guide/worksheets#worksheet-protection)** — per-sheet cell locking
- **[Cloud Compatibility](/guide/cloud-compatibility)** — kernel requirements
