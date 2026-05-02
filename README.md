# Invoice Generator

A cross-platform .NET 8 console app that generates invoices by filling placeholders in a `.docx` template and converting the result to PDF via LibreOffice.

## Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/download)
- [LibreOffice](https://www.libreoffice.org/) (for PDF conversion)

```bash
brew install --cask libreoffice
```

## Usage

### 1. Configure `appsettings.local.json`

Copy the template and fill in your personal details (only needed once):

```bash
cp appsettings.local.template.json appsettings.local.json
```

Then edit `appsettings.local.json` with your name, bank details, and customer address.

**The only field you need to update each month is `Amount`** (if your rate changes). Everything else is auto-computed:

| Field | How it's computed |
|---|---|
| `Date` | 10th of the invoice month |
| `DueDate` | 25th of the invoice month |
| `Month` | Full month name |
| `InvoiceNo` | Anchor (122 = Feb 2025) + months elapsed |
| Output filename | `{Prefix} YYYYMM` from `OutputFilePrefix` in local config |

### 2. Generate invoices

**Interactive (recommended):**
```bash
./generate.sh
```
The script will ask whether to generate the current month or let you enter one or more specific months.

**Direct CLI:**
```bash
# Current month
dotnet run

# One specific month
dotnet run -- 202605

# Multiple months at once
dotnet run -- 202603 202604 202605
```

Generated `.docx` files are saved to the `output/` folder.

### 3. Convert to PDF

**Interactive:**
```bash
./convert-to-pdf.sh
```
Lists all `.docx` files in `output/` and lets you convert all or pick specific ones.

> Requires LibreOffice: `brew install --cask libreoffice`

## Invoice Number History

| Period | Invoice No |
|---|---|
| Feb 2025 | 122 |
| Mar 2025 | 123 |
| Apr 2025 | 124 |
| May 2025 | 125 |
| Nov 2025 | 131 |
| Dec 2025 | 132 |
| Jan 2026 | 133 |
| Feb 2026 | 134 |
| Mar 2026 | 135 |
| Apr 2026 | 136 |
| **May 2026** | **137** |
| Jun 2026 | 138 |
| Jul 2026 | 139 |

## Template Placeholders

The following placeholders are defined in the `.docx` template. All are configured in `appsettings.json` under `Invoice.Placeholders`:

| Placeholder | Description |
|---|---|
| `{{Date}}` | Invoice date |
| `{{InvoiceNo}}` | Invoice number |
| `{{Amount}}` | Amount in SEK |
| `{{DueDate}}` | Payment due date |
| `{{Month}}` | Month name for the description line |
| `{{FullName}}` | Account holder full name |
| `{{Designation}}` | Job title / role |
| `{{Address}}` | Street address |
| `{{City}}` | City |
| `{{Province}}` | Province / region |
| `{{ZipCode}}` | Postal / ZIP code |
| `{{PhoneNumber}}` | Contact phone number |
| `{{EmailAddress}}` | Contact email address |
| `{{AccountNumber}}` | Bank account number |
| `{{SwiftCode}}` | Bank SWIFT/BIC code |
| `{{CustomerAddressLine01}}` | Customer address line 1 |
| `{{CustomerAddressLine02}}` | Customer address line 2 |
| `{{CustomerAddressLine03}}` | Customer address line 3 |
| `{{CustomerAddressLine04}}` | Customer address line 4 |

## Project Structure

```
invoice-generator/
├── Program.cs                       # Main application logic
├── InvoiceGenerator.csproj          # Project file
├── appsettings.json                 # Generic config (committed)
├── appsettings.local.json           # Personal config — gitignored, create from template
├── appsettings.local.template.json  # Blank personal config template (committed)
├── InvoiceTemplate.docx             # Word template with placeholders
├── generate.sh                      # Interactive invoice generation script
├── convert-to-pdf.sh                # Interactive batch PDF conversion script
├── output/                          # Generated invoices (gitignored)
└── old-examples/                    # Reference PDFs from previous months
```
