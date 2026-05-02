# Invoice Generator

A cross-platform .NET 8 console app that generates invoices by filling placeholders in a `.docx` template and converting the result to PDF via LibreOffice.

## Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/download)
- [LibreOffice](https://www.libreoffice.org/) (for PDF conversion)

```bash
brew install --cask libreoffice
```

## Usage

### 1. Edit `appsettings.json`

The following fields are **auto-computed from the current system date** every run — you never need to touch them:

| Field | How it's computed |
|---|---|
| `Date` | 10th of the current month |
| `DueDate` | 25th of the current month |
| `Month` | Full month name from system clock |
| `InvoiceNo` | Anchor (122 = Feb 2025) + months elapsed |
| `OutputFileName` | `Shazni.S YYYYMM` from system clock |

**The only field you need to update each month is `Amount`** (if it changes).

### 2. Run

```bash
dotnet run
```

The generated `.docx` and `.pdf` are saved to the `output/` folder.

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

## Project Structure

```
invoice-generator/
├── Program.cs                  # Main application logic
├── InvoiceGenerator.csproj     # Project file
├── appsettings.json            # Invoice values to edit each month
├── InvoiceTemplateV2 1.docx    # Word template with placeholders
├── output/                     # Generated invoices (gitignored)
└── old-examples/               # Reference PDFs from previous months
```
