# Xero → Airwallex payment converter

Converts the monthly Xero payment export (CSV) into an Airwallex **batch transfer**
workbook (XLSX) that can be uploaded directly.

The official Airwallex template is embedded in the binary, so the tool is
self-contained — it preserves Airwallex's hidden validation and version sheets
that the importer relies on.

## Install

```sh
go install github.com/ivanvanderbyl/xero2airwallex@latest
```

This puts `xero2airwallex` in `$(go env GOBIN)` (or `$(go env GOPATH)/bin`); add
that to your `PATH`. The Airwallex template is embedded, so the installed binary
needs nothing else on disk.

## Build from source

```sh
go build -o bin/xero2airwallex .
```

The binary lands in `bin/` (git-ignored).

## Monthly use

```sh
xero2airwallex PaymentExport_CompanyLimited_Monthly_XXXXXX_DD-Mon-YYYY.csv
```

Writes `<input>_airwallex.xlsx` next to the input. Then upload that file in the
Airwallex batch-transfer screen.

### Flags

| Flag               | Default          | Purpose                                              |
| ------------------ | ---------------- | ---------------------------------------------------- |
| `-o`               | `<input>_airwallex.xlsx` | Output path                                  |
| `-fee`             | `Payer`          | Who pays the transfer fee (`Payer` / `Recipient`)    |
| `-purpose`         | `Wages / salary` | Airwallex transfer purpose for every row             |
| `-currency`        | `NZD`            | Currency paid and received                           |
| `-country`         | `New Zealand`    | Destination (`Transfer to`)                          |
| `-recipient-type`  | `Individual`     | `Individual` / `Business`                            |
| `-reference`       | _(per-row `Salary <pay period>`)_ | Fixed reference for every row; otherwise `Salary <pay period>` with a short month (e.g. `Salary Jan 2026`) derived from the payment date |

## Field mapping

The Xero export has **no header row**. Columns used:

```
amount, bank_account, payee_name, date, ...(remaining columns ignored)
```

| Airwallex column                    | Source                                            |
| ----------------------------------- | ------------------------------------------------- |
| Transfer to                         | `New Zealand` (`-country`)                         |
| Transfer method                     | `Bank Transfer`                                    |
| Currency recipient gets / you pay   | `NZD` (`-currency`)                                |
| Transfer amount in currency you pay | Xero column 1                                      |
| Fee paid by                         | `Payer` (`-fee`)                                   |
| Account name                        | Xero column 3                                      |
| Account number                      | Xero column 2 (spaces/hyphens stripped, kept as text so leading zeros survive) |
| Transfer purpose                    | `Wages / salary` (`-purpose`)                      |
| Reference                           | `Salary <pay period>` (short month) from Xero column 4, e.g. `Salary May 2026` (`-reference` to override) |
| Description / Request ID            | left blank                                         |
| Recipient type                      | `Individual` (`-recipient-type`)                   |

## Updating the embedded template

If Airwallex changes the template, replace `template.xlsx`
(and `internal/convert/testdata/template.xlsx`) and rebuild.

## Test

```sh
go test ./...
```
