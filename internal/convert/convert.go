// Package convert turns a Xero monthly payment export (CSV) into an Airwallex
// batch-transfer workbook (XLSX).
//
// The Xero export has no header row. Each record is laid out as:
//
//	amount, bank_account, payee_name, particulars, code, reference, <payer-side repeat...>
//
// e.g.
//
//	1234.56,0100010000001000,Ada Lovelace,31-May-2026,Pay Ended,Salary/Wages,31-May-2026,Pay Ended,Salary/Wages
//
// The Airwallex side is produced by writing rows into the official batch-transfer
// template (sheet "Airwallex batch transfer"), starting at row 2 which overwrites
// the template's "e.g., ..." placeholder. Writing into the template — rather than
// building a workbook from scratch — preserves the hidden Validation /
// DropdownListMapping / StorageUsedInternally (version) sheets that Airwallex's
// importer relies on.
package convert

import (
	"bytes"
	"encoding/csv"
	"fmt"
	"io"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// SheetName is the visible worksheet the Airwallex importer reads.
const SheetName = "Airwallex batch transfer"

// firstDataRow is the row that holds the template's placeholder; real data
// overwrites it and continues downward.
const firstDataRow = 2

// Options controls the fixed (non-payee-specific) values written to every row,
// plus how the recipient statement reference is derived.
type Options struct {
	Country        string // "Transfer to", e.g. "New Zealand"
	TransferMethod string // e.g. "Bank Transfer"
	Currency       string // both "Currency recipient gets" and "Currency you pay"
	FeePaidBy      string // "Payer" or "Recipient"
	Purpose        string // "Transfer purpose", e.g. "Wages / salary"
	RecipientType  string // e.g. "Individual"

	// Reference, when non-empty, is written verbatim to every row's "Reference"
	// column. When empty, the reference is derived per row as "Salary <pay
	// period>" with a short month (e.g. "Salary Jan 2026") from the payment date
	// in the Xero record.
	Reference string
}

// DefaultOptions returns the settings for the monthly salary run: NZD bank
// transfers to individuals, fees paid by the payer, purpose "Wages / salary",
// and a per-row "Month Year" reference.
func DefaultOptions() Options {
	return Options{
		Country:        "New Zealand",
		TransferMethod: "Bank Transfer",
		Currency:       "NZD",
		FeePaidBy:      "Payer",
		Purpose:        "Wages / salary",
		RecipientType:  "Individual",
		Reference:      "",
	}
}

// Payment is one parsed row from the Xero export.
type Payment struct {
	Amount        float64
	AccountNumber string
	Name          string
	PayDate       time.Time // zero if the export had no parseable date
	rawDate       string    // original date text, kept for fallback references
}

// dateLayouts are the formats Xero is known to emit for the payment date.
var dateLayouts = []string{"2-Jan-2006", "02-Jan-2006", "2/01/2006", "02/01/2006"}

// ParseXero reads a Xero payment-export CSV and returns one Payment per
// non-empty record. Blank trailing lines are ignored.
func ParseXero(r io.Reader) ([]Payment, error) {
	cr := csv.NewReader(r)
	cr.FieldsPerRecord = -1 // the export is ragged; we only need the first columns
	cr.TrimLeadingSpace = true

	var payments []Payment
	line := 0
	for {
		rec, err := cr.Read()
		if err == io.EOF {
			break
		}
		line++
		if err != nil {
			return nil, fmt.Errorf("reading CSV line %d: %w", line, err)
		}
		if isBlankRecord(rec) {
			continue
		}
		if len(rec) < 3 {
			return nil, fmt.Errorf("line %d: expected at least 3 columns (amount, account, name), got %d", line, len(rec))
		}

		amountStr := strings.TrimSpace(rec[0])
		amount, err := strconv.ParseFloat(strings.ReplaceAll(amountStr, ",", ""), 64)
		if err != nil {
			return nil, fmt.Errorf("line %d: invalid amount %q: %w", line, rec[0], err)
		}

		account := sanitiseAccount(rec[1])
		if account == "" {
			return nil, fmt.Errorf("line %d: empty bank account number", line)
		}

		name := strings.TrimSpace(rec[2])
		if name == "" {
			return nil, fmt.Errorf("line %d: empty payee name", line)
		}

		p := Payment{Amount: amount, AccountNumber: account, Name: name}
		if len(rec) > 3 {
			p.rawDate = strings.TrimSpace(rec[3])
			p.PayDate = parseDate(p.rawDate)
		}
		payments = append(payments, p)
	}

	if len(payments) == 0 {
		return nil, fmt.Errorf("no payment rows found in CSV")
	}
	return payments, nil
}

// Build writes payments into a copy of the Airwallex template and returns the
// resulting workbook bytes. templateXLSX is the raw bytes of the official
// "Batch transfer template" file.
func Build(templateXLSX []byte, payments []Payment, opts Options) ([]byte, error) {
	f, err := excelize.OpenReader(bytes.NewReader(templateXLSX))
	if err != nil {
		return nil, fmt.Errorf("opening template: %w", err)
	}
	defer f.Close()

	if _, err := f.GetSheetIndex(SheetName); err != nil {
		return nil, fmt.Errorf("template is missing sheet %q: %w", SheetName, err)
	}

	// Styles: account number must stay text so the leading zero survives;
	// amount renders with two decimals; everything else is left general.
	textStyle, err := f.NewStyle(&excelize.Style{NumFmt: 49}) // 49 = "@" (text)
	if err != nil {
		return nil, fmt.Errorf("creating text style: %w", err)
	}
	amountStyle, err := f.NewStyle(&excelize.Style{NumFmt: 2}) // 2 = "0.00"
	if err != nil {
		return nil, fmt.Errorf("creating amount style: %w", err)
	}
	generalStyle, err := f.NewStyle(&excelize.Style{})
	if err != nil {
		return nil, fmt.Errorf("creating general style: %w", err)
	}

	for i, p := range payments {
		row := firstDataRow + i

		set := func(col string, v any) error {
			cell := fmt.Sprintf("%s%d", col, row)
			if err := f.SetCellValue(SheetName, cell, v); err != nil {
				return fmt.Errorf("row %d col %s: %w", row, col, err)
			}
			return f.SetCellStyle(SheetName, cell, cell, generalStyle)
		}

		if err := set("A", opts.Country); err != nil {
			return nil, err
		}
		if err := set("B", opts.TransferMethod); err != nil {
			return nil, err
		}
		if err := set("C", opts.Currency); err != nil {
			return nil, err
		}
		if err := set("D", opts.Currency); err != nil {
			return nil, err
		}

		// Amount — numeric, two decimals.
		amtCell := fmt.Sprintf("E%d", row)
		if err := f.SetCellFloat(SheetName, amtCell, p.Amount, 2, 64); err != nil {
			return nil, fmt.Errorf("row %d amount: %w", row, err)
		}
		if err := f.SetCellStyle(SheetName, amtCell, amtCell, amountStyle); err != nil {
			return nil, err
		}

		if err := set("F", opts.FeePaidBy); err != nil {
			return nil, err
		}
		if err := set("G", p.Name); err != nil {
			return nil, err
		}

		// Account number — text, to preserve leading zeros.
		acctCell := fmt.Sprintf("H%d", row)
		if err := f.SetCellStr(SheetName, acctCell, p.AccountNumber); err != nil {
			return nil, fmt.Errorf("row %d account: %w", row, err)
		}
		if err := f.SetCellStyle(SheetName, acctCell, acctCell, textStyle); err != nil {
			return nil, err
		}

		if err := set("I", opts.Purpose); err != nil {
			return nil, err
		}
		if err := set("J", reference(p, opts)); err != nil {
			return nil, err
		}
		if err := set("K", ""); err != nil { // Description (optional)
			return nil, err
		}
		if err := set("L", opts.RecipientType); err != nil {
			return nil, err
		}
		if err := set("M", ""); err != nil { // Request ID (optional)
			return nil, err
		}
	}

	// The template ships with a fixed worksheet dimension of "A1:M2". excelize
	// preserves it verbatim and does not recalculate it when rows are appended,
	// so without this the stale dimension makes Airwallex's importer stop after
	// the first data row. Expand it to cover every row we wrote.
	lastRow := firstDataRow + len(payments) - 1
	if err := f.SetSheetDimension(SheetName, fmt.Sprintf("A1:M%d", lastRow)); err != nil {
		return nil, fmt.Errorf("setting sheet dimension: %w", err)
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		return nil, fmt.Errorf("serialising workbook: %w", err)
	}
	return buf.Bytes(), nil
}

// reference returns the value for the "Reference" column. A fixed override wins;
// otherwise it is "Salary <pay period>" with a short month (e.g. "Salary Jan
// 2026") derived from the payment date, falling back to the raw date text when
// the date cannot be parsed.
func reference(p Payment, opts Options) string {
	if opts.Reference != "" {
		return opts.Reference
	}
	period := p.rawDate
	if !p.PayDate.IsZero() {
		period = p.PayDate.Format("Jan 2006")
	}
	return strings.TrimSpace("Salary " + period)
}

func isBlankRecord(rec []string) bool {
	for _, f := range rec {
		if strings.TrimSpace(f) != "" {
			return false
		}
	}
	return true
}

// sanitiseAccount strips spaces and hyphens, leaving the bare account digits as
// Airwallex expects (e.g. "01-0001-0000001-000" -> "0100010000001000").
func sanitiseAccount(s string) string {
	r := strings.NewReplacer(" ", "", "-", "")
	return r.Replace(strings.TrimSpace(s))
}

func parseDate(s string) time.Time {
	for _, layout := range dateLayouts {
		if t, err := time.Parse(layout, s); err == nil {
			return t
		}
	}
	return time.Time{}
}
