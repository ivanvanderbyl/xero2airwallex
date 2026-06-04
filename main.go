// Command xero2airwallex converts a monthly Xero payment export (CSV) into an
// Airwallex batch-transfer workbook (XLSX).
//
// Usage:
//
//	xero2airwallex [flags] <xero-export.csv>
//
// The Airwallex batch-transfer template is embedded in the binary, so the tool
// is self-contained: hand it the Xero CSV each month and it writes a ready-to-
// upload .xlsx.
package main

import (
	_ "embed"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/ivanvanderbyl/xero2airwallex/internal/convert"
)

// template is the official Airwallex "Batch transfer template", embedded so the
// generated workbook keeps Airwallex's hidden validation and version sheets.
//
//go:embed template.xlsx
var template []byte

func main() {
	if err := run(); err != nil {
		fmt.Fprintln(os.Stderr, "xero2airwallex:", err)
		os.Exit(1)
	}
}

func run() error {
	opts := convert.DefaultOptions()

	out := flag.String("o", "", "output .xlsx path (default: <input>_airwallex.xlsx)")
	flag.StringVar(&opts.FeePaidBy, "fee", opts.FeePaidBy, `who pays the transfer fee ("Payer" or "Recipient")`)
	flag.StringVar(&opts.Purpose, "purpose", opts.Purpose, "Airwallex transfer purpose for every row")
	flag.StringVar(&opts.Currency, "currency", opts.Currency, "currency paid and received")
	flag.StringVar(&opts.Country, "country", opts.Country, `destination country ("Transfer to")`)
	flag.StringVar(&opts.RecipientType, "recipient-type", opts.RecipientType, `recipient type ("Individual" or "Business")`)
	flag.StringVar(&opts.Reference, "reference", "", `fixed reference for every row (default: per-row "Salary <pay period>", e.g. "Salary May 2026")`)

	flag.Usage = func() {
		fmt.Fprintf(flag.CommandLine.Output(), "Usage: %s [flags] <xero-export.csv>\n\nFlags:\n", filepath.Base(os.Args[0]))
		flag.PrintDefaults()
	}
	flag.Parse()

	if flag.NArg() != 1 {
		flag.Usage()
		return fmt.Errorf("expected exactly one input CSV path")
	}
	in := flag.Arg(0)

	outPath := *out
	if outPath == "" {
		base := strings.TrimSuffix(filepath.Base(in), filepath.Ext(in))
		outPath = filepath.Join(filepath.Dir(in), base+"_airwallex.xlsx")
	}

	csvFile, err := os.Open(in)
	if err != nil {
		return fmt.Errorf("opening input: %w", err)
	}
	defer csvFile.Close()

	payments, err := convert.ParseXero(csvFile)
	if err != nil {
		return err
	}

	// Surface any rows where the pay period could not be derived, so a bad
	// reference is never silently shipped to the bank.
	if opts.Reference == "" {
		for _, p := range payments {
			if p.PayDate.IsZero() {
				fmt.Fprintf(os.Stderr, "warning: could not parse a pay date for %q; reference left as-is\n", p.Name)
			}
		}
	}

	xlsx, err := convert.Build(template, payments, opts)
	if err != nil {
		return err
	}
	if err := os.WriteFile(outPath, xlsx, 0o644); err != nil {
		return fmt.Errorf("writing output: %w", err)
	}

	var total float64
	for _, p := range payments {
		total += p.Amount
	}
	fmt.Printf("Wrote %d payment(s) totalling %s %.2f to %s\n", len(payments), opts.Currency, total, outPath)
	return nil
}
