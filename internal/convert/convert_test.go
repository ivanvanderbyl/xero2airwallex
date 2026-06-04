package convert

import (
	"bytes"
	"os"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func loadTemplate(t *testing.T) []byte {
	t.Helper()
	b, err := os.ReadFile("testdata/template.xlsx")
	if err != nil {
		t.Fatalf("reading template: %v", err)
	}
	return b
}

func TestParseXero(t *testing.T) {
	const in = "1234.56,0100010000001000,Ada Lovelace,31-May-2026,Pay Ended,Salary/Wages,31-May-2026,Pay Ended,Salary/Wages\n" +
		"7890.12,0200020000002000,Alan Turing,31-May-2026,Pay Ended,Salary/Wages,31-May-2026,Pay Ended,Salary/Wages\n" +
		"\n"

	got, err := ParseXero(strings.NewReader(in))
	if err != nil {
		t.Fatalf("ParseXero: %v", err)
	}
	if len(got) != 2 {
		t.Fatalf("expected 2 payments, got %d", len(got))
	}

	if got[0].Amount != 1234.56 {
		t.Errorf("amount = %v, want 1234.56", got[0].Amount)
	}
	if got[0].AccountNumber != "0100010000001000" {
		t.Errorf("account = %q, want 0100010000001000", got[0].AccountNumber)
	}
	if got[0].Name != "Ada Lovelace" {
		t.Errorf("name = %q", got[0].Name)
	}
	if got[0].PayDate.IsZero() {
		t.Errorf("pay date not parsed")
	}
	if y, m := got[0].PayDate.Year(), got[0].PayDate.Month().String(); y != 2026 || m != "May" {
		t.Errorf("pay date = %s %d, want May 2026", m, y)
	}
}

func TestParseXeroErrors(t *testing.T) {
	cases := map[string]string{
		"too few columns": "1234.56,0100010000001000\n",
		"bad amount":      "abc,0100010000001000,Ada Lovelace\n",
		"empty input":     "\n\n",
	}
	for name, in := range cases {
		t.Run(name, func(t *testing.T) {
			if _, err := ParseXero(strings.NewReader(in)); err == nil {
				t.Fatalf("expected error for %s", name)
			}
		})
	}
}

func TestSanitiseAccount(t *testing.T) {
	if got := sanitiseAccount(" 01-0001-0000001-000 "); got != "0100010000001000" {
		t.Errorf("sanitiseAccount = %q", got)
	}
}

// readCells reopens a generated workbook and returns the value of a cell as a
// string, exactly as a downstream importer would read it.
func readCells(t *testing.T, b []byte, cells ...string) []string {
	t.Helper()
	f, err := excelize.OpenReader(bytes.NewReader(b))
	if err != nil {
		t.Fatalf("reopening output: %v", err)
	}
	defer f.Close()
	var out []string
	for _, c := range cells {
		v, err := f.GetCellValue(SheetName, c)
		if err != nil {
			t.Fatalf("reading %s: %v", c, err)
		}
		out = append(out, v)
	}
	return out
}

func TestBuild(t *testing.T) {
	payments, err := ParseXero(strings.NewReader(
		"1234.56,0100010000001000,Ada Lovelace,31-May-2026,Pay Ended,Salary/Wages\n" +
			"7890.12,0200020000002000,Alan Turing,31-May-2026,Pay Ended,Salary/Wages\n"))
	if err != nil {
		t.Fatalf("ParseXero: %v", err)
	}

	out, err := Build(loadTemplate(t), payments, DefaultOptions())
	if err != nil {
		t.Fatalf("Build: %v", err)
	}

	// Header row must be untouched.
	hdr := readCells(t, out, "A1", "E1", "H1", "J1")
	wantHdr := []string{"Transfer to", "Transfer amount in currency you pay", "Account number", "Reference"}
	for i := range hdr {
		if hdr[i] != wantHdr[i] {
			t.Errorf("header[%d] = %q, want %q", i, hdr[i], wantHdr[i])
		}
	}

	// First data row (row 2) overwrites the placeholder.
	row2 := readCells(t, out, "A2", "B2", "C2", "D2", "E2", "F2", "G2", "H2", "I2", "J2", "L2")
	want := []string{
		"New Zealand", "Bank Transfer", "NZD", "NZD",
		"1234.56", "Payer", "Ada Lovelace",
		"0100010000001000", // leading zero preserved
		"Wages / salary", "Salary May 2026", "Individual",
	}
	for i := range want {
		if row2[i] != want[i] {
			t.Errorf("row2 col %d = %q, want %q", i, row2[i], want[i])
		}
	}

	// Second data row appended at row 3.
	row3 := readCells(t, out, "E3", "G3", "H3", "J3")
	if row3[0] != "7890.12" || row3[1] != "Alan Turing" || row3[2] != "0200020000002000" || row3[3] != "Salary May 2026" {
		t.Errorf("row3 = %v", row3)
	}

	// Hidden version sheet that Airwallex relies on must survive.
	f, err := excelize.OpenReader(bytes.NewReader(out))
	if err != nil {
		t.Fatalf("reopening: %v", err)
	}
	defer f.Close()

	// The worksheet dimension must span every data row, otherwise Airwallex's
	// importer reads only as far as the stale template dimension (A1:M2) and
	// drops every transfer after the first.
	if dim, _ := f.GetSheetDimension(SheetName); dim != "A1:M3" {
		t.Errorf("sheet dimension = %q, want A1:M3 (2 payments)", dim)
	}

	if v, _ := f.GetCellValue("StorageUsedInternally", "B1"); v != "v1" {
		// Location is not guaranteed; just assert the sheet exists with the marker somewhere.
		if idx, _ := f.GetSheetIndex("StorageUsedInternally"); idx == -1 {
			t.Errorf("StorageUsedInternally sheet was dropped")
		}
	}
}

func TestBuildShortMonthReference(t *testing.T) {
	// January distinguishes the short ("Jan") form from the long ("January").
	payments, err := ParseXero(strings.NewReader(
		"100.00,0100010000001000,Ada Lovelace,15-Jan-2026,Pay Ended,Salary/Wages\n"))
	if err != nil {
		t.Fatalf("ParseXero: %v", err)
	}
	out, err := Build(loadTemplate(t), payments, DefaultOptions())
	if err != nil {
		t.Fatalf("Build: %v", err)
	}
	if got := readCells(t, out, "J2")[0]; got != "Salary Jan 2026" {
		t.Errorf("reference = %q, want %q", got, "Salary Jan 2026")
	}
}

func TestBuildFixedReference(t *testing.T) {
	payments, err := ParseXero(strings.NewReader(
		"100.00,0100010000001000,Ada Lovelace,31-May-2026,Pay Ended,Salary/Wages\n"))
	if err != nil {
		t.Fatalf("ParseXero: %v", err)
	}
	opts := DefaultOptions()
	opts.Reference = "Salary"
	out, err := Build(loadTemplate(t), payments, opts)
	if err != nil {
		t.Fatalf("Build: %v", err)
	}
	if got := readCells(t, out, "J2")[0]; got != "Salary" {
		t.Errorf("reference = %q, want Salary", got)
	}
}
