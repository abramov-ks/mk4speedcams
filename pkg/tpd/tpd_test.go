package tpd

import (
	"bytes"
	"os"
	"testing"
)

// TestRebuildDiscTables regenerates the speedcam tables of the 2018 disc mod
// from their own URL data and compares the result with the original files.
func TestRebuildDiscTables(t *testing.T) {
	for _, c := range []struct{ idx, url, cat, form string }{
		{"testdata/0009.IDX", "testdata/0010.URL", "0013", "ENG/SE/SF_0013.HTM"},
		{"testdata/0011.IDX", "testdata/0012.URL", "0014", "ENG/SE/SF_0014.HTM"},
	} {
		url, _ := os.ReadFile(c.url)
		idx, _ := os.ReadFile(c.idx)
		pois, err := ReadURL(url)
		if err != nil {
			t.Fatal(err)
		}
		gi, gu := Table{Category: c.cat, FormPath: c.form, NameLen: 8}.Build(pois)
		if !bytes.Equal(gu, url) {
			for i := range gu {
				if i >= len(url) || gu[i] != url[i] {
					t.Errorf("%s differs at %d: %q vs %q", c.url, i, gu[imax(0, i-20):imin(len(gu), i+20)], url[imax(0, i-20):imin(len(url), i+20)])
					break
				}
			}
		}
		if len(gi) != len(idx) {
			t.Errorf("%s: size %d want %d", c.idx, len(gi), len(idx))
			continue
		}
		if !bytes.Equal(gi, idx) {
			t.Errorf("%s differs", c.idx)
		}
		t.Logf("%s: %d records, identical", c.idx, len(pois))
	}
}

func imin(a, b int) int {
	if a < b {
		return a
	}
	return b
}

func imax(a, b int) int {
	if a > b {
		return a
	}
	return b
}
