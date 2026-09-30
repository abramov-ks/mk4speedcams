package carindb

import (
	"bytes"
	"io"
	"os"
	"path/filepath"
	"testing"
)

// sparseCopy creates dst with the size of src and copies only src[from:].
func sparseCopy(t *testing.T, src, dst string, from, n int64) {
	in, err := os.Open(src)
	if err != nil {
		t.Fatal(err)
	}
	defer in.Close()
	st, _ := in.Stat()
	out, err := os.Create(dst)
	if err != nil {
		t.Fatal(err)
	}
	defer out.Close()
	if err := out.Truncate(st.Size()); err != nil {
		t.Fatal(err)
	}
	if n < 0 {
		n = st.Size() - from
	}
	if _, err := out.Seek(from, io.SeekStart); err != nil {
		t.Fatal(err)
	}
	if _, err := io.Copy(out, io.NewSectionReader(in, from, n)); err != nil {
		t.Fatal(err)
	}
}

// TestSaveRoundTrip copies the test DB, rewrites the layer with the same points
// and checks the result: replace mode, same size, same points.
func TestSaveRoundTrip(t *testing.T) {
	src := os.Getenv("CARINDB_TEST_DIR")
	if src == "" {
		t.Skip("CARINDB_TEST_DIR not set")
	}
	// only the directory/dataset (start of DB_0) and the layer (end of DB_1) are needed
	orig, err := Open(src, false)
	if err != nil {
		t.Fatal(err)
	}
	ol, err := orig.ReadLayer()
	if err != nil {
		t.Fatal(err)
	}
	orig.Close()
	dir := t.TempDir()
	sparseCopy(t, filepath.Join(src, "DB_0"), filepath.Join(dir, "DB_0"), 0, 64<<10)
	sparseCopy(t, filepath.Join(src, "DB_1"), filepath.Join(dir, "DB_1"), int64(ol.Lookup.Sector())*SectorSize, -1)
	db, err := Open(dir, true)
	if err != nil {
		t.Fatal(err)
	}
	l, err := db.ReadLayer()
	if err != nil {
		t.Fatal(err)
	}
	before := len(l.Points)
	size := db.FileSize(1)
	oldPrimary, _ := db.Read(l.Lookup)
	res, err := db.Save(l)
	if err != nil {
		t.Fatal(err)
	}
	t.Logf("%+v", res)
	if res.Mode != Replace || db.FileSize(1) != size {
		t.Errorf("mode %v size %d want replace/%d", res.Mode, db.FileSize(1), size)
	}
	newPrimary, _ := db.Read(res.Lookup)
	if !bytes.Equal(oldPrimary[4:16], newPrimary[4:16]) {
		t.Errorf("primary header differs % x vs % x", oldPrimary[:16], newPrimary[:16])
	}
	db.Close()

	db, _ = Open(dir, false)
	defer db.Close()
	l2, err := db.ReadLayer()
	if err != nil {
		t.Fatal(err)
	}
	if len(l2.Points) != before {
		t.Fatalf("points %d want %d", len(l2.Points), before)
	}
	// name offsets (record bytes 12..13) depend on the block layout, ignore them
	cnt := map[Point]int{}
	for _, p := range l.Points {
		p.Rec[12], p.Rec[13] = 0, 0
		cnt[p]++
	}
	for _, p := range l2.Points {
		p.Rec[12], p.Rec[13] = 0, 0
		cnt[p]--
	}
	for p, v := range cnt {
		if v != 0 {
			t.Fatalf("point mismatch %+v %d", p, v)
		}
	}
}
