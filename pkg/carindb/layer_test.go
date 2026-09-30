package carindb

import (
	"os"
	"testing"
)

// Set CARINDB_TEST_DIR to a folder with DB_0/DB_1 to run the tests against a real disc.
func testDB(t *testing.T, writable bool) *DB {
	dir := os.Getenv("CARINDB_TEST_DIR")
	if dir == "" {
		t.Skip("CARINDB_TEST_DIR not set")
	}
	db, err := Open(dir, writable)
	if err != nil {
		t.Fatal(err)
	}
	return db
}

// TestPlanMatchesDisc rebuilds the split plan from the points on disc and checks
// that every cell is divided into exactly the same blocks POISoN wrote.
func TestPlanMatchesDisc(t *testing.T) {
	db := testDB(t, false)
	defer db.Close()
	l, err := db.ReadLayer()
	if err != nil {
		t.Fatal(err)
	}
	t.Logf("points %d, blocks %d, width %d, cell %x", len(l.Points), len(l.Blocks), l.Width, l.CellSize)

	// rects of existing point blocks, per cell, in secondary order
	want := map[int][]Rect{}
	wantN := map[int]int{}
	// walk the primary lookup to keep cell numbers
	cell := 0
	for id := l.Lookup; cell < l.Width*l.Width; id = db.Next(id) {
		b, _ := db.Read(id)
		off, n := int(be.Uint16(b[8:])), int(be.Uint16(b[10:]))
		for i := 0; i < n && cell+i < l.Width*l.Width; i++ {
			s := ID(be.Uint32(b[off+4*i:]))
			if s == 0 {
				continue
			}
			sb, _ := db.Read(s)
			wantN[cell+i] = int(be.Uint16(sb[10:]))
			lo, ln := int(be.Uint16(sb[12:])), int(be.Uint16(sb[14:]))
			for k := 0; k < ln; k++ {
				pb, _ := db.Read(ID(be.Uint32(sb[lo+4*k:])))
				want[cell+i] = append(want[cell+i], Rect{int32(be.Uint32(pb[16:])), int32(be.Uint32(pb[20:])), int32(be.Uint32(pb[24:])), int32(be.Uint32(pb[28:]))})
			}
		}
		cell += n
	}
	plans, err := l.plan()
	if err != nil {
		t.Fatal(err)
	}
	if len(plans) != len(want) {
		t.Errorf("cells: got %d want %d", len(plans), len(want))
	}
	bad := 0
	for _, p := range plans {
		w := want[p.cell]
		ok := len(w) == len(p.leaves) && wantN[p.cell] == p.n*p.n
		for k := 0; ok && k < len(w); k++ {
			ok = w[k] == p.leaves[k].r
		}
		if !ok {
			bad++
			if bad < 5 {
				t.Logf("cell %d: quads got %d want %d, leaves got %d want %d", p.cell, p.n*p.n, wantN[p.cell], len(p.leaves), len(w))
			}
		}
	}
	if bad > 0 {
		t.Errorf("%d of %d cells differ", bad, len(plans))
	}
}
