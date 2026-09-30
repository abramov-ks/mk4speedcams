package carindb

import (
	"bytes"
	"testing"
)

// TestBlocksMatchDisc encodes every point and secondary block and compares bytes
// with the blocks on disc (block IDs taken from the disc).
func TestBlocksMatchDisc(t *testing.T) {
	db := testDB(t, false)
	defer db.Close()
	l, err := db.ReadLayer()
	if err != nil {
		t.Fatal(err)
	}
	plans, _ := l.plan()
	secOf := map[int]ID{}
	cell := 0
	for id := l.Lookup; cell < l.Width*l.Width; id = db.Next(id) {
		b, _ := db.Read(id)
		off, n := int(be.Uint16(b[8:])), int(be.Uint16(b[10:]))
		for i := 0; i < n && cell+i < l.Width*l.Width; i++ {
			if s := ID(be.Uint32(b[off+4*i:])); s != 0 {
				secOf[cell+i] = s
			}
		}
		cell += n
	}
	badP, badS, shown := 0, 0, 0
	for _, p := range plans {
		sid := secOf[p.cell]
		sb, _ := db.Read(sid)
		lo := int(be.Uint16(sb[12:]))
		links := make([]ID, len(p.leaves))
		for k := range p.leaves {
			links[k] = ID(be.Uint32(sb[lo+4*k:]))
			disc, _ := db.Read(links[k])
			enc := encodePoints(p.leaves[k].r, p.leaves[k].pts)
			be.PutUint32(enc, uint32(links[k]))
			full := make([]byte, len(disc))
			copy(full, enc)
			if SectorsFor(len(enc)) != links[k].Sectors() || !samePoints(full, disc) {
				badP++
				if shown < 3 {
					shown++
					for i := range full {
						if full[i] != disc[i] {
							t.Logf("point block %s differs at %#x: got % x want % x (sectors %d vs %d)", links[k], i, full[i:imin(i+16, len(full))], disc[i:imin(i+16, len(disc))], SectorsFor(len(enc)), links[k].Sectors())
							break
						}
					}
				}
			}
		}
		enc := l.encodeSecondary(p, links)
		be.PutUint32(enc, uint32(sid))
		full := make([]byte, len(sb))
		copy(full, enc)
		if SectorsFor(len(enc)) != sid.Sectors() || !bytes.Equal(full, sb) {
			badS++
			if badS < 3 {
				t.Logf("secondary %s differs\n got % x\nwant % x", sid, full[:32], sb[:32])
			}
		}
	}
	t.Logf("point blocks differ: %d, secondary differ: %d", badP, badS)
	if badP+badS > 0 {
		t.Fail()
	}
}

func nameAt(b []byte, o int) string {
	if o == 0 {
		return ""
	}
	e := o
	for e < len(b) && b[e] != 0 {
		e++
	}
	return string(b[o:e])
}

// samePoints compares two point blocks record by record; name offsets may differ.
func samePoints(a, b []byte) bool {
	if !bytes.Equal(a[:pointHeader], b[:pointHeader]) {
		return false
	}
	n := int(be.Uint16(a[10:]))
	key := func(blk []byte, i int) string {
		r := blk[pointHeader+i*pointRecSize:][:pointRecSize]
		return string(r[:12]) + string(r[14:]) + nameAt(blk, int(be.Uint16(r[12:])))
	}
	cnt := map[string]int{}
	var prevX uint16
	for i := 0; i < n; i++ {
		cnt[key(a, i)]++
		cnt[key(b, i)]--
		x := be.Uint16(b[pointHeader+i*pointRecSize+6:])
		if x < prevX { // disc block must be ordered by x as well
			return false
		}
		prevX = x
	}
	for _, v := range cnt {
		if v != 0 {
			return false
		}
	}
	return true
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
