package carindb

import (
	"errors"
	"fmt"
	"math"
	"sort"
)

// Layout of the map dataset block (type 7). Layer records are 28 bytes:
// u32 lookup block ID, 4 x i32 bounding box, 8 bytes of flags.
const (
	datasetLayers    = 0x14
	datasetLayerSize = 28
	datasetCatStride = 6
	poiLayerIndex    = 1

	// point block (type 6)
	pointHeader  = 32
	pointRecSize = 28

	// primary lookup block (type 8)
	primaryHeader = 16

	// CatSpeedcam is the category POISoN uses for its points ("military base").
	CatSpeedcam = 28
	// catReplaced is the category whose menu slot POISoN takes over ("city centre").
	catReplaced = 48
)

// Coordinates of the POI layer: 5e7/9 units per degree, longitude shifted by 30.
const unitsPerDegree = 5e7 / 9

// FromWGS converts WGS84 degrees to layer units.
func FromWGS(lon, lat float64) (x, y int32) {
	return int32(math.Round((lon + 30) * unitsPerDegree)), int32(math.Round(lat * unitsPerDegree))
}

// ToWGS converts layer units to WGS84 degrees.
func ToWGS(x, y int32) (lon, lat float64) {
	return float64(x)/unitsPerDegree - 30, float64(y) / unitsPerDegree
}

// Point is one internal POI.
type Point struct {
	X, Y int32
	Cat  uint16
	Name string
	// Rec keeps the original 28-byte record so unknown fields survive a rewrite.
	Rec [pointRecSize]byte
}

// Rect is a bounding box in layer units, x1/y1 inclusive, x2/y2 exclusive.
type Rect struct{ X1, Y1, X2, Y2 int32 }

// Layer is the internal POI layer of the map.
type Layer struct {
	DatasetID ID
	Dataset   []byte
	Lookup    ID
	Bounds    Rect
	CellSize  int32
	Width     int // cells per side
	Points    []Point
	// Blocks lists every block of the layer as read (primary, secondary, points).
	Blocks []ID
}

// ReadLayer loads the POI layer.
func (db *DB) ReadLayer() (*Layer, error) {
	ds := db.Tables[TypeDataset].First
	data, err := db.Read(ds)
	if err != nil {
		return nil, err
	}
	if Type(data) != TypeDataset || data[6] != 0 {
		return nil, errors.New("unexpected dataset block")
	}
	l := &Layer{DatasetID: ds, Dataset: data}
	rec := data[datasetLayers+poiLayerIndex*datasetLayerSize:]
	l.Lookup = ID(be.Uint32(rec))
	l.Bounds = Rect{int32(be.Uint32(rec[4:])), int32(be.Uint32(rec[8:])), int32(be.Uint32(rec[12:])), int32(be.Uint32(rec[16:]))}
	if l.Lookup == 0 {
		return nil, errors.New("POI layer not found, empty lookup block ID")
	}

	// primary lookup: Width*Width links to secondary blocks, x-major order
	var secondary []struct {
		cell int
		id   ID
	}
	total, cell := -1, 0
	for id := l.Lookup; id != 0 && (total < 0 || cell < total); id = db.Next(id) {
		b, err := db.Read(id)
		if err != nil {
			return nil, err
		}
		if Type(b) != TypePrimary {
			return nil, fmt.Errorf("block %s: type %d, want primary lookup", id, Type(b))
		}
		if total < 0 {
			l.CellSize = int32(be.Uint32(b[12:]))
			if l.CellSize <= 0 {
				return nil, errors.New("bad POI cell size")
			}
			l.Width = int((l.Bounds.X2 - l.Bounds.X1) / l.CellSize)
			total = l.Width * l.Width
		}
		l.Blocks = append(l.Blocks, id)
		off, n := int(be.Uint16(b[8:])), int(be.Uint16(b[10:]))
		if n > total-cell {
			n = total - cell
		}
		for i := 0; i < n; i++ {
			if s := ID(be.Uint32(b[off+4*i:])); s != 0 {
				secondary = append(secondary, struct {
					cell int
					id   ID
				}{cell + i, s})
			}
		}
		cell += n
	}
	if cell != total {
		return nil, fmt.Errorf("primary lookup ended after %d of %d cells", cell, total)
	}

	for _, s := range secondary {
		b, err := db.Read(s.id)
		if err != nil {
			return nil, err
		}
		if Type(b) != TypeSecondary {
			return nil, fmt.Errorf("block %s: type %d, want secondary lookup", s.id, Type(b))
		}
		l.Blocks = append(l.Blocks, s.id)
		off, n := int(be.Uint16(b[12:])), int(be.Uint16(b[14:]))
		for i := 0; i < n; i++ {
			pid := ID(be.Uint32(b[off+4*i:]))
			if pid == 0 {
				continue
			}
			if err := l.readPoints(db, pid); err != nil {
				return nil, err
			}
		}
	}
	return l, nil
}

func (l *Layer) readPoints(db *DB, id ID) error {
	b, err := db.Read(id)
	if err != nil {
		return err
	}
	if Type(b) != TypePoints {
		return fmt.Errorf("block %s: type %d, want points", id, Type(b))
	}
	l.Blocks = append(l.Blocks, id)
	off, n := int(be.Uint16(b[8:])), int(be.Uint16(b[10:]))
	x1, y1 := int32(be.Uint32(b[16:])), int32(be.Uint32(b[20:]))
	for i := 0; i < n; i++ {
		r := b[off+i*pointRecSize:]
		p := Point{
			X:   x1 + int32(be.Uint16(r[6:]))<<6,
			Y:   y1 + int32(be.Uint16(r[8:]))<<6,
			Cat: be.Uint16(r[10:]),
		}
		copy(p.Rec[:], r[:pointRecSize])
		if no := int(be.Uint16(r[12:])); no != 0 && no < len(b) {
			e := no
			for e < len(b) && b[e] != 0 {
				e++
			}
			p.Name = string(b[no:e])
		}
		l.Points = append(l.Points, p)
	}
	return nil
}

// Stats returns the number of points per category.
func (l *Layer) Stats() map[uint16]int {
	m := map[uint16]int{}
	for _, p := range l.Points {
		m[p.Cat]++
	}
	return m
}

// ReplaceCategory drops all points of category cat and adds the given ones.
// Points outside the layer bounds are skipped; the number added is returned.
func (l *Layer) ReplaceCategory(cat uint16, pts []Point) int {
	keep := l.Points[:0]
	for _, p := range l.Points {
		if p.Cat != cat {
			keep = append(keep, p)
		}
	}
	l.Points = keep
	added := 0
	for _, p := range pts {
		if p.X < l.Bounds.X1 || p.X >= l.Bounds.X2 || p.Y < l.Bounds.Y1 || p.Y >= l.Bounds.Y2 {
			continue
		}
		p.Cat = cat
		l.Points = append(l.Points, p)
		added++
	}
	return added
}

// ----------------------------------------------------------------------------
// building

type leaf struct {
	r     Rect
	depth int
	pts   []Point
}

// pointsSize is the byte size of a point block; equal names are stored once.
func pointsSize(pts []Point) int {
	n := pointHeader + len(pts)*pointRecSize
	seen := map[string]bool{}
	for _, p := range pts {
		if p.Name != "" && !seen[p.Name] {
			seen[p.Name] = true
			n += len(p.Name) + 1
		}
	}
	return n
}

// split divides a cell the way POISoN does: binary halves, y first, then x,
// until the points of a part fit into one MaxBlockSectors block.
func split(r Rect, depth int, pts []Point, out *[]leaf) error {
	if pointsSize(pts) <= MaxBlockSectors*SectorSize {
		*out = append(*out, leaf{r, depth, pts})
		return nil
	}
	if depth >= 16 {
		return fmt.Errorf("too many points at one place (%d)", len(pts))
	}
	var lo, hi []Point
	a, b := r, r
	if depth%2 == 0 {
		mid := r.Y1 + (r.Y2-r.Y1)/2
		a.Y2, b.Y1 = mid, mid
		for _, p := range pts {
			if p.Y < mid {
				lo = append(lo, p)
			} else {
				hi = append(hi, p)
			}
		}
	} else {
		mid := r.X1 + (r.X2-r.X1)/2
		a.X2, b.X1 = mid, mid
		for _, p := range pts {
			if p.X < mid {
				lo = append(lo, p)
			} else {
				hi = append(hi, p)
			}
		}
	}
	if err := split(a, depth+1, lo, out); err != nil {
		return err
	}
	return split(b, depth+1, hi, out)
}

type cellPlan struct {
	cell   int
	rect   Rect
	leaves []leaf // only non-empty
	n      int    // quads per side
	all    []leaf
}

func (l *Layer) plan() ([]cellPlan, error) {
	byCell := map[int][]Point{}
	for _, p := range l.Points {
		cx := int((p.X - l.Bounds.X1) / l.CellSize)
		cy := int((p.Y - l.Bounds.Y1) / l.CellSize)
		c := cx*l.Width + cy
		byCell[c] = append(byCell[c], p)
	}
	cells := make([]int, 0, len(byCell))
	for c := range byCell {
		cells = append(cells, c)
	}
	sort.Ints(cells)
	plans := make([]cellPlan, 0, len(cells))
	for _, c := range cells {
		cx, cy := int32(c/l.Width), int32(c%l.Width)
		r := Rect{l.Bounds.X1 + cx*l.CellSize, l.Bounds.Y1 + cy*l.CellSize, 0, 0}
		r.X2, r.Y2 = r.X1+l.CellSize, r.Y1+l.CellSize
		var lv []leaf
		if err := split(r, 0, byCell[c], &lv); err != nil {
			return nil, err
		}
		maxDepth := 0
		for _, x := range lv {
			if x.depth > maxDepth {
				maxDepth = x.depth
			}
		}
		p := cellPlan{cell: c, rect: r, n: 1 << ((maxDepth + 1) / 2), all: lv}
		for _, x := range lv {
			if len(x.pts) > 0 {
				p.leaves = append(p.leaves, x)
			}
		}
		// POISoN lists the blocks of a cell ordered by their lower-left corner
		sort.SliceStable(p.leaves, func(i, j int) bool {
			a, b := p.leaves[i].r, p.leaves[j].r
			if a.X1 != b.X1 {
				return a.X1 < b.X1
			}
			return a.Y1 < b.Y1
		})
		plans = append(plans, p)
	}
	return plans, nil
}

func encodePoints(r Rect, pts []Point) []byte {
	sort.SliceStable(pts, func(i, j int) bool {
		if pts[i].X != pts[j].X {
			return pts[i].X < pts[j].X
		}
		return pts[i].Y < pts[j].Y
	})
	b := make([]byte, pointsSize(pts))
	be.PutUint16(b[4:], TypePoints)
	be.PutUint16(b[8:], pointHeader)
	be.PutUint16(b[10:], uint16(len(pts)))
	be.PutUint32(b[16:], uint32(r.X1))
	be.PutUint32(b[20:], uint32(r.Y1))
	be.PutUint32(b[24:], uint32(r.X2))
	be.PutUint32(b[28:], uint32(r.Y2))
	names := pointHeader + len(pts)*pointRecSize
	offs := map[string]int{}
	for i, p := range pts {
		rec := b[pointHeader+i*pointRecSize:]
		copy(rec, p.Rec[:])
		be.PutUint16(rec[6:], uint16((p.X-r.X1)>>6))
		be.PutUint16(rec[8:], uint16((p.Y-r.Y1)>>6))
		be.PutUint16(rec[10:], p.Cat)
		be.PutUint16(rec[12:], 0)
		if p.Name != "" {
			o, ok := offs[p.Name]
			if !ok {
				o = names
				offs[p.Name] = o
				names += copy(b[names:], p.Name) + 1
			}
			be.PutUint16(rec[12:], uint16(o))
		}
	}
	return b
}

func (l *Layer) encodeSecondary(p cellPlan, links []ID) []byte {
	nq := p.n * p.n
	linksOff := 0x14 + (2*nq+3)&^3
	b := make([]byte, linksOff+4*len(links))
	be.PutUint16(b[4:], TypeSecondary)
	be.PutUint16(b[8:], 0x14)
	be.PutUint16(b[10:], uint16(nq))
	be.PutUint16(b[12:], uint16(linksOff))
	be.PutUint16(b[14:], uint16(len(links)))
	qw := l.CellSize / int32(p.n)
	be.PutUint32(b[16:], uint32(qw))
	for k, lf := range p.leaves {
		be.PutUint32(b[linksOff+4*k:], uint32(links[k]))
		qx1, qx2 := int((lf.r.X1-p.rect.X1)/qw), int((lf.r.X2-p.rect.X1)/qw)
		qy1, qy2 := int((lf.r.Y1-p.rect.Y1)/qw), int((lf.r.Y2-p.rect.Y1)/qw)
		for qx := qx1; qx < qx2; qx++ {
			for qy := qy1; qy < qy2; qy++ {
				be.PutUint16(b[0x14+2*(qx*p.n+qy):], uint16(linksOff+4*k))
			}
		}
	}
	return b
}

// SaveMode tells where the new layer is written.
type SaveMode int

const (
	// Append writes the layer after the end of the last DB file.
	Append SaveMode = iota
	// Replace overwrites a layer that an earlier run appended to the end of the DB.
	Replace
)

func (m SaveMode) String() string {
	if m == Replace {
		return "replace"
	}
	return "append"
}

// DetectSaveMode checks whether the current layer occupies the tail of the last DB file.
func (db *DB) DetectSaveMode(l *Layer) (SaveMode, uint32) {
	last := db.LastFile()
	start := l.Lookup.Sector()
	var total uint32
	for _, id := range l.Blocks {
		if id.File() != last || id.Sector() < start {
			return Append, 0
		}
		total += uint32(id.Sectors())
	}
	end := uint32(db.sizes[last] / SectorSize)
	if l.Lookup.File() == last && start+total == end {
		return Replace, start
	}
	return Append, 0
}

// SaveResult describes a written layer.
type SaveResult struct {
	Mode                            SaveMode
	Lookup                          ID
	Primary, Secondary, PointBlocks int
	FirstSector, EndSector          uint32
	File                            int
	CategorySlotChanged             bool
}

// Save writes the layer back into the database and points the dataset block at it.
func (db *DB) Save(l *Layer) (*SaveResult, error) {
	plans, err := l.plan()
	if err != nil {
		return nil, err
	}
	mode, start := db.DetectSaveMode(l)
	file := db.LastFile()
	if mode == Append {
		start = uint32((db.sizes[file] + SectorSize - 1) / SectorSize)
	}
	res := &SaveResult{Mode: mode, File: file, FirstSector: start}

	// allocate: primaries, then secondaries, then point blocks
	total := l.Width * l.Width
	perPrimary := (MaxBlockSectors*SectorSize - primaryHeader) / 4
	nPrimary := (total + perPrimary - 1) / perPrimary
	sec := start
	alloc := func(size int) (ID, error) {
		n := SectorsFor(size)
		if n > 255 {
			return 0, fmt.Errorf("block of %d bytes is too big", size)
		}
		if sec+uint32(n) > maxSector {
			return 0, errors.New("unable to allocate block in DVD database, database too big")
		}
		id := MakeID(file, sec, n)
		sec += uint32(n)
		return id, nil
	}
	primaryIDs := make([]ID, nPrimary)
	for i := range primaryIDs {
		n := perPrimary
		if i == nPrimary-1 {
			n = total - i*perPrimary
		}
		if primaryIDs[i], err = alloc(primaryHeader + 4*n); err != nil {
			return nil, err
		}
	}
	secBlocks := make([][]byte, len(plans))
	secIDs := make([]ID, len(plans))
	links := make([][]ID, len(plans))
	for i, p := range plans {
		links[i] = make([]ID, len(p.leaves))
		size := len(l.encodeSecondary(p, links[i]))
		if secIDs[i], err = alloc(size); err != nil {
			return nil, err
		}
	}
	for i, p := range plans {
		for k, lf := range p.leaves {
			if links[i][k], err = alloc(pointsSize(lf.pts)); err != nil {
				return nil, err
			}
		}
	}
	res.EndSector = sec

	// write
	if mode == Replace {
		if err := db.Truncate(file, start); err != nil {
			return nil, err
		}
	}
	for i, p := range plans {
		for k, lf := range p.leaves {
			if err := db.Write(links[i][k], encodePoints(lf.r, lf.pts)); err != nil {
				return nil, err
			}
			res.PointBlocks++
		}
		secBlocks[i] = l.encodeSecondary(p, links[i])
		if err := db.Write(secIDs[i], secBlocks[i]); err != nil {
			return nil, err
		}
		res.Secondary++
	}
	cellLink := make(map[int]ID, len(plans))
	for i, p := range plans {
		cellLink[p.cell] = secIDs[i]
	}
	for i, id := range primaryIDs {
		first := i * perPrimary
		n := perPrimary
		if first+n > total {
			n = total - first
		}
		b := make([]byte, primaryHeader+4*n)
		be.PutUint16(b[4:], TypePrimary)
		be.PutUint16(b[8:], primaryHeader)
		be.PutUint16(b[10:], uint16(n))
		be.PutUint32(b[12:], uint32(l.CellSize))
		for c := 0; c < n; c++ {
			be.PutUint32(b[primaryHeader+4*c:], uint32(cellLink[first+c]))
		}
		if err := db.Write(id, b); err != nil {
			return nil, err
		}
		res.Primary++
	}
	if err := db.Truncate(file, sec); err != nil {
		return nil, err
	}

	// dataset block: new lookup, editable category slot
	ds := append([]byte{}, l.Dataset...)
	be.PutUint32(ds[datasetLayers+poiLayerIndex*datasetLayerSize:], uint32(primaryIDs[0]))
	changed, err := ensureCategorySlot(ds, CatSpeedcam)
	if err != nil {
		return nil, err
	}
	res.CategorySlotChanged = changed
	if err := db.Write(l.DatasetID, ds); err != nil {
		return nil, err
	}
	l.Dataset, l.Lookup = ds, primaryIDs[0]
	res.Lookup = primaryIDs[0]
	return res, nil
}

// CategorySlots returns the category list of the dataset block (POI menu).
func CategorySlots(ds []byte) []uint16 {
	off, n := int(be.Uint16(ds[8:])), int(be.Uint16(ds[10:]))
	out := make([]uint16, n)
	for i := range out {
		out[i] = be.Uint16(ds[off+i*datasetCatStride:])
	}
	return out
}

// ensureCategorySlot makes cat visible in the POI menu, taking the "city centre"
// slot like POISoN does. Returns true if the dataset block changed.
func ensureCategorySlot(ds []byte, cat uint16) (bool, error) {
	off, n := int(be.Uint16(ds[8:])), int(be.Uint16(ds[10:]))
	for i := 0; i < n; i++ {
		if be.Uint16(ds[off+i*datasetCatStride:]) == cat {
			return false, nil
		}
	}
	for i := 0; i < n; i++ {
		if be.Uint16(ds[off+i*datasetCatStride:]) == catReplaced {
			be.PutUint16(ds[off+i*datasetCatStride:], cat)
			return true, nil
		}
	}
	return false, errors.New("unable to set editable category in dataset block")
}
