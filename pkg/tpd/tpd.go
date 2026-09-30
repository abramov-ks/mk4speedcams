// Package tpd writes Mapscape TPD search tables (*.IDX / *.URL) used by the
// BMW MK4 travel guide / POI search, e.g. TPD/<product>/ENG/TABLES/0/0009.IDX.
//
// URL table:
//
//	<cat>URL-<form>\r\n
//	POSWGS:S:20|NAME:S:8|NT:S:1\0\r\n
//	fixed width records padded with \0, each ends with \r\n
//
// IDX table:
//
//	Glambda- <cat>IDX-<form>_\r\n
//	ID:I:<w>|POS:P:8|SELNAME:S:<n>|NAME:S:<n>\r\n
//	sparse index: 4 x \0 + lon(rec 0), then "|" + u32 rec + lon(rec) every <step> records, then \0\r\n
//	records: right aligned ID, 8 bytes POS (u32 BE lon, lat), SELNAME, NAME, \r\n
//
// POSWGS / POS units: lon = (deg + 30) * 5e7/9, lat = deg * 5e7/9.
package tpd

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"math"
	"os"
	"sort"
	"strconv"
)

// POI is one entry of a table.
type POI struct {
	Lon, Lat float64
	Name     string
}

// Encode converts degrees to POSWGS units.
func Encode(lon, lat float64) (x, y uint32) {
	return uint32(math.Round((lon + 30) * 5e7 / 9)), uint32(math.Round(lat * 5e7 / 9))
}

// Decode converts POSWGS units to degrees.
func Decode(x, y uint32) (lon, lat float64) {
	return float64(x)*9/5e7 - 30, float64(y) * 9 / 5e7
}

// Table describes one IDX/URL pair.
type Table struct {
	Category string // e.g. "0013"
	FormPath string // e.g. "ENG/SE/SF_0013.HTM"
	NameLen  int    // width of NAME / SELNAME, 0 = longest name
}

type rec struct {
	x, y uint32
	name string
}

func (t Table) prepare(pois []POI) ([]rec, int) {
	rs := make([]rec, len(pois))
	w := t.NameLen
	for i, p := range pois {
		x, y := Encode(p.Lon, p.Lat)
		rs[i] = rec{x, y, p.Name}
		if t.NameLen == 0 && len(p.Name) > w {
			w = len(p.Name)
		}
	}
	// the search engine expects the table ordered by longitude
	sort.SliceStable(rs, func(i, j int) bool { return rs[i].x < rs[j].x })
	return rs, w
}

func pad(s string, n int) []byte {
	b := make([]byte, n)
	copy(b, s)
	return b
}

// indexStep is the distance between entries of the sparse index line.
func indexStep(n int) int {
	switch i := n / 5; {
	case i < 20:
		return 20
	case i < 30:
		return 30
	case i < 40:
		return 40
	default:
		return 50
	}
}

// Build returns the IDX and URL file contents.
func (t Table) Build(pois []POI) (idx, url []byte) {
	rs, nameLen := t.prepare(pois)
	num := t.Category
	if len(num) == 4 && num[:2] == "00" {
		num = num[2:]
	}

	var u bytes.Buffer
	fmt.Fprintf(&u, "%sURL-%s\r\n", num, t.FormPath)
	fmt.Fprintf(&u, "POSWGS:S:20|NAME:S:%d|NT:S:1\x00\r\n", nameLen)
	for _, r := range rs {
		u.Write(pad(fmt.Sprintf("%d,%d", r.x, r.y), 20))
		u.Write(pad(r.name, nameLen))
		u.WriteString("1\r\n")
	}

	var x bytes.Buffer
	idw := len(strconv.Itoa(len(rs)))
	fmt.Fprintf(&x, "Glambda- %sIDX-%s_\r\n", t.Category, t.FormPath)
	fmt.Fprintf(&x, "ID:I:%d|POS:P:8|SELNAME:S:%d|NAME:S:%d\r\n", idw, nameLen, nameLen)
	var b4 [4]byte
	if len(rs) > 0 {
		x.Write([]byte{0, 0, 0, 0})
		binary.BigEndian.PutUint32(b4[:], rs[0].x)
		x.Write(b4[:])
		step := indexStep(len(rs))
		// same sampling as the 2018 speedcam mod tables that are known to work
		for r := step; r < len(rs); r += step {
			x.WriteByte('|')
			binary.BigEndian.PutUint32(b4[:], uint32(r))
			x.Write(b4[:])
			binary.BigEndian.PutUint32(b4[:], rs[r].x)
			x.Write(b4[:])
		}
	}
	x.WriteString("\x00\r\n")
	for i, r := range rs {
		fmt.Fprintf(&x, "%*d", idw, i)
		binary.BigEndian.PutUint32(b4[:], r.x)
		x.Write(b4[:])
		binary.BigEndian.PutUint32(b4[:], r.y)
		x.Write(b4[:])
		x.Write(pad(r.name, nameLen))
		x.Write(pad(r.name, nameLen))
		x.WriteString("\r\n")
	}
	return x.Bytes(), u.Bytes()
}

// Write stores the table as <dir>/<idxFile> and <dir>/<urlFile>.
func (t Table) Write(pois []POI, idxPath, urlPath string) error {
	idx, url := t.Build(pois)
	if err := os.WriteFile(idxPath, idx, 0o644); err != nil {
		return err
	}
	return os.WriteFile(urlPath, url, 0o644)
}

// ReadURL parses POSWGS and NAME of a URL table.
func ReadURL(data []byte) ([]POI, error) {
	l1 := bytes.IndexByte(data, '\n') + 1
	l2 := l1 + bytes.IndexByte(data[l1:], '\n') + 1
	if l1 <= 0 || l2 <= l1 {
		return nil, fmt.Errorf("bad URL header")
	}
	spec := bytes.TrimRight(data[l1:l2], "\r\n\x00")
	type field struct {
		name string
		w    int
	}
	var fs []field
	total := 0
	for _, f := range bytes.Split(spec, []byte("|")) {
		p := bytes.Split(f, []byte(":"))
		if len(p) != 3 {
			return nil, fmt.Errorf("bad field %q", f)
		}
		w, _ := strconv.Atoi(string(p[2]))
		fs = append(fs, field{string(p[0]), w})
		total += w
	}
	var out []POI
	body := data[l2:]
	for len(body) >= total {
		r := body[:total]
		body = bytes.TrimLeft(body[total:], "\r\n")
		var poi POI
		o := 0
		for _, f := range fs {
			v := string(bytes.TrimRight(r[o:o+f.w], "\x00"))
			o += f.w
			switch f.name {
			case "POSWGS":
				var x, y uint32
				if _, err := fmt.Sscanf(v, "%d,%d", &x, &y); err != nil {
					return nil, fmt.Errorf("bad POSWGS %q", v)
				}
				poi.Lon, poi.Lat = Decode(x, y)
			case "NAME":
				poi.Name = v
			}
		}
		out = append(out, poi)
	}
	return out, nil
}
