// poipatch writes speed cameras into a BMW MK4 map DVD, replacing POISoN Patcher.
//
// Cameras go into two places:
//   - the internal POI layer of DB/DB_0..DB_1 (category "military base", the one
//     POISoN uses; the nav warns about it with sound when patched software is used);
//   - optionally the TPD travel guide tables (ENG/TABLES/0/0009,0010,0013,0014),
//     which show up in the POI search list.
//
// Usage:
//
//	poipatch -iso ./iso -points output/poison.txt
//	poipatch -iso ./iso -export current.txt        # dump cameras from the disc
package main

import (
	"bufio"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"

	"github.com/abramov-ks/mk4speedcams/pkg/carindb"
	"github.com/abramov-ks/mk4speedcams/pkg/tpd"
)

var categoryNames = map[uint16]string{
	11: "bmw dealer/service", 12: "petrol station", 13: "car rental", 14: "car park",
	15: "car park + public transport", 16: "rest area", 17: "automobile clubs", 20: "tourist attraction",
	21: "hotel", 22: "restaurant", 23: "bank", 24: "community centre", 25: "library", 26: "law courts",
	27: "fire station", 28: "military base (speedcams)", 29: "consulate/embassy", 30: "cash dispenser",
	31: "tourist information", 32: "museum", 33: "theatre / civic centre", 34: "civic centre",
	35: "sports centre", 36: "church", 37: "monument", 38: "amusement park", 39: "park / recreation / fitness",
	40: "exhibition/convention centre", 41: "hospital", 42: "police", 43: "town hall", 44: "post office",
	45: "medical assistance", 46: "chemist", 47: "shopping centre", 48: "city centre",
	49: "theatre / civic centre", 50: "golf course", 51: "railway station", 52: "airport", 53: "ferry",
	54: "bus station", 55: "marina", 56: "college/university", 57: "entertainment", 58: "border crossing",
	59: "motorcycles", 60: "car dealers", 61: "trading estate",
}

func readPoints(path string) ([]tpd.POI, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	var out []tpd.POI
	sc := bufio.NewScanner(f)
	line := 0
	for sc.Scan() {
		line++
		s := strings.TrimSpace(sc.Text())
		if s == "" || strings.HasPrefix(s, "#") {
			continue
		}
		p := strings.Split(s, ",")
		if len(p) < 2 {
			return nil, fmt.Errorf("%s:%d: want lon,lat", path, line)
		}
		lon, err1 := strconv.ParseFloat(strings.TrimSpace(p[0]), 64)
		lat, err2 := strconv.ParseFloat(strings.TrimSpace(p[1]), 64)
		if err1 != nil || err2 != nil || lat < -90 || lat > 90 || lon < -180 || lon > 180 {
			return nil, fmt.Errorf("%s:%d: bad coordinates %q", path, line, s)
		}
		out = append(out, tpd.POI{Lon: lon, Lat: lat, Name: "SpeedCam"})
	}
	return out, sc.Err()
}

// dedup removes exact duplicates (same position after rounding to the DB grid).
func dedup(pts []tpd.POI) []tpd.POI {
	seen := map[[2]int32]bool{}
	out := pts[:0]
	for _, p := range pts {
		x, y := carindb.FromWGS(p.Lon, p.Lat)
		k := [2]int32{x >> 6, y >> 6}
		if !seen[k] {
			seen[k] = true
			out = append(out, p)
		}
	}
	return out
}

func findDir(root string, names ...string) string {
	for _, n := range names {
		m, _ := filepath.Glob(filepath.Join(root, n))
		if len(m) > 0 {
			return m[0]
		}
	}
	return ""
}

func main() {
	iso := flag.String("iso", "", "folder with the extracted DVD (contains DB and TPD)")
	points := flag.String("points", "", "cameras, one \"lon,lat\" per line (poison.txt format)")
	export := flag.String("export", "", "write cameras found on the disc to this file and exit")
	noTPD := flag.Bool("no-tpd", false, "do not touch the TPD travel guide tables")
	noDB := flag.Bool("no-db", false, "do not touch the internal POI layer")
	dry := flag.Bool("dry-run", false, "only show what would be done")
	flag.Parse()
	if *iso == "" || (*points == "" && *export == "") {
		flag.Usage()
		os.Exit(2)
	}
	dbDir := findDir(*iso, "DB", "db")
	if dbDir == "" {
		log.Fatalf("no DB folder in %s", *iso)
	}

	db, err := carindb.Open(dbDir, !*dry && *export == "" && !*noDB)
	if err != nil {
		log.Fatal(err)
	}
	defer db.Close()
	layer, err := db.ReadLayer()
	if err != nil {
		log.Fatal(err)
	}
	stats := layer.Stats()
	log.Printf("POI layer: %d points, %d blocks, grid %dx%d", len(layer.Points), len(layer.Blocks), layer.Width, layer.Width)
	cats := make([]int, 0, len(stats))
	for c := range stats {
		cats = append(cats, int(c))
	}
	sort.Ints(cats)
	for _, c := range cats {
		log.Printf("  %3d %-30s %8d", c, categoryNames[uint16(c)], stats[uint16(c)])
	}
	mode, _ := db.DetectSaveMode(layer)
	log.Printf("save mode: %s", mode)

	if *export != "" {
		f, err := os.Create(*export)
		if err != nil {
			log.Fatal(err)
		}
		w := bufio.NewWriter(f)
		n := 0
		for _, p := range layer.Points {
			if p.Cat == carindb.CatSpeedcam {
				lon, lat := carindb.ToWGS(p.X, p.Y)
				fmt.Fprintf(w, "%.6f,%.6f\n", lon, lat)
				n++
			}
		}
		w.Flush()
		f.Close()
		log.Printf("exported %d cameras to %s", n, *export)
		return
	}

	pois, err := readPoints(*points)
	if err != nil {
		log.Fatal(err)
	}
	total := len(pois)
	pois = dedup(pois)
	log.Printf("input: %d cameras (%d after removing duplicates)", total, len(pois))

	if !*noDB {
		pts := make([]carindb.Point, len(pois))
		for i, p := range pois {
			pts[i].X, pts[i].Y = carindb.FromWGS(p.Lon, p.Lat)
		}
		old := stats[carindb.CatSpeedcam]
		added := layer.ReplaceCategory(carindb.CatSpeedcam, pts)
		log.Printf("internal layer: %d old cameras replaced by %d (%d outside the map skipped)", old, added, len(pts)-added)
		if *dry {
			log.Printf("dry run: DB not written")
		} else {
			res, err := db.Save(layer)
			if err != nil {
				log.Fatalf("save failed: %v", err)
			}
			log.Printf("DB saved (%s): DB_%d sectors %#x..%#x, %d primary, %d secondary, %d point blocks, lookup %s",
				res.Mode, res.File, res.FirstSector, res.EndSector, res.Primary, res.Secondary, res.PointBlocks, res.Lookup)
			if res.CategorySlotChanged {
				log.Printf("category %d took the \"city centre\" slot of the POI menu", carindb.CatSpeedcam)
			}
		}
	}

	if !*noTPD {
		eng := findDir(*iso, "TPD/*/ENG/TABLES/0")
		if eng == "" {
			log.Printf("TPD ENG tables not found, skipping")
			return
		}
		for _, t := range []struct{ idx, url, cat string }{
			{"0009.IDX", "0010.URL", "0013"},
			{"0013.IDX", "0014.URL", "0015"},
		} {
			tbl := tpd.Table{Category: t.cat, FormPath: "ENG/SE/SF_" + t.cat + ".HTM", NameLen: 8}
			if *dry {
				continue
			}
			if err := tbl.Write(pois, filepath.Join(eng, t.idx), filepath.Join(eng, t.url)); err != nil {
				log.Fatal(err)
			}
		}
		log.Printf("TPD speedcam tables written to %s (%d cameras)", eng, len(pois))
	}
}
