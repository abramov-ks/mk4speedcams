// Package carindb reads and writes the CARiN map database used by BMW MK4
// navigation DVDs (Siemens VDO / Continental), files DB/DB_0 .. DB/DB_3.
//
// Block IDs are 32-bit big-endian values:
//
//	bits 31..30  index of the database file (DB_0 .. DB_3)
//	bits 29..8   sector number inside that file
//	bits  7..0   block length in sectors
//
// Every block starts with an 8-byte header: its own ID, u16 block type and
// u16 pack info (byte 6 = pack method, byte 7 = unpacked length in sectors).
package carindb

import (
	"bytes"
	"compress/zlib"
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
)

const (
	// SectorSize is the only sector size seen on MK4 DVDs.
	SectorSize = 512
	// MaxBlockSectors is the biggest block POISoN writes (48 * 512 = 24576 bytes).
	MaxBlockSectors = 48
	maxFiles        = 4
	maxSector       = 1<<22 - 1

	// dirOffset is where the table directory lives in DB_0.
	dirOffset = 0x21c
	dirCount  = 25
)

// Block types used by the POI layer.
const (
	TypePoints    = 6
	TypeDataset   = 7
	TypePrimary   = 8
	TypeSecondary = 9
)

var be = binary.BigEndian

// ID is a block identifier.
type ID uint32

// MakeID builds a block ID.
func MakeID(file int, sector uint32, sectors int) ID {
	return ID(uint32(file)<<30 | (sector&maxSector)<<8 | uint32(sectors&0xff))
}

// File returns the database file index.
func (id ID) File() int { return int(uint32(id) >> 30) }

// Sector returns the first sector inside the file.
func (id ID) Sector() uint32 { return uint32(id) >> 8 & maxSector }

// Sectors returns the block length in sectors.
func (id ID) Sectors() int { return int(uint32(id) & 0xff) }

func (id ID) String() string { return fmt.Sprintf("%08x", uint32(id)) }

// Table is a directory entry of DB_0.
type Table struct {
	ID            int
	Lookup, First ID
	Last          ID
}

// DB is an opened database directory.
type DB struct {
	Dir    string
	files  [maxFiles]*os.File
	sizes  [maxFiles]int64
	names  [maxFiles]string
	Tables map[int]Table
}

// Open opens the database stored in dir (the folder that contains DB_0...).
func Open(dir string, writable bool) (*DB, error) {
	db := &DB{Dir: dir, Tables: map[int]Table{}}
	entries, err := os.ReadDir(dir)
	if err != nil {
		return nil, err
	}
	flag := os.O_RDONLY
	if writable {
		flag = os.O_RDWR
	}
	for _, e := range entries {
		n := strings.ToLower(e.Name())
		if len(n) != 4 || !strings.HasPrefix(n, "db_") || n[3] < '0' || n[3] > '3' {
			continue
		}
		i := int(n[3] - '0')
		f, err := os.OpenFile(filepath.Join(dir, e.Name()), flag, 0)
		if err != nil {
			db.Close()
			return nil, err
		}
		st, _ := f.Stat()
		db.files[i], db.sizes[i], db.names[i] = f, st.Size(), e.Name()
	}
	if db.files[0] == nil {
		return nil, fmt.Errorf("DB_0 not found in %s", dir)
	}
	if err := db.readDirectory(); err != nil {
		db.Close()
		return nil, err
	}
	return db, nil
}

// Close closes all files.
func (db *DB) Close() {
	for i, f := range db.files {
		if f != nil {
			f.Close()
			db.files[i] = nil
		}
	}
}

// FileSize returns the size of DB_n in bytes.
func (db *DB) FileSize(file int) int64 { return db.sizes[file] }

// LastFile returns the index of the last existing DB file.
func (db *DB) LastFile() int {
	for i := maxFiles - 1; i >= 0; i-- {
		if db.files[i] != nil {
			return i
		}
	}
	return 0
}

func (db *DB) readDirectory() error {
	buf := make([]byte, dirCount*16)
	if _, err := db.files[0].ReadAt(buf, dirOffset); err != nil {
		return err
	}
	for i := 0; i < dirCount; i++ {
		e := buf[i*16:]
		t := Table{ID: int(be.Uint16(e)), Lookup: ID(be.Uint32(e[4:])), First: ID(be.Uint32(e[8:])), Last: ID(be.Uint32(e[12:]))}
		db.Tables[t.ID] = t
	}
	if _, ok := db.Tables[TypeDataset]; !ok {
		return errors.New("dataset table (7) not found in directory")
	}
	return nil
}

// ReadRaw reads the stored bytes of a block without unpacking.
func (db *DB) ReadRaw(id ID) ([]byte, error) {
	f := db.files[id.File()]
	if f == nil {
		return nil, fmt.Errorf("block %s: DB_%d is missing", id, id.File())
	}
	n := id.Sectors()
	if n == 0 {
		return nil, fmt.Errorf("block %s: zero length", id)
	}
	buf := make([]byte, n*SectorSize)
	if _, err := f.ReadAt(buf, int64(id.Sector())*SectorSize); err != nil {
		return nil, fmt.Errorf("block %s: %w", id, err)
	}
	if ID(be.Uint32(buf)) != id {
		return nil, fmt.Errorf("block %s: header has id %08x", id, be.Uint32(buf))
	}
	return buf, nil
}

// Read reads a block and unpacks it. The returned slice includes the 8-byte header.
func (db *DB) Read(id ID) ([]byte, error) {
	raw, err := db.ReadRaw(id)
	if err != nil {
		return nil, err
	}
	switch raw[6] {
	case 0:
		return raw, nil
	case 2, 15:
		zr, err := zlib.NewReader(bytes.NewReader(raw[8:]))
		if err != nil {
			return nil, fmt.Errorf("block %s: %w", id, err)
		}
		body, err := io.ReadAll(zr)
		if err != nil && !errors.Is(err, io.ErrUnexpectedEOF) {
			return nil, fmt.Errorf("block %s: %w", id, err)
		}
		return append(append([]byte{}, raw[:8]...), body...), nil
	default:
		return nil, fmt.Errorf("block %s: unsupported pack method %d", id, raw[6])
	}
}

// Type returns the block type from a block header.
func Type(b []byte) int { return int(be.Uint16(b[4:])) }

// Next returns the ID of the block stored right after id, or 0 at end of file.
func (db *DB) Next(id ID) ID {
	f := db.files[id.File()]
	sec := id.Sector() + uint32(id.Sectors())
	if int64(sec)*SectorSize >= db.sizes[id.File()] {
		return 0
	}
	var h [4]byte
	if _, err := f.ReadAt(h[:], int64(sec)*SectorSize); err != nil {
		return 0
	}
	n := ID(be.Uint32(h[:]))
	if n.File() != id.File() || n.Sector() != sec || n.Sectors() == 0 {
		return 0
	}
	return n
}

// Write writes an unpacked block (header included) padded to its sector count.
func (db *DB) Write(id ID, data []byte) error {
	f := db.files[id.File()]
	if f == nil {
		return fmt.Errorf("block %s: DB_%d is missing", id, id.File())
	}
	size := id.Sectors() * SectorSize
	if len(data) > size {
		return fmt.Errorf("block %s: %d bytes do not fit %d sectors", id, len(data), id.Sectors())
	}
	buf := make([]byte, size)
	copy(buf, data)
	be.PutUint32(buf, uint32(id))
	off := int64(id.Sector()) * SectorSize
	if _, err := f.WriteAt(buf, off); err != nil {
		return err
	}
	if end := off + int64(size); end > db.sizes[id.File()] {
		db.sizes[id.File()] = end
	}
	return nil
}

// Truncate cuts DB_n to the given sector.
func (db *DB) Truncate(file int, sector uint32) error {
	size := int64(sector) * SectorSize
	if err := db.files[file].Truncate(size); err != nil {
		return err
	}
	db.sizes[file] = size
	return nil
}

// SectorsFor returns how many sectors n bytes need.
func SectorsFor(n int) int { return (n + SectorSize - 1) / SectorSize }
