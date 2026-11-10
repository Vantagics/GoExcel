package gospreadsheet

import (
	"encoding/binary"
	"errors"
	"unicode/utf16"
)

// Compound File Binary (MS-CFB) reader and a writer sufficient for an
// ECMA-376 encrypted package: EncryptionInfo in the mini stream and
// EncryptedPackage in the regular stream.

const (
	cfbMagic0 = 0xE011CFD0
	cfbMagic1 = 0xE11AB1A1

	cfbFreeSect = 0xFFFFFFFF
	cfbEndChain = 0xFFFFFFFE
	cfbFatSect  = 0xFFFFFFFD
	cfbDifSect  = 0xFFFFFFFC

	cfbDirEmpty  = 0
	cfbDirStream = 2
	cfbDirRoot   = 5
)

var oleMagic = []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

func isOLE(data []byte) bool {
	return len(data) >= 8 && string(data[:8]) == string(oleMagic)
}

type cfbFile struct {
	sectorSize int
	miniSize   int
	cutoff     uint32
	fat        []uint32
	miniFAT    []uint32
	miniStream []byte
	dirs       []cfbDir
}

type cfbDir struct {
	name  string
	typ   byte
	left  uint32
	right uint32
	child uint32
	start uint32
	size  uint64
}

func openCFB(data []byte) (*cfbFile, error) {
	if !isOLE(data) {
		return nil, errors.New("not an OLE compound file")
	}
	if len(data) < 512 {
		return nil, errors.New("OLE compound file is truncated")
	}
	sectorShift := binary.LittleEndian.Uint16(data[0x1E:])
	miniShift := binary.LittleEndian.Uint16(data[0x20:])
	if sectorShift < 9 || sectorShift > 12 || miniShift == 0 || miniShift > 12 {
		return nil, errors.New("OLE compound file has an unsupported sector size")
	}
	f := &cfbFile{
		sectorSize: 1 << sectorShift,
		miniSize:   1 << miniShift,
		cutoff:     binary.LittleEndian.Uint32(data[0x38:]),
	}
	if f.cutoff == 0 {
		f.cutoff = 4096
	}
	numFAT := binary.LittleEndian.Uint32(data[0x2C:])
	firstDir := binary.LittleEndian.Uint32(data[0x30:])
	firstMini := binary.LittleEndian.Uint32(data[0x3C:])
	numMini := binary.LittleEndian.Uint32(data[0x40:])
	firstDIFAT := binary.LittleEndian.Uint32(data[0x44:])
	numDIFAT := binary.LittleEndian.Uint32(data[0x48:])

	difat := make([]uint32, 0, numFAT)
	for i := 0; i < 109 && uint32(len(difat)) < numFAT; i++ {
		v := binary.LittleEndian.Uint32(data[0x4C+i*4:])
		if v == cfbFreeSect {
			break
		}
		difat = append(difat, v)
	}
	next := firstDIFAT
	for n := uint32(0); n < numDIFAT && next < cfbDifSect; n++ {
		sec, err := readSector(data, f.sectorSize, next)
		if err != nil {
			return nil, err
		}
		entries := f.sectorSize/4 - 1
		for i := 0; i < entries && uint32(len(difat)) < numFAT; i++ {
			difat = append(difat, binary.LittleEndian.Uint32(sec[i*4:]))
		}
		next = binary.LittleEndian.Uint32(sec[entries*4:])
	}
	fatSectors := make([]byte, 0, int(numFAT)*f.sectorSize)
	for _, id := range difat {
		sec, err := readSector(data, f.sectorSize, id)
		if err != nil {
			return nil, err
		}
		fatSectors = append(fatSectors, sec...)
	}
	f.fat = make([]uint32, len(fatSectors)/4)
	for i := range f.fat {
		f.fat[i] = binary.LittleEndian.Uint32(fatSectors[i*4:])
	}
	dirBytes, err := f.readChain(data, firstDir, 0)
	if err != nil {
		return nil, err
	}
	for off := 0; off+128 <= len(dirBytes); off += 128 {
		f.dirs = append(f.dirs, parseDir(dirBytes[off:off+128]))
	}
	if len(f.dirs) == 0 {
		return nil, errors.New("OLE compound file has no directory")
	}
	root := f.dirs[0]
	if numMini > 0 && firstMini != cfbEndChain && firstMini != cfbFreeSect {
		miniFATBytes, err := f.readChain(data, firstMini, int(numMini)*f.sectorSize)
		if err != nil {
			return nil, err
		}
		f.miniFAT = make([]uint32, len(miniFATBytes)/4)
		for i := range f.miniFAT {
			f.miniFAT[i] = binary.LittleEndian.Uint32(miniFATBytes[i*4:])
		}
		mini, err := f.readChain(data, root.start, int(root.size))
		if err != nil {
			return nil, err
		}
		f.miniStream = mini
	}
	return f, nil
}

func parseDir(b []byte) cfbDir {
	nameLen := int(binary.LittleEndian.Uint16(b[64:]))
	if nameLen > 64 {
		nameLen = 64
	}
	if nameLen >= 2 {
		nameLen -= 2 // drop the UTF-16 NUL
	}
	u16 := make([]uint16, nameLen/2)
	for i := range u16 {
		u16[i] = binary.LittleEndian.Uint16(b[i*2:])
	}
	return cfbDir{
		name:  string(utf16.Decode(u16)),
		typ:   b[66],
		left:  binary.LittleEndian.Uint32(b[68:]),
		right: binary.LittleEndian.Uint32(b[72:]),
		child: binary.LittleEndian.Uint32(b[76:]),
		start: binary.LittleEndian.Uint32(b[116:]),
		size:  binary.LittleEndian.Uint64(b[120:]),
	}
}

func readSector(data []byte, sectorSize int, id uint32) ([]byte, error) {
	off := int(id+1) * sectorSize
	if off < 0 || off+sectorSize > len(data) {
		return nil, errors.New("OLE sector is outside the file")
	}
	return data[off : off+sectorSize], nil
}

func (f *cfbFile) readChain(data []byte, start uint32, size int) ([]byte, error) {
	if start == cfbEndChain || start == cfbFreeSect {
		return nil, nil
	}
	var out []byte
	seen := map[uint32]bool{}
	id := start
	for id != cfbEndChain && id < cfbDifSect {
		if seen[id] {
			return nil, errors.New("OLE sector chain loops")
		}
		seen[id] = true
		sec, err := readSector(data, f.sectorSize, id)
		if err != nil {
			return nil, err
		}
		out = append(out, sec...)
		if size > 0 && len(out) >= size {
			return out[:size], nil
		}
		if int(id) >= len(f.fat) {
			return nil, errors.New("OLE FAT does not cover a sector")
		}
		id = f.fat[id]
	}
	if size > 0 && len(out) >= size {
		return out[:size], nil
	}
	return out, nil
}

func (f *cfbFile) readNamedStream(data []byte, name string) ([]byte, error) {
	for _, d := range f.dirs {
		if d.typ != cfbDirStream || d.name != name {
			continue
		}
		if d.size == 0 {
			return []byte{}, nil
		}
		if d.size < uint64(f.cutoff) {
			if len(f.miniStream) == 0 || len(f.miniFAT) == 0 {
				// Some writers still place a short stream in the regular FAT.
				return f.readChain(data, d.start, int(d.size))
			}
			return f.readMini(d.start, int(d.size))
		}
		return f.readChain(data, d.start, int(d.size))
	}
	return nil, errors.New("OLE stream not found: " + name)
}

func (f *cfbFile) readMini(start uint32, size int) ([]byte, error) {
	var out []byte
	seen := map[uint32]bool{}
	id := start
	for id != cfbEndChain && id < cfbDifSect && (size == 0 || len(out) < size) {
		if seen[id] {
			return nil, errors.New("OLE mini-sector chain loops")
		}
		seen[id] = true
		off := int(id) * f.miniSize
		if off < 0 || off+f.miniSize > len(f.miniStream) {
			return nil, errors.New("OLE mini sector is outside the mini stream")
		}
		out = append(out, f.miniStream[off:off+f.miniSize]...)
		if int(id) >= len(f.miniFAT) {
			return nil, errors.New("OLE mini FAT does not cover a sector")
		}
		id = f.miniFAT[id]
	}
	if size > 0 && len(out) >= size {
		return out[:size], nil
	}
	return out, nil
}

// writeCFBPackage writes a version-3 compound file with two streams.
// info is stored in the mini stream; pkg is stored in the regular FAT.
func writeCFBPackage(info, pkg []byte) ([]byte, error) {
	if len(info) >= 4096 {
		return nil, errors.New("EncryptionInfo exceeds the mini-stream cutoff")
	}
	// Streams shorter than 4096 bytes live in the mini stream. The encrypted
	// package is easier to keep in the regular FAT: pad with zeros past the
	// cutoff, on an AES block boundary. Decryptors honor the size prefix
	// inside the stream, so the extra ciphertext is ignored.
	pkg = padEncryptedPackage(pkg)
	return writeCFBLayout(info, pkg)
}

func padEncryptedPackage(pkg []byte) []byte {
	n := len(pkg)
	if n < 4096 {
		n = 4096
	}
	if n >= 8 && (n-8)%aesBlock != 0 {
		n += aesBlock - (n-8)%aesBlock
	}
	if n == len(pkg) {
		return pkg
	}
	out := make([]byte, n)
	copy(out, pkg)
	return out
}

const aesBlock = 16

func writeCFBLayout(info, pkg []byte) ([]byte, error) {
	const sectorSize = 512
	const miniSize = 64
	miniCount := (len(info) + miniSize - 1) / miniSize
	if miniCount == 0 {
		miniCount = 1
	}
	miniStreamLen := miniCount * miniSize
	miniStreamSectors := (miniStreamLen + sectorSize - 1) / sectorSize
	pkgSectors := (len(pkg) + sectorSize - 1) / sectorSize
	if pkgSectors == 0 {
		pkgSectors = 1
	}

	// Layout: FAT..., directory, mini FAT, mini stream, package.
	nFAT := 1
	for {
		total := nFAT + 1 + 1 + miniStreamSectors + pkgSectors
		need := (total + 127) / 128
		if need <= nFAT {
			break
		}
		nFAT = need
		if nFAT > 109 {
			return nil, errors.New("encrypted package needs a DIFAT chain this writer does not emit")
		}
	}
	dirAt := nFAT
	miniFATAt := nFAT + 1
	miniAt := miniFATAt + 1
	pkgAt := miniAt + miniStreamSectors
	total := pkgAt + pkgSectors

	fat := make([]uint32, nFAT*128)
	for i := range fat {
		fat[i] = cfbFreeSect
	}
	for i := 0; i < nFAT; i++ {
		fat[i] = cfbFatSect
	}
	fat[dirAt] = cfbEndChain
	fat[miniFATAt] = cfbEndChain
	for i := 0; i < miniStreamSectors; i++ {
		fat[miniAt+i] = uint32(miniAt + i + 1)
	}
	fat[miniAt+miniStreamSectors-1] = cfbEndChain
	for i := 0; i < pkgSectors; i++ {
		fat[pkgAt+i] = uint32(pkgAt + i + 1)
	}
	fat[pkgAt+pkgSectors-1] = cfbEndChain

	miniFAT := make([]uint32, 128)
	for i := range miniFAT {
		miniFAT[i] = cfbFreeSect
	}
	for i := 0; i < miniCount; i++ {
		miniFAT[i] = uint32(i + 1)
	}
	miniFAT[miniCount-1] = cfbEndChain

	out := make([]byte, (total+1)*sectorSize)
	copy(out[:8], oleMagic)
	binary.LittleEndian.PutUint16(out[0x18:], 0x003E)
	binary.LittleEndian.PutUint16(out[0x1A:], 3) // major, 512-byte sectors
	binary.LittleEndian.PutUint16(out[0x1C:], 0xFFFE)
	binary.LittleEndian.PutUint16(out[0x1E:], 9)
	binary.LittleEndian.PutUint16(out[0x20:], 6)
	binary.LittleEndian.PutUint32(out[0x2C:], uint32(nFAT))
	binary.LittleEndian.PutUint32(out[0x30:], uint32(dirAt))
	binary.LittleEndian.PutUint32(out[0x38:], 4096)
	binary.LittleEndian.PutUint32(out[0x3C:], uint32(miniFATAt))
	binary.LittleEndian.PutUint32(out[0x40:], 1)
	binary.LittleEndian.PutUint32(out[0x44:], cfbEndChain)
	for i := 0; i < 109; i++ {
		v := uint32(cfbFreeSect)
		if i < nFAT {
			v = uint32(i)
		}
		binary.LittleEndian.PutUint32(out[0x4C+i*4:], v)
	}
	for i := 0; i < nFAT; i++ {
		sec := out[(i+1)*sectorSize : (i+2)*sectorSize]
		for j := 0; j < 128; j++ {
			binary.LittleEndian.PutUint32(sec[j*4:], fat[i*128+j])
		}
	}
	dirSec := out[(dirAt+1)*sectorSize : (dirAt+2)*sectorSize]
	for i := 0; i < 4; i++ {
		for j := 0; j < 3; j++ {
			binary.LittleEndian.PutUint32(dirSec[i*128+68+j*4:], cfbFreeSect)
		}
	}
	// Root child is EncryptionInfo; EncryptedPackage hangs on its left
	// (EncryptedPackage < EncryptionInfo in the CFB name order).
	writeDir(dirSec[0:128], "Root Entry", cfbDirRoot, 1, cfbFreeSect, cfbFreeSect, 1, uint32(miniAt), uint64(miniStreamLen))
	writeDir(dirSec[128:256], "EncryptionInfo", cfbDirStream, 1, 2, cfbFreeSect, cfbFreeSect, 0, uint64(len(info)))
	writeDir(dirSec[256:384], "EncryptedPackage", cfbDirStream, 0, cfbFreeSect, cfbFreeSect, cfbFreeSect, uint32(pkgAt), uint64(len(pkg)))

	miniSec := out[(miniFATAt+1)*sectorSize : (miniFATAt+2)*sectorSize]
	for i, v := range miniFAT {
		binary.LittleEndian.PutUint32(miniSec[i*4:], v)
	}
	copy(out[(miniAt+1)*sectorSize:], append(append([]byte{}, info...), make([]byte, miniStreamLen-len(info))...))
	copy(out[(pkgAt+1)*sectorSize:], pkg)
	return out, nil
}

func writeDir(b []byte, name string, typ, color byte, left, right, child, start uint32, size uint64) {
	u := utf16.Encode([]rune(name))
	for i, v := range u {
		binary.LittleEndian.PutUint16(b[i*2:], v)
	}
	binary.LittleEndian.PutUint16(b[64:], uint16(len(u)*2+2))
	b[66] = typ
	b[67] = color
	binary.LittleEndian.PutUint32(b[68:], left)
	binary.LittleEndian.PutUint32(b[72:], right)
	binary.LittleEndian.PutUint32(b[76:], child)
	binary.LittleEndian.PutUint32(b[116:], start)
	binary.LittleEndian.PutUint64(b[120:], size)
}
