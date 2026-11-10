package gospreadsheet

import (
	"bytes"
	"crypto/aes"
	"crypto/cipher"
	"crypto/hmac"
	"crypto/md5"
	"crypto/rand"
	"crypto/rc4"
	"crypto/sha1"
	"crypto/sha256"
	"crypto/sha512"
	"encoding/base64"
	"encoding/binary"
	"encoding/xml"
	"errors"
	"fmt"
	"hash"
	"io"
	"unicode/utf16"
)

// ErrPasswordRequired is returned when an OLE-encrypted workbook is opened
// without a password.
var ErrPasswordRequired = errors.New("gospreadsheet: workbook is encrypted; a password is required")

// ErrWrongPassword is returned when a password does not unlock an encrypted workbook.
var ErrWrongPassword = errors.New("gospreadsheet: incorrect password")

// Agile block keys from MS-OFFCRYPTO 2.3.4.13 / 2.3.4.14.
var (
	blkVerifierInput = []byte{0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79}
	blkVerifierHash  = []byte{0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E}
	blkKeyValue      = []byte{0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6}
	blkIntegrity1    = []byte{0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6}
	blkIntegrity2    = []byte{0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33}
)

// encryptStandardPackage writes Office 2007 standard encryption (AES-128, SHA-1).
// Excel 2010 and later prefer Agile; this exists so both schemes can be read
// and so writers can be checked against msoffcrypto-tool.
func encryptStandardPackage(plain []byte, password string) ([]byte, error) {
	salt := make([]byte, 16)
	verifier := make([]byte, 16)
	if _, err := io.ReadFull(rand.Reader, salt); err != nil {
		return nil, err
	}
	if _, err := io.ReadFull(rand.Reader, verifier); err != nil {
		return nil, err
	}
	key := standardKey(password, salt, 128)
	encVerifier, err := aesECB(key, verifier, true)
	if err != nil {
		return nil, err
	}
	sum := sha1.Sum(verifier)
	hashBuf := make([]byte, 32)
	copy(hashBuf, sum[:])
	encHash, err := aesECB(key, hashBuf, true)
	if err != nil {
		return nil, err
	}
	csp := append(utf16LE("Microsoft Enhanced RSA and AES Cryptographic Provider"), 0, 0)
	header := make([]byte, 32+len(csp))
	binary.LittleEndian.PutUint32(header[0:], 0x24)    // flags
	binary.LittleEndian.PutUint32(header[8:], 0x660E)  // AES-128
	binary.LittleEndian.PutUint32(header[12:], 0x8004) // SHA-1
	binary.LittleEndian.PutUint32(header[16:], 128)    // key size
	binary.LittleEndian.PutUint32(header[20:], 0x18)   // provider type
	copy(header[32:], csp)
	ver := make([]byte, 4+16+16+4+32)
	binary.LittleEndian.PutUint32(ver[0:], 16)
	copy(ver[4:], salt)
	copy(ver[20:], encVerifier)
	binary.LittleEndian.PutUint32(ver[36:], 20)
	copy(ver[40:], encHash)
	info := make([]byte, 4+4+4+len(header)+len(ver))
	binary.LittleEndian.PutUint16(info[0:], 4)
	binary.LittleEndian.PutUint16(info[2:], 2)
	binary.LittleEndian.PutUint32(info[4:], 0x24) // EncryptionHeaderFlags
	binary.LittleEndian.PutUint32(info[8:], uint32(len(header)))
	copy(info[12:], header)
	copy(info[12+len(header):], ver)

	body, err := aesECB(key, plain, true)
	if err != nil {
		return nil, err
	}
	pkg := make([]byte, 8+len(body))
	binary.LittleEndian.PutUint64(pkg[:8], uint64(len(plain)))
	copy(pkg[8:], body)
	return writeCFBPackage(info, pkg)
}

// DecryptPackage decrypts an ECMA-376 encrypted OOXML package.
// Agile (Office 2010+) and Standard (Office 2007, AES) encryption are supported.
func DecryptPackage(data []byte, password string) ([]byte, error) {
	if !isOLE(data) {
		return nil, errors.New("gospreadsheet: not an encrypted OOXML package")
	}
	if password == "" {
		return nil, ErrPasswordRequired
	}
	f, err := openCFB(data)
	if err != nil {
		return nil, err
	}
	info, err := f.readNamedStream(data, "EncryptionInfo")
	if err != nil {
		return nil, err
	}
	pkg, err := f.readNamedStream(data, "EncryptedPackage")
	if err != nil {
		return nil, err
	}
	plain, err := decryptEncryptedPackage(info, pkg, password)
	if err != nil {
		return nil, err
	}
	if len(plain) < 4 || plain[0] != 'P' || plain[1] != 'K' {
		return nil, ErrWrongPassword
	}
	return plain, nil
}

// EncryptPackage encrypts an OOXML package with ECMA-376 Agile encryption
// (AES-256-CBC, SHA-512, 100000 spins), the scheme Excel 2010 and later write.
func EncryptPackage(data []byte, password string) ([]byte, error) {
	if password == "" {
		return nil, errors.New("gospreadsheet: encryption requires a password")
	}
	info, pkg, err := agileEncrypt(data, password)
	if err != nil {
		return nil, err
	}
	return writeCFBPackage(info, pkg)
}

func decryptEncryptedPackage(info, pkg []byte, password string) ([]byte, error) {
	if len(info) < 8 {
		return nil, errors.New("gospreadsheet: EncryptionInfo is truncated")
	}
	major := binary.LittleEndian.Uint16(info[0:])
	minor := binary.LittleEndian.Uint16(info[2:])
	switch {
	case major == 4 && minor == 4:
		return decryptAgile(info[8:], pkg, password)
	case major == 1 && minor == 1:
		return decryptBinaryRC4(info, pkg, password)
	case (major == 2 || major == 3 || major == 4) && minor == 2:
		return decryptVersion22(info[4:], pkg, password)
	case (major == 3 || major == 4) && minor == 3:
		return nil, errors.New("gospreadsheet: extensible encryption is not supported: the cipher is supplied by an external provider and is not defined by ECMA-376")
	default:
		return nil, fmt.Errorf("gospreadsheet: unsupported EncryptionInfo version %d.%d", major, minor)
	}
}

// decryptVersion22 handles Office encryption version *.2. The AES flag selects
// ECMA-376 standard encryption; otherwise the cipher is RC4 CryptoAPI.
func decryptVersion22(rest, pkg []byte, password string) ([]byte, error) {
	if len(rest) < 8 {
		return nil, errors.New("gospreadsheet: encryption header is truncated")
	}
	flags := binary.LittleEndian.Uint32(rest[0:])
	if flags&0x10 != 0 {
		return nil, errors.New("gospreadsheet: extensible encryption (fExternal) is not supported: the cipher is supplied by an external provider")
	}
	if flags&0x20 != 0 {
		return decryptStandard(rest, pkg, password)
	}
	return decryptCryptoAPI(rest, pkg, password)
}

type agileInfo struct {
	keySalt   []byte
	keyHash   string
	keyBlock  int
	keyBits   int
	spin      int
	passSalt  []byte
	passHash  string
	passBits  int
	encInput  []byte
	encHash   []byte
	encKey    []byte
	hmacKey   []byte
	hmacValue []byte
}

func decryptAgile(xmlBytes, pkg []byte, password string) ([]byte, error) {
	info, err := parseAgile(xmlBytes)
	if err != nil {
		return nil, err
	}
	if info.keyBlock == 0 {
		info.keyBlock = 16
	}
	h := hashByName(info.passHash)
	if h == nil {
		return nil, fmt.Errorf("gospreadsheet: unsupported hash %s", info.passHash)
	}
	iter := iteratePassword(h, info.passSalt, password, info.spin)
	keyBytes := info.passBits / 8
	if !validAESLen(keyBytes) {
		return nil, fmt.Errorf("gospreadsheet: unsupported key length %d", info.passBits)
	}
	key1 := deriveAgileKey(h, iter, blkVerifierInput, keyBytes)
	key2 := deriveAgileKey(h, iter, blkVerifierHash, keyBytes)
	iv := normalizeKey(info.passSalt, info.keyBlock)
	hashInput, err := aesCBC(key1, iv, info.encInput, false)
	if err != nil {
		return nil, ErrWrongPassword
	}
	hh := hashByName(info.passHash)
	hh.Write(hashInput)
	actual := hh.Sum(nil)
	expectedRaw, err := aesCBC(key2, iv, info.encHash, false)
	if err != nil {
		return nil, ErrWrongPassword
	}
	hashSize := hh.Size()
	if len(expectedRaw) < hashSize || !hmac.Equal(actual, expectedRaw[:hashSize]) {
		return nil, ErrWrongPassword
	}
	key3 := deriveAgileKey(h, iter, blkKeyValue, keyBytes)
	secret, err := aesCBC(key3, iv, info.encKey, false)
	if err != nil {
		return nil, err
	}
	pkgBits := info.keyBits
	if pkgBits == 0 {
		pkgBits = info.passBits
	}
	secret = normalizeKey(secret, pkgBits/8)
	if !validAESLen(len(secret)) {
		return nil, errors.New("gospreadsheet: encrypted key has an unsupported length")
	}
	plain, err := agileDecryptPayload(secret, info.keySalt, info.keyHash, pkg)
	if err != nil {
		return nil, err
	}
	return plain, nil
}

func parseAgile(data []byte) (agileInfo, error) {
	var info agileInfo
	dec := xml.NewDecoder(bytes.NewReader(data))
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return info, fmt.Errorf("gospreadsheet: parsing agile EncryptionInfo: %w", err)
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "keyData":
			info.keySalt = b64attr(se, "saltValue")
			info.keyHash = attrLocal(se, "hashAlgorithm")
			info.keyBlock = atoiAttr(se, "blockSize")
			info.keyBits = atoiAttr(se, "keyBits")
		case "dataIntegrity":
			info.hmacKey = b64attr(se, "encryptedHmacKey")
			info.hmacValue = b64attr(se, "encryptedHmacValue")
		case "encryptedKey":
			info.spin = atoiAttr(se, "spinCount")
			info.passSalt = b64attr(se, "saltValue")
			info.passHash = attrLocal(se, "hashAlgorithm")
			info.passBits = atoiAttr(se, "keyBits")
			info.encInput = b64attr(se, "encryptedVerifierHashInput")
			info.encHash = b64attr(se, "encryptedVerifierHashValue")
			info.encKey = b64attr(se, "encryptedKeyValue")
		}
	}
	if len(info.passSalt) == 0 || len(info.encKey) == 0 || len(info.keySalt) == 0 {
		return info, errors.New("gospreadsheet: agile EncryptionInfo is missing key data")
	}
	if info.spin < 0 || info.spin > 10_000_000 {
		return info, errors.New("gospreadsheet: agile spinCount is out of range")
	}
	return info, nil
}

func agileDecryptPayload(key, salt []byte, hashName string, pkg []byte) ([]byte, error) {
	if len(pkg) < 8 {
		return nil, errors.New("gospreadsheet: encrypted package is truncated")
	}
	total := binary.LittleEndian.Uint64(pkg[:8])
	rest := pkg[8:]
	h := hashByName(hashName)
	if h == nil {
		return nil, fmt.Errorf("gospreadsheet: unsupported hash %s", hashName)
	}
	blockSize := aes.BlockSize
	var out []byte
	remaining := total
	for i := 0; len(rest) > 0 && remaining > 0; i++ {
		n := 4096
		if n > len(rest) {
			n = len(rest)
		}
		if n%blockSize != 0 {
			return nil, errors.New("gospreadsheet: encrypted segment is not block aligned")
		}
		h.Reset()
		h.Write(salt)
		var idx [4]byte
		binary.LittleEndian.PutUint32(idx[:], uint32(i))
		h.Write(idx[:])
		iv := h.Sum(nil)
		if len(iv) < blockSize {
			iv = normalizeKey(iv, blockSize)
		} else {
			iv = iv[:blockSize]
		}
		plain, err := aesCBC(key, iv, rest[:n], false)
		if err != nil {
			return nil, err
		}
		if uint64(len(plain)) > remaining {
			plain = plain[:remaining]
		}
		out = append(out, plain...)
		remaining -= uint64(len(plain))
		rest = rest[n:]
	}
	return out, nil
}

func agileEncrypt(plain []byte, password string) (infoXML, encrypted []byte, err error) {
	const (
		spin     = 100000
		keyBits  = 256
		hashName = "SHA512"
		saltSize = 16
		hashSize = 64
	)
	passSalt := make([]byte, saltSize)
	keySalt := make([]byte, saltSize)
	secret := make([]byte, keyBits/8)
	verifier := make([]byte, saltSize)
	if _, err = io.ReadFull(rand.Reader, passSalt); err != nil {
		return nil, nil, err
	}
	if _, err = io.ReadFull(rand.Reader, keySalt); err != nil {
		return nil, nil, err
	}
	if _, err = io.ReadFull(rand.Reader, secret); err != nil {
		return nil, nil, err
	}
	if _, err = io.ReadFull(rand.Reader, verifier); err != nil {
		return nil, nil, err
	}
	h := sha512.New()
	iter := iteratePassword(h, passSalt, password, spin)
	keyBytes := keyBits / 8
	key1 := deriveAgileKey(h, iter, blkVerifierInput, keyBytes)
	key2 := deriveAgileKey(h, iter, blkVerifierHash, keyBytes)
	key3 := deriveAgileKey(h, iter, blkKeyValue, keyBytes)
	encInput, err := aesCBC(key1, passSalt, verifier, true)
	if err != nil {
		return nil, nil, err
	}
	h.Reset()
	h.Write(verifier)
	verifierHash := h.Sum(nil)
	encHash, err := aesCBC(key2, passSalt, verifierHash, true)
	if err != nil {
		return nil, nil, err
	}
	encSecret, err := aesCBC(key3, passSalt, secret, true)
	if err != nil {
		return nil, nil, err
	}
	payload, err := agileEncryptPayload(secret, keySalt, hashName, plain)
	if err != nil {
		return nil, nil, err
	}
	hmacRand := make([]byte, hashSize)
	if _, err = io.ReadFull(rand.Reader, hmacRand); err != nil {
		return nil, nil, err
	}
	iv1 := agileIV(h, keySalt, blkIntegrity1, saltSize)
	iv2 := agileIV(h, keySalt, blkIntegrity2, saltSize)
	encHmacKey, err := aesCBC(secret, iv1, hmacRand, true)
	if err != nil {
		return nil, nil, err
	}
	mac := hmac.New(sha512.New, hmacRand)
	mac.Write(payload)
	encHmacVal, err := aesCBC(secret, iv2, mac.Sum(nil), true)
	if err != nil {
		return nil, nil, err
	}
	xmlDoc := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+"\n"+
		`<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">`+
		`<keyData saltSize="%d" blockSize="16" keyBits="%d" hashSize="%d" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="%s" saltValue="%s"/>`+
		`<dataIntegrity encryptedHmacKey="%s" encryptedHmacValue="%s"/>`+
		`<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">`+
		`<p:encryptedKey spinCount="%d" saltSize="%d" blockSize="16" keyBits="%d" hashSize="%d" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="%s" saltValue="%s" encryptedVerifierHashInput="%s" encryptedVerifierHashValue="%s" encryptedKeyValue="%s"/>`+
		`</keyEncryptor></keyEncryptors></encryption>`,
		saltSize, keyBits, hashSize, hashName, b64(keySalt),
		b64(encHmacKey), b64(encHmacVal),
		spin, saltSize, keyBits, hashSize, hashName, b64(passSalt), b64(encInput), b64(encHash), b64(encSecret),
	)
	header := make([]byte, 8+len(xmlDoc))
	binary.LittleEndian.PutUint16(header[0:], 4)
	binary.LittleEndian.PutUint16(header[2:], 4)
	binary.LittleEndian.PutUint32(header[4:], 0x40)
	copy(header[8:], xmlDoc)
	return header, payload, nil
}

func agileEncryptPayload(key, salt []byte, hashName string, plain []byte) ([]byte, error) {
	h := hashByName(hashName)
	var body []byte
	for i := 0; i*4096 < len(plain) || (i == 0 && len(plain) == 0); i++ {
		start := i * 4096
		if start > len(plain) {
			break
		}
		end := start + 4096
		if end > len(plain) {
			end = len(plain)
		}
		chunk := plain[start:end]
		iv := agileIV(h, salt, uint32Bytes(uint32(i)), aes.BlockSize)
		enc, err := aesCBC(key, iv, chunk, true)
		if err != nil {
			return nil, err
		}
		body = append(body, enc...)
		if end == len(plain) {
			break
		}
	}
	out := make([]byte, 8+len(body))
	binary.LittleEndian.PutUint64(out[:8], uint64(len(plain)))
	copy(out[8:], body)
	return out, nil
}

func agileIV(h hash.Hash, salt, blockKey []byte, blockSize int) []byte {
	h.Reset()
	h.Write(salt)
	h.Write(blockKey)
	return normalizeKey(h.Sum(nil), blockSize)
}

func uint32Bytes(v uint32) []byte {
	var b [4]byte
	binary.LittleEndian.PutUint32(b[:], v)
	return b[:]
}

func decryptBinaryRC4(info, pkg []byte, password string) ([]byte, error) {
	if len(info) < 4+48 {
		return nil, errors.New("gospreadsheet: binary RC4 EncryptionInfo is truncated")
	}
	salt := info[4:20]
	encVerifier := info[20:36]
	encHash := info[36:52]
	password = clipPassword(password)
	key := binaryRC4BlockKey(password, salt, 0)
	c, err := rc4.NewCipher(key)
	if err != nil {
		return nil, err
	}
	verifier := make([]byte, 16)
	gotHash := make([]byte, 16)
	c.XORKeyStream(verifier, encVerifier)
	c.XORKeyStream(gotHash, encHash)
	sum := md5.Sum(verifier)
	if !hmac.Equal(sum[:], gotHash) {
		return nil, ErrWrongPassword
	}
	plain, err := decryptRC4Chunks(pkg, func(block int) []byte {
		return binaryRC4BlockKey(password, salt, block)
	})
	if err != nil {
		return nil, err
	}
	return plain, nil
}

func binaryRC4BlockKey(password string, salt []byte, block int) []byte {
	h := md5.New()
	h.Write(utf16LE(password))
	trunc := h.Sum(nil)[:5]
	h.Reset()
	for i := 0; i < 16; i++ {
		h.Write(trunc)
		h.Write(salt)
	}
	trunc = h.Sum(nil)[:5]
	h.Reset()
	h.Write(trunc)
	var b [4]byte
	binary.LittleEndian.PutUint32(b[:], uint32(block))
	h.Write(b[:])
	return h.Sum(nil)[:16]
}

func decryptCryptoAPI(rest, pkg []byte, password string) ([]byte, error) {
	if len(rest) < 8 {
		return nil, errors.New("gospreadsheet: RC4 CryptoAPI EncryptionInfo is truncated")
	}
	headerSize := binary.LittleEndian.Uint32(rest[4:])
	if uint32(len(rest)) < 8+headerSize+40 {
		return nil, errors.New("gospreadsheet: RC4 CryptoAPI EncryptionInfo is truncated")
	}
	header := rest[8 : 8+headerSize]
	verifier := rest[8+headerSize:]
	if len(header) < 20 || len(verifier) < 40 {
		return nil, errors.New("gospreadsheet: RC4 CryptoAPI EncryptionInfo is truncated")
	}
	keySize := binary.LittleEndian.Uint32(header[16:])
	if keySize == 0 {
		keySize = 40
	}
	saltSize := binary.LittleEndian.Uint32(verifier[0:])
	if saltSize != 16 || len(verifier) < int(4+saltSize+16+4+20) {
		return nil, errors.New("gospreadsheet: RC4 CryptoAPI verifier is truncated")
	}
	salt := append([]byte(nil), verifier[4:4+saltSize]...)
	encVerifier := verifier[4+saltSize : 4+saltSize+16]
	hashSize := binary.LittleEndian.Uint32(verifier[4+saltSize+16:])
	hashOff := 4 + saltSize + 16 + 4
	if hashSize == 0 || int(hashOff+hashSize) > len(verifier) {
		hashSize = 20
	}
	if int(hashOff+hashSize) > len(verifier) {
		return nil, errors.New("gospreadsheet: RC4 CryptoAPI verifier is truncated")
	}
	encHash := verifier[hashOff : hashOff+hashSize]
	password = clipPassword(password)
	key := cryptoAPIBlockKey(password, salt, int(keySize), 0)
	c, err := rc4.NewCipher(key)
	if err != nil {
		return nil, err
	}
	verifierPlain := make([]byte, len(encVerifier))
	hashPlain := make([]byte, len(encHash))
	c.XORKeyStream(verifierPlain, encVerifier)
	c.XORKeyStream(hashPlain, encHash)
	sum := sha1.Sum(verifierPlain)
	n := len(hashPlain)
	if n > sha1.Size {
		n = sha1.Size
	}
	if !hmac.Equal(sum[:n], hashPlain[:n]) {
		return nil, ErrWrongPassword
	}
	return decryptRC4Chunks(pkg, func(block int) []byte {
		return cryptoAPIBlockKey(password, salt, int(keySize), block)
	})
}

func cryptoAPIBlockKey(password string, salt []byte, keyBits, block int) []byte {
	h := sha1.New()
	h.Write(salt)
	h.Write(utf16LE(password))
	secret := h.Sum(nil)
	h.Reset()
	h.Write(secret)
	var b [4]byte
	binary.LittleEndian.PutUint32(b[:], uint32(block))
	h.Write(b[:])
	sum := h.Sum(nil)
	n := keyBits / 8
	if n <= 0 || n > len(sum) {
		n = 5
	}
	if keyBits == 40 {
		key := make([]byte, 16)
		copy(key, sum[:5])
		return key
	}
	return append([]byte(nil), sum[:n]...)
}

// decryptRC4Chunks decrypts an EncryptedPackage. ECMA-376 and POI's binary
// RC4 writer prefix the ciphertext with the plaintext size. Some CryptoAPI
// writers omit that prefix; both layouts are accepted.
func decryptRC4Chunks(pkg []byte, keyFn func(block int) []byte) ([]byte, error) {
	if len(pkg) >= 8 {
		total := binary.LittleEndian.Uint64(pkg[:8])
		if total > 0 && total < 1<<32 && int(total) <= len(pkg)-8 {
			out, err := rc4Chunk(pkg[8:], keyFn, int(total))
			if err == nil && looksLikeZip(out) {
				return out, nil
			}
		}
	}
	out, err := rc4Chunk(pkg, keyFn, len(pkg))
	if err != nil {
		return nil, err
	}
	if !looksLikeZip(out) {
		return nil, ErrWrongPassword
	}
	return out, nil
}

func rc4Chunk(cipherText []byte, keyFn func(block int) []byte, total int) ([]byte, error) {
	if total > len(cipherText) {
		total = len(cipherText)
	}
	out := make([]byte, 0, total)
	for block, off := 0, 0; off < total; block++ {
		n := 512
		if n > total-off {
			n = total - off
		}
		c, err := rc4.NewCipher(keyFn(block))
		if err != nil {
			return nil, err
		}
		chunk := make([]byte, n)
		c.XORKeyStream(chunk, cipherText[off:off+n])
		out = append(out, chunk...)
		off += n
	}
	return out, nil
}

func looksLikeZip(b []byte) bool {
	return len(b) >= 4 && b[0] == 'P' && b[1] == 'K'
}

func clipPassword(password string) string {
	r := []rune(password)
	if len(r) > 255 {
		return string(r[:255])
	}
	return password
}

func encryptBinaryRC4Package(plain []byte, password string) ([]byte, error) {
	salt := make([]byte, 16)
	verifier := make([]byte, 16)
	if _, err := io.ReadFull(rand.Reader, salt); err != nil {
		return nil, err
	}
	if _, err := io.ReadFull(rand.Reader, verifier); err != nil {
		return nil, err
	}
	password = clipPassword(password)
	key := binaryRC4BlockKey(password, salt, 0)
	c, err := rc4.NewCipher(key)
	if err != nil {
		return nil, err
	}
	encVerifier := make([]byte, 16)
	c.XORKeyStream(encVerifier, verifier)
	sum := md5.Sum(verifier)
	encHash := make([]byte, 16)
	c.XORKeyStream(encHash, sum[:])
	info := make([]byte, 4+48)
	binary.LittleEndian.PutUint16(info[0:], 1)
	binary.LittleEndian.PutUint16(info[2:], 1)
	copy(info[4:], salt)
	copy(info[20:], encVerifier)
	copy(info[36:], encHash)
	body := rc4EncryptChunks(plain, func(block int) []byte {
		return binaryRC4BlockKey(password, salt, block)
	})
	return writeCFBPackage(info, body)
}

func encryptCryptoAPIPackage(plain []byte, password string) ([]byte, error) {
	const keyBits = 40
	salt := make([]byte, 16)
	verifier := make([]byte, 16)
	if _, err := io.ReadFull(rand.Reader, salt); err != nil {
		return nil, err
	}
	if _, err := io.ReadFull(rand.Reader, verifier); err != nil {
		return nil, err
	}
	password = clipPassword(password)
	key := cryptoAPIBlockKey(password, salt, keyBits, 0)
	c, err := rc4.NewCipher(key)
	if err != nil {
		return nil, err
	}
	encVerifier := make([]byte, 16)
	c.XORKeyStream(encVerifier, verifier)
	sum := sha1.Sum(verifier)
	encHash := make([]byte, sha1.Size)
	c.XORKeyStream(encHash, sum[:])
	csp := append(utf16LE("Microsoft Enhanced Cryptographic Provider v1.0"), 0, 0)
	header := make([]byte, 32+len(csp))
	binary.LittleEndian.PutUint32(header[0:], 0x4)
	binary.LittleEndian.PutUint32(header[8:], 0x6801)
	binary.LittleEndian.PutUint32(header[12:], 0x8004)
	binary.LittleEndian.PutUint32(header[16:], keyBits)
	binary.LittleEndian.PutUint32(header[20:], 1)
	copy(header[32:], csp)
	ver := make([]byte, 4+16+16+4+sha1.Size)
	binary.LittleEndian.PutUint32(ver[0:], 16)
	copy(ver[4:], salt)
	copy(ver[20:], encVerifier)
	binary.LittleEndian.PutUint32(ver[36:], sha1.Size)
	copy(ver[40:], encHash)
	info := make([]byte, 4+4+4+len(header)+len(ver))
	binary.LittleEndian.PutUint16(info[0:], 4)
	binary.LittleEndian.PutUint16(info[2:], 2)
	binary.LittleEndian.PutUint32(info[4:], 0x4)
	binary.LittleEndian.PutUint32(info[8:], uint32(len(header)))
	copy(info[12:], header)
	copy(info[12+len(header):], ver)
	body := rc4EncryptChunks(plain, func(block int) []byte {
		return cryptoAPIBlockKey(password, salt, keyBits, block)
	})
	return writeCFBPackage(info, body)
}

func rc4EncryptChunks(plain []byte, keyFn func(block int) []byte) []byte {
	body := make([]byte, len(plain))
	for block, off := 0, 0; off < len(plain); block++ {
		n := 512
		if n > len(plain)-off {
			n = len(plain) - off
		}
		c, err := rc4.NewCipher(keyFn(block))
		if err != nil {
			return nil
		}
		c.XORKeyStream(body[off:off+n], plain[off:off+n])
		off += n
	}
	out := make([]byte, 8+len(body))
	binary.LittleEndian.PutUint64(out[:8], uint64(len(plain)))
	copy(out[8:], body)
	return out
}

func decryptStandard(rest, pkg []byte, password string) ([]byte, error) {
	if len(rest) < 8 {
		return nil, errors.New("gospreadsheet: standard EncryptionInfo is truncated")
	}
	// rest begins at Flags, which sits in front of EncryptionHeaderSize.
	headerSize := binary.LittleEndian.Uint32(rest[4:])
	if int(8+headerSize) > len(rest) {
		return nil, errors.New("gospreadsheet: standard encryption header is truncated")
	}
	header := rest[8 : 8+headerSize]
	verifier := rest[8+headerSize:]
	if len(header) < 32 || len(verifier) < 16+16+4+16 {
		return nil, errors.New("gospreadsheet: standard encryption header is truncated")
	}
	algID := binary.LittleEndian.Uint32(header[8:])
	keySize := binary.LittleEndian.Uint32(header[16:])
	switch algID {
	case 0x660E:
		if keySize == 0 {
			keySize = 128
		}
	case 0x660F:
		if keySize == 0 {
			keySize = 192
		}
	case 0x6610:
		if keySize == 0 {
			keySize = 256
		}
	default:
		return nil, fmt.Errorf("gospreadsheet: unsupported standard cipher algId %#x", algID)
	}
	saltSize := binary.LittleEndian.Uint32(verifier[0:])
	if saltSize != 16 || len(verifier) < 12+16+int(saltSize) {
		return nil, errors.New("gospreadsheet: standard verifier is truncated")
	}
	salt := verifier[4 : 4+saltSize]
	encVerifier := verifier[4+saltSize : 4+saltSize+16]
	hashOff := 4 + saltSize + 16 + 4
	encHash := verifier[hashOff:]
	if len(encHash) > 32 {
		encHash = encHash[:32]
	}
	key := standardKey(password, salt, int(keySize))
	verPlain, err := aesECB(key, encVerifier, false)
	if err != nil {
		return nil, ErrWrongPassword
	}
	sum := sha1.Sum(verPlain)
	gotHash, err := aesECB(key, encHash, false)
	if err != nil || len(gotHash) < sha1.Size || !hmac.Equal(sum[:], gotHash[:sha1.Size]) {
		return nil, ErrWrongPassword
	}
	if len(pkg) < 8 {
		return nil, errors.New("gospreadsheet: encrypted package is truncated")
	}
	total := binary.LittleEndian.Uint64(pkg[:8])
	body, err := aesECB(key, pkg[8:], false)
	if err != nil {
		return nil, err
	}
	if uint64(len(body)) > total {
		body = body[:total]
	}
	return body, nil
}

// standardKey is the ECMA-376 standard (Office 2007) key derivation:
// 50,000 SHA-1 spins followed by the CryptoAPI 0x36/0x5C expansion.
func standardKey(password string, salt []byte, keySizeBits int) []byte {
	const spins = 50000
	h := sha1.New()
	h.Write(salt)
	h.Write(utf16LE(password))
	sum := h.Sum(nil)
	var ib [4]byte
	for i := 0; i < spins; i++ {
		binary.LittleEndian.PutUint32(ib[:], uint32(i))
		h.Reset()
		h.Write(ib[:])
		h.Write(sum)
		sum = h.Sum(nil)
	}
	h.Reset()
	h.Write(sum)
	binary.LittleEndian.PutUint32(ib[:], 0)
	h.Write(ib[:])
	hfinal := h.Sum(nil)
	buf1 := bytes.Repeat([]byte{0x36}, 64)
	for i := 0; i < len(hfinal) && i < len(buf1); i++ {
		buf1[i] ^= hfinal[i]
	}
	buf2 := bytes.Repeat([]byte{0x5C}, 64)
	for i := 0; i < len(hfinal) && i < len(buf2); i++ {
		buf2[i] ^= hfinal[i]
	}
	x1 := sha1.Sum(buf1)
	x2 := sha1.Sum(buf2)
	derived := append(append([]byte{}, x1[:]...), x2[:]...)
	n := keySizeBits / 8
	if n > len(derived) {
		n = len(derived)
	}
	return derived[:n]
}

func iteratePassword(h hash.Hash, salt []byte, password string, spin int) []byte {
	h.Reset()
	h.Write(salt)
	h.Write(utf16LE(password))
	sum := h.Sum(nil)
	var ib [4]byte
	for i := 0; i < spin; i++ {
		binary.LittleEndian.PutUint32(ib[:], uint32(i))
		h.Reset()
		h.Write(ib[:])
		h.Write(sum)
		sum = h.Sum(nil)
	}
	return sum
}

func deriveAgileKey(h hash.Hash, iterated, blockKey []byte, keyBytes int) []byte {
	h.Reset()
	h.Write(iterated)
	h.Write(blockKey)
	return normalizeKey(h.Sum(nil), keyBytes)
}

func normalizeKey(b []byte, n int) []byte {
	if len(b) == n {
		return b
	}
	if len(b) > n {
		return b[:n]
	}
	out := make([]byte, n)
	copy(out, b)
	for i := len(b); i < n; i++ {
		out[i] = 0x36
	}
	return out
}

func hashByName(name string) hash.Hash {
	switch name {
	case "SHA1", "SHA-1", "sha1":
		return sha1.New()
	case "SHA256", "SHA-256", "sha256":
		return sha256.New()
	case "SHA384", "SHA-384", "sha384":
		return sha512.New384()
	case "SHA512", "SHA-512", "sha512":
		return sha512.New()
	default:
		return nil
	}
}

func utf16LE(s string) []byte {
	u := utf16.Encode([]rune(s))
	b := make([]byte, len(u)*2)
	for i, v := range u {
		binary.LittleEndian.PutUint16(b[i*2:], v)
	}
	return b
}

func validAESLen(n int) bool {
	return n == 16 || n == 24 || n == 32
}

func aesCBC(key, iv, data []byte, encrypt bool) ([]byte, error) {
	block, err := aes.NewCipher(key)
	if err != nil {
		return nil, err
	}
	if len(iv) != block.BlockSize() {
		return nil, errors.New("bad AES IV length")
	}
	buf := data
	if encrypt {
		n := block.BlockSize()
		if len(buf)%n != 0 {
			padded := make([]byte, (len(buf)+n-1)/n*n)
			copy(padded, buf)
			buf = padded
		}
		if len(buf) == 0 {
			buf = make([]byte, n)
		}
	}
	if len(buf)%block.BlockSize() != 0 {
		return nil, errors.New("ciphertext is not a multiple of the AES block size")
	}
	out := make([]byte, len(buf))
	var mode cipher.BlockMode
	if encrypt {
		mode = cipher.NewCBCEncrypter(block, iv)
	} else {
		mode = cipher.NewCBCDecrypter(block, iv)
	}
	mode.CryptBlocks(out, buf)
	return out, nil
}

func aesECB(key, data []byte, encrypt bool) ([]byte, error) {
	block, err := aes.NewCipher(key)
	if err != nil {
		return nil, err
	}
	buf := data
	if encrypt && len(buf)%block.BlockSize() != 0 {
		padded := make([]byte, (len(buf)+block.BlockSize()-1)/block.BlockSize()*block.BlockSize())
		copy(padded, buf)
		buf = padded
	}
	if len(buf)%block.BlockSize() != 0 || len(buf) == 0 {
		return nil, errors.New("ciphertext is not a multiple of the AES block size")
	}
	out := make([]byte, len(buf))
	for i := 0; i < len(buf); i += block.BlockSize() {
		if encrypt {
			block.Encrypt(out[i:], buf[i:])
		} else {
			block.Decrypt(out[i:], buf[i:])
		}
	}
	return out, nil
}

func b64(b []byte) string { return base64.StdEncoding.EncodeToString(b) }

func b64attr(se xml.StartElement, local string) []byte {
	v := attrLocal(se, local)
	if v == "" {
		return nil
	}
	out, err := base64.StdEncoding.DecodeString(v)
	if err != nil {
		return nil
	}
	return out
}

func atoiAttr(se xml.StartElement, local string) int {
	n := 0
	fmt.Sscanf(attrLocal(se, local), "%d", &n)
	return n
}
