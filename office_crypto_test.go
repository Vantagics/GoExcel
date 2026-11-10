package gospreadsheet

import (
	"bytes"
	"crypto/sha512"
	"fmt"
	"testing"
)

func TestAgileVerifierVector(t *testing.T) {
	salt := []byte{0xcb, 0xca, 0x1c, 0x99, 0x93, 'C', 0xfb, 0xad, 0x92, 0x07, 'V', '4', 0x15, 0x00, '4', 0xb0}
	encIn, _ := hexDecode("39eea54e26e514798c284bc7714d38ac")
	encHash, _ := hexDecode("14376d6d817334e6b0ff4fd8221a7c678e5d8a784e8f999f4c188930c36a4b29c5b333605b5cd403b05003adcf18cca8cbab8debe373c65604a0becfae5c0ad0")
	h := sha512.New()
	iter := iteratePassword(h, salt, "Password1234_", 100000)
	key1 := deriveAgileKey(h, iter, blkVerifierInput, 32)
	wantIter := "244a8f097708ba992f489adc323b5e52a2ea8c5589e7b1d37ed0bec67a6ec91b0603689853ac9e3e1572c54f2c733aa0d6192c91c1c78b46a8ca65e2ea66d954"
	wantKey := "d762f55fe88a9f840eba11f8984a03072128e205c8fd68b409f4fd93a98a2c8a"
	if hexEncode(iter) != wantIter {
		t.Fatalf("iter %s", hexEncode(iter))
	}
	if hexEncode(key1) != wantKey {
		t.Fatalf("key1 %s", hexEncode(key1))
	}
	hashInput, err := aesCBC(key1, salt, encIn, false)
	if err != nil {
		t.Fatal(err)
	}
	key2 := deriveAgileKey(h, iter, blkVerifierHash, 32)
	hh := sha512.New()
	hh.Write(hashInput)
	actual := hh.Sum(nil)
	expected, err := aesCBC(key2, salt, encHash, false)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(actual, expected[:sha512.Size]) {
		t.Fatalf("verifier mismatch\nactual %x\nexpect %x", actual, expected)
	}
}

func hexDecode(s string) ([]byte, error) {
	out := make([]byte, len(s)/2)
	for i := 0; i < len(out); i++ {
		fmt.Sscanf(s[i*2:i*2+2], "%02x", &out[i])
	}
	return out, nil
}

func hexEncode(b []byte) string {
	const hexdigits = "0123456789abcdef"
	out := make([]byte, len(b)*2)
	for i, v := range b {
		out[i*2] = hexdigits[v>>4]
		out[i*2+1] = hexdigits[v&0x0f]
	}
	return string(out)
}

func TestStandardKeyVector(t *testing.T) {
	salt := []byte{0xe8, 0x82, 'f', 'I', 0x0c, '[', 0xd1, 0xee, 0xbd, '+', 'C', 0x94, 0xe3, 0xf8, '0', 0xef}
	got := standardKey("Password1234_", salt, 128)
	want := []byte{'@', 0xb1, ':', 'q', 0xf9, 0x0b, 0x96, 'n', '7', 'T', 0x08, 0xf2, 0xd1, 0x81, 0xa1, 0xaa}
	if !bytes.Equal(got, want) {
		t.Fatalf("standard key = %x, want %x", got, want)
	}
}

func TestAgileCryptoWithoutCFB(t *testing.T) {
	plain := []byte("PK\x03\x04agile-payload-not-a-real-zip-but-has-the-magic")
	info, pkg, err := agileEncrypt(plain, "s3cret")
	if err != nil {
		t.Fatal(err)
	}
	got, err := decryptEncryptedPackage(info, pkg, "s3cret")
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(got, plain) {
		t.Fatalf("got %q", got)
	}
}

func TestAgileRoundTripPackage(t *testing.T) {
	plain := []byte("PK\x03\x04agile-payload-not-a-real-zip-but-has-the-magic")
	enc, err := EncryptPackage(plain, "s3cret")
	if err != nil {
		t.Fatal(err)
	}
	if !isOLE(enc) {
		t.Fatal("encrypted package is not an OLE compound file")
	}
	got, err := DecryptPackage(enc, "s3cret")
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(got, plain) {
		t.Fatalf("round trip = %q", got)
	}
	if _, err := DecryptPackage(enc, "nope"); err != ErrWrongPassword {
		t.Fatalf("wrong password error = %v", err)
	}
	if _, err := DecryptPackage(enc, ""); err != ErrPasswordRequired {
		t.Fatalf("empty password error = %v", err)
	}
}

func TestStandardRoundTripPackage(t *testing.T) {
	plain := []byte("PK\x03\x04standard-payload")
	enc, err := encryptStandardPackage(plain, "s3cret")
	if err != nil {
		t.Fatal(err)
	}
	got, err := DecryptPackage(enc, "s3cret")
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(got, plain) {
		t.Fatalf("round trip = %q", got)
	}
}
