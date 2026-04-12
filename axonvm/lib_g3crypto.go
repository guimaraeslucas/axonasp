/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * Attribution Notice:
 * If this software is used in other projects, the name "AxonASP Server"
 * must be cited in the documentation or "About" section.
 *
 * Contribution Policy:
 * Modifications to the core source code of AxonASP Server must be
 * made available under this same license terms.
 */
package axonvm

import (
	"crypto/hmac"
	"crypto/md5"
	"crypto/rand"
	"crypto/sha1"
	"crypto/sha256"
	"crypto/sha512"
	"encoding/base64"
	"encoding/hex"
	"strconv"
	"strings"

	"golang.org/x/crypto/bcrypt"
	"golang.org/x/crypto/blake2b"
	"golang.org/x/crypto/pbkdf2"
	"golang.org/x/crypto/sha3"
)

// G3Crypto stores the runtime state for one native crypto object instance.
type G3Crypto struct {
	defaultAlgorithm string
	lastHash         []byte
	bcryptCost       int
}

// NewG3Crypto creates one crypto object with default bcrypt configuration.
func NewG3Crypto() *G3Crypto {
	return &G3Crypto{bcryptCost: bcrypt.DefaultCost}
}

// NewG3CryptoWithAlgorithm creates one crypto object with a default hash algorithm.
func NewG3CryptoWithAlgorithm(algorithm string) *G3Crypto {
	lib := NewG3Crypto()
	lib.defaultAlgorithm = g3cryptoNormalizeAlgorithm(algorithm)
	return lib
}

// DispatchMethod executes method calls and method-like property behavior.
func (c *G3Crypto) DispatchMethod(methodName string, args []Value) Value {
	method := strings.ToLower(strings.TrimSpace(methodName))

	switch method {
	case "uuid", "guid":
		return NewString(c.uuid())
	case "hashpassword", "hash", "bcrypthash":
		if len(args) < 1 {
			return NewString("")
		}
		return NewString(c.hashPassword(args[0].String()))
	case "verifypassword", "verify", "bcryptverify":
		if len(args) < 2 {
			return NewBool(false)
		}
		return NewBool(c.verifyPassword(args[0].String(), args[1].String()))
	case "setbcryptcost", "configurebcryptcost":
		if len(args) < 1 {
			return NewBool(false)
		}
		return NewBool(c.setBcryptCost(g3cryptoValueToInt(args[0])))
	case "getbcryptcost":
		return NewInteger(int64(c.bcryptCost))
	case "randombytes", "randbytes":
		size := 32
		if len(args) > 0 {
			size = g3cryptoValueToInt(args[0])
		}
		return c.randomBytes(size)
	case "randomhex", "randhex":
		size := 32
		if len(args) > 0 {
			size = g3cryptoValueToInt(args[0])
		}
		return NewString(c.randomHex(size))
	case "randombase64", "randbase64":
		size := 32
		if len(args) > 0 {
			size = g3cryptoValueToInt(args[0])
		}
		return NewString(c.randomBase64(size))

	case "md5":
		return NewString(c.hashHex("md5", args))
	case "sha1":
		return NewString(c.hashHex("sha1", args))
	case "sha256":
		return NewString(c.hashHex("sha256", args))
	case "sha384":
		return NewString(c.hashHex("sha384", args))
	case "sha512":
		return NewString(c.hashHex("sha512", args))
	case "sha3_256", "sha3256":
		return NewString(c.hashHex("sha3_256", args))
	case "sha3_512", "sha3512":
		return NewString(c.hashHex("sha3_512", args))
	case "blake2b256", "blake2_256":
		return NewString(c.hashHex("blake2b256", args))
	case "blake2b512", "blake2_512":
		return NewString(c.hashHex("blake2b512", args))

	case "md5bytes":
		return c.hashBytesAsVBArray("md5", args)
	case "sha1bytes":
		return c.hashBytesAsVBArray("sha1", args)
	case "sha256bytes":
		return c.hashBytesAsVBArray("sha256", args)
	case "sha384bytes":
		return c.hashBytesAsVBArray("sha384", args)
	case "sha512bytes":
		return c.hashBytesAsVBArray("sha512", args)
	case "sha3_256bytes", "sha3256bytes":
		return c.hashBytesAsVBArray("sha3_256", args)
	case "sha3_512bytes", "sha3512bytes":
		return c.hashBytesAsVBArray("sha3_512", args)

	case "hmacsha256":
		if len(args) < 2 {
			return NewString("")
		}
		return NewString(c.hmacHex("sha256", args[0].String(), args[1].String()))
	case "hmacsha512":
		if len(args) < 2 {
			return NewString("")
		}
		return NewString(c.hmacHex("sha512", args[0].String(), args[1].String()))
	case "pbkdf2sha256":
		if len(args) < 2 {
			return NewString("")
		}
		iterations := 100000
		if len(args) > 2 {
			iterations = g3cryptoValueToInt(args[2])
		}
		keyLength := 32
		if len(args) > 3 {
			keyLength = g3cryptoValueToInt(args[3])
		}
		return NewString(c.pbkdf2SHA256(args[0].String(), args[1].String(), iterations, keyLength))
	case "computehash":
		algorithm := c.defaultAlgorithm
		if len(args) > 1 {
			algorithm = args[1].String()
		}
		if algorithm == "" {
			algorithm = "sha256"
		}
		return c.hashBytesAsVBArray(algorithm, args)
	case "initialize", "clear", "dispose":
		c.lastHash = nil
		return Value{Type: VTEmpty}
	}

	if g3cryptoIsLegacyDotNetAlias(method) {
		algorithm := g3cryptoNormalizeAlgorithm(method)
		if algorithm != "" {
			return c.hashBytesAsVBArray(algorithm, args)
		}
	}

	return Value{Type: VTEmpty}
}

// DispatchPropertyGet resolves property reads for one crypto object.
func (c *G3Crypto) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(strings.TrimSpace(propertyName)) {
	case "hash":
		if len(c.lastHash) == 0 {
			return g3cryptoBytesToVBArray(nil)
		}
		return g3cryptoBytesToVBArray(c.lastHash)
	case "hashsize":
		if c.defaultAlgorithm != "" {
			return NewInteger(int64(g3cryptoHashSizeBits(c.defaultAlgorithm)))
		}
		if len(c.lastHash) > 0 {
			return NewInteger(int64(len(c.lastHash) * 8))
		}
		return NewInteger(0)
	case "bcryptcost", "cost":
		return NewInteger(int64(c.bcryptCost))
	case "canreusetransform", "cantransformmultipleblocks":
		return NewBool(true)
	default:
		return Value{Type: VTEmpty}
	}
}

// DispatchPropertySet applies writable properties for one crypto object.
func (c *G3Crypto) DispatchPropertySet(propertyName string, val Value) {
	switch strings.ToLower(strings.TrimSpace(propertyName)) {
	case "bcryptcost", "cost":
		_ = c.setBcryptCost(g3cryptoValueToInt(val))
	}
}

// uuid returns one RFC 4122 version 4 UUID using cryptographic randomness.
func (c *G3Crypto) uuid() string {
	b := make([]byte, 16)
	if _, err := rand.Read(b); err != nil {
		return ""
	}
	b[6] = (b[6] & 0x0f) | 0x40
	b[8] = (b[8] & 0x3f) | 0x80

	out := make([]byte, 36)
	hex.Encode(out[0:8], b[0:4])
	out[8] = '-'
	hex.Encode(out[9:13], b[4:6])
	out[13] = '-'
	hex.Encode(out[14:18], b[6:8])
	out[18] = '-'
	hex.Encode(out[19:23], b[8:10])
	out[23] = '-'
	hex.Encode(out[24:36], b[10:16])

	return string(out)
}

// hashPassword hashes a password with bcrypt using the configured cost.
func (c *G3Crypto) hashPassword(password string) string {
	bytes, err := bcrypt.GenerateFromPassword([]byte(password), c.bcryptCost)
	if err != nil {
		return ""
	}
	return string(bytes)
}

// verifyPassword validates one bcrypt hash against a plain-text password.
func (c *G3Crypto) verifyPassword(password string, hash string) bool {
	return bcrypt.CompareHashAndPassword([]byte(hash), []byte(password)) == nil
}

// setBcryptCost updates bcrypt cost when it is inside the accepted range.
func (c *G3Crypto) setBcryptCost(cost int) bool {
	if cost < bcrypt.MinCost || cost > bcrypt.MaxCost {
		return false
	}
	c.bcryptCost = cost
	return true
}

// randomBytes returns a VBScript byte array with secure random content.
func (c *G3Crypto) randomBytes(size int) Value {
	if size < 0 {
		size = 0
	}
	buf := make([]byte, size)
	if _, err := rand.Read(buf); err != nil {
		return g3cryptoBytesToVBArray(nil)
	}
	return g3cryptoBytesToVBArray(buf)
}

// randomHex returns random bytes encoded as lowercase hexadecimal.
func (c *G3Crypto) randomHex(size int) string {
	if size < 0 {
		size = 0
	}
	buf := make([]byte, size)
	if _, err := rand.Read(buf); err != nil {
		return ""
	}
	return hex.EncodeToString(buf)
}

// randomBase64 returns random bytes encoded in base64.
func (c *G3Crypto) randomBase64(size int) string {
	if size < 0 {
		size = 0
	}
	buf := make([]byte, size)
	if _, err := rand.Read(buf); err != nil {
		return ""
	}
	return base64.StdEncoding.EncodeToString(buf)
}

// hashHex computes one hash and returns the output encoded as hexadecimal.
func (c *G3Crypto) hashHex(algorithm string, args []Value) string {
	hash := c.hashBytes(algorithm, g3cryptoExtractHashInput(args))
	if len(hash) == 0 {
		return ""
	}
	return hex.EncodeToString(hash)
}

// hashBytesAsVBArray computes one hash and returns bytes as a VM VBArray.
func (c *G3Crypto) hashBytesAsVBArray(algorithm string, args []Value) Value {
	hash := c.hashBytes(algorithm, g3cryptoExtractHashInput(args))
	if len(hash) == 0 {
		return g3cryptoBytesToVBArray(nil)
	}
	c.lastHash = hash
	return g3cryptoBytesToVBArray(hash)
}

// hmacHex computes an HMAC digest and returns lowercase hexadecimal output.
func (c *G3Crypto) hmacHex(algorithm string, data string, key string) string {
	algorithm = g3cryptoNormalizeHashAlgorithm(algorithm)
	dataBytes := []byte(data)
	keyBytes := []byte(key)

	var sum []byte
	switch algorithm {
	case "sha256":
		mac := hmac.New(sha256.New, keyBytes)
		_, _ = mac.Write(dataBytes)
		sum = mac.Sum(nil)
	case "sha512":
		mac := hmac.New(sha512.New, keyBytes)
		_, _ = mac.Write(dataBytes)
		sum = mac.Sum(nil)
	default:
		return ""
	}

	c.lastHash = sum
	return hex.EncodeToString(sum)
}

// pbkdf2SHA256 derives one key using PBKDF2-HMAC-SHA256 and returns hexadecimal output.
func (c *G3Crypto) pbkdf2SHA256(password string, salt string, iterations int, keyLength int) string {
	if iterations <= 0 {
		iterations = 100000
	}
	if keyLength <= 0 {
		keyLength = 32
	}
	derived := pbkdf2.Key([]byte(password), []byte(salt), iterations, keyLength, sha256.New)
	c.lastHash = derived
	return hex.EncodeToString(derived)
}

// hashBytes computes one digest for the requested algorithm.
func (c *G3Crypto) hashBytes(algorithm string, input Value) []byte {
	algorithm = g3cryptoNormalizeHashAlgorithm(algorithm)
	if algorithm == "" {
		return nil
	}

	data := g3cryptoNormalizeInput(input)

	switch algorithm {
	case "md5":
		sum := md5.Sum(data)
		return sum[:]
	case "sha1":
		sum := sha1.Sum(data)
		return sum[:]
	case "sha256":
		sum := sha256.Sum256(data)
		return sum[:]
	case "sha384":
		sum := sha512.Sum384(data)
		return sum[:]
	case "sha512":
		sum := sha512.Sum512(data)
		return sum[:]
	case "sha3_256":
		sum := sha3.Sum256(data)
		return sum[:]
	case "sha3_512":
		sum := sha3.Sum512(data)
		return sum[:]
	case "blake2b256":
		sum := blake2b.Sum256(data)
		return sum[:]
	case "blake2b512":
		sum := blake2b.Sum512(data)
		return sum[:]
	}

	return nil
}

// g3cryptoResolveProgID maps CreateObject ProgIDs to crypto activation and default algorithm.
func g3cryptoResolveProgID(progID string) (string, bool) {
	trimmed := strings.TrimSpace(progID)
	if trimmed == "" {
		return "", false
	}

	if strings.EqualFold(trimmed, "G3Crypto") || strings.EqualFold(trimmed, "G3.Crypto") {
		return "", true
	}

	methodName := strings.ToLower(trimmed)
	if g3cryptoIsLegacyDotNetAlias(methodName) {
		algorithm := g3cryptoNormalizeHashAlgorithm(trimmed)
		if algorithm != "" {
			return algorithm, true
		}
	}

	return "", false
}

// g3cryptoNormalizeHashAlgorithm canonicalizes hash algorithm identifiers and aliases.
func g3cryptoNormalizeHashAlgorithm(name string) string {
	clean := strings.ToLower(strings.TrimSpace(name))
	clean = strings.ReplaceAll(clean, "-", "")
	clean = strings.ReplaceAll(clean, "_", "")
	clean = strings.ReplaceAll(clean, ".", "")

	switch clean {
	case "md5", "md5cryptoservice", "md5cryptoserviceprovider", "systemsecuritycryptographymd5cryptoserviceprovider":
		return "md5"
	case "sha1", "sha1cryptoservice", "sha1cryptoserviceprovider", "systemsecuritycryptographysha1cryptoserviceprovider":
		return "sha1"
	case "sha256", "sha256cryptoservice", "sha256cryptoserviceprovider", "systemsecuritycryptographysha256cryptoserviceprovider":
		return "sha256"
	case "sha384", "sha384cryptoservice", "sha384cryptoserviceprovider", "systemsecuritycryptographysha384cryptoserviceprovider":
		return "sha384"
	case "sha512", "sha512cryptoservice", "sha512cryptoserviceprovider", "systemsecuritycryptographysha512cryptoserviceprovider":
		return "sha512"
	case "sha3256", "sha3256cryptoservice", "sha3256cryptoserviceprovider":
		return "sha3_256"
	case "sha3512", "sha3512cryptoservice", "sha3512cryptoserviceprovider":
		return "sha3_512"
	case "blake2b256", "blake2256":
		return "blake2b256"
	case "blake2b512", "blake2512":
		return "blake2b512"
	default:
		return ""
	}
}

// g3cryptoNormalizeAlgorithm canonicalizes public algorithm names and aliases.
func g3cryptoNormalizeAlgorithm(name string) string {
	return g3cryptoNormalizeHashAlgorithm(name)
}

// g3cryptoIsLegacyDotNetAlias checks whether one identifier matches legacy crypto service aliases.
func g3cryptoIsLegacyDotNetAlias(name string) bool {
	return strings.Contains(name, "cryptoservice") || strings.Contains(name, "cryptoserviceprovider")
}

// g3cryptoHashSizeBits returns hash output size in bits for the selected algorithm.
func g3cryptoHashSizeBits(algorithm string) int {
	switch g3cryptoNormalizeHashAlgorithm(algorithm) {
	case "md5":
		return 128
	case "sha1":
		return 160
	case "sha256", "sha3_256", "blake2b256":
		return 256
	case "sha384":
		return 384
	case "sha512", "sha3_512", "blake2b512":
		return 512
	default:
		return 0
	}
}

// g3cryptoExtractHashInput extracts the first argument used as hash input.
func g3cryptoExtractHashInput(args []Value) Value {
	if len(args) == 0 {
		return NewString("")
	}
	return args[0]
}

// g3cryptoNormalizeInput converts supported VM values into a byte slice for hashing.
func g3cryptoNormalizeInput(value Value) []byte {
	if value.Type == VTArray && value.Arr != nil {
		count := len(value.Arr.Values)
		buf := make([]byte, count)
		for i := 0; i < count; i++ {
			byteVal := g3cryptoValueToInt(value.Arr.Values[i])
			if byteVal < 0 {
				byteVal = 0
			} else if byteVal > 255 {
				byteVal = 255
			}
			buf[i] = byte(byteVal)
		}
		return buf
	}

	return []byte(value.String())
}

// g3cryptoBytesToVBArray converts bytes into a zero-based VTArray of VTInteger values.
func g3cryptoBytesToVBArray(data []byte) Value {
	if len(data) == 0 {
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
	}

	values := make([]Value, len(data))
	for i := 0; i < len(data); i++ {
		values[i] = NewInteger(int64(data[i]))
	}
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}
}

// g3cryptoValueToInt converts one VM value into an integer using VBScript-like fallback.
func g3cryptoValueToInt(v Value) int {
	switch v.Type {
	case VTBool, VTInteger, VTDate, VTNativeObject, VTBuiltin:
		return int(v.Num)
	case VTDouble:
		return int(v.Flt)
	case VTString:
		parsed, err := strconv.Atoi(strings.TrimSpace(v.Str))
		if err == nil {
			return parsed
		}
		return 0
	default:
		return 0
	}
}
