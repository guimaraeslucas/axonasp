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
	"bytes"
	"crypto/aes"
	"crypto/cipher"
	"crypto/rand"
	"crypto/sha256"
	"encoding/binary"
	"encoding/json"
	"errors"
	"fmt"
	"hash/crc32"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"

	"g3pix.com.br/axonasp/vbscript"
	"github.com/fxamacker/cbor/v2"
	"github.com/google/uuid"
	"github.com/klauspost/compress/zstd"
	"github.com/klauspost/reedsolomon"
	"golang.org/x/crypto/pbkdf2"
)

// ===================================================================================
// 1. CONSTANTS AND DATA STRUCTURES
// ===================================================================================

const (
	G3FCMagicNumber      = "G3FC"
	G3FCFooterMagic      = "G3CE"
	G3FCHeaderSize       = 331
	G3FCFooterSize       = 40
	G3FCCreatingSystem   = "G3Pix GoLang G3FC Archiver"
	G3FCSoftwareVersion  = "1.1.3"
	G3FCMaxFECLibShards  = 255
	G3FCMinFECShards     = 1
	G3FCMaxFECShards     = 254
	G3FCAESNonceSize     = 12
	G3FCAESTagSize       = 16
	G3FCDotNetEpochTicks = 621355968000000000
)

type g3fcMainHeader struct {
	MagicNumber           [4]byte
	FormatVersionMajor    uint16
	FormatVersionMinor    uint16
	ContainerUUID         [16]byte
	CreationTimestamp     int64
	ModificationTimestamp int64
	EditVersion           uint32
	CreatingSystem        [32]byte
	SoftwareVersion       [32]byte
	FileIndexOffset       uint64
	FileIndexLength       uint64
	FileIndexCompression  byte
	GlobalCompression     byte
	EncryptionMode        byte
	ReadSalt              [64]byte
	WriteSalt             [64]byte
	KDFIterations         uint32
	FECScheme             byte
	FECLevel              byte
	FECDataOffset         uint64
	FECDataLength         uint64
	HeaderChecksum        uint32
	Reserved              [50]byte
}

type g3fcFooter struct {
	MainIndexOffset        uint64
	MainIndexLength        uint64
	MetadataFECBlockOffset uint64
	MetadataFECBlockLength uint64
	FooterChecksum         uint32
	FooterMagic            [4]byte
}

type g3fcFileEntry struct {
	Path             string `cbor:"path"`
	Type             string `cbor:"type"`
	UUID             []byte `cbor:"uuid"`
	CreationTime     int64  `cbor:"creation_time"`
	ModificationTime int64  `cbor:"modification_time"`
	Permissions      uint16 `cbor:"permissions"`
	Status           byte   `cbor:"status"`
	OriginalFilename string `cbor:"original_filename"`
	UncompressedSize uint64 `cbor:"uncompressed_size"`
	Checksum         uint32 `cbor:"checksum"`
	DataOffset       uint64 `cbor:"data_offset"`
	DataSize         uint64 `cbor:"data_size"`
	Compression      byte   `cbor:"compression"`
	BlockFileIndex   uint32 `cbor:"block_file_index"`
	ChunkGroupId     []byte `cbor:"chunk_group_id"`
	ChunkIndex       uint32 `cbor:"chunk_index"`
	TotalChunks      uint32 `cbor:"total_chunks"`
}

type g3fcConfig struct {
	CompressionLevel  int
	GlobalCompression bool
	EncryptionMode    byte
	ReadPassword      string
	KDFIterations     uint32
	FECScheme         byte
	FECLevel          byte
	SplitSize         int64
}

// G3FC implements the G3FC library as a VM native object.
type G3FC struct {
	vm *VM
}

// newG3FCObject creates a native G3FC instance.
func (vm *VM) newG3FCObject() Value {
	obj := &G3FC{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3fcItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet routes property access to method aliases.
func (z *G3FC) DispatchPropertyGet(propertyName string) Value {
	return z.DispatchMethod(propertyName, nil)
}

// DispatchMethod resolves all G3FC methods.
func (z *G3FC) DispatchMethod(methodName string, args []Value) Value {
	method := strings.ToLower(strings.TrimSpace(methodName))
	switch method {
	case "create":
		return z.methodCreate(args)
	case "extract":
		return z.methodExtract(args)
	case "list":
		return z.methodList(args)
	case "info":
		return z.methodInfo(args)
	case "find":
		return z.methodFind(args)
	case "extractsingle", "extract-single", "extract_single":
		return z.methodExtractSingle(args)
	}
	return NewEmpty()
}

func (z *G3FC) methodCreate(args []Value) Value {
	if len(args) < 2 {
		return NewBool(false)
	}
	outputRel := args[0].String()
	output, ok := z.vm.fsoResolvePath(outputRel)
	if !ok {
		return NewBool(false)
	}

	var sourcePaths []Value
	if args[1].Type == VTArray && args[1].Arr != nil {
		sourcePaths = args[1].Arr.Values
	} else {
		sourcePaths = []Value{args[1]}
	}

	config := g3fcConfig{
		CompressionLevel: 6,
		KDFIterations:    100000,
	}

	if len(args) >= 3 {
		pass := args[2].String()
		if pass != "" {
			config.ReadPassword = pass
			config.EncryptionMode = 1
		}
	}

	if len(args) >= 4 && args[3].Type == VTNativeObject {
		if _, ok := z.vm.dictionaryItems[args[3].Num]; ok {
			// Get options from dictionary
			if val, ok := z.vm.dispatchDictionaryMethod(args[3].Num, "Item", []Value{NewString("CompressionLevel")}); ok && val.Type != VTEmpty {
				config.CompressionLevel = int(z.vm.asInt(val))
			}
			if val, ok := z.vm.dispatchDictionaryMethod(args[3].Num, "Item", []Value{NewString("GlobalCompression")}); ok && val.Type != VTEmpty {
				config.GlobalCompression = z.vm.asBool(val)
			}
			if val, ok := z.vm.dispatchDictionaryMethod(args[3].Num, "Item", []Value{NewString("FECLevel")}); ok && val.Type != VTEmpty {
				config.FECLevel = byte(z.vm.asInt(val))
				if config.FECLevel > 0 {
					config.FECScheme = 1
				}
			}
			if val, ok := z.vm.dispatchDictionaryMethod(args[3].Num, "Item", []Value{NewString("SplitSize")}); ok && val.Type != VTEmpty {
				split, _ := g3fcParseSize(val.String())
				config.SplitSize = split
			}
		}
	}

	var resolvedSources []string
	for _, p := range sourcePaths {
		if resolved, ok := z.vm.fsoResolvePath(p.String()); ok {
			resolvedSources = append(resolvedSources, resolved)
		}
	}

	err := g3fcCreateArchive(output, resolvedSources, config)
	if err != nil {
		z.vm.raise(vbscript.InternalError, err.Error())
		return NewBool(false)
	}
	return NewBool(true)
}

func (z *G3FC) methodExtract(args []Value) Value {
	if len(args) < 2 {
		return NewBool(false)
	}
	archiveRel := args[0].String()
	outputRel := args[1].String()
	archive, ok1 := z.vm.fsoResolvePath(archiveRel)
	output, ok2 := z.vm.fsoResolvePath(outputRel)
	if !ok1 || !ok2 {
		return NewBool(false)
	}

	password := ""
	if len(args) >= 3 {
		password = args[2].String()
	}

	err := g3fcExtractArchive(archive, output, password)
	if err != nil {
		z.vm.raise(vbscript.InternalError, err.Error())
		return NewBool(false)
	}
	return NewBool(true)
}

func (z *G3FC) methodList(args []Value) Value {
	if len(args) < 1 {
		return NewEmpty()
	}
	archiveRel := args[0].String()
	archive, ok := z.vm.fsoResolvePath(archiveRel)
	if !ok {
		return NewEmpty()
	}

	password := ""
	if len(args) >= 2 {
		password = args[1].String()
	}
	unit := "KB"
	if len(args) >= 3 {
		unit = args[2].String()
	}
	details := false
	if len(args) >= 4 {
		details = z.vm.asBool(args[3])
	}

	fileIndex, _, err := g3fcReadFileIndex(archive, password)
	if err != nil {
		z.vm.raise(vbscript.InternalError, err.Error())
		return NewEmpty()
	}

	values := make([]Value, 0, len(fileIndex))
	for _, entry := range fileIndex {
		dictVal := z.vm.newDictionaryObject()
		z.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("Path"), NewString(entry.Path)})
		z.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("Size"), NewInteger(int64(entry.UncompressedSize))})
		z.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("FormattedSize"), NewString(g3fcFormatSize(entry.UncompressedSize, unit))})
		z.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("Type"), NewString(entry.Type)})
		if details {
			z.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("Permissions"), NewString(fmt.Sprintf("0o%o", entry.Permissions))})
			z.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("CreationTime"), NewString(g3fcNetTicksToTime(entry.CreationTime).Format(time.RFC3339))})
			z.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("Checksum"), NewString(fmt.Sprintf("%08X", entry.Checksum))})
		}
		values = append(values, dictVal)
	}
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}
}

func (z *G3FC) methodInfo(args []Value) Value {
	if len(args) < 2 {
		return NewBool(false)
	}
	archiveRel := args[0].String()
	outputRel := args[1].String()
	archive, ok1 := z.vm.fsoResolvePath(archiveRel)
	output, ok2 := z.vm.fsoResolvePath(outputRel)
	if !ok1 || !ok2 {
		return NewBool(false)
	}

	password := ""
	if len(args) >= 3 {
		password = args[2].String()
	}

	err := g3fcExportInfo(archive, password, output)
	if err != nil {
		z.vm.raise(vbscript.InternalError, err.Error())
		return NewBool(false)
	}
	return NewBool(true)
}

func (z *G3FC) methodFind(args []Value) Value {
	if len(args) < 2 {
		return NewEmpty()
	}
	archiveRel := args[0].String()
	archive, ok := z.vm.fsoResolvePath(archiveRel)
	if !ok {
		return NewEmpty()
	}
	pattern := args[1].String()

	password := ""
	if len(args) >= 3 {
		password = args[2].String()
	}
	useRegex := false
	if len(args) >= 4 {
		useRegex = z.vm.asBool(args[3])
	}

	fileIndex, _, err := g3fcReadFileIndex(archive, password)
	if err != nil {
		z.vm.raise(vbscript.InternalError, err.Error())
		return NewEmpty()
	}

	var values []Value
	var regex *regexp.Regexp
	if useRegex {
		regex, _ = regexp.Compile("(?i)" + pattern)
	}

	for _, entry := range fileIndex {
		match := false
		if useRegex && regex != nil {
			match = regex.MatchString(entry.Path)
		} else {
			match = strings.Contains(strings.ToLower(entry.Path), strings.ToLower(pattern))
		}

		if match {
			dictVal := z.vm.newDictionaryObject()
			z.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("Path"), NewString(entry.Path)})
			z.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("Size"), NewInteger(int64(entry.UncompressedSize))})
			values = append(values, dictVal)
		}
	}
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}
}

func (z *G3FC) methodExtractSingle(args []Value) Value {
	if len(args) < 3 {
		return NewBool(false)
	}
	archiveRel := args[0].String()
	filePath := args[1].String()
	outputRel := args[2].String()
	archive, ok1 := z.vm.fsoResolvePath(archiveRel)
	output, ok2 := z.vm.fsoResolvePath(outputRel)
	if !ok1 || !ok2 {
		return NewBool(false)
	}

	password := ""
	if len(args) >= 4 {
		password = args[3].String()
	}

	err := g3fcExtractSingleFile(archive, password, filePath, output)
	if err != nil {
		z.vm.raise(vbscript.InternalError, err.Error())
		return NewBool(false)
	}
	return NewBool(true)
}

// ===================================================================================
// 2. HELPER METHODS (Internal)
// ===================================================================================

var g3fcCRC32Table = crc32.MakeTable(crc32.IEEE)

func g3fcCrc32Compute(data []byte) uint32 {
	return crc32.Checksum(data, g3fcCRC32Table)
}

func g3fcDeriveKey(password string, salt []byte, iterations int) []byte {
	return pbkdf2.Key([]byte(password), salt, iterations, 32, sha256.New)
}

func g3fcEncryptAESGCM(plaintext, key []byte) ([]byte, error) {
	block, err := aes.NewCipher(key)
	if err != nil {
		return nil, err
	}
	gcm, err := cipher.NewGCM(block)
	if err != nil {
		return nil, err
	}
	nonce := make([]byte, G3FCAESNonceSize)
	if _, err := io.ReadFull(rand.Reader, nonce); err != nil {
		return nil, err
	}
	sealed := gcm.Seal(nil, nonce, plaintext, nil)
	ciphertext := sealed[:len(plaintext)]
	tag := sealed[len(plaintext):]
	result := make([]byte, 0, G3FCAESNonceSize+G3FCAESTagSize+len(ciphertext))
	result = append(result, nonce...)
	result = append(result, tag...)
	result = append(result, ciphertext...)
	return result, nil
}

func g3fcDecryptAESGCM(payload, key []byte) ([]byte, error) {
	if len(payload) < G3FCAESNonceSize+G3FCAESTagSize {
		return nil, errors.New("invalid encryption data: payload too short")
	}
	block, err := aes.NewCipher(key)
	if err != nil {
		return nil, err
	}
	gcm, err := cipher.NewGCM(block)
	if err != nil {
		return nil, err
	}
	nonce := payload[:G3FCAESNonceSize]
	tag := payload[G3FCAESNonceSize : G3FCAESNonceSize+G3FCAESTagSize]
	ciphertext := payload[G3FCAESNonceSize+G3FCAESTagSize:]
	ciphertextAndTag := append(ciphertext, tag...)
	plaintext, err := gcm.Open(nil, nonce, ciphertextAndTag, nil)
	if err != nil {
		return nil, errors.New("decryption failed: the password may be incorrect or the data corrupted")
	}
	return plaintext, nil
}

func g3fcCreateFEC(data []byte, fecLevel byte) ([]byte, error) {
	if len(data) == 0 || fecLevel == 0 {
		return []byte{}, nil
	}
	parityShardsCount := (int(fecLevel) * (G3FCMaxFECLibShards - 1)) / 100
	if parityShardsCount < G3FCMinFECShards {
		parityShardsCount = G3FCMinFECShards
	}
	if parityShardsCount > G3FCMaxFECShards {
		parityShardsCount = G3FCMaxFECShards
	}
	dataShardsCount := G3FCMaxFECLibShards - parityShardsCount
	if dataShardsCount <= 0 {
		dataShardsCount = 1
	}
	enc, err := reedsolomon.New(dataShardsCount, parityShardsCount)
	if err != nil {
		return nil, err
	}
	shards, err := enc.Split(data)
	if err != nil {
		return nil, err
	}
	if err := enc.Encode(shards); err != nil {
		return nil, err
	}
	var parityBytes bytes.Buffer
	for _, shard := range shards[dataShardsCount:] {
		parityBytes.Write(shard)
	}
	return parityBytes.Bytes(), nil
}

func g3fcSerializeIndex(fileIndex []g3fcFileEntry) ([]byte, error) { return cbor.Marshal(fileIndex) }
func g3fcDeserializeIndex(data []byte) ([]g3fcFileEntry, error) {
	var fileIndex []g3fcFileEntry
	err := cbor.Unmarshal(data, &fileIndex)
	return fileIndex, err
}

func g3fcTimeToNetTicks(t time.Time) int64 {
	return t.UnixNano()/100 + G3FCDotNetEpochTicks
}

func g3fcNetTicksToTime(ticks int64) time.Time {
	return time.Unix(0, (ticks-G3FCDotNetEpochTicks)*100)
}

func g3fcParseSize(sizeStr string) (int64, error) {
	if sizeStr == "" {
		return 0, nil
	}
	re := regexp.MustCompile(`^(\d+)(MB|GB)$`)
	matches := re.FindStringSubmatch(strings.ToUpper(sizeStr))
	if len(matches) != 3 {
		return 0, fmt.Errorf("invalid size format. Use a number followed by MB or GB (e.g., 100MB)")
	}
	size, _ := strconv.ParseInt(matches[1], 10, 64)
	unit := matches[2]
	if unit == "MB" {
		return size * 1024 * 1024, nil
	}
	if unit == "GB" {
		return size * 1024 * 1024 * 1024, nil
	}
	return 0, nil
}

func g3fcFormatSize(bytes uint64, unit string) string {
	size := float64(bytes)
	switch strings.ToUpper(unit) {
	case "TB":
		return fmt.Sprintf("%.2f TB", size/(1024*1024*1024*1024))
	case "GB":
		return fmt.Sprintf("%.2f GB", size/(1024*1024*1024))
	case "MB":
		return fmt.Sprintf("%.2f MB", size/(1024*1024))
	case "KB":
		return fmt.Sprintf("%.2f KB", size/1024)
	default:
		return fmt.Sprintf("%d B", bytes)
	}
}

// ===================================================================================
// 3. CORE ARCHIVER LOGIC
// ===================================================================================

func g3fcCreateArchive(outputFilePath string, sourcePaths []string, config g3fcConfig) error {
	var filesToProcess []struct{ FullPath, RelativePath string }
	for _, path := range sourcePaths {
		info, err := os.Stat(path)
		if os.IsNotExist(err) {
			continue
		}
		if !info.IsDir() {
			filesToProcess = append(filesToProcess, struct{ FullPath, RelativePath string }{path, filepath.Base(path)})
		} else {
			baseDir := filepath.Dir(path)
			filepath.Walk(path, func(p string, i os.FileInfo, err error) error {
				if err != nil {
					return err
				}
				if !i.IsDir() {
					relPath, _ := filepath.Rel(baseDir, p)
					filesToProcess = append(filesToProcess, struct{ FullPath, RelativePath string }{p, relPath})
				}
				return nil
			})
		}
	}
	if len(filesToProcess) == 0 {
		return errors.New("no valid files found in the input paths")
	}

	var fileIndex []g3fcFileEntry
	dataBlockStream := new(bytes.Buffer)
	zstdEncoder, _ := zstd.NewWriter(nil, zstd.WithEncoderLevel(zstd.EncoderLevelFromZstd(config.CompressionLevel)))
	for _, file := range filesToProcess {
		fileData, err := os.ReadFile(file.FullPath)
		if err != nil {
			continue
		}
		fileInfo, _ := os.Stat(file.FullPath)
		permissions := uint16(fileInfo.Mode().Perm() & 0777)
		modTimeTicks := g3fcTimeToNetTicks(fileInfo.ModTime())
		newUUID, _ := uuid.NewRandom()
		entry := g3fcFileEntry{
			Path:             filepath.ToSlash(file.RelativePath),
			Type:             "file",
			UUID:             newUUID[:],
			CreationTime:     modTimeTicks,
			ModificationTime: modTimeTicks,
			Permissions:      permissions,
			Status:           0,
			OriginalFilename: fileInfo.Name(),
			UncompressedSize: uint64(len(fileData)),
			Checksum:         g3fcCrc32Compute(fileData),
			ChunkGroupId:     make([]byte, 0),
		}
		var dataToAdd []byte
		if config.GlobalCompression {
			dataToAdd = fileData
			entry.Compression = 0
		} else {
			dataToAdd = zstdEncoder.EncodeAll(fileData, nil)
			entry.Compression = 1
		}
		entry.DataOffset = uint64(dataBlockStream.Len())
		entry.DataSize = uint64(len(dataToAdd))
		dataBlockStream.Write(dataToAdd)
		fileIndex = append(fileIndex, entry)
	}

	var readKey, readSalt, writeSalt []byte
	if config.EncryptionMode > 0 {
		readSalt = make([]byte, 64)
		rand.Read(readSalt)
		readKey = g3fcDeriveKey(config.ReadPassword, readSalt, int(config.KDFIterations))
		writeSalt = readSalt
	}

	if config.SplitSize > 0 {
		return g3fcWriteSplitArchive(outputFilePath, fileIndex, dataBlockStream.Bytes(), config, readKey, readSalt, writeSalt)
	}
	return g3fcWriteSingleArchive(outputFilePath, fileIndex, dataBlockStream.Bytes(), config, readKey, readSalt, writeSalt)
}

func g3fcWriteSingleArchive(outputFilePath string, fileIndex []g3fcFileEntry, fileDataBlockBytes []byte, config g3fcConfig, readKey, readSalt, writeSalt []byte) error {
	var err error
	if config.GlobalCompression {
		zstdEncoder, _ := zstd.NewWriter(nil, zstd.WithEncoderLevel(zstd.EncoderLevelFromZstd(config.CompressionLevel)))
		fileDataBlockBytes = zstdEncoder.EncodeAll(fileDataBlockBytes, nil)
	}
	if config.EncryptionMode > 0 {
		fileDataBlockBytes, err = g3fcEncryptAESGCM(fileDataBlockBytes, readKey)
		if err != nil {
			return err
		}
	}
	uncompressedIndexBytes, _ := g3fcSerializeIndex(fileIndex)
	zstdEncoder, _ := zstd.NewWriter(nil)
	compressedIndexBytes := zstdEncoder.EncodeAll(uncompressedIndexBytes, nil)
	indexBlockBytes := compressedIndexBytes
	if config.EncryptionMode > 0 {
		indexBlockBytes, err = g3fcEncryptAESGCM(indexBlockBytes, readKey)
		if err != nil {
			return err
		}
	}
	header := g3fcCreateHeader(config, readSalt, writeSalt)
	currentOffset := uint64(G3FCHeaderSize)
	header.FileIndexOffset = currentOffset
	header.FileIndexLength = uint64(len(indexBlockBytes))
	currentOffset += header.FileIndexLength
	currentOffset += uint64(len(fileDataBlockBytes))
	header.FECDataOffset = currentOffset
	var dataFECBytes []byte
	if config.FECScheme == 1 {
		dataFECBytes, err = g3fcCreateFEC(fileDataBlockBytes, config.FECLevel)
		if err != nil {
			return err
		}
	}
	header.FECDataLength = uint64(len(dataFECBytes))
	currentOffset += header.FECDataLength
	var metadataFECBytes []byte
	if config.FECScheme == 1 {
		var tempHeaderBuf bytes.Buffer
		binary.Write(&tempHeaderBuf, binary.LittleEndian, header)
		metadataToProtect := append(tempHeaderBuf.Bytes(), uncompressedIndexBytes...)
		metadataFECBytes, err = g3fcCreateFEC(metadataToProtect, 10)
		if err != nil {
			return err
		}
	}
	footer := g3fcFooter{
		MainIndexOffset:        header.FileIndexOffset,
		MainIndexLength:        header.FileIndexLength,
		MetadataFECBlockOffset: currentOffset,
		MetadataFECBlockLength: uint64(len(metadataFECBytes)),
	}
	copy(footer.FooterMagic[:], []byte(G3FCFooterMagic))
	var footerChecksumBuf bytes.Buffer
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MainIndexOffset)
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MainIndexLength)
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MetadataFECBlockOffset)
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MetadataFECBlockLength)
	footer.FooterChecksum = g3fcCrc32Compute(footerChecksumBuf.Bytes())
	header.ModificationTimestamp = g3fcTimeToNetTicks(time.Now())
	var headerBuf bytes.Buffer
	binary.Write(&headerBuf, binary.LittleEndian, &header)
	headerBytes := headerBuf.Bytes()
	header.HeaderChecksum = g3fcCrc32Compute(headerBytes[:277])
	f, err := os.Create(outputFilePath)
	if err != nil {
		return err
	}
	defer f.Close()
	if err := binary.Write(f, binary.LittleEndian, &header); err != nil {
		return err
	}
	if _, err := f.Write(indexBlockBytes); err != nil {
		return err
	}
	if _, err := f.Write(fileDataBlockBytes); err != nil {
		return err
	}
	if _, err := f.Write(dataFECBytes); err != nil {
		return err
	}
	if _, err := f.Write(metadataFECBytes); err != nil {
		return err
	}
	if err := binary.Write(f, binary.LittleEndian, &footer); err != nil {
		return err
	}
	return nil
}

func g3fcWriteSplitArchive(outputFilePath string, originalFileIndex []g3fcFileEntry, combinedData []byte, config g3fcConfig, readKey, readSalt, writeSalt []byte) error {
	splitSize := config.SplitSize
	blockIndex := 0
	var finalFileIndex []g3fcFileEntry
	currentBlockStream := new(bytes.Buffer)
	for _, entry := range originalFileIndex {
		entryData := combinedData[entry.DataOffset : entry.DataOffset+entry.DataSize]
		chunkGroupId, _ := uuid.NewRandom()
		entryDataOffset := int64(0)
		chunkIndex := uint32(0)
		totalChunks := uint32((int64(len(entryData)) + splitSize - 1) / splitSize)
		if totalChunks == 0 && len(entryData) > 0 {
			totalChunks = 1
		}
		for entryDataOffset < int64(len(entryData)) || (len(entryData) == 0 && chunkIndex == 0) {
			spaceInCurrentBlock := splitSize - int64(currentBlockStream.Len())
			if spaceInCurrentBlock <= 0 && currentBlockStream.Len() > 0 {
				g3fcWriteDataBlock(outputFilePath, blockIndex, currentBlockStream.Bytes(), config, readKey)
				blockIndex++
				currentBlockStream.Reset()
				spaceInCurrentBlock = splitSize
			}
			bytesToWrite := g3fcMin64(int64(len(entryData))-entryDataOffset, spaceInCurrentBlock)
			chunkEntry := entry
			chunkEntry.BlockFileIndex = uint32(blockIndex)
			chunkEntry.DataOffset = uint64(currentBlockStream.Len())
			chunkEntry.DataSize = uint64(bytesToWrite)
			chunkEntry.ChunkGroupId = chunkGroupId[:]
			chunkEntry.ChunkIndex = chunkIndex
			chunkEntry.TotalChunks = totalChunks
			finalFileIndex = append(finalFileIndex, chunkEntry)
			currentBlockStream.Write(entryData[entryDataOffset : entryDataOffset+bytesToWrite])
			entryDataOffset += bytesToWrite
			chunkIndex++
			if len(entryData) == 0 {
				break
			}
		}
	}
	if currentBlockStream.Len() > 0 {
		g3fcWriteDataBlock(outputFilePath, blockIndex, currentBlockStream.Bytes(), config, readKey)
	}
	uncompressedIndexBytes, _ := g3fcSerializeIndex(finalFileIndex)
	zstdEncoder, _ := zstd.NewWriter(nil)
	compressedIndexBytes := zstdEncoder.EncodeAll(uncompressedIndexBytes, nil)
	indexBlockBytes := compressedIndexBytes
	var err error
	if config.EncryptionMode > 0 {
		indexBlockBytes, err = g3fcEncryptAESGCM(indexBlockBytes, readKey)
		if err != nil {
			return err
		}
	}
	header := g3fcCreateHeader(config, readSalt, writeSalt)
	header.FileIndexOffset = G3FCHeaderSize
	header.FileIndexLength = uint64(len(indexBlockBytes))
	header.FECDataOffset = 0
	header.FECDataLength = 0
	currentOffset := uint64(G3FCHeaderSize) + header.FileIndexLength
	var metadataFECBytes []byte
	if config.FECScheme == 1 {
		var tempHeaderBuf bytes.Buffer
		binary.Write(&tempHeaderBuf, binary.LittleEndian, header)
		metadataToProtect := append(tempHeaderBuf.Bytes(), uncompressedIndexBytes...)
		metadataFECBytes, err = g3fcCreateFEC(metadataToProtect, 10)
		if err != nil {
			return err
		}
	}
	footer := g3fcFooter{
		MainIndexOffset:        header.FileIndexOffset,
		MainIndexLength:        header.FileIndexLength,
		MetadataFECBlockOffset: currentOffset,
		MetadataFECBlockLength: uint64(len(metadataFECBytes)),
	}
	copy(footer.FooterMagic[:], []byte(G3FCFooterMagic))
	var footerChecksumBuf bytes.Buffer
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MainIndexOffset)
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MainIndexLength)
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MetadataFECBlockOffset)
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MetadataFECBlockLength)
	footer.FooterChecksum = g3fcCrc32Compute(footerChecksumBuf.Bytes())
	header.ModificationTimestamp = g3fcTimeToNetTicks(time.Now())
	var headerBuf bytes.Buffer
	binary.Write(&headerBuf, binary.LittleEndian, &header)
	headerBytes := headerBuf.Bytes()
	header.HeaderChecksum = g3fcCrc32Compute(headerBytes[:277])
	f, err := os.Create(outputFilePath)
	if err != nil {
		return err
	}
	defer f.Close()
	binary.Write(f, binary.LittleEndian, &header)
	f.Write(indexBlockBytes)
	f.Write(metadataFECBytes)
	binary.Write(f, binary.LittleEndian, &footer)
	return nil
}

func g3fcWriteDataBlock(baseFilePath string, blockIndex int, data []byte, config g3fcConfig, readKey []byte) {
	blockPath := fmt.Sprintf("%s%d", baseFilePath, blockIndex)
	var err error
	if config.GlobalCompression {
		zstdEncoder, _ := zstd.NewWriter(nil, zstd.WithEncoderLevel(zstd.EncoderLevelFromZstd(config.CompressionLevel)))
		data = zstdEncoder.EncodeAll(data, nil)
	}
	if config.EncryptionMode > 0 {
		data, err = g3fcEncryptAESGCM(data, readKey)
		if err != nil {
			return
		}
	}
	os.WriteFile(blockPath, data, 0644)
}

func g3fcCreateHeader(config g3fcConfig, readSalt, writeSalt []byte) g3fcMainHeader {
	ticksNow := g3fcTimeToNetTicks(time.Now())
	containerUUID, _ := uuid.NewRandom()
	header := g3fcMainHeader{
		FormatVersionMajor:    1,
		FormatVersionMinor:    0,
		CreationTimestamp:     ticksNow,
		ModificationTimestamp: ticksNow,
		EditVersion:           1,
		FileIndexCompression:  1,
		GlobalCompression:     0,
		EncryptionMode:        config.EncryptionMode,
		KDFIterations:         config.KDFIterations,
		FECScheme:             config.FECScheme,
		FECLevel:              config.FECLevel,
	}
	copy(header.MagicNumber[:], []byte(G3FCMagicNumber))
	copy(header.ContainerUUID[:], containerUUID[:])
	copy(header.CreatingSystem[:], []byte(G3FCCreatingSystem))
	copy(header.SoftwareVersion[:], []byte(G3FCSoftwareVersion))
	if config.GlobalCompression {
		header.GlobalCompression = 1
	}
	if readSalt != nil {
		copy(header.ReadSalt[:], readSalt)
	}
	if writeSalt != nil {
		copy(header.WriteSalt[:], writeSalt)
	}
	return header
}

func g3fcMin64(a, b int64) int64 {
	if a < b {
		return a
	}
	return b
}

func g3fcReadFileIndex(filePath, password string) ([]g3fcFileEntry, g3fcMainHeader, error) {
	f, err := os.Open(filePath)
	if err != nil {
		return nil, g3fcMainHeader{}, err
	}
	defer f.Close()
	var header g3fcMainHeader
	if err := binary.Read(f, binary.LittleEndian, &header); err != nil {
		return nil, g3fcMainHeader{}, fmt.Errorf("failed to read header: %w", err)
	}
	if string(header.MagicNumber[:]) != G3FCMagicNumber {
		return nil, g3fcMainHeader{}, errors.New("invalid header magic number")
	}
	indexBlockBytes := make([]byte, header.FileIndexLength)
	_, err = f.ReadAt(indexBlockBytes, int64(header.FileIndexOffset))
	if err != nil {
		return nil, header, fmt.Errorf("failed to read index block: %w", err)
	}
	if header.EncryptionMode > 0 {
		if password == "" {
			return nil, header, errors.New("password required for this archive")
		}
		key := g3fcDeriveKey(password, header.ReadSalt[:], int(header.KDFIterations))
		indexBlockBytes, err = g3fcDecryptAESGCM(indexBlockBytes, key)
		if err != nil {
			return nil, header, fmt.Errorf("failed to decrypt index: %w", err)
		}
	}
	if header.FileIndexCompression == 1 {
		zstdDecoder, _ := zstd.NewReader(nil)
		indexBlockBytes, err = zstdDecoder.DecodeAll(indexBlockBytes, nil)
		if err != nil {
			return nil, header, fmt.Errorf("failed to decompress index: %w", err)
		}
	}
	fileIndex, err := g3fcDeserializeIndex(indexBlockBytes)
	if err != nil {
		return nil, header, fmt.Errorf("failed to deserialize index: %w", err)
	}
	return fileIndex, header, nil
}

func g3fcExtractArchive(archivePath, destDir, password string) error {
	fileIndex, header, err := g3fcReadFileIndex(archivePath, password)
	if err != nil {
		return err
	}
	fileGroups := make(map[string][]g3fcFileEntry)
	for _, entry := range fileIndex {
		var groupID string
		if len(entry.ChunkGroupId) == 16 {
			groupID = string(entry.ChunkGroupId)
		} else {
			groupID = string(entry.UUID)
		}
		if _, ok := fileGroups[groupID]; !ok {
			fileGroups[groupID] = make([]g3fcFileEntry, 0)
		}
		fileGroups[groupID] = append(fileGroups[groupID], entry)
	}
	for _, chunks := range fileGroups {
		sort.Slice(chunks, func(i, j int) bool { return chunks[i].ChunkIndex < chunks[j].ChunkIndex })
		err := g3fcExtractFileFromChunks(archivePath, destDir, chunks, header, password)
		if err != nil {
			continue
		}
	}
	return nil
}

func g3fcExtractFileFromChunks(archivePath, destDir string, chunks []g3fcFileEntry, header g3fcMainHeader, password string) error {
	if len(chunks) == 0 {
		return nil
	}
	firstChunk := chunks[0]
	var readKey []byte
	if header.EncryptionMode > 0 {
		readKey = g3fcDeriveKey(password, header.ReadSalt[:], int(header.KDFIterations))
	}
	reassembledStream := new(bytes.Buffer)
	isSplit := header.FECDataOffset == 0 && header.FECDataLength == 0
	dataBlocksCache := make(map[uint32][]byte)
	zstdDecoder, _ := zstd.NewReader(nil)
	for _, chunk := range chunks {
		dataBlock, cached := dataBlocksCache[chunk.BlockFileIndex]
		if !cached {
			var rawDataBlock []byte
			var err error
			if isSplit {
				blockPath := fmt.Sprintf("%s%d", archivePath, chunk.BlockFileIndex)
				rawDataBlock, err = os.ReadFile(blockPath)
				if err != nil {
					return fmt.Errorf("data block not found: %s", blockPath)
				}
			} else {
				f, err := os.Open(archivePath)
				if err != nil {
					return err
				}
				dataBlockStart := int64(header.FileIndexOffset + header.FileIndexLength)
				dataBlockLength := int64(header.FECDataOffset) - dataBlockStart
				rawDataBlock = make([]byte, dataBlockLength)
				_, err = f.ReadAt(rawDataBlock, dataBlockStart)
				f.Close()
				if err != nil {
					return err
				}
			}
			if header.EncryptionMode > 0 {
				rawDataBlock, err = g3fcDecryptAESGCM(rawDataBlock, readKey)
				if err != nil {
					return err
				}
			}
			if header.GlobalCompression == 1 {
				rawDataBlock, err = zstdDecoder.DecodeAll(rawDataBlock, nil)
				if err != nil {
					return err
				}
			}
			dataBlock = rawDataBlock
			dataBlocksCache[chunk.BlockFileIndex] = dataBlock
		}
		chunkData := dataBlock[chunk.DataOffset : chunk.DataOffset+chunk.DataSize]
		reassembledStream.Write(chunkData)
	}
	finalData := reassembledStream.Bytes()
	var err error
	uncompressedSize := firstChunk.UncompressedSize
	if header.GlobalCompression == 0 && firstChunk.Compression == 1 {
		finalData, err = zstdDecoder.DecodeAll(finalData, make([]byte, 0, uncompressedSize))
		if err != nil {
			return err
		}
	}
	if g3fcCrc32Compute(finalData) != firstChunk.Checksum {
		return fmt.Errorf("checksum mismatch for file %s", firstChunk.OriginalFilename)
	}
	destDirAbs, err := filepath.Abs(destDir)
	if err != nil {
		return fmt.Errorf("could not determine absolute destination path: %w", err)
	}
	destPath := filepath.Join(destDirAbs, firstChunk.Path)
	if !strings.HasPrefix(destPath, destDirAbs+string(os.PathSeparator)) && destPath != destDirAbs {
		return fmt.Errorf("path traversal attempt detected: '%s' tries to escape the destination directory", firstChunk.Path)
	}
	if err := os.MkdirAll(filepath.Dir(destPath), 0755); err != nil {
		return err
	}
	perm := os.FileMode(firstChunk.Permissions)
	if perm == 0 {
		perm = 0644
	}
	if err := os.WriteFile(destPath, finalData, perm); err != nil {
		return err
	}
	return nil
}

func g3fcExportInfo(archivePath, password, outputJsonPath string) error {
	fileIndex, _, err := g3fcReadFileIndex(archivePath, password)
	if err != nil {
		return fmt.Errorf("error exporting info: %w", err)
	}

	type fileEntryJsonExport struct {
		Path             string `json:"Path"`
		Type             string `json:"Type"`
		UUID             string `json:"UUID"`
		CreationTime     string `json:"CreationTime"`
		ModificationTime string `json:"ModificationTime"`
		Permissions      string `json:"Permissions"`
		Status           byte   `json:"Status"`
		OriginalFilename string `json:"OriginalFilename"`
		UncompressedSize uint64 `json:"UncompressedSize"`
		Checksum         uint32 `json:"Checksum"`
		BlockFileIndex   uint32 `json:"BlockFileIndex"`
		ChunkGroupId     string `json:"ChunkGroupId"`
		ChunkIndex       uint32 `json:"ChunkIndex"`
		TotalChunks      uint32 `json:"TotalChunks"`
	}

	jsonEntries := make([]fileEntryJsonExport, len(fileIndex))
	for i, entry := range fileIndex {
		chunkGroupIdStr := "N/A"
		if len(entry.ChunkGroupId) == 16 {
			if parsedUUID, err := uuid.FromBytes(entry.ChunkGroupId); err == nil {
				chunkGroupIdStr = parsedUUID.String()
			}
		}
		parsedUUID, _ := uuid.FromBytes(entry.UUID)
		jsonEntries[i] = fileEntryJsonExport{
			Path:             entry.Path,
			Type:             entry.Type,
			UUID:             parsedUUID.String(),
			CreationTime:     g3fcNetTicksToTime(entry.CreationTime).Format(time.RFC3339),
			ModificationTime: g3fcNetTicksToTime(entry.ModificationTime).Format(time.RFC3339),
			Permissions:      fmt.Sprintf("0o%o", entry.Permissions),
			Status:           entry.Status,
			OriginalFilename: entry.OriginalFilename,
			UncompressedSize: entry.UncompressedSize,
			Checksum:         entry.Checksum,
			BlockFileIndex:   entry.BlockFileIndex,
			ChunkGroupId:     chunkGroupIdStr,
			ChunkIndex:       entry.ChunkIndex,
			TotalChunks:      entry.TotalChunks,
		}
	}
	jsonData, err := json.MarshalIndent(jsonEntries, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to marshal index to JSON: %w", err)
	}
	err = os.WriteFile(outputJsonPath, jsonData, 0644)
	if err != nil {
		return fmt.Errorf("failed to write JSON file: %w", err)
	}
	return nil
}

func g3fcExtractSingleFile(archivePath, password, filePathInArchive, destinationDir string) error {
	fileIndex, header, err := g3fcReadFileIndex(archivePath, password)
	if err != nil {
		return fmt.Errorf("error extracting single file: %w", err)
	}
	var chunksToExtract []g3fcFileEntry
	for _, entry := range fileIndex {
		if entry.Path == filePathInArchive {
			chunksToExtract = append(chunksToExtract, entry)
		}
	}
	if len(chunksToExtract) == 0 {
		return fmt.Errorf("file '%s' not found in the archive", filePathInArchive)
	}
	if chunksToExtract[0].Type == "directory" {
		destPath := filepath.Join(destinationDir, chunksToExtract[0].Path)
		if err := os.MkdirAll(destPath, 0755); err != nil {
			return fmt.Errorf("failed to create directory %s: %w", destPath, err)
		}
		return nil
	}
	err = g3fcExtractFileFromChunks(archivePath, destinationDir, chunksToExtract, header, password)
	if err != nil {
		return err
	}
	return nil
}
