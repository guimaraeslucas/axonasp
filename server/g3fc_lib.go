/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 */
package server

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
	MagicNumber      = "G3FC"
	FooterMagic      = "G3CE"
	HeaderSize       = 331
	FooterSize       = 40
	CreatingSystem   = "G3Pix GoLang G3FC Archiver"
	SoftwareVersion  = "1.1.3"
	MaxFECLibShards  = 255
	MinFECShards     = 1
	MaxFECShards     = 254
	AESNonceSize     = 12
	AESTagSize       = 16
	DotNetEpochTicks = 621355968000000000
)

type MainHeader struct {
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

type Footer struct {
	MainIndexOffset        uint64
	MainIndexLength        uint64
	MetadataFECBlockOffset uint64
	MetadataFECBlockLength uint64
	FooterChecksum         uint32
	FooterMagic            [4]byte
}

type FileEntry struct {
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

type FileEntryJsonExport struct {
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

type Config struct {
	CompressionLevel  int
	GlobalCompression bool
	EncryptionMode    byte
	ReadPassword      string
	KDFIterations     uint32
	FECScheme         byte
	FECLevel          byte
	SplitSize         int64
}

// ===================================================================================
// 2. HELPER METHODS (G3FCHelpers)
// ===================================================================================

var crc32Table = crc32.MakeTable(crc32.IEEE)

func Crc32Compute(data []byte) uint32 {
	return crc32.Checksum(data, crc32Table)
}

func DeriveKey(password string, salt []byte, iterations int) []byte {
	return pbkdf2.Key([]byte(password), salt, iterations, 32, sha256.New)
}

func EncryptAESGCM(plaintext, key []byte) ([]byte, error) {
	block, err := aes.NewCipher(key)
	if err != nil {
		return nil, err
	}
	gcm, err := cipher.NewGCM(block)
	if err != nil {
		return nil, err
	}
	nonce := make([]byte, AESNonceSize)
	if _, err := io.ReadFull(rand.Reader, nonce); err != nil {
		return nil, err
	}
	sealed := gcm.Seal(nil, nonce, plaintext, nil)
	ciphertext := sealed[:len(plaintext)]
	tag := sealed[len(plaintext):]
	result := make([]byte, 0, AESNonceSize+AESTagSize+len(ciphertext))
	result = append(result, nonce...)
	result = append(result, tag...)
	result = append(result, ciphertext...)
	return result, nil
}

func DecryptAESGCM(payload, key []byte) ([]byte, error) {
	if len(payload) < AESNonceSize+AESTagSize {
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
	nonce := payload[:AESNonceSize]
	tag := payload[AESNonceSize : AESNonceSize+AESTagSize]
	ciphertext := payload[AESNonceSize+AESTagSize:]
	ciphertextAndTag := append(ciphertext, tag...)
	plaintext, err := gcm.Open(nil, nonce, ciphertextAndTag, nil)
	if err != nil {
		return nil, errors.New("decryption failed: the password may be incorrect or the data corrupted")
	}
	return plaintext, nil
}

func CreateFEC(data []byte, fecLevel byte) ([]byte, error) {
	if len(data) == 0 || fecLevel == 0 {
		return []byte{}, nil
	}
	parityShardsCount := (int(fecLevel) * (MaxFECLibShards - 1)) / 100
	if parityShardsCount < MinFECShards {
		parityShardsCount = MinFECShards
	}
	if parityShardsCount > MaxFECShards {
		parityShardsCount = MaxFECShards
	}
	dataShardsCount := MaxFECLibShards - parityShardsCount
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

func SerializeIndex(fileIndex []FileEntry) ([]byte, error) { return cbor.Marshal(fileIndex) }
func DeserializeIndex(data []byte) ([]FileEntry, error) {
	var fileIndex []FileEntry
	err := cbor.Unmarshal(data, &fileIndex)
	return fileIndex, err
}

func TimeToNetTicks(t time.Time) int64 {
	return t.UnixNano()/100 + DotNetEpochTicks
}

func NetTicksToTime(ticks int64) time.Time {
	return time.Unix(0, (ticks-DotNetEpochTicks)*100)
}

func parseSize(sizeStr string) (int64, error) {
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

// ===================================================================================
// 3. G3FC WRITER
// ===================================================================================

func CreateG3FCArchive(outputFilePath string, sourcePaths []string, config Config) error {
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

	var fileIndex []FileEntry
	dataBlockStream := new(bytes.Buffer)
	zstdEncoder, _ := zstd.NewWriter(nil, zstd.WithEncoderLevel(zstd.EncoderLevelFromZstd(config.CompressionLevel)))
	for _, file := range filesToProcess {
		fileData, err := os.ReadFile(file.FullPath)
		if err != nil {
			continue
		}
		fileInfo, _ := os.Stat(file.FullPath)
		permissions := uint16(fileInfo.Mode().Perm() & 0777)
		modTimeTicks := TimeToNetTicks(fileInfo.ModTime())
		newUUID, _ := uuid.NewRandom()
		entry := FileEntry{
			Path:             filepath.ToSlash(file.RelativePath),
			Type:             "file",
			UUID:             newUUID[:],
			CreationTime:     modTimeTicks,
			ModificationTime: modTimeTicks,
			Permissions:      permissions,
			Status:           0,
			OriginalFilename: fileInfo.Name(),
			UncompressedSize: uint64(len(fileData)),
			Checksum:         Crc32Compute(fileData),
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
		readKey = DeriveKey(config.ReadPassword, readSalt, int(config.KDFIterations))
		writeSalt = readSalt
	}

	if config.SplitSize > 0 {
		return writeSplitArchive(outputFilePath, fileIndex, dataBlockStream.Bytes(), config, readKey, readSalt, writeSalt)
	}
	return writeSingleArchive(outputFilePath, fileIndex, dataBlockStream.Bytes(), config, readKey, readSalt, writeSalt)
}

func createHeader(config Config, readSalt, writeSalt []byte) MainHeader {
	ticksNow := TimeToNetTicks(time.Now())
	containerUUID, _ := uuid.NewRandom()
	header := MainHeader{
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
	copy(header.MagicNumber[:], []byte(MagicNumber))
	copy(header.ContainerUUID[:], containerUUID[:])
	copy(header.CreatingSystem[:], []byte(CreatingSystem))
	copy(header.SoftwareVersion[:], []byte(SoftwareVersion))
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

func writeSingleArchive(outputFilePath string, fileIndex []FileEntry, fileDataBlockBytes []byte, config Config, readKey, readSalt, writeSalt []byte) error {
	var err error
	if config.GlobalCompression {
		zstdEncoder, _ := zstd.NewWriter(nil, zstd.WithEncoderLevel(zstd.EncoderLevelFromZstd(config.CompressionLevel)))
		fileDataBlockBytes = zstdEncoder.EncodeAll(fileDataBlockBytes, nil)
	}
	if config.EncryptionMode > 0 {
		fileDataBlockBytes, err = EncryptAESGCM(fileDataBlockBytes, readKey)
		if err != nil {
			return err
		}
	}
	uncompressedIndexBytes, _ := SerializeIndex(fileIndex)
	zstdEncoder, _ := zstd.NewWriter(nil)
	compressedIndexBytes := zstdEncoder.EncodeAll(uncompressedIndexBytes, nil)
	indexBlockBytes := compressedIndexBytes
	if config.EncryptionMode > 0 {
		indexBlockBytes, err = EncryptAESGCM(indexBlockBytes, readKey)
		if err != nil {
			return err
		}
	}
	header := createHeader(config, readSalt, writeSalt)
	currentOffset := uint64(HeaderSize)
	header.FileIndexOffset = currentOffset
	header.FileIndexLength = uint64(len(indexBlockBytes))
	currentOffset += header.FileIndexLength
	currentOffset += uint64(len(fileDataBlockBytes))
	header.FECDataOffset = currentOffset
	var dataFECBytes []byte
	if config.FECScheme == 1 {
		dataFECBytes, err = CreateFEC(fileDataBlockBytes, config.FECLevel)
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
		metadataFECBytes, err = CreateFEC(metadataToProtect, 10)
		if err != nil {
			return err
		}
	}
	footer := Footer{
		MainIndexOffset:        header.FileIndexOffset,
		MainIndexLength:        header.FileIndexLength,
		MetadataFECBlockOffset: currentOffset,
		MetadataFECBlockLength: uint64(len(metadataFECBytes)),
	}
	copy(footer.FooterMagic[:], []byte(FooterMagic))
	var footerChecksumBuf bytes.Buffer
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MainIndexOffset)
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MainIndexLength)
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MetadataFECBlockOffset)
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MetadataFECBlockLength)
	footer.FooterChecksum = Crc32Compute(footerChecksumBuf.Bytes())
	header.ModificationTimestamp = TimeToNetTicks(time.Now())
	var headerBuf bytes.Buffer
	binary.Write(&headerBuf, binary.LittleEndian, &header)
	headerBytes := headerBuf.Bytes()
	header.HeaderChecksum = Crc32Compute(headerBytes[:277])
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

func writeSplitArchive(outputFilePath string, originalFileIndex []FileEntry, combinedData []byte, config Config, readKey, readSalt, writeSalt []byte) error {
	splitSize := config.SplitSize
	blockIndex := 0
	var finalFileIndex []FileEntry
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
				writeDataBlock(outputFilePath, blockIndex, currentBlockStream.Bytes(), config, readKey)
				blockIndex++
				currentBlockStream.Reset()
				spaceInCurrentBlock = splitSize
			}
			bytesToWrite := min(int64(len(entryData))-entryDataOffset, spaceInCurrentBlock)
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
		writeDataBlock(outputFilePath, blockIndex, currentBlockStream.Bytes(), config, readKey)
	}
	uncompressedIndexBytes, _ := SerializeIndex(finalFileIndex)
	zstdEncoder, _ := zstd.NewWriter(nil)
	compressedIndexBytes := zstdEncoder.EncodeAll(uncompressedIndexBytes, nil)
	indexBlockBytes := compressedIndexBytes
	var err error
	if config.EncryptionMode > 0 {
		indexBlockBytes, err = EncryptAESGCM(indexBlockBytes, readKey)
		if err != nil {
			return err
		}
	}
	header := createHeader(config, readSalt, writeSalt)
	header.FileIndexOffset = HeaderSize
	header.FileIndexLength = uint64(len(indexBlockBytes))
	header.FECDataOffset = 0
	header.FECDataLength = 0
	currentOffset := uint64(HeaderSize) + header.FileIndexLength
	var metadataFECBytes []byte
	if config.FECScheme == 1 {
		var tempHeaderBuf bytes.Buffer
		binary.Write(&tempHeaderBuf, binary.LittleEndian, header)
		metadataToProtect := append(tempHeaderBuf.Bytes(), uncompressedIndexBytes...)
		metadataFECBytes, err = CreateFEC(metadataToProtect, 10)
		if err != nil {
			return err
		}
	}
	footer := Footer{
		MainIndexOffset:        header.FileIndexOffset,
		MainIndexLength:        header.FileIndexLength,
		MetadataFECBlockOffset: currentOffset,
		MetadataFECBlockLength: uint64(len(metadataFECBytes)),
	}
	copy(footer.FooterMagic[:], []byte(FooterMagic))
	var footerChecksumBuf bytes.Buffer
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MainIndexOffset)
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MainIndexLength)
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MetadataFECBlockOffset)
	binary.Write(&footerChecksumBuf, binary.LittleEndian, footer.MetadataFECBlockLength)
	footer.FooterChecksum = Crc32Compute(footerChecksumBuf.Bytes())
	header.ModificationTimestamp = TimeToNetTicks(time.Now())
	var headerBuf bytes.Buffer
	binary.Write(&headerBuf, binary.LittleEndian, &header)
	headerBytes := headerBuf.Bytes()
	header.HeaderChecksum = Crc32Compute(headerBytes[:277])
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

func writeDataBlock(baseFilePath string, blockIndex int, data []byte, config Config, readKey []byte) {
	blockPath := fmt.Sprintf("%s%d", baseFilePath, blockIndex)
	var err error
	if config.GlobalCompression {
		zstdEncoder, _ := zstd.NewWriter(nil, zstd.WithEncoderLevel(zstd.EncoderLevelFromZstd(config.CompressionLevel)))
		data = zstdEncoder.EncodeAll(data, nil)
	}
	if config.EncryptionMode > 0 {
		data, err = EncryptAESGCM(data, readKey)
		if err != nil {
			return
		}
	}
	os.WriteFile(blockPath, data, 0644)
}

func min(a, b int64) int64 {
	if a < b {
		return a
	}
	return b
}

// ===================================================================================
// 4. G3FC READER
// ===================================================================================

func ReadFileIndex(filePath, password string) ([]FileEntry, MainHeader, error) {
	f, err := os.Open(filePath)
	if err != nil {
		return nil, MainHeader{}, err
	}
	defer f.Close()
	var header MainHeader
	if err := binary.Read(f, binary.LittleEndian, &header); err != nil {
		return nil, MainHeader{}, fmt.Errorf("failed to read header: %w", err)
	}
	if string(header.MagicNumber[:]) != MagicNumber {
		return nil, MainHeader{}, errors.New("invalid header magic number")
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
		key := DeriveKey(password, header.ReadSalt[:], int(header.KDFIterations))
		indexBlockBytes, err = DecryptAESGCM(indexBlockBytes, key)
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
	fileIndex, err := DeserializeIndex(indexBlockBytes)
	if err != nil {
		return nil, header, fmt.Errorf("failed to deserialize index: %w", err)
	}
	return fileIndex, header, nil
}

func ExtractArchive(archivePath, destDir, password string) error {
	fileIndex, header, err := ReadFileIndex(archivePath, password)
	if err != nil {
		return err
	}
	fileGroups := make(map[string][]FileEntry)
	for _, entry := range fileIndex {
		var groupID string
		if len(entry.ChunkGroupId) == 16 {
			groupID = string(entry.ChunkGroupId)
		} else {
			groupID = string(entry.UUID)
		}
		if _, ok := fileGroups[groupID]; !ok {
			fileGroups[groupID] = make([]FileEntry, 0)
		}
		fileGroups[groupID] = append(fileGroups[groupID], entry)
	}
	for _, chunks := range fileGroups {
		sort.Slice(chunks, func(i, j int) bool { return chunks[i].ChunkIndex < chunks[j].ChunkIndex })
		err := extractFileFromChunks(archivePath, destDir, chunks, header, password)
		if err != nil {
			continue
		}
	}
	return nil
}

func extractFileFromChunks(archivePath, destDir string, chunks []FileEntry, header MainHeader, password string) error {
	if len(chunks) == 0 {
		return nil
	}
	firstChunk := chunks[0]
	var readKey []byte
	if header.EncryptionMode > 0 {
		readKey = DeriveKey(password, header.ReadSalt[:], int(header.KDFIterations))
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
				rawDataBlock, err = DecryptAESGCM(rawDataBlock, readKey)
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
	if Crc32Compute(finalData) != firstChunk.Checksum {
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

// ===================================================================================
// 5. NEW COMMANDS LOGIC
// ===================================================================================

func ListFilesInContainer(archivePath, password, sizeUnit string, showDetails bool) error {
	fileIndex, _, err := ReadFileIndex(archivePath, password)
	if err != nil {
		return fmt.Errorf("error listing files: %w", err)
	}
	logicalFiles := make(map[string]FileEntry)
	for _, entry := range fileIndex {
		if _, exists := logicalFiles[entry.Path]; !exists {
			logicalFiles[entry.Path] = entry
		}
	}
	return nil
}

func ExportInfo(archivePath, password, outputJsonPath string) error {
	fileIndex, _, err := ReadFileIndex(archivePath, password)
	if err != nil {
		return fmt.Errorf("error exporting info: %w", err)
	}
	jsonEntries := make([]FileEntryJsonExport, len(fileIndex))
	for i, entry := range fileIndex {
		chunkGroupIdStr := "N/A"
		if len(entry.ChunkGroupId) == 16 {
			if parsedUUID, err := uuid.FromBytes(entry.ChunkGroupId); err == nil {
				chunkGroupIdStr = parsedUUID.String()
			}
		}
		parsedUUID, _ := uuid.FromBytes(entry.UUID)
		jsonEntries[i] = FileEntryJsonExport{
			Path:             entry.Path,
			Type:             entry.Type,
			UUID:             parsedUUID.String(),
			CreationTime:     NetTicksToTime(entry.CreationTime).Format(time.RFC3339),
			ModificationTime: NetTicksToTime(entry.ModificationTime).Format(time.RFC3339),
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

func FindFilesInContainer(archivePath, password, pattern string, useRegex bool, sizeUnit string) error {
	return nil
}

func ExtractSingleFile(archivePath, password, filePathInArchive, destinationDir string) error {
	fileIndex, header, err := ReadFileIndex(archivePath, password)
	if err != nil {
		return fmt.Errorf("error extracting single file: %w", err)
	}
	var chunksToExtract []FileEntry
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
	err = extractFileFromChunks(archivePath, destinationDir, chunksToExtract, header, password)
	if err != nil {
		return err
	}
	return nil
}

func formatSize(bytes uint64, unit string) string {
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

// G3FC implements Component interface for G3FC operations
type G3FC struct {
	ctx *ExecutionContext
}

func (z *G3FC) GetProperty(name string) interface{} {
	return nil
}

func (z *G3FC) SetProperty(name string, value interface{}) {}

func (z *G3FC) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)

	// Handle explicit CallMethod call (e.g. obj.CallMethod("MethodName", args))
	if method == "callmethod" && len(args) > 0 {
		actualMethod := fmt.Sprintf("%v", args[0])
		return z.CallMethod(actualMethod, args[1:]...)
	}

	getStr := func(i int) string {
		if i >= len(args) || args[i] == nil {
			return ""
		}
		return fmt.Sprintf("%v", args[i])
	}

	switch method {
	case "create":
		if len(args) < 2 {
			return false
		}
		outputRel := getStr(0)
		output := z.ctx.Server_MapPath(outputRel)

		var sourcePaths []string
		// Check if it's an array/slice
		switch v := args[1].(type) {
		case []interface{}:
			for _, p := range v {
				sourcePaths = append(sourcePaths, z.ctx.Server_MapPath(fmt.Sprintf("%v", p)))
			}
		case []string:
			for _, p := range v {
				sourcePaths = append(sourcePaths, z.ctx.Server_MapPath(p))
			}
		default:
			sourcePaths = append(sourcePaths, z.ctx.Server_MapPath(fmt.Sprintf("%v", args[1])))
		}

		config := Config{
			CompressionLevel: 6,
			KDFIterations:    100000,
		}

		if len(args) >= 3 {
			// Password
			pass := getStr(2)
			if pass != "" {
				config.ReadPassword = pass
				config.EncryptionMode = 1
			}
		}

		if len(args) >= 4 {
			// Options dictionary or object
			if dict, ok := args[3].(*Dictionary); ok {
				if val := dict.Item([]interface{}{"CompressionLevel"}); val != nil {
					if i, err := strconv.Atoi(fmt.Sprintf("%v", val)); err == nil {
						config.CompressionLevel = i
					}
				}
				if val := dict.Item([]interface{}{"GlobalCompression"}); val != nil {
					config.GlobalCompression = strings.EqualFold(fmt.Sprintf("%v", val), "true")
				}
				if val := dict.Item([]interface{}{"FECLevel"}); val != nil {
					if i, err := strconv.Atoi(fmt.Sprintf("%v", val)); err == nil {
						config.FECLevel = byte(i)
						if config.FECLevel > 0 {
							config.FECScheme = 1
						}
					}
				}
				if val := dict.Item([]interface{}{"SplitSize"}); val != nil {
					split, _ := parseSize(fmt.Sprintf("%v", val))
					config.SplitSize = split
				}
			}
		}

		err := CreateG3FCArchive(output, sourcePaths, config)
		return err == nil

	case "extract":
		if len(args) < 2 {
			return false
		}
		archive := z.ctx.Server_MapPath(getStr(0))
		output := z.ctx.Server_MapPath(getStr(1))
		password := getStr(2)

		err := ExtractArchive(archive, output, password)
		return err == nil

	case "list":
		if len(args) < 1 {
			return nil
		}
		archive := z.ctx.Server_MapPath(getStr(0))
		password := getStr(1)
		unit := "KB"
		if len(args) >= 3 {
			unit = getStr(2)
		}
		details := false
		if len(args) >= 4 {
			details = strings.EqualFold(getStr(3), "true")
		}

		fileIndex, _, err := ReadFileIndex(archive, password)
		if err != nil {
			return nil
		}

		var result []interface{}
		for _, entry := range fileIndex {
			dict := NewDictionary(z.ctx)
			dict.CallMethod("Add", "Path", entry.Path)
			dict.CallMethod("Add", "Size", entry.UncompressedSize)
			dict.CallMethod("Add", "FormattedSize", formatSize(entry.UncompressedSize, unit))
			dict.CallMethod("Add", "Type", entry.Type)
			if details {
				dict.CallMethod("Add", "Permissions", fmt.Sprintf("0o%o", entry.Permissions))
				dict.CallMethod("Add", "CreationTime", NetTicksToTime(entry.CreationTime).Format(time.RFC3339))
				dict.CallMethod("Add", "Checksum", fmt.Sprintf("%08X", entry.Checksum))
			}
			result = append(result, dict)
		}
		return result

	case "info":
		if len(args) < 2 {
			return false
		}
		archive := z.ctx.Server_MapPath(getStr(0))
		output := z.ctx.Server_MapPath(getStr(1))
		password := getStr(2)

		err := ExportInfo(archive, password, output)
		return err == nil

	case "find":
		if len(args) < 2 {
			return nil
		}
		archive := z.ctx.Server_MapPath(getStr(0))
		pattern := getStr(1)
		password := getStr(2)
		useRegex := false
		if len(args) >= 4 {
			useRegex = strings.EqualFold(getStr(3), "true")
		}

		fileIndex, _, err := ReadFileIndex(archive, password)
		if err != nil {
			return nil
		}

		var result []interface{}
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
				dict := NewDictionary(z.ctx)
				dict.CallMethod("Add", "Path", entry.Path)
				dict.CallMethod("Add", "Size", entry.UncompressedSize)
				result = append(result, dict)
			}
		}
		return result

	case "extractsingle", "extract-single", "extract_single":
		if len(args) < 3 {
			return false
		}
		archive := z.ctx.Server_MapPath(getStr(0))
		filePath := getStr(1)
		output := z.ctx.Server_MapPath(getStr(2))
		password := getStr(3)

		err := ExtractSingleFile(archive, password, filePath, output)
		return err == nil
	}

	return nil
}

// CLI_Main renamed from main() for CLI usage
func CLI_Main() {
	if len(os.Args) < 2 || os.Args[1] == "-h" || os.Args[1] == "--help" {
		showHelp()
		return
	}
	// ... CLI logic ...
}

func showHelp() {
	fmt.Printf("G3FC Archiver Tool - Go Version %s\n", SoftwareVersion)
	fmt.Printf("Usage: %s <command> [options] [arguments...]\n", filepath.Base(os.Args[0]))
}
