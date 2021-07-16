package idxdbfile

import (
	"bufio"
	"fmt"
	"github.com/abramov-ks/mk4speedcams/pkg/speedcamonline"
	"github.com/abramov-ks/mk4speedcams/pkg/utils"
	"os"
	"strconv"
)

var PoiSelName = "SpeedCam"
var PoiName = "SpeedCam"

type IdxFile struct {
	file        os.File
	lineCounter int
}

// New Create new file
func New(resultFile string) (*IdxFile, error) {
	file, err := os.Create(resultFile)
	if err != nil {
		return nil, err
	}

	idxFile := IdxFile{file: *file, lineCounter: 0}

	return &idxFile, nil
}

// AppendHeaderFromFile Append headers
func (idxFile *IdxFile) AppendHeaderFromFile(headerFilePath string) error {
	file, err := os.Open(headerFilePath)
	if err != nil {
		return err
	}
	defer file.Close()

	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		idxFile.file.WriteString(scanner.Text())
	}

	return nil
}

func createDbRowFromRecord(line speedcamonline.SpeedcamRecord, counter int) [4]string {
	var result [4]string
	result[1] = utils.ConvertPosWgsToIndex(line.LatAsNum(), line.LongAsNum())
	result[2] = PoiSelName
	result[3] = PoiName

	return result
}

// AppendLine Add record to file
func (idxFile *IdxFile) AppendLine(line speedcamonline.SpeedcamRecord) {

	var newLineCols = createDbRowFromRecord(line, idxFile.lineCounter)
	var newLineString = fmt.Sprintf("%5s%8s%8s%8s\r\n", strconv.Itoa(idxFile.lineCounter), newLineCols[1], newLineCols[2], newLineCols[3])
	// формируем строку
	idxFile.file.WriteString(newLineString)
	idxFile.lineCounter++
}

func (idxFile *IdxFile) Close() {
	idxFile.file.Close()
}
