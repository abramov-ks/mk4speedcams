package urldbfile

import (
	"fmt"
	"github.com/abramov-ks/mk4speedcams/pkg/speedcamonline"
	"os"
)

var PoiName = "SpeedCam"

type UrlFile struct {
	file        os.File
	lineCounter int
}

// New Create new file
func New(resultFile string) (*UrlFile, error) {
	file, err := os.Create(resultFile)
	if err != nil {
		return nil, err
	}

	urlFile := UrlFile{file: *file, lineCounter: 0}

	return &urlFile, nil
}

// AppendHeaderFromFile Append headers
func (urlFile *UrlFile) AppendHeaderFromFile(headerFilePath string) error {
	data, err := os.ReadFile(headerFilePath)
	if err != nil {
		return err
	}

	urlFile.file.Write(data)
	urlFile.file.WriteString("\n")
	return nil
}

func createDbRowFromRecord(line speedcamonline.SpeedcamRecord, counter int) [2]string {
	var result [2]string

	result[0] = fmt.Sprintf("%d,%d", line.LatAsNum(), line.LongAsNum()) + string(0x00)
	result[1] = PoiName

	return result
}

// AppendLine Add record to file
func (urlFile *UrlFile) AppendLine(line speedcamonline.SpeedcamRecord) {

	var newLineCols = createDbRowFromRecord(line, urlFile.lineCounter)
	var newLineString = fmt.Sprintf("%20s%8s%1s\r\n", newLineCols[0], newLineCols[1], "1")
	// формируем строку
	urlFile.file.WriteString(newLineString)
	urlFile.lineCounter++
}

func (urlFile *UrlFile) Close() {
	urlFile.file.Close()
}
