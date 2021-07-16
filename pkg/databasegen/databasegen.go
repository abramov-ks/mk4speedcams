package databasegen

import (
	"fmt"
	"github.com/abramov-ks/mk4speedcams/pkg/idxdbfile"
	"github.com/abramov-ks/mk4speedcams/pkg/speedcamonline"
	"github.com/abramov-ks/mk4speedcams/pkg/urldbfile"
	"log"
	"os"
	"path"
)

type Config struct {
	OutputDir     string `yaml:"output_dir"`
	IdxHeaderFile string `yaml:"idx_header_file"`
	UrlHeaderFile string `yaml:"url_header_file"`
}

func (config *Config) getOutputDirPath() string {
	execPath, _ := os.Executable()
	return path.Join(execPath, "../../"+config.OutputDir)
}

func (config *Config) getIdxHeaderFilePath() string {
	execPath, _ := os.Executable()
	return path.Join(execPath, "../../"+config.IdxHeaderFile)
}

func (config *Config) getUrlHeaderFilePath() string {
	execPath, _ := os.Executable()
	return path.Join(execPath, "../../"+config.UrlHeaderFile)
}

type Result struct {
	IdxFile      string
	UrlFile      string
	PoisonFile   string
	RecordsCount int
}

func (config Config) GenerateFromDatabase(database []speedcamonline.SpeedcamRecord) (*Result, error) {
	var result = Result{}

	log.Println("Start DB generating")

	// читаем заголовки в новые файлы
	result.IdxFile = config.getOutputDirPath() + "/0009.IDX"
	result.UrlFile = config.getOutputDirPath() + "/0010.URL"
	result.PoisonFile = config.getOutputDirPath() + "/poison.txt"

	idxFile, err := idxdbfile.New(result.IdxFile)
	if err != nil {
		return nil, err
	}

	headerError := idxFile.AppendHeaderFromFile(config.getIdxHeaderFilePath())

	if headerError != nil {
		return nil, headerError
	}

	urlFile, err2 := urldbfile.New(result.UrlFile)
	if err2 != nil {
		return nil, err
	}

	headerError2 := urlFile.AppendHeaderFromFile(config.getUrlHeaderFilePath())

	poisonFile, poisonFileError := os.Create(result.PoisonFile)

	if poisonFileError != nil {
		return nil, poisonFileError
	}

	if headerError2 != nil {
		return nil, headerError
	}

	for _, line := range database {
		idxFile.AppendLine(line)
		urlFile.AppendLine(line)

		poisonFile.WriteString(fmt.Sprintf("%f,%f\n", line.Long, line.Lat))

		result.RecordsCount++
	}

	idxFile.Close()
	urlFile.Close()
	poisonFile.Close()

	return &result, nil
}
