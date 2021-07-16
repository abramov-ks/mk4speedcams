package speedcamonline

import (
	"bufio"
	"errors"
	"math"
	"net/http"
	"os"
	"path"
	"sort"
	"strconv"
	"strings"
)

type Config struct {
	Url          string `yaml:"url"`
	DatabaseDump string `yaml:"database_dump"`
}

func (config *Config) getDatabaseDumpPath() string {
	execPath, _ := os.Executable()
	return path.Join(execPath, "../../"+config.DatabaseDump)
}

type SpeedcamRecord struct {
	Idx       int
	Lat       float64
	Long      float64
	Type      int
	Speed     int
	DirType   int
	Direction int
}

func (sr SpeedcamRecord) LatAsNum() int64 {
	return int64(math.Round(sr.Lat * 50000000 / 9))
}

func (sr SpeedcamRecord) LongAsNum() int64 {
	return int64(math.Round((sr.Long + 30) * 50000000 / 9))
}

func (config Config) downloadDatabase() ([]string, error) {
	var lines []string
	resp, err := http.Get(config.Url)
	if err != nil {
		return nil, err
	}

	defer resp.Body.Close()
	scanner := bufio.NewScanner(resp.Body)
	for scanner.Scan() {
		lines = append(lines, scanner.Text())
	}

	if config.DatabaseDump != "" {
		f, err := os.Create(config.getDatabaseDumpPath() + "/database.txt")
		if err != nil {
			return nil, err
		}
		for _, line := range lines {
			f.WriteString(line + "\n")
		}

		f.Close()
	}

	return lines[1:], err
}

/**

 */
func (config Config) parseDatabase(lines []string) ([]SpeedcamRecord, error) {
	var result []SpeedcamRecord
	for _, oneLine := range lines {
		lineData := strings.Split(oneLine, ",")
		if len(lineData) < 7 {
			return nil, errors.New("Cannot parse line: " + oneLine)
		}
		lineStruct := new(SpeedcamRecord)
		lineStruct.Idx, _ = strconv.Atoi(lineData[0])
		lineStruct.Lat, _ = strconv.ParseFloat(lineData[2], 64)
		lineStruct.Long, _ = strconv.ParseFloat(lineData[1], 64)
		lineStruct.Type, _ = strconv.Atoi(lineData[3])
		lineStruct.Speed, _ = strconv.Atoi(lineData[4])
		lineStruct.DirType, _ = strconv.Atoi(lineData[5])
		lineStruct.Direction, _ = strconv.Atoi(lineData[6])

		result = append(result, *lineStruct)
	}

	return result, nil
}

func SortDatabase(lines []SpeedcamRecord) []SpeedcamRecord {
	sort.SliceStable(lines, func(i, j int) bool {
		if lines[i].LatAsNum() < lines[j].LatAsNum() {
			return true
		} else if lines[i].LatAsNum() == lines[j].LatAsNum() {
			return lines[i].LongAsNum() < lines[j].LongAsNum()
		}

		return false
	})

	return lines
}

func (config Config) Load() ([]SpeedcamRecord, error) {

	lines, err := config.downloadDatabase()
	if err != nil {
		return nil, err
	}

	return config.parseDatabase(lines)
}
