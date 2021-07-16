package main

import (
	"flag"
	"fmt"
	"github.com/abramov-ks/mk4speedcams/pkg/databasegen"
	"github.com/abramov-ks/mk4speedcams/pkg/speedcamonline"
	"gopkg.in/yaml.v2"
	"log"
	"os"
)

var APP_VERSION = "0.1"

type Config struct {
	Source speedcamonline.Config `yaml:"source"`
	Output databasegen.Config    `yaml:"output"`
}

// ValidateConfigPath Валидация конфига
func ValidateConfigPath(path string) error {
	s, err := os.Stat(path)
	if err != nil {
		return err
	}
	if s.IsDir() {
		return fmt.Errorf("'%s' is a directory, not a normal file", path)
	}

	return nil
}

// NewConfig Загрузка конфига
func NewConfig(configPath string) (*Config, error) {
	config := &Config{}
	file, err := os.Open(configPath)
	if err != nil {
		return nil, err
	}
	defer file.Close()
	d := yaml.NewDecoder(file)
	if err := d.Decode(&config); err != nil {
		return nil, err
	}

	return config, nil
}

func (config Config) Run() int {
	log.Println("App started to " + config.Output.OutputDir)
	// download database

	database, err := config.Source.Load()

	if err != nil {
		log.Printf("Error on loading database: %s", err)
		return 127
	}

	database = speedcamonline.SortDatabase(database)

	generationResult, generationError := config.Output.GenerateFromDatabase(database)
	if err != nil {
		log.Printf("Error on generation: %s", generationError)
		return 127
	}

	log.Printf("Generation successful!\nTotal lines: %d\nIdx file path: %s\nUrl file path: %s", generationResult.RecordsCount, generationResult.IdxFile, generationResult.UrlFile)

	return 0
}

func main() {
	var cfgPath string
	var appVersion bool

	flag.StringVar(&cfgPath, "config", "./config.yml", "path to config file")
	flag.BoolVar(&appVersion, "version", false, "show application version")

	flag.Parse()

	if appVersion == true {
		fmt.Printf("BMW MK4 Speedcams DB Updater. Version %s\n", APP_VERSION)
		return
	}
	if err := ValidateConfigPath(cfgPath); err != nil {
		fmt.Println("No config file found " + cfgPath)
		return
	}

	cfg, err := NewConfig(cfgPath)
	if err != nil {
		log.Fatal(err)
	}

	os.Exit(cfg.Run())
}
