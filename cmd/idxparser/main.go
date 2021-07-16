package main

import (
	"bufio"
	"flag"
	"fmt"
	"log"
	"os"
)

func readFileByLine(filePath string) {
	var file, err = os.Open(filePath)

	if err != nil {
		log.Fatalf("failed opening file: %s", err)
	}

	defer file.Close()

	scanner := bufio.NewScanner(file)
	scanner.Split(bufio.ScanLines)

	for scanner.Scan() {
		fmt.Println(scanner.Text())
		bufio.NewReader(os.Stdin).ReadBytes('\n')
	}
}

func main() {
	var nFlag = flag.String("f", "", "File to parse")
	flag.Parse()
	readFileByLine(*nFlag)
}
