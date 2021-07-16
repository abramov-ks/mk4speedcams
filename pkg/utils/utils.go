package utils

import (
	"encoding/hex"
	"fmt"
	"strconv"
	"strings"
)

func ConvertPosWgsToIndex(latInt int64, lonInt int64) string {
	var result = ""

	var hexLine = fmt.Sprintf("%08x%08x", lonInt, latInt)

	for i := 0; i < len(hexLine); i = i + 2 {
		hexByte, _ := hex.DecodeString(fmt.Sprintf("%02x", (hex2int(hexLine[i : i+2]))))
		result = result + string(hexByte)
	}

	return result

}

func hex2int(hexStr string) uint64 {
	// remove 0x suffix if found in the input string
	cleaned := strings.Replace(hexStr, "0x", "", -1)

	// base 16 for hexadecimal
	result, _ := strconv.ParseUint(cleaned, 16, 64)
	return uint64(result)
}
