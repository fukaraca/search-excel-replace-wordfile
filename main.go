package main

import (
	"bufio"
	"fmt"
	"github.com/nguyenthenguyen/docx"
	"github.com/xuri/excelize/v2"
	"log"
	"regexp"
	"strings"
)

//
func main() {
	path := "word.docx"
	output := "new file.docx"
	parseAndFind(path, output)
}

//parseAndFind opens related files and searches for a pattern as in regexp dictates
func parseAndFind(path, output string) {

	r, _ := docx.ReadDocxFile(path)
	defer r.Close()
	doc := r.Editable()

	exc, err := excelize.OpenFile("excel.xlsx")
	defer exc.Close()
	if err != nil {
		log.Fatal("failed at opening excel:", err)
	}
	reader := strings.NewReader(doc.GetContent())

	reg := regexp.MustCompile(`([A-Z]+[A-Z0-9-]*){4,}\d`)
	list := make(map[string]string)
	scanner := bufio.NewScanner(reader)
	for scanner.Scan() {
		for _, str := range reg.FindAllString(scanner.Text(), -1) {
			if str != "" {
				if _, ok := list[str]; !ok {
					if changed := findAndReplace(str, exc); changed != "" {
						list[str] = changed
					}
				}
			}
		}
	}

	fmt.Println("total found:", len(list))

	for k, v := range list {
		err := doc.Replace(k, v, -1)
		log.Printf("replaced %s with %s successfully \n", k, v)
		if err != nil {
			log.Println("failed to replace:", err)
		}
	}
	err = doc.WriteToFile(output)
	if err != nil {
		log.Fatalln(err)
	}
	fmt.Println("no problem occured for", output)
}

//findAndReplace searches the excel and formats and returns found string accordingly
func findAndReplace(str string, exc *excelize.File) string {
	res, err := exc.SearchSheet("Sheet", str)
	if err != nil {
		log.Fatal("not found in excel")
	}
	for _, re := range res {
		date, _ := exc.GetCellValue("Sheet", "G"+re[1:])
		rev, _ := exc.GetCellValue("Sheet", "F"+re[1:])
		return fmt.Sprintf("%s (%s: Rev.%s)", str, date, rev)
	}
	return ""
}
