package main

import (
	"bufio"
	"fmt"
	"math/rand"
	"os"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

func main() {
	scanner := bufio.NewScanner(os.Stdin)
	fmt.Print("How's many group?: ")
	scanner.Scan()
	numberOfGroupString := scanner.Text()
	numberOfGroup, err := strconv.Atoi(numberOfGroupString)
	if err != nil || numberOfGroup <= 0 {
		fmt.Println("Number of group should be a valid number and is higher than 0")
	}
	indexToGroup := make([]*string, 0)
	for i := 0; i < numberOfGroup; i++ {
		fmt.Printf("What's the name of group#%d: ", i+1)
		scanner.Scan()
		currentGroupName := scanner.Text()
		fmt.Printf("How's many people in this group: ")
		scanner.Scan()
		numberOfPeopleInGroupString := scanner.Text()
		numberOfPeopleInGroup, err := strconv.Atoi(numberOfPeopleInGroupString)
		if err != nil || numberOfPeopleInGroup <= 0 {
			fmt.Println("Number of people should be a valid number and is higher than 0")
		}
		for j := 0; j < numberOfPeopleInGroup; j++ {
			indexToGroup = append(indexToGroup, &currentGroupName)
		}
	}
	rand.Shuffle(len(indexToGroup), func(i, j int) {
		indexToGroup[i], indexToGroup[j] = indexToGroup[j], indexToGroup[i]
	})
	excelFile := excelize.NewFile()
	excelFile.SetCellValue("Sheet1", "A1", "No.")
	excelFile.SetCellValue("Sheet1", "B1", "Group Name")
	for i, groupName := range indexToGroup {
		axis, _ := excelize.CoordinatesToCellName(1, i+2)
		excelFile.SetCellValue("Sheet1", axis, i+1)
		axis, _ = excelize.CoordinatesToCellName(2, i+2)
		excelFile.SetCellValue("Sheet1", axis, *groupName)
	}
	fmt.Printf("What's the output file name (without .xlsx): ")
	scanner.Scan()
	fileName := scanner.Text()
	err = excelFile.SaveAs(fileName + ".xlsx")
	if err != nil {
		fmt.Printf("Cannot save file: %s\n", err.Error())
	}
}
