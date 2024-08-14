package go_excel

import (
	"testing"
)

type Person struct {
	Name string `excel:"姓名"`
	Age  int    `excel:"年龄"`
}

func TestExcel_ExportToFile(t *testing.T) {
	people := []Person{{Name: "Jason", Age: 20}, {Name: "Jack", Age: 25}}
	err := New("people", "people").ExportToFile(&people)
	if err != nil {
		t.Error(err)
	}
}

func Test_refType(t *testing.T) {
	plist := make([]any, 0)

	// p1 := []Person{{Name: "Jason", Age: 20}, {Name: "Jack", Age: 21}}
	// p2 := []*Person{{Name: "Jack", Age: 22}, {Name: "Jack", Age: 23}}
	// plist = append(plist, p1, p2)

	p3 := []Person{{Name: "Jason", Age: 20}, {Name: "Jack", Age: 21}}
	p4 := []*Person{{Name: "Jack", Age: 22}, {Name: "Jack", Age: 23}}
	plist = append(plist, &p3, &p4)

	for _, v := range plist {
		val, ok := refType(v)
		if !ok {
			t.Error("refType err")
		}
		t.Log(val)

		colMap, err := getField(v)
		if err != nil {
			t.Error(err)
		}
		t.Log(colMap)
	}
}
