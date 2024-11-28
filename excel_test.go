package go_excel

import (
	"os"
	"reflect"
	"testing"
)

type Person struct {
	Name string `excel:"姓名"`
	Age  int    `excel:"年龄"`
}

func TestExcel_ExportToFile(t *testing.T) {
	people := []Person{{Name: "Jason", Age: 20}, {Name: "Jack", Age: 25}}
	err := New(&DefaultOption{
		SheetName: "Ye",
		Title:     "Y01",
	}).ExportToFile(&people)
	if err != nil {
		t.Error(err)
	}
}

func Test_Import(t *testing.T) {
	dataByte, err := os.ReadFile("people.xlsx")
	if err != nil {
		t.Error(err)
	}
	list := make([]*Person, 0)
	err = New(&DefaultOption{"people", "People"}).Import(dataByte, &list)
	if err != nil {
		t.Error(err)
	}
	for _, v := range list {
		t.Log(v)
	}
}

func Test_refType(t *testing.T) {
	type args struct {
		value any
	}
	type St struct {
		Name string `excel:"姓名"`
	}
	tests := []struct {
		name   string
		args   args
		wantIs bool
	}{
		{"01", args{value: []St{}}, true},
		{"02", args{value: []*St{}}, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			gotVal, gotIs := refType(tt.args.value)
			if gotIs != tt.wantIs {
				t.Errorf("refType() gotIs = %v, want %v", gotIs, tt.wantIs)
			}
			t.Logf("%+v", gotVal)

			t.Logf("%+v", reflect.TypeOf(gotVal).Name())

			for i := 0; i < reflect.TypeOf(gotVal).NumField(); i++ {
				t.Log(reflect.TypeOf(gotVal).Field(i).Tag)
			}
		})
	}
}
