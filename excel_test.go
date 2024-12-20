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

func TestExcel_ReportFromFile(t *testing.T) {
	type Po struct {
		OrderNo string
		Items   []struct {
			Id   int
			Name string
		}
	}
	po1 := Po{
		OrderNo: "A112A2",
		Items: []struct {
			Id   int
			Name string
		}{
			{2, "螺丝"},
		},
	}

	type args struct {
		filePath string
		data     any
	}
	tests := []struct {
		name string
		args args
	}{
		{"test01", args{
			filePath: "temp.xlsx",
			data:     po1,
		}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			e := &Excel{}
			e.ReportFromFile(tt.args.filePath, tt.args.data)
			err := e.ReportToFile("result.xlsx")
			if err != nil {
				t.Error(err)
			}
		})
	}
}

func TestExcel_processCellTemplate(t *testing.T) {
	type args struct {
		cellContent string
		Data        any
	}
	d01 := struct {
		OrderNo string
	}{OrderNo: "A112A"}

	tests := []struct {
		name    string
		args    args
		wantErr bool
	}{
		{"01", args{cellContent: "{{.OrderNo}}", Data: d01}, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			e := &Excel{
				Data: tt.args.Data,
			}
			got, err := e.processCellTemplate(tt.args.cellContent)
			if (err != nil) != tt.wantErr {
				t.Errorf("processCellTemplate() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			t.Logf("%+v", got)
		})
	}
}

func TestExcel_processRangeTemplate(t *testing.T) {
	type args struct {
		cellContent string
		data        any
	}

	type T01 struct {
		Id int
	}

	type T02 struct {
		Name string
		List []T01
	}

	d01 := []T01{
		{1}, {2}, {3},
	}
	d02 := T02{
		Name: "T02",
		List: d01,
	}
	tests := []struct {
		name    string
		args    args
		wantErr bool
	}{
		{"01", args{
			cellContent: "{{- range . }} {{ .Id }} {{- end }}",
			data:        d01,
		}, false},
		{"02", args{
			cellContent: "{{ .Name kml }} {{- range .List }} {{ .Id }} {{- end }}",
			data:        d02,
		}, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			e := &Excel{
				Data: tt.args.data,
			}
			got, err := e.processRangeTemplate(tt.args.cellContent)
			if (err != nil) != tt.wantErr {
				t.Errorf("processRangeTemplate() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			t.Logf("%+v", got)
		})
	}
}
