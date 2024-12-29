package go_excel

import (
	"io"
	"net/http"
	"testing"
)

func Test_detectFileType(t *testing.T) {
	type args struct {
		data []byte
	}
	resp, err := http.Get("https://erp.marchmetal.top/cdn/6,c19457c1f9")
	if err != nil {
		t.Error(err)
	}
	defer resp.Body.Close()
	img01, _ := io.ReadAll(resp.Body)

	tests := []struct {
		name string
		args args
		want string
	}{
		{"jpg network", args{img01}, ".jpg"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := detectFileType(tt.args.data); got != tt.want {
				t.Errorf("detectFileType() = %v, want %v", got, tt.want)
			}
		})
	}
}
