package docx

import (
	"testing"
)

func Test_splitTextIntoTexts(t *testing.T) {
	type args struct {
		text string
	}
	tests := []struct {
		name string
		args args
		want string
	}{
		{"first", args{"<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space=\"preserve\">{{nothing-to-replace}}</w:t></w:r>"}, "<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space=\"preserve\">{{nothing-to-replace}}</w:t></w:r>"},
		{"second", args{"<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space=\"preserve\">Hello {{firstname}} {{lastname}}</w:t></w:r>"}, "<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space=\"preserve\">Hello </w:t></w:r><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space=\"preserve\">{{firstname}}</w:t></w:r><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space=\"preserve\"> </w:t></w:r><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space=\"preserve\">{{lastname}}</w:t></w:r>"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got, _ := splitTextIntoTexts(tt.args.text); got != tt.want {
				t.Errorf("splitTextIntoTexts() = \n%v\n, want \n%v\n", got, tt.want)
			}
		})
	}
}
