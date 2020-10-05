package docx

import (
	"archive/zip"
	"bytes"
	"regexp"
	"testing"
)

func TestDocx_CreateNewDocx(t *testing.T) {
	var doc Docx
	check := "\nI like to move itI like to move it, move itI like to move it, move itya like to move itI like to move it, move itI like to move itI like to move it, move itya like to move itI like to move it, move itI like to move itI like to move it, move itYa like to move itI like to move it, move itI like to move itI like to move it, move itya like to move itAll girls all over the worldOriginal Mad Stuntman pon' ya case manI love how all girls a move them bodyAnd when ya move ya bodyGonna move it nice and sweet and sexy, alright?Woman ya cute and you don't need no make upOriginal cute body you a mek man mud upWoman ya cute and you don't need no make upOriginal cute body you a mek man mud upWomanMan physically fit, physically fitPhysically, physically, physicallyWoman, physically fit, physically fitPhysically, physically, physicallyWomanMan WomanMan, ya nice, sweet, fantasticBig ship 'pon di ocean that a big TitanicC'mon, ya nice, sweet, I enjoy the thingBig ship 'pon di ocean that a big TitanicWoman, ya nice, sweet, fantasticBig ship 'pon di ocean that a big TitanicC'mon, ya nice, sweet, I enjoy the thingBig ship 'pon di ocean that a big TitanicWomanI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itWoman ya cute and you don't need no make upOriginal cute body you a mek man mud upWoman ya cute and you don't need no make upOriginal cute body you a mek man mud upEyeliner 'pon ya face a mek man mud upNose powder 'pon ya face a mek man mud upPluck ya eyebrow 'pon ya face a mek man mud upGal ya lipstick 'pon ya face a mek man mud upWoman, ya nice, broad faceAnd ya nice hipMake man flip and bust up them lipWoman, ya nice and energeticBig ship 'pon de ocean that a big TitanicWoman, ya nice, broad faceAnd ya nice hipMake man flip and bust up them lipWoman, ya nice and energeticBig ship 'pon de ocean that a big TitanicWomanI like to move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move itI like to move it, move itI like to move it, move itya like to move it"
	var testData = map[string]string{
		"iliketo":  "I like to move it",
		"yaliketo": "ya like to move it",
		"subject":  "Woman",
		"object":   "Man",
	}
	doc.LoadDocx("testdata/iliketomoveit.docx")
	err := doc.Replace(testData)
	if err != nil {
		t.Fatalf("Error: %s\n", err.Error())
	}
	doc.CreateNewDocx()
	doc.SaveDocxToFile("/tmp/iliketo.docx")
	extractedText := extractRawTextFromDocxXml(doc.replaced)
	if check != extractedText {
		t.Error("Error: Text not matching\n\n")
		t.Error(extractedText)
		t.Error(check)
	}
	// unzip the generated docx
	r, err := zip.NewReader(bytes.NewReader(doc.result), int64(len(doc.result)))
	if err != nil {
		t.Fatalf("Error: %s\n", err.Error())
	}
	// extract document.xml
	// Iterate through the files in the archive
	for _, f := range r.File {
		if f.FileHeader.Name == "word/document.xml" {
			readCloser, err := f.Open()
			if err != nil {
				t.Errorf("Error: %s\n", err.Error())
			}
			buf := new(bytes.Buffer)
			buf.ReadFrom(readCloser)
			replacedText := extractRawTextFromDocxXml(string(buf.Bytes()))
			if check != replacedText {
				t.Error("Error: Text not matching\n\n")
			}
			if readCloser != nil {
				readCloser.Close()
			}
			break
		}
	}
}

func extractRawTextFromDocxXml(document string) (result string) {
	// from commandlinefu
	// sed -e 's/<\/w:p>/\n/g; s/<[^>]\{1,\}>//g; s/[^[:print:]\n]\{1,\}//g'
	//one := regexp.MustCompile("<\\/w:p>")
	two := regexp.MustCompile("<[^>]+>")
	three := regexp.MustCompilePOSIX("[^[:print:]\n]+")
	result = document
	//result = one.ReplaceAllString(result,"\n")
	result = two.ReplaceAllString(result, "")
	result = three.ReplaceAllString(result, "")
	return
}
