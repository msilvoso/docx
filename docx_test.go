package docx

import (
	"archive/zip"
	"bytes"
	"regexp"
	"testing"
)

func TestDocx_CreateNewDocx(t *testing.T) {
	var checks [3]string
	checks[0] = "\nI like to move itI like to move it, move itI like to move it, move itya like to move itI like to move it, move itI like to move itI like to move it, move itya like to move itI like to move it, move itI like to move itI like to move it, move itYa like to move itI like to move it, move itI like to move itI like to move it, move itya like to move itAll girls all over the worldOriginal Mad Stuntman pon' ya case manI love how all girls a move them bodyAnd when ya move ya bodyGonna move it nice and sweet and sexy, alright?Woman ya cute and you don't need no make upOriginal cute body you a mek man mud upWoman ya cute and you don't need no make upOriginal cute body you a mek man mud upWomanMan physically fit, physically fitPhysically, physically, physicallyWoman, physically fit, physically fitPhysically, physically, physicallyWomanMan WomanMan, ya nice, sweet, fantasticBig ship 'pon di ocean that a big TitanicC'mon, ya nice, sweet, I enjoy the thingBig ship 'pon di ocean that a big TitanicWoman, ya nice, sweet, fantasticBig ship 'pon di ocean that a big TitanicC'mon, ya nice, sweet, I enjoy the thingBig ship 'pon di ocean that a big TitanicWomanI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itWoman ya cute and you don't need no make upOriginal cute body you a mek man mud upWoman ya cute and you don't need no make upOriginal cute body you a mek man mud upEyeliner 'pon ya face a mek man mud upNose powder 'pon ya face a mek man mud upPluck ya eyebrow 'pon ya face a mek man mud upGal ya lipstick 'pon ya face a mek man mud upWoman, ya nice, broad faceAnd ya nice hipMake man flip and bust up them lipWoman, ya nice and energeticBig ship 'pon de ocean that a big TitanicWoman, ya nice, broad faceAnd ya nice hipMake man flip and bust up them lipWoman, ya nice and energeticBig ship 'pon de ocean that a big TitanicWomanI like to move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move itI like to move it, move itI like to move it, move itya like to move it"
	checks[1] = "\nI like to move itI like to move it, move itI like to move it, move it&amp; ya like to move itI like to move it, move itI like to move itI like to move it, move it&amp; ya like to move itI like to move it, move itI like to move itI like to move it, move itYa like to move itI like to move it, move itI like to move itI like to move it, move it&amp; ya like to move itAll girls all over the worldOriginal Mad Stuntman pon' ya case manI love how all girls a move them bodyAnd when ya move ya bodyGonna move it nice and sweet and sexy, alright?&#x9;Woman ya cute and you don't need no make upOriginal cute body you a mek man mud up&#x9;Woman ya cute and you don't need no make upOriginal cute body you a mek man mud up&#x9;WomanMan&lt;&gt; physically fit, physically fitPhysically, physically, physicallyWoman, physically fit, physically fitPhysically, physically, physically&#x9;WomanMan&lt;&gt; &#x9;WomanMan&lt;&gt;, ya nice, sweet, fantasticBig ship 'pon di ocean that a big TitanicC'mon, ya nice, sweet, I enjoy the thingBig ship 'pon di ocean that a big TitanicWoman, ya nice, sweet, fantasticBig ship 'pon di ocean that a big TitanicC'mon, ya nice, sweet, I enjoy the thingBig ship 'pon di ocean that a big TitanicWomanI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itWoman ya cute and you don't need no make upOriginal cute body you a mek man mud upWoman ya cute and you don't need no make upOriginal cute body you a mek man mud upEyeliner 'pon ya face a mek man mud upNose powder 'pon ya face a mek man mud upPluck ya eyebrow 'pon ya face a mek man mud upGal ya lipstick 'pon ya face a mek man mud upWoman, ya nice, broad faceAnd ya nice hipMake man flip and bust up them lipWoman, ya nice and energeticBig ship 'pon de ocean that a big TitanicWoman, ya nice, broad faceAnd ya nice hipMake man flip and bust up them lipWoman, ya nice and energeticBig ship 'pon de ocean that a big TitanicWomanI like to move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move itI like to move it, move itI like to move it, move it&amp; ya like to move it"
	checks[2] = "\nI like to move itI like to move it, move itI like to move it, move it&amp; ya like to move itI like to move it, move itI like to move itI like to move it, move it&amp; ya like to move itI like to move it, move itI like to move itI like to move it, move itYa like to move itI like to move it, move itI like to move itI like to move it, move it&amp; ya like to move itAll girls all over the worldOriginal Mad Stuntman pon' ya case manI love how all girls a move them bodyAnd when ya move ya bodyGonna move it nice and sweet and sexy, alright?Woman ya cute and you don't need no make upOriginal cute body you a mek man mud upWoman ya cute and you don't need no make upOriginal cute body you a mek man mud upWomanMan&lt;&gt; physically fit, physically fitPhysically, physically, physicallyWoman, physically fit, physically fitPhysically, physically, physicallyWomanMan&lt;&gt; WomanMan&lt;&gt;, ya nice, sweet, fantasticBig ship 'pon di ocean that a big TitanicC'mon, ya nice, sweet, I enjoy the thingBig ship 'pon di ocean that a big TitanicWoman, ya nice, sweet, fantasticBig ship 'pon di ocean that a big TitanicC'mon, ya nice, sweet, I enjoy the thingBig ship 'pon di ocean that a big TitanicWomanI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itWoman ya cute and you don't need no make upOriginal cute body you a mek man mud upWoman ya cute and you don't need no make upOriginal cute body you a mek man mud upEyeliner 'pon ya face a mek man mud upNose powder 'pon ya face a mek man mud upPluck ya eyebrow 'pon ya face a mek man mud upGal ya lipstick 'pon ya face a mek man mud upWoman, ya nice, broad faceAnd ya nice hipMake man flip and bust up them lipWoman, ya nice and energeticBig ship 'pon de ocean that a big TitanicWoman, ya nice, broad faceAnd ya nice hipMake man flip and bust up them lipWoman, ya nice and energeticBig ship 'pon de ocean that a big TitanicWomanI like to move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move it, move itI like to move it, move itI like to move it, move itYa like to move itI like to move itI like to move it, move itI like to move it, move it&amp; ya like to move it"
	var testData [2]map[string]string
	testData[0] = map[string]string{
		"iliketo":  "I like to move it",
		"yaliketo": "ya like to move it",
		"subject":  "Woman",
		"object":   "Man",
	}
	testData[1] = map[string]string{
		"iliketo":  "I like to move it",
		"yaliketo": "& ya like to move it",
		"subject":  "\tWoman",
		"object":   "Man<>",
	}
	for testNb, check := range checks {
		doc, err := New("testdata/iliketomoveit.docx")
		if err != nil {
			t.Fatalf("Error: %s\n", err.Error())
		}
		switch testNb {
		case 0:
			err = doc.replace(testData[0])
		case 1:
			err = doc.Replace(testData[1])
		case 2:
			testDataRepls := map[string]Replacement{}
			for k, v := range testData[1] {
				replaced := true
				if k == "subject" {
					replaced = false
				}
				testDataRepls[k] = Replacement{ v, replaced}
			}
			err = doc.ReplaceCond(testDataRepls)
		}
		if err != nil {
			t.Fatalf("Error: %s\n", err.Error())
		}
		doc.CreateNewDocx()
		//doc.SaveDocxToFile("/tmp/iliketo.docx")
		extractedText := extractRawTextFromDocxXml(doc.documentPartsReplaced["word/document.xml"])
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
}

func extractRawTextFromDocxXml(document string) string {
	// from commandlinefu
	// sed -e 's/<\/w:p>/\n/g; s/<[^>]\{1,\}>//g; s/[^[:print:]\n]\{1,\}//g'
	first := regexp.MustCompile("<[^>]+>")
	second := regexp.MustCompilePOSIX("[^[:print:]\n]+")
	result := first.ReplaceAllString(document, "")
	return second.ReplaceAllString(result, "")
}
