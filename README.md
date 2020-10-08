# docx
Docx documents template processor (not old format doc files)

Uses [text/template](https://golang.org/pkg/text/template/ "Golang text/template") to do replacements

Example:
```go
package main

import (
	"github.com/msilvoso/docx"
        "log"
)

func main() {
    replacements := map[string]string {
        "placeholder1": "Hello",
        "placeholder2": "World",
        "placeholder3": "!",
    }   

    // load word document
    d, err := docx.New("wordfile.docx")
    if err != nil {
        log.Fatalln("%s\n", err.Error())
    }
    
    // replace placeholders
    d.Replace(replacements)

    // Save resulting docx to file
    d.SaveDocxToFile("replaced.docx")
 
    d.Close()
}
```
*{{.placeholder1}} {{.placeholder2}}{{.placeholder3}}*

becomes

*Hello World!*

in the word docx file