package docx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io/ioutil"
	"regexp"
	"strings"
	"text/template"
)

type Docx struct {
	ref                   *zip.ReadCloser
	documentParts         map[string]string
	documentPartsReplaced map[string]string
	result                []byte
}

// load the docx file an extract the document.xml file
// TODO: replace file by io.reader?
func (d *Docx) LoadDocx(path string) error {
	var err error
	// docx documents are zip archives
	d.ref, err = zip.OpenReader(path)
	if err != nil {
		return err
	}
	// find the filenames of the document, header and footer files
	err = d.findXmlNames()
	if err != nil {
		return err
	}
	// extract the file parts
	// Iterate through the files in the archive
zipFileIter:
	for _, f := range d.ref.File {
		var content string
		for name, _ := range d.documentParts {
			if f.FileHeader.Name == name {
				content, err = extractFileFromZip(f)
				d.documentParts[name] = cleanDocXml(content)
				continue zipFileIter
			}
		}
		if err != nil {
			return err
		}
	}
	return nil
}

// replace placeholders using the text/template package
func (d *Docx) Replace(replacements map[string]string) (err error) {
	d.documentPartsReplaced = map[string]string{}
	for name, content := range d.documentParts {
		buf := strings.Builder{}
		tmpl, err := template.New("docx").Option("missingkey=zero").Parse(content)
		if err != nil {
			return err
		}
		err = tmpl.Execute(&buf, replacements)
		if err != nil {
			return err
		}
		d.documentPartsReplaced[name] = buf.String()
	}
	return nil
}

// create the resulting docx and store the byte slice to result
func (d *Docx) CreateNewDocx() error {
	buf := new(bytes.Buffer)
	w := zip.NewWriter(buf)
	// Iterate through the files in the archive and save every file
docxWrite:
	for _, f := range d.ref.File {
		fWriter, err := w.Create(f.Name)
		if err != nil {
			return err
		}
		for name, content := range d.documentPartsReplaced {
			if f.FileHeader.Name == name {
				fWriter.Write([]byte(content))
				continue docxWrite
			}
		}
		readCloser, err := f.Open()
		if err != nil {
			return err
		}
		b := new(bytes.Buffer)
		b.ReadFrom(readCloser)
		fWriter.Write(b.Bytes())
		readCloser.Close()
	}
	w.Close()
	d.result = buf.Bytes()
	return nil
}

// create the resulting docx, store the byte slice to result and return it
func (d *Docx) NewDocx() []byte {
	d.CreateNewDocx()
	return d.result
}

// Save the resulting docx to a file
func (d *Docx) SaveDocxToFile(path string) error {
	return ioutil.WriteFile(path, d.result, 0644)
}

func (d *Docx) Close() {
	d.ref.Close()
}

// find the internal xml document names
func (d *Docx) findXmlNames() error {
	d.documentParts = map[string]string{}
	documentRegex := regexp.MustCompile(`PartName="/(word/document.*?\.xml)`)
	headersRegex := regexp.MustCompile(`PartName="/(word/header.*?\.xml)`)
	footersRegex := regexp.MustCompile(`PartName="/(word/footer.*?\.xml)`)

	// iterate through all filenames in the zip archive and look for [Content_Types].xml
	var contentTypes string
	for _, f := range d.ref.File {
		if f.FileHeader.Name == "[Content_Types].xml" {
			contentTypes, _ = extractFileFromZip(f)
			break
		}
	}
	documentMatch := documentRegex.FindAllStringSubmatch(contentTypes, -1)
	if len(documentMatch) == 0 {
		return fmt.Errorf("no document.xml in the zip")
	}
	// set the name as index of the map
	d.documentParts[documentMatch[0][1]] = ""
	headersMatch := headersRegex.FindAllStringSubmatch(contentTypes, -1)
	footersMatch := footersRegex.FindAllStringSubmatch(contentTypes, -1)
	for _, m := range headersMatch {
		d.documentParts[m[1]] = ""
	}
	for _, m := range footersMatch {
		d.documentParts[m[1]] = ""
	}
	return nil
}

func extractFileFromZip(f *zip.File) (string, error) {
	readCloser, err := f.Open()
	if err != nil {
		return "", err
	}
	buf := new(bytes.Buffer)
	buf.ReadFrom(readCloser)
	content := cleanDocXml(string(buf.Bytes()))
	readCloser.Close()
	return content, nil
}
