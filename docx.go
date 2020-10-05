package docx

import (
	"archive/zip"
	"bytes"
	"io/ioutil"
	"strings"
	"text/template"
)

type Docx struct {
	ref      *zip.ReadCloser
	document string
	headers  []string
	footers  []string
	replaced string
	result   []byte
}

// load the docx file an extract the document.xml file
func (d *Docx) LoadDocx(path string) error {
	var err error
	// Open a zip archive for reading.
	d.ref, err = zip.OpenReader(path)
	if err != nil {
		return err
	}
	// extract document.xml
	// Iterate through the files in the archive
	for _, f := range d.ref.File {
		if f.FileHeader.Name == "word/document.xml" {
			readCloser, err := f.Open()
			if err != nil {
				return err
			}
			buf := new(bytes.Buffer)
			buf.ReadFrom(readCloser)
			d.document = cleanDocXml(string(buf.Bytes()))
			readCloser.Close()
			break
		}
	}
	return nil
}

// replace placeholders
func (d *Docx) Replace(replacements map[string]string) (err error) {
	buf := strings.Builder{}
	tmpl, err := template.New("docx").Option("missingkey=zero").Parse(d.document)
	if err != nil {
		return err
	}
	err = tmpl.Execute(&buf, replacements)
	if err != nil {
		return err
	}
	d.replaced = buf.String()
	return nil
}

func (d *Docx) CreateNewDocx() error {
	buf := new(bytes.Buffer)
	w := zip.NewWriter(buf)
	// Iterate through the files in the archive and save every file
	for _, f := range d.ref.File {
		fWriter, err := w.Create(f.Name)
		if err != nil {
			return err
		}
		switch f.Name {
		case "word/document.xml":
			fWriter.Write([]byte(d.replaced))
		default:
			readCloser, err := f.Open()
			if err != nil {
				return err
			}
			b := new(bytes.Buffer)
			b.ReadFrom(readCloser)
			fWriter.Write(b.Bytes())
			readCloser.Close()
		}
	}
	w.Close()
	d.result = buf.Bytes()
	return nil
}

func (d *Docx) NewDocx() []byte {
	d.CreateNewDocx()
	return d.result
}

func (d *Docx) SaveDocxToFile(path string) error {
	return ioutil.WriteFile(path, d.result, 0644)
}

func (d *Docx) Close() {
	d.ref.Close()
}
