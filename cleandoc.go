package docx

import (
	"fmt"
	"regexp"
	"strings"
)

// find {{...}} placeholder and remove unneeded xml tags
var cleanupRegexes = []string{ // find {{..}}
	`}<[^{}]*?}`,  // clean between first } and second }
	`{<[^{}]*?{`,  // clean between first { and second {
	`{{[^{}]*?}}`, // clean between second { and first }
}

// placeholder is {{..}}
var macrosRegexp = regexp.MustCompile(`{{.*?}}`)
var needsSplittingRegexp = regexp.MustCompile(`[^>]{{|}}[^<]`)
var splitTextIntoTextsOpening = [2]string{`{{`, `</w:t></w:r><w:r>%s<w:t xml:space="preserve">{{`}
var splitTextIntoTextsClosing = [2]string{`}}`, `}}</w:t></w:r><w:r>%s<w:t xml:space="preserve">`}

// function to clean the docx xml
// Shamelesly stolen from: https://github.com/PHPOffice/PHPWord
func cleanDocXml(content string) string {
	cleanContent := findPlaceholdersAndRemoveWordTags(content)

	for _, find := range findContainingXmlBlockForMacro(cleanContent) {
		replace := splitTextIntoTexts(find)
		cleanContent = strings.ReplaceAll(cleanContent, find, replace)
	}
	return cleanContent
}

// findPlaceholdersAndRemoveWordTags removes unneeded xml tags in the placeholders
// Word tends to put a lot of xml tags in the placeholders during edits of the document
// thus making it impossible to do sensible replacements by the templating engine
func findPlaceholdersAndRemoveWordTags(content string) string {
	var cleanRegex = regexp.MustCompile("<.*?>")

	for _, matchRegex := range cleanupRegexes {
		regex := regexp.MustCompile(matchRegex)
		matches := regex.FindAllString(content, -1)
		for _, match := range matches {
			cleanString := cleanRegex.ReplaceAllString(match, "")
			content = strings.ReplaceAll(content, match, cleanString)
		}
	}
	return content
}

// findContainingXmlBlockForMacro has been translated as is from the PhpWord project
func findContainingXmlBlockForMacro(text string) []string {
	blocks := regexp.MustCompile(`(?i)<w:r[ >].*?</w:r>`)

	var result []string
	for _, b := range blocks.FindAllString(text, -1) {
		if macrosRegexp.MatchString(b) {
			result = append(result, b)
		}
	}
	return result
}

// splitTextIntoTexts isolates placeholders in their own <w:t> element
// translated from PhpWord project
// and adds xml:space="preserve" in case the replacement contains spaces
// See https://github.com/PHPOffice/PHPWord/issues/590
// returns "true" as second value when something needs replacing
func splitTextIntoTexts(text string) string {
	if !needsSplittingRegexp.MatchString(text) {
		return text
	}
	extractedStyles := regexp.MustCompile(`(?i)<w:rPr.*?</w:rPr>`)
	styleMatches := extractedStyles.FindAllString(text, -1)
	var style string
	if len(styleMatches) > 0 {
		style = styleMatches[0]
	}
	nonformatted := regexp.MustCompile(`>\s+<`)
	nonformattedText := nonformatted.ReplaceAllString(text, `><`)
	result := strings.ReplaceAll(nonformattedText, splitTextIntoTextsOpening[0], fmt.Sprintf(splitTextIntoTextsOpening[1], style))
	result = strings.ReplaceAll(result, splitTextIntoTextsClosing[0], fmt.Sprintf(splitTextIntoTextsClosing[1], style))
	result = strings.ReplaceAll(result, fmt.Sprintf("<w:r>%s<w:t xml:space=\"preserve\"></w:t></w:r>", style), "")
	result = strings.ReplaceAll(result, "<w:r><w:t xml:space=\"preserve\"></w:t></w:r>", "")
	result = strings.ReplaceAll(result, "<w:t>", "<w:t xml:space=\"preserve\">")

	return result
}
