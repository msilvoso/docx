package docx

import (
	"fmt"
	"regexp"
	"strings"
)

// function to clean the docx xml
// Shamelesly stolen from: https://github.com/PHPOffice/PHPWord
func cleanDocXml(content string) string {
	replacements := findPlaceholdersAndRemoveWordTags(content)

	cleanContent := replacePlaceholdersByCleanVersion(replacements, content)

	blocks := map[string]string{}
	for _, s := range findContainingXmlBlockForMacro(cleanContent) {
		blocks[s] = splitTextIntoTexts(s)
	}
	processedContent := replacePlaceholdersByCleanVersion(blocks, cleanContent)

	return processedContent
}

func findPlaceholdersAndRemoveWordTags(content string) map[string]string {
	var cleanRegex = regexp.MustCompile("<.*?>")

	replacements := make(map[string]string)
	matchRegexes := []string{ // find {{..}}
		`}<[^{}]*?}`, // clean between first } and second }
		`{<[^{}]*?{`, // clean between first { and second {
		`{[^{}]*?}`,  // clean between second { and first }
	}
	for _, matchRegex := range matchRegexes {
		regex := regexp.MustCompile(matchRegex)
		matches := regex.FindAllString(content, -1)
		for _, match := range matches {
			cleanString := cleanRegex.ReplaceAllString(match, "")
			replacements[match] = fmt.Sprintf("%s", cleanString)
		}
	}
	return replacements
}

func replacePlaceholdersByCleanVersion(replacements map[string]string, content string) string {
	for find, replace := range replacements {
		content = strings.ReplaceAll(content, find, replace)
	}
	return content
}

func findContainingXmlBlockForMacro(text string) []string {
	blocks := regexp.MustCompile(`(?i)<w:r[ >].*?</w:r>`)
	macros := regexp.MustCompile(`{{.*?}}`)
	var result []string
	for _, b := range blocks.FindAllString(text, -1) {
		if macros.MatchString(b) {
			result = append(result, b)
		}
	}
	return result
}

func splitTextIntoTexts(text string) string {
	if !textNeedsSplitting(text) {
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
	result := strings.ReplaceAll(nonformattedText, `{{`, fmt.Sprintf("</w:t></w:r><w:r>%s<w:t xml:space=\"preserve\">{{", style))
	result = strings.ReplaceAll(result, `}}`, fmt.Sprintf("}}</w:t></w:r><w:r>%s<w:t xml:space=\"preserve\">", style))
	result = strings.ReplaceAll(result, fmt.Sprintf("<w:r>%s<w:t xml:space=\"preserve\"></w:t></w:r>", style), "")
	result = strings.ReplaceAll(result, "<w:r><w:t xml:space=\"preserve\"></w:t></w:r>", "")
	result = strings.ReplaceAll(result, "<w:t>", "<w:t xml:space=\"preserve\">")

	return result
}

func textNeedsSplitting(text string) bool {
	needsSplitting := regexp.MustCompile(`[^>]{{|}}[^<]`)
	return needsSplitting.MatchString(text)
}
