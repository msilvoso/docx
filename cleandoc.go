package docx

import (
	"fmt"
	"regexp"
	"strings"
)

var cleanRegex = regexp.MustCompile("<.*?>")

func cleanDocXml(content string) string {
	replacements := findPlaceholdersAndRemoveWordTags(content)

	cleanContent := replacePlaceholdersByCleanVersion(replacements, content)

	// From: https://github.com/PHPOffice/PHPWord/issues/590
	// $this->tempDocumentMainPart = preg_replace('/(<w:t*)>/', '$1 xml:space="preserve">', $this->tempDocumentMainPart);
	preserveSpaces := regexp.MustCompile("(<w:t*?)>")
	cleanContent = preserveSpaces.ReplaceAllString(cleanContent, `$1 xml:space="preserve">`)

	return cleanContent
}

func findPlaceholdersAndRemoveWordTags(content string) map[string]string {
	replacements := make(map[string]string)
	// matchRegexes := []string{`\{(\{|[^{]*\>)\{[^}]*\}(\}|[^{}]*\>)\}`} // find {{..}}
	matchRegexes := []string{ // find {{..}}
		`\}<[^{}]*?\}`, // clean between first } and second }
		`\{<[^{}]*?\{`, // clean between first { and second {
		`\{[^{}]*?\}`,   // clean between second { and first }
	}
	for _, matchRegex := range matchRegexes {
		regex := regexp.MustCompile(matchRegex)
		matches := regex.FindAllString(content, -1)
		for _, match := range matches {
			cleanString := cleanRegex.ReplaceAllString(match, "")
			replacements[match] = fmt.Sprintf("%s", cleanString) // remplace {...} by {{...}}
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
