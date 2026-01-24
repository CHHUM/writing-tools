/**
 * Google Docs Formatter & Link Checker
 *
 * Provides tools for maintaining document quality including link checking,
 * formatting consistency, and text case transformations.
 *
 * LINK CHECKING:
 * - Checks all HTTP/HTTPS links for broken links (404 errors)
 * - Highlights broken links in red
 * - Highlights Apple.com links in yellow
 * - Detects underlined text without links (possible missing links) in purple
 * - Automatically formats all links with proper blue text and underline
 *
 * Valid non-HTTP protocols are skipped (not flagged as errors):
 * - mailto: (email links)
 * - tel: (phone numbers)
 * - sms: (SMS links)
 * - ftp: / sftp: (file transfer)
 * - file: (local files)
 *
 * Invalid/malformed links are flagged in orange:
 * - Links that don't start with any recognized protocol
 * - Typos like htp:// or htps://
 * - Weird formats or broken URLs
 *
 * DOCUMENT FORMATTING:
 * - Applies consistent typography throughout the document
 * - Heading 1: Helvetica Neue Bold 24pt
 * - Heading 2: Helvetica Neue Bold 14pt
 * - Normal Text: Helvetica Neue 11pt
 *
 * TEXT CASE TOOLS:
 * - Lower case: Converts selected text to lowercase
 * - Upper case: Converts selected text to uppercase
 * - Initial Caps: Capitalizes the first letter of each word
 * - Sentence case: Capitalizes only the first letter of the selection
 * - Title Case (Chicago Style): Proper title capitalization following Chicago Manual of Style
 *
 * NOTES:
 * - The script may take a while if you have many links, as it checks each one individually
 * - Google Apps Script has rate limits on external URL fetches
 * - For documents with 100+ links, you might need to run it multiple times or add delays
 *
 * HOW TO SET IT UP:
 * 1. Open your Google Doc
 * 2. Go to Extensions > Apps Script
 * 3. Delete any existing code and paste this script
 * 4. Click the Save icon (💾)
 * 5. Close the Apps Script tab and refresh your Google Doc
 * 6. You'll now see a "Document Tools" menu in your document
 *
 * TO USE IT:
 * - Click Document Tools > Check Links > In Entire Document (or In Active Tab)
 * - Click Document Tools > Fix Document Formatting
 * - Select text and use Document Tools > Text Case submenu for case changes
 */

// ============================================================================
// CONFIGURATION CONSTANTS
// ============================================================================

const COLORS = {
    BROKEN_LINK: '#FF0000',        // Red
    APPLE_LINK: '#FFFF00',         // Yellow
    INVALID_LINK: '#FFA500',       // Orange
    MISSING_LINK: '#DDA0DD',       // Purple/Plum
    LINK_BLUE: '#1155CC'           // Google Docs standard link color
};

const VALID_NON_HTTP_PROTOCOLS = ['mailto:', 'tel:', 'sms:', 'ftp:', 'sftp:', 'file:'];

const DOCUMENT_STYLES = {
    heading1: {
        [DocumentApp.Attribute.FONT_FAMILY]: 'Helvetica Neue',
        [DocumentApp.Attribute.FONT_SIZE]: 24,
        [DocumentApp.Attribute.BOLD]: true
    },
    heading2: {
        [DocumentApp.Attribute.FONT_FAMILY]: 'Helvetica Neue',
        [DocumentApp.Attribute.FONT_SIZE]: 14,
        [DocumentApp.Attribute.BOLD]: true
    },
    normal: {
        [DocumentApp.Attribute.FONT_FAMILY]: 'Helvetica Neue',
        [DocumentApp.Attribute.FONT_SIZE]: 11,
        [DocumentApp.Attribute.BOLD]: false
    }
};

const TITLE_CASE_LOWERCASE_WORDS = new Set([
    'a', 'an', 'the',  // articles
    'and', 'but', 'or', 'nor', 'for', 'so', 'yet',  // coordinating conjunctions
    'as', 'at', 'by', 'for', 'from', 'in', 'into', 'of', 'off', 'on',
    'onto', 'out', 'over', 'to', 'up', 'with', 'via'  // common prepositions
]);

// ============================================================================
// UI & MENU FUNCTIONS
// ============================================================================

/**
 * Creates custom menu when document is opened
 */
function onOpen() {
    DocumentApp.getUi()
        .createMenu('Document Tools')
        .addSubMenu(DocumentApp.getUi().createMenu('Check Links')
            .addItem('In Active Tab', 'checkLinksInActiveTab')
            .addItem('In Entire Document', 'checkLinksInDocument'))
        .addSeparator()
        .addItem('Fix Document Formatting', 'fixDocumentFormatting')
        .addSeparator()
        .addSubMenu(DocumentApp.getUi().createMenu('Text Case')
            .addItem('Lower case', 'convertToLowerCase')
            .addItem('Upper case', 'convertToUpperCase')
            .addItem('Initial Caps', 'convertToInitialCaps')
            .addItem('Sentence case', 'convertToSentenceCase')
            .addItem('Title Case (Chicago Style)', 'convertToTitleCase'))
        .addToUi();
}

function showAlert(title, message) {
    DocumentApp.getUi().alert(title, message, DocumentApp.getUi().ButtonSet.OK);
}

// ============================================================================
// CORE UTILITY FUNCTIONS
// ============================================================================

/**
 * Gets the document body based on scope (tab or entire document)
 */
function getDocumentBody(scope) {
    const doc = DocumentApp.getActiveDocument();

    if (scope === 'tab') {
        const activeTab = doc.getActiveTab();
        if (!activeTab) {
            showAlert(
                'No Active Tab',
                'Could not find an active tab. The document might not have tabs, or you might need to click into a tab first.'
            );
            return null;
        }
        return activeTab.asDocumentTab().getBody();
    }

    return doc.getBody();
}

/**
 * Recursively extracts all links from document elements
 */
function extractLinks(element, links) {
    if (element.getType() === DocumentApp.ElementType.TEXT) {
        const text = element.asText();
        const textString = text.getText();
        const indices = text.getTextAttributeIndices();

        for (let i = 0; i < indices.length; i++) {
            const startOffset = indices[i];
            const endOffset = i + 1 < indices.length ? indices[i + 1] : textString.length;
            const url = text.getLinkUrl(startOffset);

            if (url) {
                links.push({
                    url: url,
                    startOffset: startOffset,
                    endOffset: endOffset,
                    element: text
                });
            }
        }
    }

    // Recursively process child elements
    if (element.getNumChildren) {
        const numChildren = element.getNumChildren();
        for (let i = 0; i < numChildren; i++) {
            extractLinks(element.getChild(i), links);
        }
    }
}

/**
 * Recursively extracts underlined text that is NOT a link
 */
function extractUnderlinedText(element, underlinedText) {
    if (element.getType() === DocumentApp.ElementType.TEXT) {
        const text = element.asText();
        const textString = text.getText();
        const indices = text.getTextAttributeIndices();

        for (let i = 0; i < indices.length; i++) {
            const startOffset = indices[i];
            const endOffset = i + 1 < indices.length ? indices[i + 1] : textString.length;

            const isUnderlined = text.isUnderline(startOffset);
            const hasLink = text.getLinkUrl(startOffset) !== null;

            if (isUnderlined && !hasLink) {
                const textContent = textString.substring(startOffset, endOffset).trim();
                if (textContent.length > 0) {
                    underlinedText.push({
                        startOffset: startOffset,
                        endOffset: endOffset,
                        element: text
                    });
                }
            }
        }
    }

    // Recursively process child elements
    if (element.getNumChildren) {
        const numChildren = element.getNumChildren();
        for (let i = 0; i < numChildren; i++) {
            extractUnderlinedText(element.getChild(i), underlinedText);
        }
    }
}

/**
 * Processes selected text with a transformation function while preserving formatting
 */
function processSelectedText(transformFn, errorMessage = 'Please select some text first.') {
    const selection = DocumentApp.getActiveDocument().getSelection();

    if (!selection) {
        showAlert('No Selection', errorMessage);
        return;
    }

    const elements = selection.getRangeElements();

    for (let i = 0; i < elements.length; i++) {
        const element = elements[i];

        if (element.getElement().editAsText) {
            const text = element.getElement().editAsText();
            const startOffset = element.isPartial() ? element.getStartOffset() : 0;
            const endOffset = element.isPartial() ? element.getEndOffsetInclusive() : text.getText().length - 1;

            const selectedText = text.getText().substring(startOffset, endOffset + 1);
            const transformedText = transformFn(selectedText);

            // Update character by character to preserve formatting
            for (let j = 0; j <= endOffset - startOffset; j++) {
                const pos = startOffset + j;
                const originalChar = selectedText.charAt(j);
                const newChar = transformedText.charAt(j);

                if (originalChar !== newChar) {
                    text.deleteText(pos, pos);
                    text.insertText(pos, newChar);
                }
            }
        }
    }
}

// ============================================================================
// LINK CHECKING FUNCTIONS
// ============================================================================

/**
 * Applies proper link formatting (blue text with underline)
 * Returns true if formatting was needed, false if already correct
 */
function applyLinkFormatting(element, startOffset, endOffset) {
    const currentColor = element.getForegroundColor(startOffset);
    const currentUnderline = element.isUnderline(startOffset);

    const needsFormatting = currentColor !== COLORS.LINK_BLUE || !currentUnderline;

    element.setForegroundColor(startOffset, endOffset - 1, COLORS.LINK_BLUE);
    element.setUnderline(startOffset, endOffset - 1, true);

    return needsFormatting;
}

/**
 * Checks if a link is broken (returns 404)
 */
function isLinkBroken(url) {
    try {
        const response = UrlFetchApp.fetch(url, {
            'muteHttpExceptions': true,
            'followRedirects': false
        });
        return response.getResponseCode() === 404;
    } catch (e) {
        return true; // Consider unfetchable URLs as broken
    }
}

/**
 * Processes a single link and returns classification
 */
function processLink(link) {
    const { url, startOffset, endOffset, element } = link;
    const urlLower = url.toLowerCase();

    // Apply proper link formatting
    const needsFormatting = applyLinkFormatting(element, startOffset, endOffset);

    // Check if it's a valid non-HTTP protocol
    const isValidNonHttp = VALID_NON_HTTP_PROTOCOLS.some(protocol => urlLower.startsWith(protocol));

    if (isValidNonHttp) {
        return { type: 'skipped', needsFormatting };
    }

    // Check if it's HTTP/HTTPS
    const isHttp = urlLower.startsWith('http://') || urlLower.startsWith('https://');

    if (!isHttp) {
        element.setBackgroundColor(startOffset, endOffset - 1, COLORS.INVALID_LINK);
        return { type: 'invalid', needsFormatting };
    }

    // Check if it's an Apple.com link
    if (urlLower.includes('apple.com')) {
        element.setBackgroundColor(startOffset, endOffset - 1, COLORS.APPLE_LINK);

        // Also check if it's broken
        if (isLinkBroken(url)) {
            element.setBackgroundColor(startOffset, endOffset - 1, COLORS.BROKEN_LINK);
            return { type: 'broken', needsFormatting };
        }

        return { type: 'apple', needsFormatting };
    }

    // Check if link is broken
    if (isLinkBroken(url)) {
        element.setBackgroundColor(startOffset, endOffset - 1, COLORS.BROKEN_LINK);
        return { type: 'broken', needsFormatting };
    }

    return { type: 'valid', needsFormatting };
}

/**
 * Main link checking function
 */
function checkLinks(scope) {
    const body = getDocumentBody(scope);
    if (!body) return;

    const scopeText = scope === 'tab' ? ' in active tab' : ' in document';

    // Extract all links and underlined text
    const links = [];
    const underlinedText = [];
    const numChildren = body.getNumChildren();

    for (let i = 0; i < numChildren; i++) {
        const child = body.getChild(i);
        extractLinks(child, links);
        extractUnderlinedText(child, underlinedText);
    }

    // Process links and count results
    const counts = {
        broken: 0,
        apple: 0,
        skipped: 0,
        invalid: 0,
        formatted: 0,
        missingLink: 0
    };

    links.forEach(link => {
        const result = processLink(link);
        counts[result.type]++;
        if (result.needsFormatting) counts.formatted++;
    });

    // Highlight underlined text without links
    underlinedText.forEach(item => {
        item.element.setBackgroundColor(item.startOffset, item.endOffset - 1, COLORS.MISSING_LINK);
        counts.missingLink++;
    });

    // Build result message
    let message = `Found ${counts.broken} broken link(s) (highlighted in red)\n` +
        `Found ${counts.apple} Apple.com link(s) (highlighted in yellow)\n` +
        `Skipped ${counts.skipped} valid non-HTTP link(s) (email, phone, etc.)\n` +
        `Fixed formatting on ${counts.formatted} link(s)${scopeText}`;

    if (counts.invalid > 0) {
        message += `\nFound ${counts.invalid} invalid/malformed link(s) (highlighted in orange)`;
    }

    if (counts.missingLink > 0) {
        message += `\nFound ${counts.missingLink} underlined text(s) without links (highlighted in purple)`;
    }

    showAlert('Link Check Complete', message);
}

function checkLinksInDocument() {
    checkLinks('document');
}

function checkLinksInActiveTab() {
    checkLinks('tab');
}

// ============================================================================
// DOCUMENT FORMATTING FUNCTIONS
// ============================================================================

/**
 * Applies consistent typography throughout the document
 * Heading 1: Helvetica Neue Bold 24pt
 * Heading 2: Helvetica Neue Bold 14pt
 * Normal Text: Helvetica Neue 11pt
 */
function fixDocumentFormatting() {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    const counts = { h1: 0, h2: 0, normal: 0 };
    const numChildren = body.getNumChildren();

    for (let i = 0; i < numChildren; i++) {
        const child = body.getChild(i);

        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const paragraph = child.asParagraph();
            const heading = paragraph.getHeading();

            if (heading === DocumentApp.ParagraphHeading.HEADING1) {
                paragraph.setAttributes(DOCUMENT_STYLES.heading1);
                counts.h1++;
            } else if (heading === DocumentApp.ParagraphHeading.HEADING2) {
                paragraph.setAttributes(DOCUMENT_STYLES.heading2);
                counts.h2++;
            } else if (heading === DocumentApp.ParagraphHeading.NORMAL) {
                paragraph.setAttributes(DOCUMENT_STYLES.normal);
                counts.normal++;
            }
        }
    }

    const message = `Document formatting updated:\n` +
        `${counts.h1} Heading 1 paragraph(s) formatted (Helvetica Neue Bold 24pt)\n` +
        `${counts.h2} Heading 2 paragraph(s) formatted (Helvetica Neue Bold 14pt)\n` +
        `${counts.normal} Normal text paragraph(s) formatted (Helvetica Neue 11pt)`;

    showAlert('Formatting Complete', message);
}

// ============================================================================
// TEXT CASE TRANSFORMATION FUNCTIONS
// ============================================================================

function convertToLowerCase() {
    processSelectedText(text => text.toLowerCase());
}

function convertToUpperCase() {
    processSelectedText(text => text.toUpperCase());
}

function convertToInitialCaps() {
    processSelectedText(text => {
        return text.replace(/\b\w/g, char => char.toUpperCase());
    });
}

function convertToSentenceCase() {
    processSelectedText(text => {
        const lower = text.toLowerCase();
        return lower.length > 0 ? lower.charAt(0).toUpperCase() + lower.slice(1) : lower;
    });
}

function convertToTitleCase() {
    processSelectedText(text => {
        const words = text.split(/(\s+|[-—:])/);
        let afterColonOrDash = false;

        const titleCaseWords = words.map((word, index) => {
            // Preserve whitespace and punctuation
            if (/^\s+$/.test(word) || word === '-' || word === '—') {
                return word;
            }

            if (word === ':') {
                afterColonOrDash = true;
                return word;
            }

            // Extract word without leading/trailing punctuation
            const match = word.match(/^(\W*)(\w+)(\W*)$/);
            if (!match) return word;

            const [, prefix, actualWord, suffix] = match;
            const lowerWord = actualWord.toLowerCase();

            // Determine if we should capitalize
            const isFirstWord = words.slice(0, index).every(w => /^\s+$/.test(w) || w === '-' || w === '—' || w === ':');
            const isLastWord = words.slice(index + 1).every(w => /^\s+$/.test(w) || w === '-' || w === '—' || w === ':');

            const shouldCapitalize = isFirstWord || isLastWord || afterColonOrDash ||
                !TITLE_CASE_LOWERCASE_WORDS.has(lowerWord);

            if (afterColonOrDash && /\w/.test(word)) {
                afterColonOrDash = false;
            }

            const capitalizedWord = shouldCapitalize
                ? lowerWord.charAt(0).toUpperCase() + lowerWord.slice(1)
                : lowerWord;

            return prefix + capitalizedWord + suffix;
        });

        return titleCaseWords.join('');
    });
}