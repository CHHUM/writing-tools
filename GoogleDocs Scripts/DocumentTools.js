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
 * - Trims leading/trailing spaces from linked text (spaces remain as normal text)
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
 * - Preserves character-level formatting (bold, underline, italic, etc.)
 * - Only updates elements that need changes
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
 * - Link formatting count only includes links that actually needed formatting changes
 * - Space trimming preserves the spaces in the document but removes them from the link
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
 *
 * VERSION: 2.1.0
 * CHANGELOG:
 * - v2.1.0: Fixed list item formatting, preserved character-level formatting,
 *           only updates elements that need changes
 * - v2.0.0: Added automatic trimming of leading/trailing spaces from links
 *           Fixed link formatting count to only include links that needed changes
 *           Improved accuracy of link formatting detection
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
 * Removes link formatting from leading/trailing spaces in linked text
 * Keeps the spaces in the document as normal text
 * Returns object with trimmedUrl and spacesTrimmed flag
 */
function trimLinkSpaces(link) {
    const { url, startOffset, endOffset, element } = link;

    // Get the actual text that is linked
    const fullText = element.getText();
    const linkedText = fullText.substring(startOffset, endOffset);

    // Find where the actual non-space content starts and ends
    let trimStart = 0;
    let trimEnd = linkedText.length;

    // Count leading spaces
    while (trimStart < linkedText.length && linkedText[trimStart] === ' ') {
        trimStart++;
    }

    // Count trailing spaces
    while (trimEnd > trimStart && linkedText[trimEnd - 1] === ' ') {
        trimEnd--;
    }

    // If there are no spaces to trim, return early
    if (trimStart === 0 && trimEnd === linkedText.length) {
        return { trimmedUrl: url, spacesTrimmed: false };
    }

    // Calculate the new link boundaries (excluding spaces)
    const newStartOffset = startOffset + trimStart;
    const newEndOffset = startOffset + trimEnd;

    // Only proceed if we actually have content left after trimming
    if (newStartOffset >= newEndOffset) {
        return { trimmedUrl: url, spacesTrimmed: false };
    }

    // Remove the link from the entire original range
    element.setLinkUrl(startOffset, endOffset - 1, null);

    // Remove underline from the spaces (but keep the spaces themselves)
    if (trimStart > 0) {
        // Remove underline from leading spaces
        element.setUnderline(startOffset, newStartOffset - 1, false);
    }
    if (trimEnd < linkedText.length) {
        // Remove underline from trailing spaces
        element.setUnderline(newEndOffset, endOffset - 1, false);
    }

    // Re-apply the link only to the trimmed text (not the spaces)
    element.setLinkUrl(newStartOffset, newEndOffset - 1, url);

    // Update the link object for subsequent processing
    link.startOffset = newStartOffset;
    link.endOffset = newEndOffset;

    return { trimmedUrl: url, spacesTrimmed: true };
}

/**
 * Applies proper link formatting (blue text with underline)
 * Returns true if formatting was needed, false if already correct
 */
function applyLinkFormatting(element, startOffset, endOffset) {
    // Check the current formatting
    let needsFormatting = false;

    // Check a few points across the range to see if formatting is needed
    const checkPoints = [startOffset, Math.floor((startOffset + endOffset) / 2), endOffset - 1];

    for (const pos of checkPoints) {
        if (pos >= startOffset && pos < endOffset) {
            const currentColor = element.getForegroundColor(pos);
            const currentUnderline = element.isUnderline(pos);

            // Check if this position has correct formatting
            const isCorrectColor = currentColor === COLORS.LINK_BLUE ||
                currentColor === '#1155cc' ||
                currentColor === '#1155CC';

            if (!isCorrectColor || !currentUnderline) {
                needsFormatting = true;
                break;
            }
        }
    }

    // Only apply formatting if it's actually needed
    if (needsFormatting) {
        element.setForegroundColor(startOffset, endOffset - 1, COLORS.LINK_BLUE);
        element.setUnderline(startOffset, endOffset - 1, true);
    }

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
 * Validates and formats a link (without trimming spaces - that's done separately)
 */
function processLinkValidation(link) {
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

    // FIRST PASS: Trim spaces from links (in reverse order to preserve offsets)
    let spacesTrimmedCount = 0;
    for (let i = links.length - 1; i >= 0; i--) {
        const link = links[i];
        const { spacesTrimmed } = trimLinkSpaces(link);
        if (spacesTrimmed) {
            spacesTrimmedCount++;
        }
    }

    // SECOND PASS: Process links for validation and formatting
    const counts = {
        broken: 0,
        apple: 0,
        skipped: 0,
        invalid: 0,
        formatted: 0,
        missingLink: 0,
        spacesTrimmed: spacesTrimmedCount
    };

    links.forEach(link => {
        const result = processLinkValidation(link);
        counts[result.type]++;
        if (result.needsFormatting) {
            counts.formatted++;
        }
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

    if (counts.spacesTrimmed > 0) {
        message += `\nTrimmed spaces from ${counts.spacesTrimmed} link(s)`;
    }

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
 * Checks if an element needs formatting by comparing current font settings
 */
function needsFormatting(element, targetStyle) {
    const currentFont = element.getAttributes()[DocumentApp.Attribute.FONT_FAMILY];
    const currentSize = element.getAttributes()[DocumentApp.Attribute.FONT_SIZE];

    // For headings that should be bold, check if paragraph-level bold is set
    const targetBold = targetStyle[DocumentApp.Attribute.BOLD];
    const currentBold = element.getAttributes()[DocumentApp.Attribute.BOLD];

    return currentFont !== targetStyle[DocumentApp.Attribute.FONT_FAMILY] ||
        currentSize !== targetStyle[DocumentApp.Attribute.FONT_SIZE] ||
        (targetBold !== undefined && currentBold !== targetBold);
}

/**
 * Applies font formatting while preserving character-level formatting
 */
function applyFontFormatting(element, style) {
    const text = element.editAsText();
    const textContent = text.getText();

    // Apply paragraph-level font family and size
    element.setAttributes({
        [DocumentApp.Attribute.FONT_FAMILY]: style[DocumentApp.Attribute.FONT_FAMILY],
        [DocumentApp.Attribute.FONT_SIZE]: style[DocumentApp.Attribute.FONT_SIZE]
    });

    // For headings that should be bold, apply bold at paragraph level
    if (style[DocumentApp.Attribute.BOLD] === true) {
        element.setAttributes({
            [DocumentApp.Attribute.BOLD]: true
        });
    } else if (style[DocumentApp.Attribute.BOLD] === false) {
        // For normal text, remove paragraph-level bold but preserve character-level bold
        // We need to check each character and preserve its individual bold state
        const indices = text.getTextAttributeIndices();

        // First, remove paragraph-level bold
        element.setAttributes({
            [DocumentApp.Attribute.BOLD]: false
        });

        // Then restore character-level bold where it existed
        for (let i = 0; i < indices.length; i++) {
            const startOffset = indices[i];
            const endOffset = i + 1 < indices.length ? indices[i + 1] - 1 : textContent.length - 1;

            // Get the bold state before we changed paragraph formatting
            // We need to re-check after setting paragraph attributes
            const wasBold = text.isBold(startOffset);
            if (wasBold && endOffset >= startOffset) {
                text.setBold(startOffset, endOffset, true);
            }
        }
    }
}

/**
 * Applies consistent typography throughout the document
 * Heading 1: Helvetica Neue Bold 24pt
 * Heading 2: Helvetica Neue Bold 14pt
 * Normal Text: Helvetica Neue 11pt
 * Preserves character-level formatting (bold, underline, italic, etc.)
 */
function fixDocumentFormatting() {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    const counts = { h1: 0, h2: 0, normal: 0, listItems: 0 };
    const numChildren = body.getNumChildren();

    for (let i = 0; i < numChildren; i++) {
        const child = body.getChild(i);
        const childType = child.getType();

        if (childType === DocumentApp.ElementType.PARAGRAPH) {
            const paragraph = child.asParagraph();
            const heading = paragraph.getHeading();

            if (heading === DocumentApp.ParagraphHeading.HEADING1) {
                if (needsFormatting(paragraph, DOCUMENT_STYLES.heading1)) {
                    applyFontFormatting(paragraph, DOCUMENT_STYLES.heading1);
                    counts.h1++;
                }
            } else if (heading === DocumentApp.ParagraphHeading.HEADING2) {
                if (needsFormatting(paragraph, DOCUMENT_STYLES.heading2)) {
                    applyFontFormatting(paragraph, DOCUMENT_STYLES.heading2);
                    counts.h2++;
                }
            } else if (heading === DocumentApp.ParagraphHeading.NORMAL) {
                if (needsFormatting(paragraph, DOCUMENT_STYLES.normal)) {
                    applyFontFormatting(paragraph, DOCUMENT_STYLES.normal);
                    counts.normal++;
                }
            }
        } else if (childType === DocumentApp.ElementType.LIST_ITEM) {
            const listItem = child.asListItem();
            if (needsFormatting(listItem, DOCUMENT_STYLES.normal)) {
                applyFontFormatting(listItem, DOCUMENT_STYLES.normal);
                counts.listItems++;
            }
        }
    }

    const totalChanges = counts.h1 + counts.h2 + counts.normal + counts.listItems;

    if (totalChanges === 0) {
        showAlert('Formatting Check Complete', 'All paragraphs and list items already have correct formatting. No changes needed.');
        return;
    }

    const message = `Document formatting updated:\n` +
        `${counts.h1} Heading 1 paragraph(s) formatted (Helvetica Neue Bold 24pt)\n` +
        `${counts.h2} Heading 2 paragraph(s) formatted (Helvetica Neue Bold 14pt)\n` +
        `${counts.normal} Normal text paragraph(s) formatted (Helvetica Neue 11pt)\n` +
        `${counts.listItems} List item(s) formatted (Helvetica Neue 11pt)`;

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