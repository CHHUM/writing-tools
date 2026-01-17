/**
 * Google Docs Formatter & Link Checker
 *
 * This script provides tools to maintain document quality:
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
 * - Click Document Tools > Check Links in Entire Document (or Active Tab)
 * - Click Document Tools > Fix Document Formatting
 * - Review the results shown in the dialog
 *
 */

function checkLinks(scope) {
    const doc = DocumentApp.getActiveDocument();
    let body;
    let scopeText = '';

    if (scope === 'tab') {
        // Get the active tab
        const activeTab = doc.getActiveTab();
        if (!activeTab) {
            DocumentApp.getUi().alert(
                'No Active Tab',
                'Could not find an active tab. The document might not have tabs, or you might need to click into a tab first.',
                DocumentApp.getUi().ButtonSet.OK
            );
            return;
        }
        body = activeTab.asDocumentTab().getBody();
        scopeText = ' in active tab';
    } else {
        // Get the entire document body
        body = doc.getBody();
        scopeText = ' in document';
    }

    const text = body.editAsText();

    // Get all links in the document
    const links = [];
    const underlinedText = [];
    const numChildren = body.getNumChildren();

    for (let i = 0; i < numChildren; i++) {
        const child = body.getChild(i);
        extractLinks(child, links);
        extractUnderlinedText(child, underlinedText);
    }

    // Check each link
    let brokenCount = 0;
    let appleCount = 0;
    let skippedCount = 0;
    let invalidCount = 0;
    let formattedCount = 0;
    let missingLinkCount = 0;

    // Valid non-HTTP protocols that should be skipped
    const validProtocols = ['mailto:', 'tel:', 'sms:', 'ftp:', 'sftp:', 'file:'];

    links.forEach(link => {
        const url = link.url;
        const startOffset = link.startOffset;
        const endOffset = link.endOffset;
        const element = link.element;
        const urlLower = url.toLowerCase();

        // Apply proper link formatting (blue text with underline)
        const needsFormatting = applyLinkFormatting(element, startOffset, endOffset);
        if (needsFormatting) {
            formattedCount++;
        }

        // Check if it's a valid non-HTTP protocol
        const isValidNonHttp = validProtocols.some(protocol => urlLower.startsWith(protocol));

        if (isValidNonHttp) {
            skippedCount++;
            return; // Skip valid non-HTTP links
        }

        // Check if it's HTTP/HTTPS
        const isHttp = urlLower.startsWith('http://') || urlLower.startsWith('https://');

        if (!isHttp) {
            // This is an invalid/unknown protocol or malformed link
            element.setBackgroundColor(startOffset, endOffset - 1, '#FFA500'); // Orange for invalid
            invalidCount++;
            return;
        }

        // Check if it's an Apple.com link
        if (urlLower.includes('apple.com')) {
            element.setBackgroundColor(startOffset, endOffset - 1, '#FFFF00'); // Yellow
            appleCount++;
        }

        // Check if link is broken (404)
        try {
            const response = UrlFetchApp.fetch(url, {
                'muteHttpExceptions': true,
                'followRedirects': false
            });

            if (response.getResponseCode() === 404) {
                element.setBackgroundColor(startOffset, endOffset - 1, '#FF0000'); // Red
                brokenCount++;
            }
        } catch (e) {
            // If we can't fetch the URL, consider it potentially broken
            element.setBackgroundColor(startOffset, endOffset - 1, '#FF0000'); // Red
            brokenCount++;
        }
    });

    // Check for underlined text without links (possible missing links)
    underlinedText.forEach(item => {
        const element = item.element;
        const startOffset = item.startOffset;
        const endOffset = item.endOffset;

        // Highlight in purple as possible missing link
        element.setBackgroundColor(startOffset, endOffset - 1, '#DDA0DD'); // Purple/Plum
        missingLinkCount++;
    });

    // Show results
    let message = `Found ${brokenCount} broken link(s) (highlighted in red)\n` +
        `Found ${appleCount} Apple.com link(s) (highlighted in yellow)\n` +
        `Skipped ${skippedCount} valid non-HTTP link(s) (email, phone, etc.)\n` +
        `Fixed formatting on ${formattedCount} link(s)${scopeText}`;

    if (invalidCount > 0) {
        message += `\nFound ${invalidCount} invalid/malformed link(s) (highlighted in orange)`;
    }

    if (missingLinkCount > 0) {
        message += `\nFound ${missingLinkCount} underlined text(s) without links (highlighted in purple)`;
    }

    DocumentApp.getUi().alert(
        'Link Check Complete',
        message,
        DocumentApp.getUi().ButtonSet.OK
    );
}

/**
 * Apply proper link formatting (blue text with underline)
 * Returns true if formatting was needed, false if already correct
 */
function applyLinkFormatting(element, startOffset, endOffset) {
    let needsFormatting = false;

    // Check current formatting
    const currentColor = element.getForegroundColor(startOffset);
    const currentUnderline = element.isUnderline(startOffset);

    // Google Docs link blue color
    const linkBlue = '#1155CC';

    // Check if formatting needs to be applied
    if (currentColor !== linkBlue || !currentUnderline) {
        needsFormatting = true;
    }

    // Apply the formatting
    element.setForegroundColor(startOffset, endOffset - 1, linkBlue);
    element.setUnderline(startOffset, endOffset - 1, true);

    return needsFormatting;
}

/**
 * Recursively extract all links from document elements
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
 * Recursively extract underlined text that is NOT a link
 */
function extractUnderlinedText(element, underlinedText) {
    if (element.getType() === DocumentApp.ElementType.TEXT) {
        const text = element.asText();
        const textString = text.getText();
        const indices = text.getTextAttributeIndices();

        for (let i = 0; i < indices.length; i++) {
            const startOffset = indices[i];
            const endOffset = i + 1 < indices.length ? indices[i + 1] : textString.length;

            // Check if this text is underlined
            const isUnderlined = text.isUnderline(startOffset);

            // Check if this text has a link
            const hasLink = text.getLinkUrl(startOffset) !== null;

            // If underlined but no link, it's a possible missing link
            if (isUnderlined && !hasLink) {
                // Only flag non-empty text
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
 * Check links in the entire document
 */
function checkLinksInDocument() {
    checkLinks('document');
}

/**
 * Check links in the active tab only
 */
function checkLinksInActiveTab() {
    checkLinks('tab');
}

/**
 * Fix document formatting with standard styles
 * Heading 1: Helvetica Neue Bold 24pt
 * Heading 2: Helvetica Neue Bold 14pt
 * Normal Text: Helvetica Neue 11pt
 */
function fixDocumentFormatting() {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    // Define styles
    const heading1Style = {};
    heading1Style[DocumentApp.Attribute.FONT_FAMILY] = 'Helvetica Neue';
    heading1Style[DocumentApp.Attribute.FONT_SIZE] = 24;
    heading1Style[DocumentApp.Attribute.BOLD] = true;

    const heading2Style = {};
    heading2Style[DocumentApp.Attribute.FONT_FAMILY] = 'Helvetica Neue';
    heading2Style[DocumentApp.Attribute.FONT_SIZE] = 14;
    heading2Style[DocumentApp.Attribute.BOLD] = true;

    const normalStyle = {};
    normalStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Helvetica Neue';
    normalStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
    normalStyle[DocumentApp.Attribute.BOLD] = false;

    // Count how many paragraphs were styled
    let h1Count = 0;
    let h2Count = 0;
    let normalCount = 0;

    // Apply styles to all paragraphs in the document
    const numChildren = body.getNumChildren();

    for (let i = 0; i < numChildren; i++) {
        const child = body.getChild(i);

        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const paragraph = child.asParagraph();
            const heading = paragraph.getHeading();

            if (heading === DocumentApp.ParagraphHeading.HEADING1) {
                paragraph.setAttributes(heading1Style);
                h1Count++;
            } else if (heading === DocumentApp.ParagraphHeading.HEADING2) {
                paragraph.setAttributes(heading2Style);
                h2Count++;
            } else if (heading === DocumentApp.ParagraphHeading.NORMAL) {
                paragraph.setAttributes(normalStyle);
                normalCount++;
            }
        }
    }

    // Show results
    const message = `Document formatting updated:\n` +
        `${h1Count} Heading 1 paragraph(s) formatted (Helvetica Neue Bold 24pt)\n` +
        `${h2Count} Heading 2 paragraph(s) formatted (Helvetica Neue Bold 14pt)\n` +
        `${normalCount} Normal text paragraph(s) formatted (Helvetica Neue 11pt)`;

    DocumentApp.getUi().alert(
        'Formatting Complete',
        message,
        DocumentApp.getUi().ButtonSet.OK
    );
}

/**
 * Creates a custom menu when the document is opened
 */
function onOpen() {
    DocumentApp.getUi()
        .createMenu('Document Tools')
        .addItem('Check Links in Entire Document', 'checkLinksInDocument')
        .addItem('Check Links in Active Tab', 'checkLinksInActiveTab')
        .addSeparator()
        .addItem('Fix Document Formatting', 'fixDocumentFormatting')
        .addToUi();
}