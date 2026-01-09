/**
 * Google Docs Link Checker
 *
 * Notes:
 * Valid non-HTTP protocols are skipped (not flagged as errors):
 * mailto: (email links)
 * tel: (phone numbers)
 * sms: (SMS links)
 * ftp: / sftp: (file transfer)
 * file: (local files)
 *
 * Invalid/malformed links are flagged in orange:
 * inks that don't start with any recognized protocol
 * Typos like htp:// or htps://
 * Weird formats or broken URLs
 *
 * HTTP/HTTPS links are checked normally:
 * Red for broken (404 or unreachable)
 * Yellow for Apple.com links
 *
 * mailto:contact@example.com → Skipped ✅
 * http://broken-site.com → Checked and flagged if broken 🔴
 * htp://typo.com → Flagged as invalid in orange 🟠
 *
 * he script may take a while if you have many links, as it checks each one individually. Also, Google Apps Script has some rate limits on
 * external URL fetches, so for documents with a very large number of links (100+), you might need to run it multiple times or add delays.
 *
 * How to set it up:
 * Open your Google Doc
 * Go to Extensions > Apps Script
 * Delete any existing code and paste this script
 * Click the Save icon (💾)
 * Close the Apps Script tab and refresh your Google Doc
 * You'll now see a "Link Checker" menu in your document
 *
 * To use it:
 * Click Link Checker > Check Links
 * The script will check all links and highlight them accordingly
 * A dialog will show you how many broken and Apple.com links were found
 *
 *
 */
function checkLinks() {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const text = body.editAsText();

    // Get all links in the document
    const links = [];
    const numChildren = body.getNumChildren();

    for (let i = 0; i < numChildren; i++) {
        const child = body.getChild(i);
        extractLinks(child, links);
    }

    // Check each link
    let brokenCount = 0;
    let appleCount = 0;
    let skippedCount = 0;
    let invalidCount = 0;

    // Valid non-HTTP protocols that should be skipped
    const validProtocols = ['mailto:', 'tel:', 'sms:', 'ftp:', 'sftp:', 'file:'];

    links.forEach(link => {
        const url = link.url;
        const startOffset = link.startOffset;
        const endOffset = link.endOffset;
        const element = link.element;
        const urlLower = url.toLowerCase();

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

    // Show results
    let message = `Found ${brokenCount} broken link(s) (highlighted in red)\n` +
        `Found ${appleCount} Apple.com link(s) (highlighted in yellow)\n` +
        `Skipped ${skippedCount} valid non-HTTP link(s) (email, phone, etc.)`;

    if (invalidCount > 0) {
        message += `\nFound ${invalidCount} invalid/malformed link(s) (highlighted in orange)`;
    }

    DocumentApp.getUi().alert(
        'Link Check Complete',
        message,
        DocumentApp.getUi().ButtonSet.OK
    );
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
 * Creates a custom menu when the document is opened
 */
function onOpen() {
    DocumentApp.getUi()
        .createMenu('Link Checker')
        .addItem('Check Links', 'checkLinks')
        .addToUi();
}