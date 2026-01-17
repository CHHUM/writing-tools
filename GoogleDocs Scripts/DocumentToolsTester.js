/**
 * Document Tools Test Document Generator
 *
 * This script creates a test Google Doc with various scenarios
 * to validate the Document Tools script (Link Checker & Formatter).
 *
 * How to use:
 * 1. Open Google Apps Script (script.google.com)
 * 2. Create a new project
 * 3. Paste this code
 * 4. Run the createTestDocument() function
 * 5. Check your Google Drive for "Document Tools Test Document"
 * 6. Open that document and run your Document Tools on it
 *
 * Expected Results:
 * LINK CHECKING:
 * - Valid web links: Should be checked (no highlight if working, red if 404)
 * - Valid non-HTTP links: Should be skipped (no highlighting)
 * - Invalid/malformed links: Should be highlighted in orange
 * - Apple.com links: Should be highlighted in yellow
 * - Broken links: Should be highlighted in red
 * - Underlined text without links: Should be highlighted in purple
 * - Links should be formatted with blue text and underline
 *
 * FORMATTING:
 * - Heading 1 should become: Helvetica Neue Bold 24pt
 * - Heading 2 should become: Helvetica Neue Bold 14pt
 * - Normal text should become: Helvetica Neue 11pt
 */

function createTestDocument() {
    // Create a new document
    const doc = DocumentApp.create('Document Tools Test Document');
    const body = doc.getBody();

    // Add title with mixed formatting (will be corrected by formatter)
    const title = body.appendParagraph('Document Tools Test Document');
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    title.editAsText().setFontFamily('Arial'); // Wrong font - should be fixed
    title.editAsText().setFontSize(18); // Wrong size - should be fixed

    body.appendParagraph('This document contains various scenarios to test the Document Tools script.');
    body.appendParagraph('');

    // Section 1: Valid HTTP/HTTPS links that should work
    addSection(body, '1. Valid Working Links (should NOT be highlighted)');
    addTestLink(body, 'Google Homepage', 'https://www.google.com');
    addTestLink(body, 'Wikipedia', 'https://www.wikipedia.org');
    addTestLink(body, 'GitHub', 'https://github.com');
    body.appendParagraph('');

    // Section 2: Broken HTTP links (404)
    addSection(body, '2. Broken Links - 404 (should be RED)');
    addTestLink(body, 'Non-existent page', 'https://www.google.com/this-page-definitely-does-not-exist-12345');
    addTestLink(body, 'Another broken link', 'https://github.com/nonexistent-user-xyz-12345/nonexistent-repo-abc-67890');
    body.appendParagraph('');

    // Section 3: Apple.com links
    addSection(body, '3. Apple.com Links (should be YELLOW)');
    addTestLink(body, 'Apple Homepage', 'https://www.apple.com');
    addTestLink(body, 'Apple Support', 'https://support.apple.com');
    addTestLink(body, 'Apple Store', 'https://www.apple.com/store');
    body.appendParagraph('');

    // Section 4: Valid non-HTTP protocols (should be skipped)
    addSection(body, '4. Valid Non-HTTP Links (should be SKIPPED - no highlighting)');
    addTestLink(body, 'Email link', 'mailto:test@example.com');
    addTestLink(body, 'Phone link', 'tel:+1234567890');
    addTestLink(body, 'SMS link', 'sms:+1234567890');
    addTestLink(body, 'FTP link', 'ftp://ftp.example.com/file.txt');
    addTestLink(body, 'SFTP link', 'sftp://secure.example.com/file.txt');
    body.appendParagraph('');

    // Section 5: Invalid/malformed links
    addSection(body, '5. Invalid/Malformed Links (should be ORANGE)');
    addTestLink(body, 'Missing protocol', 'www.example.com');
    addTestLink(body, 'Typo in protocol', 'htp://example.com');
    addTestLink(body, 'Another typo', 'htps://example.com');
    addTestLink(body, 'Weird protocol', 'xyz://example.com');
    addTestLink(body, 'Just text', 'not-a-real-link');
    body.appendParagraph('');

    // Section 6: Underlined text without links (NEW!)
    addSection(body, '6. Underlined Text Without Links (should be PURPLE)');

    const underlinedPara1 = body.appendParagraph('• This is underlined text but has no link');
    underlinedPara1.editAsText().setUnderline(2, 40, true); // Underline "This is underlined text but has no link"

    const underlinedPara2 = body.appendParagraph('• Another example of underlined text');
    underlinedPara2.editAsText().setUnderline(2, 34, true);

    const mixedPara = body.appendParagraph('• This paragraph has ');
    const underlinedText = mixedPara.appendText('underlined text');
    underlinedText.setUnderline(true);
    mixedPara.appendText(' and ');
    appendInlineLink(mixedPara, 'a real link', 'https://www.google.com');
    mixedPara.appendText(' mixed together.');

    body.appendParagraph('');

    // Section 7: Links with incorrect formatting (should be auto-fixed to blue + underline)
    addSection(body, '7. Links with Wrong Formatting (should be auto-fixed to blue + underline)');

    const wrongColorLink = body.appendParagraph('• ');
    const redLink = wrongColorLink.appendText('This link is red instead of blue');
    redLink.setLinkUrl('https://www.google.com');
    redLink.setForegroundColor('#FF0000'); // Wrong color
    redLink.setUnderline(true);

    const noUnderlineLink = body.appendParagraph('• ');
    const plainLink = noUnderlineLink.appendText('This link has no underline');
    plainLink.setLinkUrl('https://www.wikipedia.org');
    plainLink.setForegroundColor('#000000'); // Wrong color
    plainLink.setUnderline(false); // No underline

    const wrongBothLink = body.appendParagraph('• ');
    const greenLink = wrongBothLink.appendText('This link is green and not underlined');
    greenLink.setLinkUrl('https://github.com');
    greenLink.setForegroundColor('#00FF00'); // Wrong color
    greenLink.setUnderline(false); // No underline

    body.appendParagraph('');

    // Section 8: Mixed scenarios
    addSection(body, '8. Mixed Scenarios');
    const para1 = body.appendParagraph('This paragraph has multiple links: ');
    appendInlineLink(para1, 'a working link', 'https://www.google.com');
    para1.appendText(', ');
    appendInlineLink(para1, 'an email', 'mailto:contact@example.com');
    para1.appendText(', ');
    appendInlineLink(para1, 'a broken link', 'https://www.google.com/nonexistent-page-xyz');
    para1.appendText(', and ');
    const underlined = para1.appendText('underlined text with no link');
    underlined.setUnderline(true);
    para1.appendText('.');

    body.appendParagraph('');

    const para2 = body.appendParagraph('Here is ');
    appendInlineLink(para2, 'an Apple link', 'https://www.apple.com/iphone');
    para2.appendText(' and ');
    appendInlineLink(para2, 'an invalid link', 'htp://typo.com');
    para2.appendText('.');

    body.appendParagraph('');

    // Section 9: Edge cases
    addSection(body, '9. Edge Cases');
    addTestLink(body, 'HTTPS with subdomain', 'https://docs.google.com');
    addTestLink(body, 'HTTP (not HTTPS)', 'http://example.com');
    addTestLink(body, 'URL with query params', 'https://www.google.com/search?q=test');
    addTestLink(body, 'URL with anchor', 'https://en.wikipedia.org/wiki/Main_Page#mp-tfa');
    body.appendParagraph('');

    // Section 10: Font formatting tests
    addSection(body, '10. Font Formatting Tests (wrong fonts - should be fixed)');

    const wrongFont1 = body.appendParagraph('This is a Heading 1 in Times New Roman (should become Helvetica Neue Bold 24pt)');
    wrongFont1.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    wrongFont1.editAsText().setFontFamily('Times New Roman');
    wrongFont1.editAsText().setFontSize(20);
    wrongFont1.editAsText().setBold(false);

    const wrongFont2 = body.appendParagraph('This is a Heading 2 in Comic Sans (should become Helvetica Neue Bold 14pt)');
    wrongFont2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    wrongFont2.editAsText().setFontFamily('Comic Sans MS');
    wrongFont2.editAsText().setFontSize(16);
    wrongFont2.editAsText().setBold(false);

    const wrongFont3 = body.appendParagraph('This is normal text in Courier New 14pt (should become Helvetica Neue 11pt)');
    wrongFont3.editAsText().setFontFamily('Courier New');
    wrongFont3.editAsText().setFontSize(14);
    wrongFont3.editAsText().setBold(true);

    body.appendParagraph('');

    // Add summary
    body.appendParagraph('');
    const summary = body.appendParagraph('Test Summary');
    summary.setHeading(DocumentApp.ParagraphHeading.HEADING2);

    body.appendParagraph('After running "Check Links", you should see:');
    body.appendParagraph('• Section 1: NO highlights on working HTTP/HTTPS links');
    body.appendParagraph('• Section 2: RED highlights on broken/404 links');
    body.appendParagraph('• Section 3: YELLOW highlights on Apple.com links');
    body.appendParagraph('• Section 4: NO highlights on valid non-HTTP links (skipped)');
    body.appendParagraph('• Section 5: ORANGE highlights on invalid/malformed links');
    body.appendParagraph('• Section 6: PURPLE highlights on underlined text without links');
    body.appendParagraph('• Section 7: All links auto-fixed to blue and underlined');
    body.appendParagraph('• Section 8: Mixed highlights based on link types');
    body.appendParagraph('• Section 9: NO highlights on edge case links (all should work)');

    body.appendParagraph('');
    body.appendParagraph('After running "Fix Document Formatting", you should see:');
    body.appendParagraph('• All Heading 1 paragraphs in Helvetica Neue Bold 24pt');
    body.appendParagraph('• All Heading 2 paragraphs in Helvetica Neue Bold 14pt');
    body.appendParagraph('• All Normal text paragraphs in Helvetica Neue 11pt (not bold)');
    body.appendParagraph('• Section 10 should show the before/after difference clearly');

    // Log the document URL
    Logger.log('Test document created: ' + doc.getUrl());
    Logger.log('SUCCESS: Test document created successfully!');
    Logger.log('Document ID: ' + doc.getId());
    Logger.log('Open it here: ' + doc.getUrl());

    return doc.getUrl();
}

/**
 * Helper function to add a section heading
 */
function addSection(body, title) {
    const section = body.appendParagraph(title);
    section.setHeading(DocumentApp.ParagraphHeading.HEADING2);
}

/**
 * Helper function to add a test link as a paragraph
 */
function addTestLink(body, text, url) {
    const para = body.appendParagraph('• ' + text + ': ');
    const linkText = para.appendText(url);
    linkText.setLinkUrl(url);
}

/**
 * Helper function to append an inline link within a paragraph
 */
function appendInlineLink(paragraph, text, url) {
    const linkText = paragraph.appendText(text);
    linkText.setLinkUrl(url);
    return linkText;
}

/**
 * Creates the custom menu
 */
function onOpen() {
    const ui = DocumentApp.getUi();
    ui.createMenu('Test Generator')
        .addItem('Create Test Document', 'createTestDocument')
        .addToUi();
}

/**
 * Optional: Function to validate the test results
 * Run this AFTER running the Document Tools on the test document
 */
function validateTestResults() {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    Logger.log('=== Test Validation ===');
    Logger.log('Open the document and verify:');
    Logger.log('');
    Logger.log('LINK CHECKING RESULTS:');
    Logger.log('1. Section 2 links are RED (broken)');
    Logger.log('2. Section 3 links are YELLOW (Apple.com)');
    Logger.log('3. Section 4 links have NO highlight (skipped non-HTTP)');
    Logger.log('4. Section 5 links are ORANGE (invalid)');
    Logger.log('5. Section 6 underlined text is PURPLE (missing links)');
    Logger.log('6. Section 7 links are now BLUE and UNDERLINED (auto-fixed)');
    Logger.log('7. Section 1 and 9 links have NO highlight (working)');
    Logger.log('');
    Logger.log('FORMATTING RESULTS:');
    Logger.log('1. All Heading 1 text is Helvetica Neue Bold 24pt');
    Logger.log('2. All Heading 2 text is Helvetica Neue Bold 14pt');
    Logger.log('3. All Normal text is Helvetica Neue 11pt (not bold)');
    Logger.log('4. Section 10 paragraphs should show corrected fonts');
    Logger.log('');
    Logger.log('Check the document visually to confirm all expectations are met.');
}