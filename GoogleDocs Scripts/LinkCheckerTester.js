/**
 * Link Checker Test Document Generator
 *
 * This script creates a test Google Doc with various link scenarios
 * to validate the Link Checker script.
 *
 * How to use:
 * 1. Open Google Apps Script (script.google.com)
 * 2. Create a new project
 * 3. Paste this code
 * 4. Run the createTestDocument() function
 * 5. Check your Google Drive for "Link Checker Test Document"
 * 6. Open that document and run your Link Checker on it
 *
 * Expected Results:
 * - Valid web links: Should be checked (green if working, red if 404)
 * - Valid non-HTTP links: Should be skipped (no highlighting)
 * - Invalid/malformed links: Should be highlighted in orange
 * - Apple.com links: Should be highlighted in yellow
 * - Broken links: Should be highlighted in red
 */

function createTestDocument() {
    // Create a new document
    const doc = DocumentApp.create('Link Checker Test Document');
    const body = doc.getBody();

    // Add title
    const title = body.appendParagraph('Link Checker Test Document');
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);

    body.appendParagraph('This document contains various types of links to test the Link Checker script.');
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
    body.appendParagraph('');

    // Section 5: Invalid/malformed links
    addSection(body, '5. Invalid/Malformed Links (should be ORANGE)');
    addTestLink(body, 'Missing protocol', 'www.example.com');
    addTestLink(body, 'Typo in protocol', 'htp://example.com');
    addTestLink(body, 'Another typo', 'htps://example.com');
    addTestLink(body, 'Weird protocol', 'xyz://example.com');
    addTestLink(body, 'Just text', 'not-a-real-link');
    body.appendParagraph('');

    // Section 6: Mixed scenarios
    addSection(body, '6. Mixed Scenarios');
    const para1 = body.appendParagraph('This paragraph has multiple links: ');
    appendInlineLink(para1, 'a working link', 'https://www.google.com');
    para1.appendText(', ');
    appendInlineLink(para1, 'an email', 'mailto:contact@example.com');
    para1.appendText(', and ');
    appendInlineLink(para1, 'a broken link', 'https://www.google.com/nonexistent-page-xyz');
    para1.appendText('.');

    body.appendParagraph('');

    const para2 = body.appendParagraph('Here is ');
    appendInlineLink(para2, 'an Apple link', 'https://www.apple.com/iphone');
    para2.appendText(' and ');
    appendInlineLink(para2, 'an invalid link', 'htp://typo.com');
    para2.appendText('.');

    body.appendParagraph('');

    // Section 7: Edge cases
    addSection(body, '7. Edge Cases');
    addTestLink(body, 'HTTPS with subdomain', 'https://docs.google.com');
    addTestLink(body, 'HTTP (not HTTPS)', 'http://example.com');
    addTestLink(body, 'URL with query params', 'https://www.google.com/search?q=test');
    addTestLink(body, 'URL with anchor', 'https://en.wikipedia.org/wiki/Main_Page#mp-tfa');
    body.appendParagraph('');

    // Add summary
    body.appendParagraph('');
    const summary = body.appendParagraph('Test Summary');
    summary.setHeading(DocumentApp.ParagraphHeading.HEADING2);

    body.appendParagraph('After running Link Checker, you should see:');
    body.appendParagraph('• RED highlights on broken/404 links (Section 2)');
    body.appendParagraph('• YELLOW highlights on Apple.com links (Section 3)');
    body.appendParagraph('• ORANGE highlights on invalid/malformed links (Section 5)');
    body.appendParagraph('• NO highlights on valid non-HTTP links (Section 4)');
    body.appendParagraph('• NO highlights on working HTTP/HTTPS links (Section 1)');

    // Log the document URL
    Logger.log('Test document created: ' + doc.getUrl());

    // Show success message
    const ui = SpreadsheetApp.getUi(); // Note: This won't work in standalone scripts
    // For standalone scripts, just check the logger
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
 * Run this AFTER running the Link Checker on the test document
 */
function validateTestResults() {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    Logger.log('=== Test Validation ===');
    Logger.log('Open the document and verify:');
    Logger.log('1. Section 2 links are RED (broken)');
    Logger.log('2. Section 3 links are YELLOW (Apple.com)');
    Logger.log('3. Section 4 links have NO highlight (skipped)');
    Logger.log('4. Section 5 links are ORANGE (invalid)');
    Logger.log('5. Section 1 and 7 links have NO highlight (working)');
    Logger.log('');
    Logger.log('Check the document visually to confirm colors match expectations.');
}