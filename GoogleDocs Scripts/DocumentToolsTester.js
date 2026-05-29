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
 * - Links with leading/trailing spaces: Spaces removed from link but kept in document
 * - Only links needing formatting changes should be counted in results
 *
 * FORMATTING:
 * - Heading 1 should become: Helvetica Neue Bold 24pt
 * - Heading 2 should become: Helvetica Neue Bold 14pt
 * - Normal text should become: Helvetica Neue 11pt
 *
 * EXTRA SPACE REMOVAL (Section 13):
 * - Runs of 2 or more consecutive spaces are collapsed to a single space
 * - Applies everywhere in the document body
 * - Count reported in the "Fix Document Formatting" summary
 *
 * TEXT CASE TOOLS:
 * - Lower case: Converts selected text to lowercase
 * - Upper case: Converts selected text to uppercase
 * - Initial Caps: Capitalizes the first letter of each word
 * - Sentence case: Capitalizes only the first letter of the selection
 * - Title Case (Chicago Style): Proper title capitalization following Chicago Manual of Style
 * - All conversions preserve paragraph-level heading styles (Heading 1, Heading 2, etc.)
 * - All conversions preserve the selection highlight on the full selected range after conversion
 *
 * MULTI-TAB LINK CHECKING:
 * - The generated document has three tabs
 * - Tab 1 (default): all main test scenarios (Sections 1–13)
 * - Tab 2 ("Tab 2 – Working & Apple Links"): valid, Apple, and broken links only
 * - Tab 3 ("Tab 3 – Spaces & Underlines"): space-trimming and underlined-without-link scenarios
 * - Use "Check Links > In Active Tab" on each tab and verify only that tab's links are processed
 * - Use "Check Links > In Entire Document" and verify all three tabs are processed together
 *
 * VERSION: 3.0.0
 * CHANGELOG:
 * - v3.0.0: Added Section 13 to test extra space removal (double/triple spaces)
 *           Added multi-tab document: Tab 2 (working, Apple, broken links) and
 *           Tab 3 (space-trimming, underlined-without-link) for per-tab link checking tests
 *           Updated summary, validation, and expected counts to reflect new sections and tabs
 * - v2.1.0: Added heading style preservation tests to Section 12
 *           Added selection highlight consistency tests to Section 12
 *           Updated expected results and validation to reflect v2.4.0 of Document Tools
 * - v2.0.0: Added Section 9 to test link space trimming functionality
 *           Added validation for accurate formatting count reporting
 *           Updated test expectations and documentation
 */

// Define the link blue color constant
const COLORS = {
    LINK_BLUE: '#1155CC'
};

// ============================================================================
// MAIN ENTRY POINT
// ============================================================================

function createTestDocument() {
    // Create a new document (starts with one default tab)
    const doc = DocumentApp.create('Document Tools Test Document');
    const body = doc.getBody();

    // ── Tab 1: all main test scenarios ────────────────────────────────────────
    populateTab1(body);

    // ── Tab 2: working, Apple, and broken links only ──────────────────────────
    const tab2 = doc.addTab(DocumentApp.newTabProperties().setTitle('Tab 2 – Working & Apple Links'));
    populateTab2(tab2.asDocumentTab().getBody());

    // ── Tab 3: space-trimming and underlined-without-link scenarios ───────────
    const tab3 = doc.addTab(DocumentApp.newTabProperties().setTitle('Tab 3 – Spaces & Underlines'));
    populateTab3(tab3.asDocumentTab().getBody());

    // Log the document URL
    Logger.log('Test document created: ' + doc.getUrl());
    Logger.log('SUCCESS: Test document created successfully!');
    Logger.log('Document ID: ' + doc.getId());
    Logger.log('Open it here: ' + doc.getUrl());

    return doc.getUrl();
}

// ============================================================================
// TAB 1 – ALL MAIN TEST SCENARIOS
// ============================================================================

function populateTab1(body) {
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

    // Section 6: Underlined text without links
    addSection(body, '6. Underlined Text Without Links (should be PURPLE)');

    const underlinedPara1 = body.appendParagraph('• This is underlined text but has no link');
    underlinedPara1.editAsText().setUnderline(2, 40, true);

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

    // Section 9: Links with leading/trailing spaces
    addSection(body, '9. Links with Leading/Trailing Spaces (should be auto-trimmed)');

    const spacePara1 = body.appendParagraph('• Link with leading space: ');
    const leadingSpace = spacePara1.appendText(' Google Homepage');
    leadingSpace.setLinkUrl('https://www.google.com');
    leadingSpace.setForegroundColor(COLORS.LINK_BLUE);
    leadingSpace.setUnderline(true);

    const spacePara2 = body.appendParagraph('• Link with trailing space: ');
    const trailingSpace = spacePara2.appendText('Wikipedia ');
    trailingSpace.setLinkUrl('https://www.wikipedia.org');
    trailingSpace.setForegroundColor(COLORS.LINK_BLUE);
    trailingSpace.setUnderline(true);

    const spacePara3 = body.appendParagraph('• Link with both leading and trailing spaces: ');
    const bothSpaces = spacePara3.appendText('  GitHub  ');
    bothSpaces.setLinkUrl('https://github.com');
    bothSpaces.setForegroundColor(COLORS.LINK_BLUE);
    bothSpaces.setUnderline(true);

    const spacePara4 = body.appendParagraph('This sentence has');
    const inlineLeading = spacePara4.appendText(' a link with leading space');
    inlineLeading.setLinkUrl('https://www.google.com');
    inlineLeading.setForegroundColor(COLORS.LINK_BLUE);
    inlineLeading.setUnderline(true);
    spacePara4.appendText(' in the middle of text.');

    const spacePara5 = body.appendParagraph('This sentence has');
    const inlineTrailing = spacePara5.appendText(' a link with trailing space ');
    inlineTrailing.setLinkUrl('https://www.wikipedia.org');
    inlineTrailing.setForegroundColor(COLORS.LINK_BLUE);
    inlineTrailing.setUnderline(true);
    spacePara5.appendText('in the middle of text.');

    const spacePara6 = body.appendParagraph('Multiple spaces:');
    const multiSpaces = spacePara6.appendText('   link with three spaces each side   ');
    multiSpaces.setLinkUrl('https://github.com');
    multiSpaces.setForegroundColor(COLORS.LINK_BLUE);
    multiSpaces.setUnderline(true);
    spacePara6.appendText('end.');

    body.appendParagraph('Expected after running link checker:');
    body.appendParagraph('• Spaces should remain in document as normal text');
    body.appendParagraph('• Spaces should NOT be part of the link (not blue, not underlined)');
    body.appendParagraph('• Only the actual text should be linked');
    body.appendParagraph('• Script should report "Trimmed spaces from X link(s)"');
    body.appendParagraph('');

    // Section 10: Edge cases
    addSection(body, '10. Edge Cases');
    addTestLink(body, 'HTTPS with subdomain', 'https://docs.google.com');
    addTestLink(body, 'HTTP (not HTTPS)', 'http://example.com');
    addTestLink(body, 'URL with query params', 'https://www.google.com/search?q=test');
    addTestLink(body, 'URL with anchor', 'https://en.wikipedia.org/wiki/Main_Page#mp-tfa');
    body.appendParagraph('');

    // Section 11: Font formatting tests
    addSection(body, '11. Font Formatting Tests (wrong fonts - should be fixed)');

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

    // Section 12: Text Case Tests
    addSection(body, '12. Text Case Conversion Tests');

    body.appendParagraph('Select the text samples below and use Document Tools > Text Case menu to test conversions.');
    body.appendParagraph('All conversions should:');
    body.appendParagraph('  • Preserve character-level formatting (bold, italic, colors, underline)');
    body.appendParagraph('  • Preserve paragraph-level heading styles (Heading 1, Heading 2, etc.)');
    body.appendParagraph('  • Leave the full selected range highlighted after conversion');
    body.appendParagraph('');

    // Lower case test
    const lowerTest = body.appendParagraph('LOWER CASE TEST: ');
    const lowerSample = lowerTest.appendText('THE QUICK BROWN FOX JUMPS OVER THE LAZY DOG');
    lowerSample.setBold(true);
    lowerSample.setForegroundColor('#FF0000');
    body.appendParagraph('Expected after "Lower case": the quick brown fox jumps over the lazy dog (red, bold)');
    body.appendParagraph('');

    // Upper case test
    const upperTest = body.appendParagraph('UPPER CASE TEST: ');
    const upperSample = upperTest.appendText('the quick brown fox jumps over the lazy dog');
    upperSample.setItalic(true);
    upperSample.setForegroundColor('#0000FF');
    body.appendParagraph('Expected after "Upper case": THE QUICK BROWN FOX JUMPS OVER THE LAZY DOG (blue, italic)');
    body.appendParagraph('');

    // Initial Caps test
    const initialTest = body.appendParagraph('INITIAL CAPS TEST: ');
    const initialSample = initialTest.appendText('the quick brown fox jumps over the lazy dog');
    initialSample.setUnderline(true);
    initialSample.setForegroundColor('#00AA00');
    body.appendParagraph('Expected after "Initial Caps": The Quick Brown Fox Jumps Over The Lazy Dog (green, underlined)');
    body.appendParagraph('');

    // Sentence case test
    const sentenceTest = body.appendParagraph('SENTENCE CASE TEST: ');
    const sentenceSample = sentenceTest.appendText('THE QUICK BROWN FOX JUMPS OVER THE LAZY DOG');
    sentenceSample.setBold(true);
    sentenceSample.setItalic(true);
    body.appendParagraph('Expected after "Sentence case": The quick brown fox jumps over the lazy dog (bold, italic)');
    body.appendParagraph('');

    // Title Case test - simple
    const titleTest1 = body.appendParagraph('TITLE CASE TEST 1 (Simple): ');
    const titleSample1 = titleTest1.appendText('the quick brown fox jumps over the lazy dog');
    titleSample1.setForegroundColor('#AA00AA');
    body.appendParagraph('Expected after "Title Case": The Quick Brown Fox Jumps over the Lazy Dog (purple)');
    body.appendParagraph('');

    // Title Case test - with articles and prepositions
    const titleTest2 = body.appendParagraph('TITLE CASE TEST 2 (Articles): ');
    titleTest2.appendText('a tale of two cities');
    body.appendParagraph('Expected after "Title Case": A Tale of Two Cities');
    body.appendParagraph('');

    // Title Case test - with conjunctions
    const titleTest3 = body.appendParagraph('TITLE CASE TEST 3 (Conjunctions): ');
    titleTest3.appendText('the lord of the rings and the hobbit');
    body.appendParagraph('Expected after "Title Case": The Lord of the Rings and the Hobbit');
    body.appendParagraph('');

    // Title Case test - with colon
    const titleTest4 = body.appendParagraph('TITLE CASE TEST 4 (Colon): ');
    titleTest4.appendText('the lord of the rings: the fellowship of the ring');
    body.appendParagraph('Expected after "Title Case": The Lord of the Rings: The Fellowship of the Ring');
    body.appendParagraph('');

    // Title Case test - first and last words
    const titleTest5 = body.appendParagraph('TITLE CASE TEST 5 (First/Last): ');
    titleTest5.appendText('to be or not to be');
    body.appendParagraph('Expected after "Title Case": To Be or Not to Be');
    body.appendParagraph('');

    // Heading style preservation tests
    addSection(body, '12a. Heading Style Preservation Tests');
    body.appendParagraph('Select ONLY the text of each heading below (not this paragraph) and apply a case conversion.');
    body.appendParagraph('The heading style (size, bold) must be preserved after conversion.');
    body.appendParagraph('');

    const headingCaseTest1 = body.appendParagraph('THIS IS A HEADING ONE IN UPPER CASE');
    headingCaseTest1.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('Expected after "Lower case": this is a heading one in upper case — still Heading 1 (Helvetica Neue Bold 24pt)');
    body.appendParagraph('');

    const headingCaseTest2 = body.appendParagraph('this is a heading two in lower case');
    headingCaseTest2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Expected after "Upper case": THIS IS A HEADING TWO IN LOWER CASE — still Heading 2 (Helvetica Neue Bold 14pt)');
    body.appendParagraph('');

    const headingCaseTest3 = body.appendParagraph('the lord of the rings: the two towers');
    headingCaseTest3.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Expected after "Title Case": The Lord of the Rings: The Two Towers — still Heading 2 (Helvetica Neue Bold 14pt)');
    body.appendParagraph('');

    // Selection highlight consistency tests
    addSection(body, '12b. Selection Highlight Consistency Tests');
    body.appendParagraph('For each test below, select the sample text and apply the named conversion.');
    body.appendParagraph('After conversion, the FULL selected text should remain highlighted — not just part of it, and the cursor should not jump to the start of the word.');
    body.appendParagraph('');

    const highlightTest1 = body.appendParagraph('UPPER CASE HIGHLIGHT TEST: ');
    highlightTest1.appendText('social media');
    body.appendParagraph('Select "social media", apply Upper case. Expected: "SOCIAL MEDIA" fully highlighted after conversion.');
    body.appendParagraph('');

    const highlightTest2 = body.appendParagraph('LOWER CASE HIGHLIGHT TEST: ');
    highlightTest2.appendText('SOCIAL MEDIA');
    body.appendParagraph('Select "SOCIAL MEDIA", apply Lower case. Expected: "social media" fully highlighted after conversion.');
    body.appendParagraph('');

    const highlightTest3 = body.appendParagraph('INITIAL CAPS HIGHLIGHT TEST: ');
    highlightTest3.appendText('SOCIAL MEDIA STRATEGY');
    body.appendParagraph('Select "SOCIAL MEDIA STRATEGY", apply Initial Caps. Expected: "Social Media Strategy" fully highlighted after conversion.');
    body.appendParagraph('');

    const highlightTest4 = body.appendParagraph('MULTI-WORD HIGHLIGHT TEST: ');
    highlightTest4.appendText('Social Social Social');
    body.appendParagraph('Select all three words, apply Upper case. Expected: "SOCIAL SOCIAL SOCIAL" fully highlighted — last character must be included.');
    body.appendParagraph('');

    // Mixed formatting test
    const mixedTest = body.appendParagraph('MIXED FORMATTING TEST: ');
    const mixedSample1 = mixedTest.appendText('THIS IS ');
    mixedSample1.setBold(true);
    const mixedSample2 = mixedTest.appendText('MIXED ');
    mixedSample2.setItalic(true);
    mixedSample2.setForegroundColor('#FF0000');
    const mixedSample3 = mixedTest.appendText('FORMATTING');
    mixedSample3.setUnderline(true);
    body.appendParagraph('Expected: All case conversions preserve individual formatting of each word');
    body.appendParagraph('');

    // ── Section 13: Extra Space Removal Tests ─────────────────────────────────
    addSection(body, '13. Extra Space Removal Tests (should be fixed by Fix Document Formatting)');

    body.appendParagraph('Run "Fix Document Formatting" and verify that all extra spaces below are collapsed to a single space.');
    body.appendParagraph('');

    // Double spaces between words
    body.appendParagraph('DOUBLE SPACE TEST: This sentence  has a double  space between  some of its  words.');
    body.appendParagraph('Expected after formatting: "This sentence has a double space between some of its words."');
    body.appendParagraph('');

    // Triple spaces
    body.appendParagraph('TRIPLE SPACE TEST: This one   has triple   spaces   scattered   throughout.');
    body.appendParagraph('Expected after formatting: "This one has triple spaces scattered throughout."');
    body.appendParagraph('');

    // Mixed double and triple spaces
    body.appendParagraph('MIXED SPACE TEST: One  space here,   three there,  two again,    four at the end    .');
    body.appendParagraph('Expected after formatting: "One space here, three there, two again, four at the end ."');
    body.appendParagraph('(Note: the space before the period is intentional to verify only extra spaces are removed)');
    body.appendParagraph('');

    // A run of many spaces
    body.appendParagraph('MANY SPACES TEST: Lots          of          spaces          here.');
    body.appendParagraph('Expected after formatting: "Lots of spaces here."');
    body.appendParagraph('');

    // Double spaces in a heading (formatting also corrects the heading font)
    const doubleSpaceHeading = body.appendParagraph('Heading  With  Double  Spaces');
    doubleSpaceHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    doubleSpaceHeading.editAsText().setFontFamily('Arial');
    body.appendParagraph('Expected after formatting: "Heading With Double Spaces" — still Heading 2, corrected to Helvetica Neue Bold 14pt.');
    body.appendParagraph('');

    // Extra spaces adjacent to a link (space inside the run, not part of the link URL)
    const linkSpacePara = body.appendParagraph('Sentence with  double spaces before ');
    appendInlineLink(linkSpacePara, 'a link', 'https://www.google.com');
    linkSpacePara.appendText('  and  double spaces  after it.');
    body.appendParagraph('Expected: double spaces outside the link are collapsed; the link itself is unaffected.');
    body.appendParagraph('');

    body.appendParagraph('Reporting: the "Fix Document Formatting" summary should include a line like:');
    body.appendParagraph('"X extra space(s) removed" — where X reflects every duplicate space in this section and across the whole document.');
    body.appendParagraph('');

    // ── Summary ───────────────────────────────────────────────────────────────
    body.appendParagraph('');
    const summary = body.appendParagraph('Test Summary — Tab 1');
    summary.setHeading(DocumentApp.ParagraphHeading.HEADING2);

    body.appendParagraph('After running "Check Links > In Entire Document", you should see:');
    body.appendParagraph('• Section 1: NO highlights on working HTTP/HTTPS links');
    body.appendParagraph('• Section 2: RED highlights on broken/404 links');
    body.appendParagraph('• Section 3: YELLOW highlights on Apple.com links');
    body.appendParagraph('• Section 4: NO highlights on valid non-HTTP links (skipped)');
    body.appendParagraph('• Section 5: ORANGE highlights on invalid/malformed links');
    body.appendParagraph('• Section 6: PURPLE highlights on underlined text without links');
    body.appendParagraph('• Section 7: All links auto-fixed to blue and underlined');
    body.appendParagraph('• Section 8: Mixed highlights based on link types');
    body.appendParagraph('• Section 9: Links trimmed, spaces remain as normal text (not blue/underlined)');
    body.appendParagraph('• Section 10: NO highlights on edge case links (all should work)');

    body.appendParagraph('');
    body.appendParagraph('After running "Fix Document Formatting", you should see:');
    body.appendParagraph('• All Heading 1 paragraphs in Helvetica Neue Bold 24pt');
    body.appendParagraph('• All Heading 2 paragraphs in Helvetica Neue Bold 14pt');
    body.appendParagraph('• All Normal text paragraphs in Helvetica Neue 11pt (not bold)');
    body.appendParagraph('• Section 11 should show the before/after difference clearly');
    body.appendParagraph('• Section 13 double/triple spaces collapsed to single spaces throughout');
    body.appendParagraph('• Summary reports "X extra space(s) removed" for Section 13 cases');

    body.appendParagraph('');
    body.appendParagraph('After testing "Text Case" conversions on Section 12:');
    body.appendParagraph('• Character-level formatting (bold, italic, colors, underline) preserved');
    body.appendParagraph('• Paragraph-level heading styles (Heading 1, Heading 2) preserved — see Section 12a');
    body.appendParagraph('• Full selected range remains highlighted after conversion for ALL five case functions — see Section 12b');
    body.appendParagraph('• Lower case: All letters become lowercase');
    body.appendParagraph('• Upper case: All letters become uppercase');
    body.appendParagraph('• Initial Caps: First letter of each word capitalized');
    body.appendParagraph('• Sentence case: Only first letter capitalized');
    body.appendParagraph('• Title Case: Chicago Style rules applied (see test expectations above)');

    body.appendParagraph('');
    body.appendParagraph('Expected message counts:');
    body.appendParagraph('• "Fixed formatting on X link(s)" should only count links that actually needed formatting changes');
    body.appendParagraph('• "Trimmed spaces from 6 link(s)" should count the 6 links in Section 9');

    body.appendParagraph('');
    body.appendParagraph('Multi-tab testing (see also Tab 2 and Tab 3):');
    body.appendParagraph('• "Check Links > In Active Tab" while on Tab 1 should process ONLY Tab 1 links');
    body.appendParagraph('• "Check Links > In Active Tab" while on Tab 2 should process ONLY Tab 2 links');
    body.appendParagraph('• "Check Links > In Active Tab" while on Tab 3 should process ONLY Tab 3 links');
    body.appendParagraph('• "Check Links > In Entire Document" should process ALL tabs and report combined counts');
}

// ============================================================================
// TAB 2 – WORKING, APPLE, AND BROKEN LINKS
// ============================================================================
//
// Purpose: verify that "Check Links > In Active Tab" scopes correctly to this
// tab. When active here the checker should find only the link types below, and
// the counts should NOT include links from Tab 1 or Tab 3.
//
// Expected results when "Check Links > In Active Tab" is run on Tab 2:
//   • 3 working links  → NO highlight
//   • 3 Apple links    → YELLOW highlight
//   • 2 broken links   → RED highlight
//   • 0 invalid, 0 non-HTTP, 0 missing-link highlights
// ============================================================================

function populateTab2(body) {
    const title = body.appendParagraph('Tab 2 – Working & Apple Links');
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);

    body.appendParagraph('This tab is used to test "Check Links > In Active Tab" in isolation.');
    body.appendParagraph('Only Working, Apple, and Broken links appear here.');
    body.appendParagraph('');

    // Tab 2 – Section A: Valid working links
    addSection(body, 'A. Valid Working Links (should NOT be highlighted)');
    addTestLink(body, 'Google', 'https://www.google.com');
    addTestLink(body, 'Wikipedia', 'https://www.wikipedia.org');
    addTestLink(body, 'GitHub', 'https://github.com');
    body.appendParagraph('');

    // Tab 2 – Section B: Apple links
    addSection(body, 'B. Apple.com Links (should be YELLOW)');
    addTestLink(body, 'Apple Homepage', 'https://www.apple.com');
    addTestLink(body, 'Apple Developer', 'https://developer.apple.com');
    addTestLink(body, 'Apple Support', 'https://support.apple.com');
    body.appendParagraph('');

    // Tab 2 – Section C: Broken links
    addSection(body, 'C. Broken Links - 404 (should be RED)');
    addTestLink(body, 'Non-existent page', 'https://www.google.com/this-page-definitely-does-not-exist-tab2');
    addTestLink(body, 'Another broken link', 'https://github.com/nonexistent-user-tab2-xyz/nonexistent-repo');
    body.appendParagraph('');

    // Tab 2 – Summary
    const tab2Summary = body.appendParagraph('Tab 2 Expected Results');
    tab2Summary.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Run "Check Links > In Active Tab" while this tab is active.');
    body.appendParagraph('• Section A: NO highlights (3 working links)');
    body.appendParagraph('• Section B: YELLOW highlights (3 Apple links)');
    body.appendParagraph('• Section C: RED highlights (2 broken links)');
    body.appendParagraph('• Reported counts should reflect only these 8 links — none from Tab 1 or Tab 3.');
    body.appendParagraph('');
    body.appendParagraph('Then switch to Tab 1 or Tab 3 and run "Check Links > In Active Tab" again.');
    body.appendParagraph('The Tab 2 highlights should remain unchanged — they are not re-processed.');
}

// ============================================================================
// TAB 3 – SPACE TRIMMING AND UNDERLINED-WITHOUT-LINK
// ============================================================================
//
// Purpose: verify that "Check Links > In Active Tab" scopes correctly to this
// tab. When active here the checker should find only space-trimming cases and
// underlined-without-link text, with NO broken/Apple/invalid link processing.
//
// Expected results when "Check Links > In Active Tab" is run on Tab 3:
//   • 3 space-trimmed links → spaces stripped from link (remain as normal text)
//   • 2 underlined runs     → PURPLE highlight
//   • 2 valid links (used as anchors alongside the above) → NO highlight
//   • 0 broken, 0 Apple, 0 invalid highlights
// ============================================================================

function populateTab3(body) {
    const title = body.appendParagraph('Tab 3 – Spaces & Underlines');
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);

    body.appendParagraph('This tab is used to test "Check Links > In Active Tab" for space-trimming and underlined-without-link detection.');
    body.appendParagraph('');

    // Tab 3 – Section A: Links with leading/trailing spaces
    addSection(body, 'A. Links with Leading/Trailing Spaces (spaces should be trimmed)');

    const t3para1 = body.appendParagraph('• Leading space: ');
    const t3leading = t3para1.appendText(' Apple Developer');
    t3leading.setLinkUrl('https://developer.apple.com');
    t3leading.setForegroundColor(COLORS.LINK_BLUE);
    t3leading.setUnderline(true);

    const t3para2 = body.appendParagraph('• Trailing space: ');
    const t3trailing = t3para2.appendText('GitHub Docs ');
    t3trailing.setLinkUrl('https://docs.github.com');
    t3trailing.setForegroundColor(COLORS.LINK_BLUE);
    t3trailing.setUnderline(true);

    const t3para3 = body.appendParagraph('• Both leading and trailing spaces: ');
    const t3both = t3para3.appendText('  Wikipedia  ');
    t3both.setLinkUrl('https://www.wikipedia.org');
    t3both.setForegroundColor(COLORS.LINK_BLUE);
    t3both.setUnderline(true);

    body.appendParagraph('Expected: spaces removed from link range; text remains but is not blue/underlined.');
    body.appendParagraph('Reported count: "Trimmed spaces from 3 link(s)"');
    body.appendParagraph('');

    // Tab 3 – Section B: Underlined text without links
    addSection(body, 'B. Underlined Text Without Links (should be PURPLE)');

    const t3ul1 = body.appendParagraph('• This underlined text has no hyperlink at all');
    t3ul1.editAsText().setUnderline(2, 44, true);

    const t3ul2 = body.appendParagraph('• Another underlined phrase with no link attached');
    t3ul2.editAsText().setUnderline(2, 47, true);

    body.appendParagraph('Expected: both underlined runs highlighted in purple.');
    body.appendParagraph('');

    // Tab 3 – Section C: Mix — valid link alongside underlined text
    addSection(body, 'C. Mixed: Valid Link + Underlined Text (link: no highlight; underline: purple)');

    const t3mixed = body.appendParagraph('Visit ');
    appendInlineLink(t3mixed, 'Google', 'https://www.google.com');
    t3mixed.appendText(' or check ');
    const t3mixedUnderline = t3mixed.appendText('our own internal page');
    t3mixedUnderline.setUnderline(true);
    t3mixed.appendText(' for more.');

    body.appendParagraph('Expected: "Google" has no highlight (valid link); "our own internal page" is purple (underlined, no link).');
    body.appendParagraph('');

    // Tab 3 – Summary
    const tab3Summary = body.appendParagraph('Tab 3 Expected Results');
    tab3Summary.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Run "Check Links > In Active Tab" while this tab is active.');
    body.appendParagraph('• Section A: 3 links trimmed — "Trimmed spaces from 3 link(s)" in summary');
    body.appendParagraph('• Section B: 2 purple highlights (underlined text without links)');
    body.appendParagraph('• Section C: valid link unchanged; underlined text highlighted purple');
    body.appendParagraph('• NO red, yellow, or orange highlights — Tab 3 contains none of those link types');
    body.appendParagraph('• Counts should reflect only this tab — none from Tab 1 or Tab 2');
    body.appendParagraph('');
    body.appendParagraph('Then run "Check Links > In Entire Document" from any tab.');
    body.appendParagraph('The reported totals should be the combined sum across all three tabs.');
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

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

// ============================================================================
// MENU
// ============================================================================

/**
 * Creates the custom menu
 */
function onOpen() {
    const ui = DocumentApp.getUi();
    ui.createMenu('Test Generator')
        .addItem('Create Test Document', 'createTestDocument')
        .addToUi();
}

// ============================================================================
// VALIDATION HELPER
// ============================================================================

/**
 * Optional: Function to validate the test results.
 * Run this AFTER running the Document Tools on the test document.
 */
function validateTestResults() {
    Logger.log('=== Test Validation ===');
    Logger.log('Open the document and verify:');
    Logger.log('');

    Logger.log('── TAB 1: LINK CHECKING (Check Links > In Active Tab or In Entire Document) ──');
    Logger.log('1. Section 2 links are RED (broken)');
    Logger.log('2. Section 3 links are YELLOW (Apple.com)');
    Logger.log('3. Section 4 links have NO highlight (skipped non-HTTP)');
    Logger.log('4. Section 5 links are ORANGE (invalid)');
    Logger.log('5. Section 6 underlined text is PURPLE (missing links)');
    Logger.log('6. Section 7 links are now BLUE and UNDERLINED (auto-fixed)');
    Logger.log('7. Section 1 and 10 links have NO highlight (working)');
    Logger.log('8. Section 9 links have spaces trimmed (spaces remain but are not linked/underlined)');
    Logger.log('   → Summary reports "Trimmed spaces from 6 link(s)"');
    Logger.log('');

    Logger.log('── TAB 1: EXTRA SPACE REMOVAL (Fix Document Formatting) ──');
    Logger.log('1. Section 13 DOUBLE SPACE TEST: all double spaces collapsed to single');
    Logger.log('2. Section 13 TRIPLE SPACE TEST: all triple spaces collapsed to single');
    Logger.log('3. Section 13 MIXED SPACE TEST: all runs of 2+ spaces collapsed to single');
    Logger.log('4. Section 13 MANY SPACES TEST: long runs collapsed to single');
    Logger.log('5. Section 13 heading with double spaces: collapsed AND heading font corrected');
    Logger.log('6. Section 13 para with link: spaces outside link collapsed; link unaffected');
    Logger.log('7. Summary includes "X extra space(s) removed" — X > 0 if any were found');
    Logger.log('');

    Logger.log('── TAB 1: FORMATTING RESULTS (Fix Document Formatting) ──');
    Logger.log('1. All Heading 1 text is Helvetica Neue Bold 24pt');
    Logger.log('2. All Heading 2 text is Helvetica Neue Bold 14pt');
    Logger.log('3. All Normal text is Helvetica Neue 11pt (not bold)');
    Logger.log('4. Section 11 paragraphs show corrected fonts');
    Logger.log('');

    Logger.log('── TAB 1: TEXT CASE CONVERSION RESULTS ──');
    Logger.log('1. Character-level formatting (bold, italic, colors) preserved after conversion');
    Logger.log('2. Paragraph-level heading styles preserved after conversion (Section 12a)');
    Logger.log('3. Full selected range remains highlighted after conversion for all five case functions (Section 12b)');
    Logger.log('4. Lower case: All lowercase');
    Logger.log('5. Upper case: All uppercase');
    Logger.log('6. Initial Caps: First Letter Of Each Word');
    Logger.log('7. Sentence case: First letter only');
    Logger.log('8. Title Case: Chicago Style capitalization');
    Logger.log('');

    Logger.log('── TAB 2: CHECK LINKS > IN ACTIVE TAB (switch to Tab 2 first) ──');
    Logger.log('1. Section A (3 links): NO highlights — all working');
    Logger.log('2. Section B (3 links): YELLOW highlights — all Apple.com');
    Logger.log('3. Section C (2 links): RED highlights — all broken/404');
    Logger.log('4. NO orange, purple, or space-trim results reported');
    Logger.log('5. Counts exclude Tab 1 and Tab 3 links entirely');
    Logger.log('');

    Logger.log('── TAB 3: CHECK LINKS > IN ACTIVE TAB (switch to Tab 3 first) ──');
    Logger.log('1. Section A (3 links): spaces trimmed — "Trimmed spaces from 3 link(s)"');
    Logger.log('2. Section B (2 runs): PURPLE highlights — underlined text with no link');
    Logger.log('3. Section C: valid link unchanged (no highlight); underlined text purple');
    Logger.log('4. NO red, yellow, or orange highlights');
    Logger.log('5. Counts exclude Tab 1 and Tab 2 links entirely');
    Logger.log('');

    Logger.log('── ALL TABS: CHECK LINKS > IN ENTIRE DOCUMENT ──');
    Logger.log('1. Reported totals are the sum of all three tabs');
    Logger.log('2. Broken count = Tab 1 broken + Tab 2 broken');
    Logger.log('3. Apple count  = Tab 1 Apple  + Tab 2 Apple');
    Logger.log('4. Trimmed count = Tab 1 trimmed (6) + Tab 3 trimmed (3) = 9');
    Logger.log('5. Purple underlines appear across Tab 1 (Section 6) and Tab 3 (Section B & C)');
    Logger.log('');

    Logger.log('── REPORTING ACCURACY ──');
    Logger.log('1. "Fixed formatting on X link(s)" only counts links that needed changes');
    Logger.log('2. "Trimmed spaces from 6 link(s)" for Tab 1 Section 9 (In Active Tab on Tab 1)');
    Logger.log('3. "Trimmed spaces from 3 link(s)" for Tab 3 Section A (In Active Tab on Tab 3)');
    Logger.log('4. "Trimmed spaces from 9 link(s)" across all tabs (In Entire Document)');
    Logger.log('');

    Logger.log('Check the document visually to confirm all expectations are met.');
}