# Writing Tools

> Handy utilities for my authoring process. Mostly written with help from Claude rather than by hand.

A collection of automation scripts and tools designed to streamline the writing and editing workflow, with a focus on Google Docs. This project was developed collaboratively with Claude AI, focusing on practical solutions for common authoring challenges. The codebase prioritises functionality and ease of modification over complex architecture. I set-up a template in Google Docs that has the various Macros attached to it and work from there; at some point I'll likely repackage this as an addon.

The Utils are:  
**Check Links**:  
Checks all HTTP/HTTPS links for broken links (404 errors)  
Highlights broken links in red   
Highlights Apple.com links in yellow (Pages puts these in be default and I’m forever missing the odd one!  
Detects underlined text without links (possible missing links) in purple   
Automatically formats all links with proper blue text and underline   
Trims leading/trailing spaces from linked text (spaces remain as normal text)

Can be applied either to the active tab or the entire document.

Valid non-HTTP protocols are skipped (not flagged as errors):
mailto: (email links)  
tel: (phone numbers)  
sms: (SMS links)  
ftp: / sftp: (file transfer)  
file: (local files)  

Invalid/malformed links are flagged in orange:  
Links that don't start with any recognized protocol  
Typos like htp:// or htps://  
Weird formats or broken URLs  

**Fix Document Formatting**:  
Applies consistent typography throughout the document  
Heading 1: Helvetica Neue Bold 24pt  
Heading 2: Helvetica Neue Bold 14pt  
Normal Text: Helvetica Neue 11pt  

TEXT CASE TOOLS:  
**Lower case**: Converts selected text to lowercase  
**Upper case**: Converts selected text to uppercase  
**Initial Caps**: Capitalises the first letter of each word  
**Sentence case**: Capitalises only the first letter of the selection  
**Title Case** (Chicago Style): Proper title capitalisation following Chicago Manual of Style (more-or-less - to do this properly would be a lot of work but it has a reasonable stab at it!)


## Prerequisites

- Google account (for Google Docs Scripts)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/CHHUM/writing-tools.git
cd writing-tools
```

## Usage

### Google Docs Scripts

To use the Google Docs Scripts:

1. Open your Google Doc
2. Navigate to Extensions > Apps Script
3. Copy the relevant script from the `GoogleDocs Scripts` folder
4. Paste it into the Apps Script editor
5. Save and authorize the script
6. Run from the custom menu or trigger


### Contributing

Contributions are welcome! Whether you have:
- Bug fixes
- New utility ideas
- Documentation improvements
- Performance enhancements

Feel free to open an issue or submit a pull request.

***Note**: This is a personal toolkit and may require customization for your specific use case. Use at your own discretion.*
