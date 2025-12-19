# Apple Mail Archive Tool: Retro Macintosh 3-Pane Viewer
**Python Script for Mbox to HTML Conversion with Steve Jobs Minimalist UI**

This open-source Python utility is a specialized email archiving tool designed to convert standard Apple Mail (.mbox) exports into a high-performance, self-contained 3-column web application. It features a retro user interface (UI) inspired by early System 7 Macintosh operating systems and Steve Jobs' minimalist design philosophy.



## Search Engine Optimized Description
This local email backup utility allows users to preserve their digital history in a privacy-focused, offline format. By converting Mbox files to HTML, it eliminates the need for bulky email clients while maintaining a professional, organized structure. This script is optimized for developers and researchers looking for a lightweight email forensics or personal data management solution.

---

## Technical Specifications and Features

* **3-Column Architecture:** Features a dedicated folder sidebar, a message list pane, and a primary reading pane for efficient navigation.
* **International Encoding Support:** Robust handling for Cyrillic (Russian) characters, supporting KOI8-R and Windows-1251 encodings common in historical data.
* **Mbox to HTML Conversion:** Transforms raw mailbox data into scannable, searchable, and responsive web pages.
* **Inline Image Processing:** Automatically renders JPG, PNG, and GIF attachments directly within the email body using local file paths.
* **No External Dependencies:** Built entirely on the Python Standard Library (mailbox, email, html). No pip installation or third-party libraries required.
* **Privacy and Security:** All processing is done locally on your machine. No data is sent to the cloud, making it an ideal tool for sensitive data archiving.

---

## Installation and Usage Flow

### 1. Data Preparation: Exporting Apple Mail
* Launch the Apple Mail application on macOS.
* Select the desired mailboxes or folders.
* Navigate to **Mailbox > Export Mailbox...**
* Save the output to a local directory. This will create folders with the .mbox extension.

### 2. Execution of the Python Script
* Download the `main.py` file to your computer.
* Open your Terminal and execute the following command:
    ```bash
    python3 main.py
    ```
* When prompted for the input path, drag and drop the folder containing your .mbox files into the terminal window and press Enter.

### 3. Archive Access
* The script generates a new directory titled `Mac_Mail_Archive_3Col`.
* Launch the `index.html` file in any modern web browser (Chrome, Safari, Firefox) to view your offline archive.

---

## Troubleshooting and Maintenance

* **UTF-8 Standards:** If text characters appear incorrectly, ensure your browser is set to auto-detect UTF-8 encoding.
* **Relative Path Integrity:** Do not separate the `index.html` file from the `data` folder, as this will break the internal links and image references.
* **Scalability:** The script is capable of handling thousands of messages; however, browser performance during the initial load may vary based on your hardware specifications.

---

## Metadata
* **Version:** 3.0.0
* **Developer Context:** Designed for users seeking a minimalist, distraction-free email reading environment.
* **License:** Open Source / MIT

---

## SEO Keywords
Email archiving software, Python Mbox converter, Apple Mail backup tool, Macintosh System 7 UI, Steve Jobs design aesthetic, offline email viewer, Russian email character fix, Cyrillic email decoder, open source email forensics, local email management, Mbox to HTML script.
