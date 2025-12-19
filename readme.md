# Apple Mail Archive Tool: Retro Macintosh 3-Pane Viewer
**Python Script for Mbox to HTML Conversion with Strict Threading & Minimalist UI**

This open-source Python utility is a specialized email archiving tool designed to convert standard Apple Mail (.mbox) exports into a high-performance, self-contained 3-column web application. It features a retro user interface (UI) inspired by early System 7 Macintosh operating systems and uses advanced threading algorithms to reconstruct conversations accurately.

<img width="600" alt="Screenshot 2025-12-18 at 22 35 38" src="https://github.com/user-attachments/assets/81c96d41-5880-4c1c-8cc7-f7b4ab209de6" />

---

## Technical Specifications and Features

* **Strict Graph Threading Engine (JWZ & Exchange):** Implements a dual-layer threading algorithm. It uses the standard JWZ algorithm (Message-ID references) and falls back to Microsoft's `Thread-Index` header to accurately group Outlook/Exchange conversations, preventing "fuzzy" matching errors.
* **3-Column Architecture:** Features a dedicated folder sidebar, a thread-aware message list pane, and a primary reading pane.
* **ISO Date Formatting:** All timestamps are normalized to `YYYY-MM-DD HH:MM` (24-hour format) for clarity and international consistency.
* **Smart UI Indicators:** Threaded conversations display a message count badge *preceding* the subject line for quick scanning.
* **International Encoding Support:** Robust handling for Cyrillic (Russian) characters, supporting KOI8-R and Windows-1251 encodings common in historical data.
* **Inline Image Processing:** Automatically renders JPG, PNG, and GIF attachments directly within the email body using local file paths.
* **No External Dependencies:** Built entirely on the Python Standard Library (`mailbox`, `email`, `html`, `mimetypes`, `datetime`). No pip installation required.
* **Privacy and Security:** All processing is done locally on your machine. No data is sent to the cloud.

---

## Installation and Usage Flow

### 1. Data Preparation: Exporting Apple Mail
* Launch the Apple Mail application on macOS.
* Select the desired mailboxes or folders.
* Navigate to **Mailbox > Export Mailbox...**
* Save the output to a local directory. This will create folders with the .mbox extension (e.g., `MyExport`).

### 2. Execution of the Python Script
* Save the script as `main.py` on your computer.
* Open your Terminal and execute the following command:
    ```bash
    python3 main.py
    ```
* When prompted for the input path, drag and drop the folder containing your .mbox files (e.g., `MyExport`) into the terminal window and press Enter.
* **Note:** The script will output verbose debug logs ("Phase 1", "Phase 2") showing exactly how messages are being linked.

### 3. Archive Access
* The script generates a new directory titled with your original folder name plus a `_Debug_Threaded` suffix (e.g., `MyExport_Debug_Threaded`).
* Launch the `index.html` file inside that new folder using any modern web browser to view your offline archive.

---

## Troubleshooting and Maintenance

* **Ghost Nodes:** If you see threads that seem to start in the middle of a conversation, it is likely because the original "root" email was not present in your export. The script handles this gracefully by creating invisible "ghost" parents to keep the tree structure intact.
* **Relative Path Integrity:** Do not separate the `index.html` file from the `data` folder, as this will break the internal links and image references.
* **Browser Performance:** The script is capable of handling thousands of messages; however, extremely large archives (10,000+ threads) may require a few seconds to load the initial index.

---

## Metadata
* **Version:** 4.0.0 (Strict Threading Update)
* **Developer Context:** Designed for users seeking a minimalist, distraction-free email reading environment with forensic-level threading accuracy.
* **License:** Open Source / MIT

---


## Search Engine Optimized Description
This local email backup utility allows users to preserve their digital history in a privacy-focused, offline format. By converting Mbox files to HTML, it eliminates the need for bulky email clients while maintaining a professional, organized structure. This script is optimized for developers and researchers looking for a lightweight email forensics or personal data management solution with robust conversation threading.
## SEO Keywords
Email archiving software, Python Mbox converter, Apple Mail backup tool, JWZ threading algorithm, Microsoft Thread-Index parser, Macintosh System 7 UI, Steve Jobs design aesthetic, offline email viewer, Russian email character fix, Cyrillic email decoder, open source email forensics, local email management.
