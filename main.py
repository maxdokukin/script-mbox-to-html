import os
import shutil
import html
import mailbox
import re
import mimetypes
from email.header import decode_header
from email.utils import parsedate_to_datetime

# --- CONFIGURATION ---
OUTPUT_DIR_NAME = "Mac_Mail_Archive_3Col"

RETRO_CSS = """
<style>
    :root {
        --mac-black: #000000;
        --mac-white: #ffffff;
        --mac-grey: #cccccc;
        --mac-border: 2px solid #000;
        --mac-highlight: #000000;
        --mac-highlight-text: #ffffff;
    }
    * { box-sizing: border-box; }
    body, html {
        margin: 0; padding: 0;
        height: 100%; width: 100%;
        font-family: "Geneva", "Verdana", sans-serif;
        background-color: #777;
        background-image: radial-gradient(#999 25%, transparent 25%);
        background-size: 4px 4px;
        overflow: hidden;
    }

    .window {
        display: flex;
        flex-direction: column;
        height: 100vh;
        width: 100vw;
        background: var(--mac-white);
        border: var(--mac-border);
    }

    /* TITLE BAR */
    .title-bar {
        height: 30px;
        border-bottom: var(--mac-border);
        background: repeating-linear-gradient(0deg, #fff, #fff 2px, #999 2px, #999 4px);
        display: flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
    }
    .title-text {
        background: white;
        padding: 2px 20px;
        border: 1px solid black;
        box-shadow: 2px 2px 0 #000;
        font-weight: bold;
        font-size: 13px;
    }

    /* 3-COLUMN LAYOUT */
    .main-view {
        display: flex;
        flex: 1;
        overflow: hidden;
    }

    /* COL 1: SIDEBAR (Folders) */
    .sidebar {
        width: 220px;
        min-width: 200px;
        border-right: var(--mac-border);
        background: #eee;
        overflow-y: auto;
        display: flex;
        flex-direction: column;
    }
    .folder-item {
        padding: 8px 12px;
        cursor: pointer;
        font-size: 12px;
        border-bottom: 1px dotted #ccc;
        display: flex;
        justify-content: space-between;
    }
    .folder-item:hover { background: #ccc; }
    .folder-item.active { background: #000; color: #fff; }

    /* COL 2: LIST (Emails) */
    .list-pane {
        width: 350px;
        min-width: 300px;
        border-right: var(--mac-border);
        background: white;
        overflow-y: auto;
        display: flex;
        flex-direction: column;
    }
    .mail-row {
        cursor: pointer;
        border-bottom: 1px solid #ddd;
        padding: 8px;
        font-size: 12px;
    }
    .mail-row:hover { background: #f0f0f0; }
    .mail-row.selected { background: #000; color: #fff; }

    .mail-row-sender { font-weight: bold; margin-bottom: 2px; }
    .mail-row-subject { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .mail-row-date { font-size: 10px; color: #666; float: right; }
    .mail-row.selected .mail-row-date { color: #ccc; }

    /* COL 3: READING PANE */
    .preview-pane {
        flex: 1;
        background: #fff;
        display: flex;
        flex-direction: column;
        position: relative;
    }
    .preview-placeholder {
        display: flex;
        align-items: center;
        justify-content: center;
        height: 100%;
        color: #999;
        font-style: italic;
    }
    iframe {
        width: 100%;
        height: 100%;
        border: none;
        display: none; /* Hidden until loaded */
    }

    /* SCROLLBARS (Webkit) */
    ::-webkit-scrollbar { width: 10px; }
    ::-webkit-scrollbar-track { background: #fff; border-left: 1px solid #000; }
    ::-webkit-scrollbar-thumb { background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAYAAACp8Z5+AAAAIklEQVQIW2NkQAKrVq36zwjjgzjIHFBmAAxxDatWrfoPAQA7YAxbHYyPOAAAAABJRU5ErkJggg=='); border: 1px solid #000; }
</style>
"""


# --- EMAIL HEADER FIXES ---
def safe_decode(text_bytes, encodings=['utf-8', 'windows-1251', 'koi8-r', 'latin1']):
    if not text_bytes: return ""
    for enc in encodings:
        try:
            return text_bytes.decode(enc)
        except:
            continue
    return text_bytes.decode('utf-8', errors='replace')


def decode_header_safe(header_val):
    if not header_val: return ""
    try:
        decoded_parts = decode_header(header_val)
        header_text = ""
        for bytes_part, encoding in decoded_parts:
            if isinstance(bytes_part, bytes):
                try:
                    header_text += bytes_part.decode(encoding or 'utf-8')
                except:
                    header_text += safe_decode(bytes_part)
            else:
                header_text += str(bytes_part)
        return header_text
    except:
        return str(header_val)


def clean_filename(filename):
    if not filename: return "untitled"
    return re.sub(r'[\\/*?:"<>|]', "_", decode_header_safe(filename))[:60]


def is_image(filename):
    guess, _ = mimetypes.guess_type(filename)
    return guess and guess.startswith('image')


def extract_content(msg, attachment_dir):
    body, attachments = "", []
    if not os.path.exists(attachment_dir): os.makedirs(attachment_dir)

    if msg.is_multipart():
        for part in msg.walk():
            fname = part.get_filename()
            ctype = part.get_content_type()

            if fname or "image" in ctype:
                if not fname: fname = f"embedded_{len(attachments)}" + (mimetypes.guess_extension(ctype) or ".bin")
                safe_name = clean_filename(fname)
                payload = part.get_payload(decode=True)
                if payload:
                    with open(os.path.join(attachment_dir, safe_name), "wb") as f: f.write(payload)
                    attachments.append({"name": safe_name, "is_image": is_image(safe_name)})
                continue

            try:
                payload = part.get_payload(decode=True)
                if payload:
                    decoded = safe_decode(payload, [part.get_content_charset() or 'utf-8', 'windows-1251', 'koi8-r'])
                    if ctype == "text/html":
                        body = decoded
                    elif ctype == "text/plain" and not body:
                        body = f"<pre>{html.escape(decoded)}</pre>"
            except:
                pass
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            decoded = safe_decode(payload, [msg.get_content_charset() or 'utf-8', 'windows-1251', 'koi8-r'])
            body = decoded if msg.get_content_type() == "text/html" else f"<pre>{html.escape(decoded)}</pre>"

    return body, attachments


def main():
    print("---  Mac Mail 3-Pane Exporter ---")
    raw_input = input("Drag and drop your 'Mail Export' folder here: ").strip()
    input_path = raw_input.strip("'").strip('"')

    if not os.path.exists(input_path): return print("Folder not found.")

    output_path = os.path.join(os.path.dirname(input_path), OUTPUT_DIR_NAME)
    if os.path.exists(output_path): shutil.rmtree(output_path)
    data_dir = os.path.join(output_path, "data")
    os.makedirs(data_dir)

    all_emails = []
    folder_counts = {}

    for root, dirs, files in os.walk(input_path):
        if root.endswith(".mbox") and os.path.exists(os.path.join(root, "mbox")):
            folder_name = os.path.relpath(root, input_path).replace(".mbox", "")
            print(f"Processing: {folder_name}...")
            folder_counts[folder_name] = 0

            try:
                for msg in mailbox.mbox(os.path.join(root, "mbox")):
                    folder_counts[folder_name] += 1
                    email_id = f"msg_{len(all_emails)}"

                    subj = decode_header_safe(msg.get('subject', '(No Subject)'))
                    sender = decode_header_safe(msg.get('from', 'Unknown'))
                    date = decode_header_safe(msg.get('date', ''))

                    # Generate Email File
                    att_dir = os.path.join(data_dir, f"{email_id}_att")
                    body, atts = extract_content(msg, att_dir)

                    att_html = ""
                    if atts:
                        links = "".join([
                                            f"<li><a href='{email_id}_att/{a['name']}' target='_blank'>{html.escape(a['name'])}</a></li>"
                                            for a in atts if not a['is_image']])
                        imgs = "".join([
                                           f"<img src='{email_id}_att/{a['name']}' style='max-width:100%; border:1px solid #000; margin:10px 0;'>"
                                           for a in atts if a['is_image']])
                        if links: att_html += f"<div style='border:1px dashed #000; padding:10px; background:#eee;'><b>Attachments:</b><ul>{links}</ul></div>"
                        if imgs: att_html += f"<div>{imgs}</div>"

                    # Email HTML Template (Embedded in iframe)
                    with open(os.path.join(data_dir, f"{email_id}.html"), "w", encoding="utf-8") as f:
                        f.write(f"""
                        <!DOCTYPE html><html><head><meta charset="UTF-8">
                        <style>
                            body {{ font-family: "Geneva", sans-serif; padding: 20px; font-size: 14px; }}
                            .header {{ border-bottom: 2px solid black; padding-bottom: 10px; margin-bottom: 20px; background: #f9f9f9; padding: 15px; }}
                            h2 {{ margin: 0 0 10px 0; font-size: 18px; }}
                            .meta {{ color: #555; margin-bottom: 5px; }}
                            pre {{ white-space: pre-wrap; font-family: Courier; }}
                        </style>
                        </head><body>
                        <div class="header">
                            <h2>{html.escape(subj)}</h2>
                            <div class="meta"><b>From:</b> {html.escape(sender)}</div>
                            <div class="meta"><b>Date:</b> {html.escape(date)}</div>
                        </div>
                        {att_html}
                        <div>{body}</div>
                        </body></html>
                        """)

                    all_emails.append({
                        "id": email_id, "folder": folder_name, "subj": subj, "sender": sender, "date": date
                    })
            except Exception as e:
                print(f"Error reading {folder_name}: {e}")

    # --- GENERATE INDEX.HTML ---

    # Sort folders alphabetically
    folder_html = "".join(
        [f'<div class="folder-item" onclick="filterFolder(\'{f}\')"><span>{f}</span><span>{c}</span></div>' for f, c in
         sorted(folder_counts.items())])

    # Generate List Rows
    rows_html = ""
    for e in all_emails:
        rows_html += f"""
        <div class="mail-row" data-folder="{html.escape(e['folder'])}" onclick="loadEmail('data/{e['id']}.html', this)">
            <div class="mail-row-date">{html.escape(e['date'][:16])}</div>
            <div class="mail-row-sender">{html.escape(e['sender'][:35])}</div>
            <div class="mail-row-subject">{html.escape(e['subj'])}</div>
        </div>
        """

    index_html = f"""
    <!DOCTYPE html>
    <html>
    <head><meta charset="UTF-8"><title>Mail Archive</title>{RETRO_CSS}</head>
    <body>
        <div class="window">
            <div class="title-bar"><div class="title-text">Mail Archive - {len(all_emails)} items</div></div>
            <div class="main-view">
                <div class="sidebar">
                    <div class="folder-item active" onclick="filterFolder('all')"><span>All Mailboxes</span><span>{len(all_emails)}</span></div>
                    {folder_html}
                </div>
                <div class="list-pane" id="emailList">
                    {rows_html}
                </div>
                <div class="preview-pane">
                    <div id="placeholder" class="preview-placeholder">Select an email to view</div>
                    <iframe id="previewFrame" name="previewFrame"></iframe>
                </div>
            </div>
        </div>
        <script>
            function filterFolder(folderName) {{
                // Update Sidebar UI
                document.querySelectorAll('.folder-item').forEach(el => el.classList.remove('active'));
                event.currentTarget.classList.add('active');

                // Filter List
                const rows = document.querySelectorAll('.mail-row');
                rows.forEach(row => {{
                    row.style.display = (folderName === 'all' || row.getAttribute('data-folder') === folderName) ? 'block' : 'none';
                }});
            }}

            function loadEmail(url, el) {{
                // Update List UI (Selection State)
                document.querySelectorAll('.mail-row').forEach(row => row.classList.remove('selected'));
                el.classList.add('selected');

                // Load Iframe
                document.getElementById('placeholder').style.display = 'none';
                const frame = document.getElementById('previewFrame');
                frame.style.display = 'block';
                frame.src = url;
            }}
        </script>
    </body>
    </html>
    """

    with open(os.path.join(output_path, "index.html"), "w", encoding="utf-8") as f:
        f.write(index_html)
    print(f"\n✅ Done! Open: {output_path}/index.html")


if __name__ == "__main__":
    main()