import os
import shutil
import html
import mailbox
import re
import mimetypes
import hashlib
import datetime
import base64
import binascii
from email.header import decode_header
from email.utils import parsedate_to_datetime

# --- CONFIGURATION ---
OUTPUT_DIR_NAME = "Mac_Mail_Archive_Strict_Debug"

# CSS (No changes)
RETRO_CSS = """
<style>
    :root {
        --mac-black: #000000;
        --mac-white: #ffffff;
        --mac-grey: #cccccc;
        --mac-border: 2px solid #000;
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
    .main-view {
        display: flex;
        flex: 1;
        overflow: hidden;
    }
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
        display: flex;
        flex-direction: column;
    }
    .mail-row:hover { background: #f0f0f0; }
    .mail-row.selected { background: #000; color: #fff; }
    .mail-row-header { display: flex; justify-content: space-between; margin-bottom: 4px; }
    .mail-row-sender { font-weight: bold; }
    .mail-row-date { font-size: 10px; color: #666; }
    .mail-row.selected .mail-row-date { color: #ccc; }
    .mail-row-subject { 
        white-space: nowrap; 
        overflow: hidden; 
        text-overflow: ellipsis; 
    }
    .thread-count {
        background: #ccc;
        color: #000;
        border-radius: 8px;
        padding: 0 5px;
        font-size: 9px;
        margin-left: 5px;
        font-weight: bold;
        display: inline-block;
    }
    .mail-row.selected .thread-count { background: #fff; color: #000; }
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
        display: none;
    }
    ::-webkit-scrollbar { width: 10px; }
    ::-webkit-scrollbar-track { background: #fff; border-left: 1px solid #000; }
    ::-webkit-scrollbar-thumb { background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAYAAACp8Z5+AAAAIklEQVQIW2NkQAKrVq36zwjjgzjIHFBmAAxxDatWrfoPAQA7YAxbHYyPOAAAAABJRU5ErkJggg=='); border: 1px solid #000; }
</style>
"""


# --- UTILITIES ---
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


def parse_date_strict(date_str):
    if not date_str: return datetime.datetime.min.replace(tzinfo=datetime.timezone.utc)
    try:
        dt = parsedate_to_datetime(date_str)
        if dt.tzinfo is None: dt = dt.replace(tzinfo=datetime.timezone.utc)
        return dt
    except:
        return datetime.datetime.min.replace(tzinfo=datetime.timezone.utc)


def extract_msg_id(msg):
    mid = msg.get('Message-ID', '').strip()
    if mid.startswith('<') and mid.endswith('>'): mid = mid[1:-1]
    return mid


def extract_references(msg):
    refs = []
    # 1. References
    ref_header = msg.get('References', '')
    if ref_header:
        found = re.findall(r'<([^>]+)>', ref_header)
        refs.extend(found)
    # 2. In-Reply-To
    irt_header = msg.get('In-Reply-To', '')
    if irt_header:
        found = re.findall(r'<([^>]+)>', irt_header)
        refs.extend(found)

    seen = set()
    unique_refs = []
    for r in refs:
        if r and r not in seen:
            unique_refs.append(r)
            seen.add(r)
    return unique_refs


def extract_thread_index(msg):
    """
    Extracts the Microsoft/Exchange Thread-Index.
    This is a base64 string. The first 22 bytes (30 chars) are the GUID for the conversation.
    """
    ti = msg.get('Thread-Index', '').strip()
    if not ti: return None

    # We only care about the conversation ID (first 30 chars/22 bytes)
    # This groups "RE: Hello" and "FW: Hello" even if subjects change
    try:
        # Just return the raw string prefix as a grouping key
        return ti[:30]
    except:
        return None


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
                    decoded = safe_decode(payload, [part.get_content_charset() or 'utf-8'])
                    if ctype == "text/html":
                        body = decoded
                    elif ctype == "text/plain" and not body:
                        body = f"<pre>{html.escape(decoded)}</pre>"
            except:
                pass
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            decoded = safe_decode(payload, [msg.get_content_charset() or 'utf-8'])
            body = decoded if msg.get_content_type() == "text/html" else f"<pre>{html.escape(decoded)}</pre>"
    return body, attachments


# --- CORE THREADING CLASSES ---

class Node:
    def __init__(self, msg_id):
        self.msg_id = msg_id
        self.message = None  # None = Ghost
        self.parent = None
        self.children = []

    def add_child(self, child_node):
        if child_node.parent: return  # Already attached
        child_node.parent = self
        self.children.append(child_node)

    def get_root(self):
        curr = self
        while curr.parent:
            curr = curr.parent
        return curr

    def walk(self):
        result = []
        if self.message: result.append(self.message)
        for child in self.children:
            result.extend(child.walk())
        return result

    @property
    def date(self):
        if self.message: return self.message['dt']
        dates = [c.date for c in self.children]
        if dates: return min(dates)
        return datetime.datetime.min.replace(tzinfo=datetime.timezone.utc)


def main():
    print("--- DEBUG GRAPH THREADING ENGINE (STRICT) ---")
    raw_input = input("Drag and drop your 'Mail Export' folder here: ").strip()
    input_path = raw_input.strip("'").strip('"')

    if not os.path.exists(input_path): return print("Folder not found.")

    original_folder_name = os.path.basename(input_path.rstrip(os.sep))
    output_path = os.path.join(os.path.dirname(input_path), f"{original_folder_name}_Debug_Threaded")

    if os.path.exists(output_path): shutil.rmtree(output_path)
    data_dir = os.path.join(output_path, "data")
    os.makedirs(data_dir)

    # Global Lookups
    nodes = {}  # msg_id -> Node
    thread_index_map = {}  # thread_index_prefix -> Node (Root or parent of that thread)
    folder_counts = {}
    msg_counter = 0

    # --- PHASE 1: INGEST AND CREATE NODES ---
    print("\n[PHASE 1] Ingesting Messages & Extracting IDs...")

    for root, _, _ in os.walk(input_path):
        if root.endswith(".mbox") and os.path.exists(os.path.join(root, "mbox")):
            folder_name = os.path.relpath(root, input_path).replace(".mbox", "")
            if folder_name not in folder_counts: folder_counts[folder_name] = 0

            try:
                for msg in mailbox.mbox(os.path.join(root, "mbox")):
                    folder_counts[folder_name] += 1
                    msg_counter += 1
                    local_id = f"m{msg_counter}"

                    # Data Extraction
                    subj = decode_header_safe(msg.get('subject', '(No Subject)'))
                    date_str = decode_header_safe(msg.get('date', ''))
                    dt_obj = parse_date_strict(date_str)
                    mid = extract_msg_id(msg)
                    refs = extract_references(msg)
                    ti = extract_thread_index(msg)

                    # --- DEBUG PRINT ---
                    if msg_counter % 100 == 0: print(f"Processing {msg_counter}...")
                    # print(f"[DEBUG] Msg: {subj[:30]}... | ID: {mid} | Refs: {len(refs)} | Thread-Index: {ti}")

                    # Save Body
                    att_dir = os.path.join(data_dir, f"{local_id}_att")
                    body, atts = extract_content(msg, att_dir)

                    # HTML Frag (Same as before)
                    att_html = ""
                    if atts:
                        links = "".join([
                                            f"<li><a href='{local_id}_att/{a['name']}' target='_blank'>{html.escape(a['name'])}</a></li>"
                                            for a in atts if not a['is_image']])
                        imgs = "".join([
                                           f"<img src='{local_id}_att/{a['name']}' style='max-width:100%; border:1px solid #000; margin:10px 0;'>"
                                           for a in atts if a['is_image']])
                        if links: att_html += f"<div style='border:1px dashed #000; padding:10px; background:#eee; margin-bottom:10px;'><b>Attachments:</b><ul>{links}</ul></div>"
                        if imgs: att_html += f"<div>{imgs}</div>"

                    content_fragment = f"""
                    <div class="email-container" id="{local_id}" style="border:1px solid #ccc; margin-bottom:20px;">
                        <div class="email-header" style="background:#f4f4f4; padding:8px; border-bottom:1px solid #ddd;">
                            <div style="float:right; font-size:11px; color:#666;">{html.escape(date_str)}</div>
                            <div class="email-meta"><b>From:</b> {html.escape(decode_header_safe(msg.get('from', 'Unknown')))}</div>
                            <div class="email-meta"><b>Folder:</b> {html.escape(folder_name)}</div>
                            <div class="email-meta"><b>Subject:</b> {html.escape(subj)}</div>
                            <div class="email-meta" style="font-size:10px; color:#999;"><b>Debug ID:</b> {mid}</div>
                        </div>
                        <div class="email-body" style="padding:15px;">{att_html}{body}</div>
                    </div>
                    """

                    with open(os.path.join(data_dir, f"{local_id}.frag"), "w", encoding="utf-8") as f:
                        f.write(content_fragment)

                    # Create Node
                    if not mid:
                        hasher = hashlib.md5()
                        hasher.update((subj + str(dt_obj.timestamp())).encode('utf-8'))
                        mid = f"synth_{hasher.hexdigest()}"

                    if mid not in nodes:
                        nodes[mid] = Node(mid)

                    # Store Data
                    nodes[mid].message = {
                        'local_id': local_id,
                        'mid': mid,
                        'subj': subj,
                        'date_str': date_str,
                        'dt': dt_obj,
                        'sender': decode_header_safe(msg.get('from', '')),
                        'folder': folder_name,
                        'refs': refs,
                        'thread_index': ti
                    }

            except Exception as e:
                print(f"Skipping corrupt message: {e}")

    # --- PHASE 2: STRICT LINKING (NO FUZZY SUBJECTS) ---
    print("\n[PHASE 2] Linking via References & Thread-Index...")

    # 1. Standard JWZ (References)
    # We iterate a copy to allow adding ghosts safely
    for mid, node in list(nodes.items()):
        if not node.message: continue

        refs = node.message['refs']

        if refs:
            print(f"[DEBUG-LINK] '{node.message['subj'][:20]}' ({mid}) has {len(refs)} refs.")

            # Ensure chain exists
            prev = None
            for r in refs:
                if r not in nodes:
                    nodes[r] = Node(r)  # Create Ghost

                curr = nodes[r]
                if prev and not curr.parent and prev != curr:
                    print(f"   -> Chaining Ref {prev.msg_id} -> {curr.msg_id}")
                    prev.add_child(curr)
                prev = curr

            # Link current to last ref
            last_ref = nodes[refs[-1]]
            if last_ref != node and not node.parent:
                print(f"   -> Linking Message to Parent {last_ref.msg_id}")
                last_ref.add_child(node)

    # 2. Microsoft Thread-Index (The "Missing Data" Fix)
    # This groups messages that share the same conversation GUID but lost their References
    print("\n[PHASE 2.5] Linking via Thread-Index (Outlook Grouping)...")

    # Bucket by Thread-Index prefix
    ti_buckets = {}
    for mid, node in nodes.items():
        if node.message and node.message['thread_index']:
            ti = node.message['thread_index']
            if ti not in ti_buckets: ti_buckets[ti] = []
            ti_buckets[ti].append(node)

    for ti, node_list in ti_buckets.items():
        if len(node_list) < 2: continue

        # Sort by date
        node_list.sort(key=lambda x: x.date)

        # We have a list of messages that DEFINITELY belong together (cryptographically linked)
        # If they aren't already linked via References, link them now chronologically.
        for i in range(len(node_list) - 1):
            parent = node_list[i]
            child = node_list[i + 1]

            # Only link if child is orphan (or we are repairing a broken tree)
            if not child.parent and child != parent:
                # Check if they are already in the same tree?
                # Actually, if they share a TI, they ARE the same thread.
                # If 'child' has no parent, attach to 'parent'.
                print(
                    f"[DEBUG-MS-LINK] Linking via Thread-Index: {child.message['subj'][:20]} -> {parent.message['subj'][:20]}")
                parent.add_child(child)

    # --- PHASE 3: REMOVED (NO SUBJECT MERGING) ---
    print("\n[PHASE 3] Subject Merging DISABLED (Preventing 'Stacking')...")

    # --- PHASE 4: FLATTEN & SORT ---
    print("\n[PHASE 4] Generating Thread Views...")

    final_roots = [n for n in nodes.values() if n.parent is None]
    final_threads = []
    thread_id_counter = 0

    for root in final_roots:
        msgs = root.walk()
        if not msgs: continue

        msgs.sort(key=lambda x: x['dt'])

        folders = set(m['folder'] for m in msgs)
        folders_str = "||".join(sorted(folders))

        latest_msg = msgs[-1]
        thread_id_counter += 1
        tid = f"t{thread_id_counter}"

        html_content = ""
        for m in msgs:
            fpath = os.path.join(data_dir, f"{m['local_id']}.frag")
            if os.path.exists(fpath):
                with open(fpath, "r", encoding="utf-8") as f:
                    html_content += f.read()

        full_html = f"""
        <!DOCTYPE html><html><head><meta charset="UTF-8">
        <style>
            body {{ font-family: "Geneva", sans-serif; padding: 20px; font-size: 14px; background: #fff; }}
            pre {{ white-space: pre-wrap; font-family: Courier; }}
            img {{ max-width: 100%; height: auto; }}
        </style>
        </head><body>
        <h2 style='border-bottom: 2px solid black; padding-bottom:10px;'>Topic: {html.escape(latest_msg['subj'])}</h2>
        {html_content}
        </body></html>
        """

        with open(os.path.join(data_dir, f"{tid}.html"), "w", encoding="utf-8") as f:
            f.write(full_html)

        final_threads.append({
            'tid': tid,
            'subj': latest_msg['subj'],
            'sender': latest_msg['sender'],
            'date_str': latest_msg['date_str'],
            'sort_dt': latest_msg['dt'],
            'folders': folders_str,
            'count': len(msgs)
        })

    final_threads.sort(key=lambda x: x['sort_dt'], reverse=True)

    # --- PHASE 5: INDEX HTML ---
    folder_html = "".join(
        [f'<div class="folder-item" onclick="filterFolder(\'{f}\')"><span>{f}</span><span>{c}</span></div>' for f, c in
         sorted(folder_counts.items())])

    rows_html = ""
    for t in final_threads:
        badge = f'<span class="thread-count">{t["count"]}</span>' if t["count"] > 1 else ""
        rows_html += f"""
        <div class="mail-row" data-folders="{html.escape(t['folders'])}" onclick="loadEmail('data/{t['tid']}.html', this)">
            <div class="mail-row-header">
                <div class="mail-row-sender">{html.escape(t['sender'][:30])}</div>
                <div class="mail-row-date">{html.escape(t['date_str'][:10])}</div>
            </div>
            <div class="mail-row-subject">
                {html.escape(t['subj'])}{badge}
            </div>
        </div>
        """

    index_html = f"""
    <!DOCTYPE html>
    <html>
    <head><meta charset="UTF-8"><title>Mail Archive</title>{RETRO_CSS}</head>
    <body>
        <div class="window">
            <div class="title-bar"><div class="title-text">{original_folder_name} Archive - {len(final_threads)} Conversations</div></div>
            <div class="main-view">
                <div class="sidebar">
                    <div class="folder-item active" onclick="filterFolder('all')"><span>All Mailboxes</span><span>{sum(folder_counts.values())}</span></div>
                    {folder_html}
                </div>
                <div class="list-pane" id="emailList">
                    {rows_html}
                </div>
                <div class="preview-pane">
                    <div id="placeholder" class="preview-placeholder">Select a conversation</div>
                    <iframe id="previewFrame" name="previewFrame"></iframe>
                </div>
            </div>
        </div>
        <script>
            function filterFolder(folderName) {{
                document.querySelectorAll('.folder-item').forEach(el => el.classList.remove('active'));
                event.currentTarget.classList.add('active');

                const rows = document.querySelectorAll('.mail-row');
                rows.forEach(row => {{
                    const rowFolders = row.getAttribute('data-folders').split('||');
                    if (folderName === 'all' || rowFolders.includes(folderName)) {{
                        row.style.display = 'flex';
                    }} else {{
                        row.style.display = 'none';
                    }}
                }});
            }}
            function loadEmail(url, el) {{
                document.querySelectorAll('.mail-row').forEach(row => row.classList.remove('selected'));
                el.classList.add('selected');
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

    print(f"Done! Created STRICT archive at: {output_path}")


if __name__ == "__main__":
    main()