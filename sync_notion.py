#!/usr/bin/env python3
"""
sync_notion.py
Utility untuk membaca halaman Notion susunan bal gudang (Blok 1-8 atau Blok 10-16),
mengonversinya ke Markdown, dan memudahkan pembaruan ke create_report.py serta web.
"""

import sys
import os
import re
import json
import urllib.request

def extract_page_id(url_or_id):
    """Mengekstrak 32-karakter hex UUID dari URL Notion atau string UUID."""
    clean = url_or_id.split('?')[0].split('#')[0].rstrip('/')
    slug = clean.split('/')[-1]
    raw = slug[-32:]
    if len(raw) == 32 and all(c in '0123456789abcdefABCDEF' for c in raw):
        return f"{raw[:8]}-{raw[8:12]}-{raw[12:16]}-{raw[16:20]}-{raw[20:]}".lower()
    return None

def fetch_blocks(block_ids):
    """Mengambil rekaman blocks dari API Notion."""
    chunks = [block_ids[i:i + 100] for i in range(0, len(block_ids), 100)]
    all_results = []
    for chunk in chunks:
        req = urllib.request.Request(
            'https://www.notion.so/api/v3/getRecordValues',
            headers={'Content-Type': 'application/json', 'User-Agent': 'Mozilla/5.0'},
            data=json.dumps({'requests': [{'table': 'block', 'id': bid} for bid in chunk]}).encode()
        )
        with urllib.request.urlopen(req) as resp:
            res = json.loads(resp.read().decode())
            for r in res.get('results', []):
                val = r.get('value')
                if isinstance(val, dict):
                    if 'value' in val and isinstance(val['value'], dict):
                        all_results.append(val['value'])
                    else:
                        all_results.append(val)
                else:
                    all_results.append(None)
    return all_results

def fetch_notion_page(url_or_id):
    page_id = extract_page_id(url_or_id)
    if not page_id:
        print(f"[!] Error: Format URL/ID Notion tidak valid: '{url_or_id}'")
        return None, None

    print(f"[*] Menghubungi Notion API untuk Page ID: {page_id}...")
    roots = fetch_blocks([page_id])
    if not roots or not roots[0]:
        print(f"[!] Error: Halaman Notion tidak ditemukan atau tidak memiliki akses publik.")
        return None, None

    root = roots[0]
    title_chunks = root.get('properties', {}).get('title', [])
    page_title = ''.join([t[0] for t in title_chunks]).strip() or 'Halaman Notion'
    print(f"[OK] Judul Halaman: {page_title}")

    content_ids = root.get('content', [])
    print(f"[*] Mengambil {len(content_ids)} blok konten...")
    blocks = fetch_blocks(content_ids)

    sub_ids = []
    for b in blocks:
        if b and b.get('content'):
            sub_ids.extend(b.get('content'))

    sub_blocks_map = {}
    if sub_ids:
        print(f"[*] Mengambil {len(sub_ids)} sub-blok (baris tabel)...")
        sub_blocks = fetch_blocks(sub_ids)
        for sb in sub_blocks:
            if sb and 'id' in sb:
                sub_blocks_map[sb['id']] = sb

    # Render Markdown
    lines = [f"# {page_title}\n"]
    for b in blocks:
        if not b:
            continue
        btype = b.get('type')
        props = b.get('properties', {})

        if btype in ['header', 'sub_header', 'sub_sub_header']:
            prefix = '## ' if btype == 'sub_header' else ('### ' if btype == 'sub_sub_header' else '# ')
            text = ''.join([t[0] for t in props.get('title', [])])
            lines.append(f"\n{prefix}{text}\n")
        elif btype == 'text':
            text = ''.join([t[0] for t in props.get('title', [])])
            lines.append(f"{text}\n")
        elif btype == 'table':
            row_ids = b.get('content', [])
            table_rows = [sub_blocks_map.get(rid) for rid in row_ids if rid in sub_blocks_map]
            format_props = b.get('format', {})
            table_columns = format_props.get('table_block_column_order', [])

            for idx, row in enumerate(table_rows):
                if not row:
                    continue
                rprops = row.get('properties', {})
                row_cells = []
                for col_id in table_columns:
                    cell_val = rprops.get(col_id, [])
                    cell_text = ''.join([t[0] for t in cell_val]).strip()
                    row_cells.append(cell_text)
                lines.append("| " + " | ".join(row_cells) + " |")
                if idx == 0:
                    lines.append("| " + " | ".join(['---'] * len(row_cells)) + " |")
            lines.append("")
        else:
            text = ''.join([t[0] for t in props.get('title', [])])
            lines.append(f"{text}\n")

    markdown_content = "\n".join(lines)
    return page_title, markdown_content

if __name__ == '__main__':
    url = sys.argv[1] if len(sys.argv) > 1 else "https://app.notion.com/p/greendmalik/Blok-1-8-3d97f284d5354cb6847450fc92743e88"
    title, md = fetch_notion_page(url)
    if md:
        print("\n--- Hasil Render Markdown (Awal) ---")
        print(md[:600])
        print("...\n[OK] Sukses membaca halaman Notion!")
