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
    """Mengambil rekaman blocks dari API Notion getRecordValues sebagai fallback."""
    chunks = [block_ids[i:i + 100] for i in range(0, len(block_ids), 100)]
    all_results = []
    for chunk in chunks:
        req = urllib.request.Request(
            'https://www.notion.so/api/v3/getRecordValues',
            headers={'Content-Type': 'application/json', 'User-Agent': 'Mozilla/5.0'},
            data=json.dumps({'requests': [{'table': 'block', 'id': bid} for bid in chunk]}).encode()
        )
        try:
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
        except Exception:
            pass
    return all_results

def fetch_page_record_map(page_id):
    """Mengambil rekaman blocks via loadPageChunk (metode paling andal untuk public page)."""
    req = urllib.request.Request(
        'https://www.notion.so/api/v3/loadPageChunk',
        headers={'Content-Type': 'application/json', 'User-Agent': 'Mozilla/5.0'},
        data=json.dumps({'pageId': page_id, 'limit': 100, 'cursor': {'stack': []}, 'chunkNumber': 0, 'verticalColumns': False}).encode()
    )
    with urllib.request.urlopen(req) as resp:
        res = json.loads(resp.read().decode())
    raw_blocks = res.get('recordMap', {}).get('block', {})
    block_map = {}
    for k, v in raw_blocks.items():
        val = v.get('value')
        if isinstance(val, dict) and 'value' in val and isinstance(val['value'], dict):
            block_map[k] = val['value']
        elif isinstance(val, dict):
            block_map[k] = val
    return block_map

def fetch_notion_page(url_or_id):
    page_id = extract_page_id(url_or_id)
    if not page_id:
        print(f"[!] Error: Format URL/ID Notion tidak valid: '{url_or_id}'")
        return None, None

    print(f"[*] Menghubungi Notion API untuk Page ID: {page_id}...")
    try:
        block_map = fetch_page_record_map(page_id)
    except Exception as e:
        print(f"[*] Gagal memanggil loadPageChunk: {e}. Mencoba getRecordValues...")
        roots = fetch_blocks([page_id])
        if not roots or not roots[0]:
            print(f"[!] Error: Halaman Notion tidak ditemukan atau tidak memiliki akses publik.")
            return None, None
        root = roots[0]
        content_ids = root.get('content', [])
        blocks = fetch_blocks(content_ids)
        sub_ids = []
        for b in blocks:
            if b and b.get('content'):
                sub_ids.extend(b.get('content'))
        sub_blocks_map = {}
        if sub_ids:
            sub_blocks = fetch_blocks(sub_ids)
            for sb in sub_blocks:
                if sb and 'id' in sb:
                    sub_blocks_map[sb['id']] = sb
        block_map = {page_id: root}
        for b in blocks:
            if b and 'id' in b:
                block_map[b['id']] = b
        block_map.update(sub_blocks_map)

    root = block_map.get(page_id)
    if not root:
        print(f"[!] Error: Root block tidak ditemukan pada respon Notion.")
        return None, None

    title_chunks = root.get('properties', {}).get('title', [])
    page_title = ''.join([t[0] for t in title_chunks]).strip() or 'Halaman Notion'
    print(f"[OK] Judul Halaman: {page_title}")

    content_ids = root.get('content', [])
    print(f"[*] Memproses {len(content_ids)} blok konten...")

    # Render Markdown
    lines = [f"# {page_title}\n"]
    for cid in content_ids:
        b = block_map.get(cid)
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
            table_rows = [block_map.get(rid) for rid in row_ids if rid in block_map]
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
        target_file = f"Blok 1–8 3d97f284d5354cb6847450fc92743e88.md"
        with open(target_file, "w", encoding="utf-8") as f:
            f.write(md)
        print(f"[OK] Hasil Markdown berhasil disimpan ke: {target_file}")
        print("\n--- Hasil Render Markdown (Awal) ---")
        print(md[:600])
        print("...\n[OK] Sukses membaca dan menyinkronkan halaman Notion!")
