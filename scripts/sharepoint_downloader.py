#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SharePoint / OneDrive Sharing Link Downloader
Support password-protected and direct guest-sharing links on Microsoft 365 / SharePoint (including sharepoint.cn and sharepoint.com).

Usage:
    python sharepoint_downloader.py --url "<SHARE_URL>" --password "<PWD>" --output "<DEST_DIR>" [--concurrency 4]
"""

import os
import sys
import time
import argparse
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Any, Optional

try:
    import requests
except ImportError:
    print("Error: requests is not installed. Run `pip install requests`.")
    sys.exit(1)

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("Error: playwright is not installed. Run `pip install playwright`.")
    sys.exit(1)


def parse_args():
    parser = argparse.ArgumentParser(description="SharePoint & OneDrive folder/file downloader")
    parser.add_argument("--url", "-u", required=True, help="SharePoint or OneDrive guest sharing URL")
    parser.add_argument("--password", "-p", default="", help="Password for password-protected link (if any)")
    parser.add_argument("--output", "-o", required=True, help="Target download directory")
    parser.add_argument("--concurrency", "-c", type=int, default=4, help="Max concurrent download threads (default: 4)")
    parser.add_argument("--channel", default="msedge", help="Browser channel for Playwright ('msedge', 'chrome', etc.)")
    return parser.parse_args()


def extract_drive_and_items(share_url: str, password: str = "", channel: str = "msedge") -> Tuple[List[Dict[str, Any]], Dict[str, str], str]:
    """
    Launch headless browser to authenticate, intercept RenderListDataAsStream, and extract cookies.
    """
    print(f"[*] Navigating to {share_url} using Playwright (channel={channel})...")
    render_data = []
    
    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(channel=channel, headless=True)
        except Exception as e:
            print(f"[!] Failed to launch browser channel '{channel}': {e}. Falling back to default chromium.")
            browser = p.chromium.launch(headless=True)
            
        context = browser.new_context()
        page = context.new_page()

        def on_response(res):
            if "RenderListDataAsStream" in res.url:
                try:
                    render_data.append(res.json())
                except Exception:
                    pass

        page.on("response", on_response)
        page.goto(share_url, wait_until="networkidle", timeout=60000)

        # Check if password protection input is present
        pwd_input = page.locator("#txtPassword")
        if pwd_input.count() > 0 and password:
            print("[*] Submitting access password...")
            pwd_input.first.fill(password)
            page.click("#btnSubmitPassword")
            page.wait_for_load_state("networkidle", timeout=60000)
            time.sleep(5)
        elif pwd_input.count() > 0 and not password:
            browser.close()
            raise ValueError("This SharePoint link is protected by a password, but no --password was supplied.")

        # Allow SPA components to settle
        time.sleep(3)
        cookies = context.cookies()
        current_url = page.url
        browser.close()

    parsed_url = urllib.parse.urlparse(current_url)
    origin = f"{parsed_url.scheme}://{parsed_url.netloc}"

    return render_data, cookies, origin


def collect_file_tree(render_data: List[Dict[str, Any]], cookies: List[Dict[str, Any]], origin: str) -> List[Dict[str, Any]]:
    """
    Recursively fetch all file nodes using the SharePoint v2.0 DriveItem REST API.
    """
    session = requests.Session()
    for c in cookies:
        session.cookies.set(c['name'], c['value'], domain=c.get('domain'), path=c.get('path', '/'))

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json;odata=verbose"
    }

    # Extract target root Drive items from captured RenderListDataAsStream
    root_items = []
    drive_id = ""
    for data in render_data:
        rows = data.get("ListData", {}).get("Row", [])
        for r in rows:
            sp_url = r.get(".spItemUrl", "")
            if "/drives/" in sp_url:
                # e.g. https://.../_api/v2.0/drives/b!.../items/...
                parts = sp_url.split("/drives/")[1].split("/items/")
                drive_id = parts[0]
                root_items.append({
                    "id": parts[1].split("?")[0],
                    "name": r.get("FileLeafRef"),
                    "is_folder": str(r.get("FSObjType")) == "1",
                    "size": int(r.get("File_x0020_Size") or 0)
                })

    if not drive_id:
        raise RuntimeError("Failed to resolve drive ID from SharePoint page metadata.")

    print(f"[*] Resolved SharePoint Drive ID: {drive_id}")

    all_files = []

    def traverse(item_id: str, current_path: str = ""):
        api_url = f"{origin}/_api/v2.0/drives/{drive_id}/items/{item_id}/children"
        resp = session.get(api_url, headers=headers)
        if resp.status_code != 200:
            print(f"[!] Warning: failed to fetch children for {current_path}: {resp.status_code}")
            return
        
        children = resp.json().get("value", [])
        for child in children:
            name = child.get("name")
            rel_path = f"{current_path}/{name}" if current_path else name
            if "folder" in child:
                traverse(child.get("id"), rel_path)
            else:
                dl_url = child.get("@content.downloadUrl") or child.get("@microsoft.graph.downloadUrl")
                size = child.get("size", 0)
                all_files.append({
                    "rel_path": rel_path,
                    "name": name,
                    "size": size,
                    "download_url": dl_url
                })

    # Traverse each root item
    for item in root_items:
        if item["is_folder"]:
            traverse(item["id"], item["name"])
        else:
            # Single root file
            single_api = f"{origin}/_api/v2.0/drives/{drive_id}/items/{item['id']}"
            resp = session.get(single_api, headers=headers)
            if resp.status_code == 200:
                meta = resp.json()
                dl_url = meta.get("@content.downloadUrl") or meta.get("@microsoft.graph.downloadUrl")
                all_files.append({
                    "rel_path": item["name"],
                    "name": item["name"],
                    "size": item["size"],
                    "download_url": dl_url
                })

    # Deduplicate items by rel_path
    dedup = {}
    for f in all_files:
        dedup[f["rel_path"]] = f
    return list(dedup.values())


def download_all(files: List[Dict[str, Any]], target_dir: str, concurrency: int = 4):
    os.makedirs(target_dir, exist_ok=True)
    total_bytes = sum(f["size"] for f in files)
    print(f"\n[*] Target directory: {target_dir}")
    print(f"[*] Total files: {len(files)} | Total payload: {total_bytes / (1024*1024):.2f} MB")

    def _download_single(item):
        rel_path = item["rel_path"].replace("/", os.sep)
        dest_path = os.path.join(target_dir, rel_path)
        expected_size = item["size"]
        dl_url = item["download_url"]

        os.makedirs(os.path.dirname(dest_path), exist_ok=True)

        if os.path.exists(dest_path) and os.path.getsize(dest_path) == expected_size and expected_size > 0:
            print(f"[SKIP] {rel_path} already exists ({expected_size:,} bytes)")
            return True, rel_path, expected_size

        for attempt in range(1, 4):
            temp_path = dest_path + ".tmp"
            try:
                t0 = time.time()
                with requests.get(dl_url, stream=True, timeout=60) as r:
                    r.raise_for_status()
                    with open(temp_path, "wb") as fp:
                        for chunk in r.iter_content(chunk_size=1024 * 1024):
                            if chunk:
                                fp.write(chunk)
                
                actual_size = os.path.getsize(temp_path)
                if expected_size > 0 and actual_size != expected_size:
                    raise ValueError(f"Size mismatch: expected {expected_size}, got {actual_size}")
                
                if os.path.exists(dest_path):
                    os.remove(dest_path)
                os.rename(temp_path, dest_path)
                
                elapsed = time.time() - t0
                speed = (actual_size / (1024 * 1024)) / elapsed if elapsed > 0 else 0
                print(f"[DONE] {rel_path} ({actual_size:,} bytes) in {elapsed:.1f}s ({speed:.2f} MB/s)")
                return True, rel_path, actual_size
            except Exception as e:
                print(f"[WARN] Failed attempt {attempt}/3 for {rel_path}: {e}")
                if os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass
                time.sleep(2)

        return False, rel_path, 0

    succeeded, failed, downloaded_bytes = 0, 0, 0
    t_start = time.time()
    with ThreadPoolExecutor(max_workers=concurrency) as executor:
        futures = {executor.submit(_download_single, f): f for f in files}
        for future in as_completed(futures):
            ok, pth, sz = future.result()
            if ok:
                succeeded += 1
                downloaded_bytes += sz
            else:
                failed += 1

    total_time = time.time() - t_start
    print("\n" + "=" * 50)
    print(f"Download Summary:")
    print(f"Total: {len(files)} | Succeeded: {succeeded} | Failed: {failed}")
    print(f"Downloaded: {downloaded_bytes / (1024*1024):.2f} MB | Time: {total_time:.1f}s")
    print("=" * 50)

    if failed > 0:
        sys.exit(1)


def main():
    args = parse_args()
    render_data, cookies, origin = extract_drive_and_items(args.url, args.password, args.channel)
    files = collect_file_tree(render_data, cookies, origin)
    download_all(files, args.output, args.concurrency)


if __name__ == "__main__":
    from typing import Tuple
    main()
