#!/usr/bin/env python3
"""
批量获取红色遗址图片（支持 Wikipedia API 和百度百科 scraping）。
用法:
  python3 download_images.py                         # Wikipedia API (推荐, 稳定)
  python3 download_images.py --source baidu          # 百度百科 scraping
  python3 download_images.py --dry-run               # 只打印不操作
  python3 download_images.py --limit 10              # 只处理前10个
  python3 download_images.py --source baidu --download  # 百度百科+下载到本地
"""
import json
import os
import re
import sys
import time
import requests
from bs4 import BeautifulSoup

BASE = '/Users/mr.beef/工作台/数字档案馆小组作业'
DATA_JS = os.path.join(BASE, '网页部分/red-map/red-sites-data.js')
SITE_IMAGES_DIR = os.path.join(BASE, '网页部分/red-map/site-images')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept-Language': 'zh-CN,zh;q=0.9',
}


# ─── Wikipedia API ─────────────────────────────────────────────

def get_wikipedia_image_url(site_name):
    """通过 Wikipedia API 获取页面主图"""
    params = {
        'action': 'query',
        'format': 'json',
        'titles': site_name,
        'prop': 'pageimages',
        'pithumbsize': 800,
        'redirects': 1
    }
    try:
        resp = requests.get('https://zh.wikipedia.org/w/api.php',
                            params=params, headers=HEADERS, timeout=15)
    except requests.RequestException as e:
        return None, f'请求失败: {e}'

    if resp.status_code == 429:
        return None, 'HTTP 429 (限流)'
    if resp.status_code != 200:
        return None, f'HTTP {resp.status_code}'

    data = resp.json()
    pages = data.get('query', {}).get('pages', {})
    for page_id, page in pages.items():
        if page_id == '-1':
            # 尝试去除括号再搜
            simple = re.sub(r'[（(][^)）]*[)）]', '', site_name).strip()
            if simple != site_name and len(simple) > 2:
                return get_wikipedia_image_url(simple)
            return None, 'Wiki页面不存在'
        thumb = page.get('thumbnail')
        if thumb:
            return thumb['source'], None
        return None, 'Wiki页面无图片'

    return None, '未知错误'


# ─── Baidu Baike scraping ──────────────────────────────────────

def get_baike_image_url(site_name):
    """从百度百科 HTML 提取主图 URL"""
    url = f'https://baike.baidu.com/item/{site_name}'
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
    except requests.RequestException as e:
        return None, f'请求失败: {e}'

    if resp.status_code == 403:
        return None, 'HTTP 403 (反爬)'
    if resp.status_code == 404:
        simple = re.sub(r'[（(][^)）]*[)）]', '', site_name).strip()
        if simple != site_name and len(simple) > 2:
            return get_baike_image_url(simple)
        return None, '百科页面不存在(404)'
    if resp.status_code != 200:
        return None, f'HTTP {resp.status_code}'

    soup = BeautifulSoup(resp.text, 'html.parser')
    meta = soup.find('meta', property='og:image')
    if meta and meta.get('content'):
        return meta['content'], None
    pic = soup.select_one('.summary-pic img')
    if pic and pic.get('src'):
        return pic['src'], None
    return None, '未找到图片'


# ─── 通用逻辑 ──────────────────────────────────────────────────

def load_data():
    with open(DATA_JS, 'r', encoding='utf-8') as f:
        content = f.read()
    prefix = 'window.RED_SITES_DATA = '
    json_str = content[len(prefix):]
    if json_str.rstrip().endswith(';'):
        json_str = json_str.rstrip()[:-1]
    return json.loads(json_str)


def save_data(data):
    js_content = f"window.RED_SITES_DATA = {json.dumps(data, ensure_ascii=False, indent=2)};"
    with open(DATA_JS, 'w', encoding='utf-8') as f:
        f.write(js_content)


def has_local_image(site_id):
    num_match = re.match(r'^R-(\d+)$', site_id)
    if num_match:
        if 1 <= int(num_match.group(1)) <= 96:
            return True
    return os.path.exists(os.path.join(SITE_IMAGES_DIR, f'site-{site_id}.jpg'))


def download_image(url, save_path):
    try:
        resp = requests.get(url, headers=HEADERS, timeout=30)
        if resp.status_code == 200 and len(resp.content) > 1000:
            with open(save_path, 'wb') as f:
                f.write(resp.content)
            return True
    except requests.RequestException:
        pass
    return False


def main():
    dry_run = '--dry-run' in sys.argv
    do_download = '--download' in sys.argv
    source = 'wikipedia' if '--source' not in ' '.join(sys.argv) else (
        'baidu' if '--source baidu' in ' '.join(sys.argv) else 'wikipedia')
    limit = None
    for i, arg in enumerate(sys.argv):
        if arg == '--limit' and i + 1 < len(sys.argv):
            limit = int(sys.argv[i + 1])

    fetch_fn = get_wikipedia_image_url if source == 'wikipedia' else get_baike_image_url
    print(f'数据源: {source}')
    delay = 2 if source == 'wikipedia' else 8

    data = load_data()
    features = data['features']
    print(f'共 {len(features)} 个遗址')

    to_process = []
    for f in features:
        sid = f['properties']['基础信息']['遗址ID']
        name = f['properties']['基础信息']['遗址名称']
        if has_local_image(sid):
            continue
        to_process.append((f, sid, name))

    if limit:
        to_process = to_process[:limit]

    action = '下载' if do_download else '提取URL'
    print(f'需处理: {len(to_process)} 个 ({action})')
    if dry_run:
        print('--- DRY RUN ---')

    success = 0
    fail = 0
    os.makedirs(SITE_IMAGES_DIR, exist_ok=True)

    for i, (feature, sid, name) in enumerate(to_process):
        print(f'[{i + 1}/{len(to_process)}] {sid} {name} ...', end=' ', flush=True)
        img_url, error = fetch_fn(name)

        if error:
            print(f'失败: {error}')
            fail += 1
            time.sleep(delay)
            continue

        if dry_run:
            print(f'找到: {img_url[:90]}...')
            success += 1
            time.sleep(delay)
            continue

        if do_download:
            ext = '.jpg'
            if '.png' in img_url.lower():
                ext = '.png'
            sp = os.path.join(SITE_IMAGES_DIR, f'site-{sid}{ext}')
            if download_image(img_url, sp):
                print(f'OK ({os.path.getsize(sp) // 1024}KB)')
                success += 1
            else:
                print('下载失败')
                fail += 1
        else:
            # 存储为直链URL
            feature['properties']['史实内容']['影像资料与链接'] = [img_url]
            print('URL已存入')
            success += 1

        time.sleep(delay)

    if not dry_run and not do_download and success > 0:
        save_data(data)
        print(f'\n已保存 {success} 个URL到 red-sites-data.js')

    print(f'\n--- 完成: 成功 {success}, 失败 {fail} ---')


if __name__ == '__main__':
    main()
