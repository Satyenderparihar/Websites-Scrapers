import urllib.request
import urllib.parse
import re

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

def resolve_url(query):
    # Try Startpage / Brave / Qwant / DDG API / Wikipedia / Bing
    url = 'https://www.bing.com/search?q=' + urllib.parse.quote_plus(query)
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=5) as r:
            html = r.read().decode('utf-8', errors='ignore')
            urls = re.findall(r'href="(https?://[^"]+)"', html)
            banned = ['bing.com', 'microsoft.com', 'live.com', 'w3.org', 'schema.org', 'cloudflare.com', 'r.bat.bing.com', 'c.bing.com', 'youtube.com', 'tiktok.com', 'instagram.com', 'facebook.com', 'twitter.com', 'x.com']
            clean = []
            for u in urls:
                if not any(b in u.lower() for b in banned):
                    clean.append(u)
            return clean[:3]
    except Exception as e:
        return [str(e)]

if __name__ == '__main__':
    tests = [
        'TinyWow PDF website',
        'Vocal Remover website split audio',
        'Vectorizer ai website convert vector',
        'Planner 5D 3D home design website',
        'Falstad circuit simulator website'
    ]
    for t in tests:
        print(f"Query: {t}")
        print(f"Results: {resolve_url(t)}")
        print('-' * 40)
