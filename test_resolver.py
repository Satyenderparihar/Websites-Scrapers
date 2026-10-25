import urllib.request
import urllib.parse
import re

def get_exact_website_url(query):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }
    encoded = urllib.parse.quote_plus(query.strip() + " website")
    url = f"https://html.duckduckgo.com/html/?q={encoded}"
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=5) as resp:
            html = resp.read().decode('utf-8', errors='ignore')
            # Look for result uddg redirect URLs
            matches = re.findall(r'uddg=([^&"]+)', html)
            for m in matches:
                target = urllib.parse.unquote(m)
                if not any(x in target.lower() for x in ['youtube.com', 'facebook.com', 'twitter.com', 'instagram.com', 'tiktok.com', 'reddit.com', 'quora.com', 'pinterest.com']):
                    return target
    except Exception as e:
        pass
    return None

if __name__ == "__main__":
    queries = [
        "Split your favorite song",
        "Design your home in 3D",
        "Find almost any PDF tool",
        "Turn blurry image into vector",
        "Paper animation",
        "Fix almost anything",
        "Learn English while practicing your typing"
    ]
    for q in queries:
        print(f"Query: {q} => Exact URL: {get_exact_website_url(q)}")
