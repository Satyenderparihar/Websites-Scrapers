import os
import sys
import re
import json
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Set stdout encoding
sys.stdout.reconfigure(encoding='utf-8')

PROJECT_DIR = r"D:\Antigravity - Agent\🌐 YouTube Websites Extractor Agent"
RAW_DATA_PATH = os.path.join(PROJECT_DIR, "💾 raw_channel_videos.txt")

import subprocess

def get_channel_info(channel_target=None):
    if not channel_target or channel_target.strip() in ['@setupsaitony', 'https://www.youtube.com/@setupsaitony', 'http://www.youtube.com/@setupsaitony', '']:
        return "SetupsAI", "@setupsaitony", RAW_DATA_PATH

    clean_target = channel_target.strip()
    if clean_target.startswith('@'):
        url = f"https://www.youtube.com/{clean_target}"
    elif not clean_target.startswith('http'):
        url = f"https://www.youtube.com/@{clean_target}"
    else:
        url = clean_target

    print(f"[*] Fetching channel details for: {url}")
    try:
        cmd_name = ['yt-dlp', '--print', '%(channel)s', '--playlist-items', '1', url]
        res_name = subprocess.run(cmd_name, capture_output=True, text=True, encoding='utf-8', errors='ignore')
        lines = [l.strip() for l in res_name.stdout.splitlines() if l.strip() and not l.startswith('WARNING')]
        channel_name = lines[-1] if lines else "YouTube_Channel"
    except Exception:
        channel_name = "Custom_Channel"

    safe_name = re.sub(r'[^a-zA-Z0-9_\-]', '', channel_name)
    if not safe_name:
        safe_name = "YouTube_Channel"
    raw_path = os.path.join(PROJECT_DIR, f"💾 raw_{safe_name}_videos.txt")

    if not os.path.exists(raw_path):
        print(f"[*] Downloading playlist metadata from {url}...")
        cmd_dump = ['yt-dlp', '--flat-playlist', '--print', '%(title)s | %(id)s | %(webpage_url)s', url]
        res_dump = subprocess.run(cmd_dump, capture_output=True, text=True, encoding='utf-8', errors='ignore')
        with open(raw_path, 'w', encoding='utf-8') as f:
            f.write(res_dump.stdout)

    return safe_name, clean_target, raw_path

# Domain knowledge mapping for creators and viral website series
KNOWN_MENTIONS = {
    'replit': ('Replit', 'https://replit.com', '💻 Coding & Web Development', 'Cloud software development platform to build and launch websites/apps with AI.'),
    'meetgamma': ('Gamma App', 'https://gamma.app', '🤖 AI & Machine Learning Tools', 'AI-powered presentation, document, and webpage generator in seconds.'),
    'gamma': ('Gamma App', 'https://gamma.app', '🤖 AI & Machine Learning Tools', 'AI-powered presentation, document, and webpage generator in seconds.'),
    'kittldesign': ('Kittl', 'https://kittl.com', '🎨 Design, 3D & Graphics', 'Advanced online graphic design and logo creation tool with AI vectors and typography.'),
    'freepik': ('Freepik Pikaso', 'https://freepik.com/pikaso', '🎨 Design, 3D & Graphics', 'AI real-time sketch-to-image and photo editing generator.'),
    'mindgraspai': ('Mindgrasp AI', 'https://mindgrasp.ai', '📚 Learning, Study & Research', 'AI study assistant that reads documents, lectures, and generates notes/quizzes.'),
    'mindgrasp': ('Mindgrasp AI', 'https://mindgrasp.ai', '📚 Learning, Study & Research', 'AI study assistant that reads documents, lectures, and generates notes/quizzes.'),
    'podcastle': ('Podcastle', 'https://podcastle.ai', '🎵 Audio, Voice & Music', 'Studio-quality audio recording, background noise removal, and AI voice editing.'),
    'unriddle': ('Unriddle AI', 'https://unriddle.ai', '📚 Learning, Study & Research', 'AI research companion to read, understand, and chat with PDFs and papers quickly.'),
    'framer': ('Framer', 'https://framer.com', '💻 Coding & Web Development', 'No-code website builder that turns Figma designs or AI prompts into live web pages.'),
    'invideo': ('InVideo AI', 'https://invideo.io', '🤖 AI & Machine Learning Tools', 'Text-to-video AI generator that creates complete edited video shorts with voiceover.'),
    'invideoofficial': ('InVideo AI', 'https://invideo.io', '🤖 AI & Machine Learning Tools', 'Text-to-video AI generator that creates complete edited video shorts with voiceover.'),
    'brickit': ('Brickit App', 'https://brickit.app', '🛠️ Utilities & Web Tools', 'AI app that scans LEGO piles on the floor and gives step-by-step build ideas.'),
    'printify': ('Printify', 'https://printify.com', '⚡ Productivity & Document Tools', 'Print-on-demand platform for custom merchandise and e-commerce.'),
    'socialsonic': ('Socialsonic', 'https://socialsonic.com', '⚡ Productivity & Document Tools', 'AI social media assistant and viral content creation suite.'),
    'wegic': ('Wegic AI', 'https://wegic.ai', '💻 Coding & Web Development', 'AI conversational website designer that builds responsive websites via chat.'),
    'writesonic': ('Writesonic', 'https://writesonic.com', '🤖 AI & Machine Learning Tools', 'AI SEO content writer, copywriting, and chatbot creation platform.'),
    'filmorawondershare': ('Wondershare Filmora', 'https://filmora.wondershare.com', '🎨 Design, 3D & Graphics', 'Full-featured desktop & web video editing software with AI tools.'),
}

# Knowledge base matching for topics in video titles
TOPIC_RULES = [
    # Audio & Music
    (r'(split your favorite song|separate the song|separate out the vocals|vocals.*instruments)', 
     'Vocal Remover', 'https://vocalremover.org', '🎵 Audio, Voice & Music', 'Extract vocals and instrumental tracks from any song for free using AI.'),
    (r'(radio\s*garden|#radiogarden)', 
     'Radio Garden', 'https://radio.garden', '🎵 Audio, Voice & Music', 'Explore live radio stations worldwide by rotating a 3D digital globe.'),
    (r'(play on guitar and piano|guitar and piano)', 
     'Chordify', 'https://chordify.net', '🎵 Audio, Voice & Music', 'Extract chords and learn to play any song on guitar, piano, or ukulele.'),
    (r'(podcastle|sound professional|background noise)', 
     'Podcastle AI', 'https://podcastle.ai', '🎵 Audio, Voice & Music', 'Clean audio noise, studio recording, and AI audio enhancer.'),

    # Design & 3D
    (r'(blurry image into vector|blurry.*vector)', 
     'Vectorizer.AI', 'https://vectorizer.ai', '🎨 Design, 3D & Graphics', 'Convert raster blurry JPEG/PNG images into crisp, scalable SVG vectors using deep learning.'),
    (r'(design your home in 3d|create your dream room)', 
     'Planner 5D', 'https://planner5d.com', '🎨 Design, 3D & Graphics', 'Design interior floors, furniture, and 3D home layouts directly in your browser.'),
    (r'(paper animation)', 
     'Papercurves', 'https://papercurves.com', '🎨 Design, 3D & Graphics', 'Generate interactive 3D paper fold and flip animations for creative web designs.'),
    (r'(turn your logo into an animated 3d|3d logo)', 
     'Renderforest 3D', 'https://renderforest.com', '🎨 Design, 3D & Graphics', 'Transform 2D logos into animated 3D videos and intro sequences.'),
    (r'(turn any shape into a complete logo|design logo in seconds)', 
     'Looka / Kittl', 'https://looka.com', '🎨 Design, 3D & Graphics', 'AI logo maker and brand identity generator with custom vector assets.'),
    (r'(svg patterns|find almost any svg)', 
     'Hero Patterns', 'https://heropatterns.com', '🎨 Design, 3D & Graphics', 'Free library of repeatable, customizable SVG background patterns.'),
    (r'(holographic and cloth|visualdesign)', 
     'Silk / Interactive Generative Art', 'https://weavesilk.com', '🎨 Design, 3D & Graphics', 'Interactive generative art canvas with glowing symmetry and cloth effects.'),
    (r'(find almost any design tools|designtools)', 
     'Toools.design', 'https://toools.design', '🎨 Design, 3D & Graphics', 'Curated directory of 1,000+ modern design resources and creative utilities.'),
    (r'(ready to download peeps|peeps)', 
     'Open Peeps', 'https://openpeeps.com', '🎨 Design, 3D & Graphics', 'Hand-drawn illustration library with customizable poses and outfits.'),
    (r'(free ai photo editor|photo editor)', 
     'Clipdrop / Photopea', 'https://clipdrop.co', '🎨 Design, 3D & Graphics', 'Browser-based AI photo editing, object removal, and Photoshop alternative.'),
    (r'(custom animated switches|animated components|how to use loaders)', 
     'Uiverse.io', 'https://uiverse.io', '🎨 Design, 3D & Graphics', 'Community-created open-source CSS & Tailwind buttons, loaders, and UI elements.'),

    # Productivity & PDF / Office
    (r'(find almost any pdf tool|pdf tools|pdftools)', 
     'TinyWow', 'https://tinywow.com', '⚡ Productivity & Document Tools', 'Free 200+ online tools for PDF editing, image conversion, video editing, and AI text.'),
    (r'(create almost any kind of worksheet|worksheet)', 
     'Worksheet Genius', 'https://worksheetgenius.com', '⚡ Productivity & Document Tools', 'Free generator for educational worksheets, word searches, puzzles, and quizzes.'),
    (r'(fix almost anything|repairing)', 
     'iFixit', 'https://ifixit.com', '🛠️ Utilities & Web Tools', 'Step-by-step repair manuals, teardowns, and repair guides for every gadget and appliance.'),
    (r'(analyze your resume|generate professional resume)', 
     'Reactive Resume / Jobscan', 'https://rxresu.me', '⚡ Productivity & Document Tools', 'Free, privacy-focused open-source resume builder with ATS optimization.'),
    (r'(gpt excel|excel tips)', 
     'GPTExcel', 'https://gptexcel.io', '⚡ Productivity & Document Tools', 'AI generator for Microsoft Excel & Google Sheets formulas, scripts, and SQL queries.'),
    (r'(write essays in seconds|aithor)', 
     'Aithor', 'https://aithor.com', '⚡ Productivity & Document Tools', 'AI essay and academic paper assistant that generates structured, cited drafts.'),
    (r'(organize your study|studytips)', 
     'Notion / StudySmarter', 'https://notion.so', '⚡ Productivity & Document Tools', 'Smart workspace to organize notes, flashcards, study schedules, and project tasks.'),

    # Coding & Web Development
    (r'(build and lunch a full website|build a landing page|websitebuilder)', 
     'Framer / Replit', 'https://framer.com', '💻 Coding & Web Development', 'Fast AI-powered landing page and website generator that launches instantly.'),
    (r'(build almost anything.*replit|@replit)', 
     'Replit', 'https://replit.com', '💻 Coding & Web Development', 'Cloud software creation platform with AI coding assistance and 1-click hosting.'),
    (r'(runableai|build almost anything.*runable)', 
     'Runable AI', 'https://runable.ai', '💻 Coding & Web Development', 'Prompt-driven AI web app builder that generates full-stack web applications.'),
    (r'(electronic circuits|build and test real electronic)', 
     'Falstad Circuit Simulator', 'https://falstad.com/circuit', '💻 Coding & Web Development', 'Browser-based graphical circuit simulator demonstrating electronic components in real-time.'),

    # Maps, Geography & Earth
    (r'(minecraft world on a 3d map|explore minecraft)', 
     'BlueMap / Chunkbase', 'https://bluemap.bluecolored.de', '🗺️ Maps, 3D & Exploration', 'Renders real-time 3D interactive web maps of Minecraft worlds.'),
    (r'(generates an entire city.*street map|map poster of almost any city)', 
     'City-roads / Anvaka', 'https://anvaka.github.io/city-roads', '🗺️ Maps, 3D & Exploration', 'Visualizes every single road in any city as a minimalist high-res map.'),
    (r'(explore the world in real time|see what.*happening around the world)', 
     'EarthCam / Liveuamap', 'https://earthcam.com', '🗺️ Maps, 3D & Exploration', 'Live webcams and interactive situational maps of landmarks across the globe.'),
    (r'(instantly turn the map into a 3d model)', 
     'TouchTerrain', 'https://touchterrain.geol.iastate.edu', '🗺️ Maps, 3D & Exploration', 'Converts real-world elevation maps and terrain into 3D printable STL models.'),
    (r'(match the silhouette to its real world scale|real world scale)', 
     'The True Size Of', 'https://thetruesize.com', '🗺️ Maps, 3D & Exploration', 'Drag countries across the Mercator map to compare their real geographical sizes.'),
    (r'(ancient earth|earthcup)', 
     'Ancient Earth Globe', 'https://dinosaurpictures.org/ancient-earth', '🗺️ Maps, 3D & Exploration', 'Interactive 3D globe visualizing Earth from 750 million years ago to today.'),
    (r'(flights and satellites|flight)', 
     'Flightradar24 / Satellites', 'https://flightradar24.com', '🗺️ Maps, 3D & Exploration', 'Real-time global flight tracking and satellite orbital view in the sky.'),
    (r'(make roadtrip fun|plan your roadtrip)', 
     'Roadtrippers', 'https://roadtrippers.com', '🗺️ Maps, 3D & Exploration', 'Ultimate road trip itinerary planner discovering offbeat roadside attractions.'),

    # Science, Education & Medical
    (r'(chemistry experiments|chemical)', 
     'PhET Interactive Simulations', 'https://phet.colorado.edu', '📚 Learning, Study & Research', 'Interactive science & chemistry laboratory simulations by the University of Colorado.'),
    (r'(camera settings work together|camera.*lens)', 
     'CameraSim', 'https://camerasim.com', '📚 Learning, Study & Research', 'Interactive SLR camera simulator to learn aperture, shutter speed, and ISO.'),
    (r'(learn languages through games)', 
     'Ba Ba Dum', 'https://babadum.com', '📚 Learning, Study & Research', 'Playful browser game teaching 1,500 words across 21 languages with illustrations.'),
    (r'(medical knowledge|anatomy)', 
     'BioDigital Human / Kenhub', 'https://biodigital.com', '📚 Learning, Study & Research', 'Comprehensive interactive 3D anatomy and human body disease simulation.'),
    (r'(how things work with cool animation|simulate all kinds of machine)', 
     'Animated Engines', 'https://animatedengines.com', '📚 Learning, Study & Research', 'Animated visual breakdowns of how steam, internal combustion, and jet engines operate.'),
    (r'(practicing portraits|sketch)', 
     'Quickposes', 'https://quickposes.com', '📚 Learning, Study & Research', 'Figure drawing and gesture pose timing tool for artists and sketchers.'),
    (r'(brochures for almost any car)', 
     'Auto-Brochures', 'https://auto-brochures.com', '📚 Learning, Study & Research', 'Free PDF archive of original vehicle brochures spanning vintage to modern cars.'),
    (r'(simulate repairing your car)', 
     'How a Car Works', 'https://howacarworks.com', '📚 Learning, Study & Research', 'Visual video courses and interactive technical diagrams of car mechanics.'),
    (r'(customize almost any car in 3d)', 
     '3DTuning', 'https://3dtuning.com', '🎨 Design, 3D & Graphics', 'Photorealistic 3D car configurator with thousands of models and custom parts.'),
    (r'(typing|practicing your typing)', 
     'TypeRacer / Keybr', 'https://play.typeracer.com', '⚡ Productivity & Document Tools', 'Multiplayer typing speed test and keyboard touch-typing practice web app.'),

    # Gaming & Entertainment
    (r'(stack your ice cream)', 
     'Ice Cream Stack / Neave', 'https://neave.com', '🎮 Gaming & Fun Web Experiments', 'Charming retro casual browser game to balance scoops of ice cream.'),
    (r'(draw your racetrack)', 
     'Line Rider', 'https://linerider.com', '🎮 Gaming & Fun Web Experiments', 'Classic physics toy game where you draw sled tracks for a rider on a hill.'),
    (r'(break apart website design elements)', 
     'Kick Ass Web Destroyer', 'https://kickassapp.com', '🎮 Gaming & Fun Web Experiments', 'Asteroids-style bookmarklet and web game that lets you blast HTML elements off any page.'),
    (r'(create your own fantasy world|dnd.*map)', 
     'Azgaar\'s Fantasy Map Generator', 'https://azgaar.github.io/Fantasy-Map-Generator/', '🎮 Gaming & Fun Web Experiments', 'Generates procedural medieval & fantasy world maps with rivers, biomes, and kingdoms.'),
    (r'(driving simulate game|drivingsimulator)', 
     'Slow Roads', 'https://slowroads.io', '🎮 Gaming & Fun Web Experiments', 'Procedural, infinite zen driving simulator running smoothly in the web browser.'),
    (r'(before i play)', 
     'Before I Play', 'https://beforeiplay.com', '🎮 Gaming & Fun Web Experiments', 'Crowdsourced spoiler-free tips and advice before you start any video game.'),
    (r'(super smash brothers|unblockedgames)', 
     'Smash Bros Web Edition', 'https://smashboards.com', '🎮 Gaming & Fun Web Experiments', 'Play classic Super Smash Bros retro games directly in modern web browsers.'),

    # AI Tools & Utilities
    (r'(chat with pdf|use chatgpt for free.*coze|#coze)', 
     'Coze AI', 'https://coze.com', '🤖 AI & Machine Learning Tools', 'Next-gen AI bot building platform powered by top LLM models like GPT-4.'),
    (r'(interview candidates with ai|#interview)', 
     'Interview Warmup by Google', 'https://grow.google/certificates/interview-warmup/', '🤖 AI & Machine Learning Tools', 'Practice answering job interview questions with automated AI transcription and insights.'),
    (r'(create multiple product images|aiart)', 
     'Flair.ai / Pebblely', 'https://flair.ai', '🤖 AI & Machine Learning Tools', 'AI photo shoot studio for e-commerce products with stunning background generation.'),
    (r'(instant id)', 
     'InstantID', 'https://github.com/InstantID/InstantID', '🤖 AI & Machine Learning Tools', 'Zero-shot identity-preserving AI portrait generation technique.'),
    (r'(100 ai tools|120 mind blowing ai tools|find almost any tools)', 
     'Toolify.ai / Futurepedia', 'https://toolify.ai', '🤖 AI & Machine Learning Tools', 'The world\'s largest directory of AI websites, software, and daily tools.'),
]

def clean_title(title):
    # Remove hashtags and mentions for cleaner display
    t = re.sub(r'#[a-zA-Z0-9_\-]+', '', title)
    t = re.sub(r'@[a-zA-Z0-9_\-\.]+', '', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t

def extract_website_info(title, url, video_id, channel_name="YouTube"):
    title_lower = title.lower()

    # 1. Check Creator Mentions
    for m in re.findall(r'@([a-zA-Z0-9_\-\.]+)', title_lower):
        if m in KNOWN_MENTIONS:
            name, site_url, cat, desc = KNOWN_MENTIONS[m]
            return {
                "name": name,
                "url": site_url,
                "category": cat,
                "description": desc,
                "video_title": clean_title(title),
                "video_url": url,
                "video_id": video_id
            }

    # 2. Check Topic Rules
    for pattern, name, site_url, cat, desc in TOPIC_RULES:
        if re.search(pattern, title_lower):
            return {
                "name": name,
                "url": site_url,
                "category": cat,
                "description": desc,
                "video_title": clean_title(title),
                "video_url": url,
                "video_id": video_id
            }

    # 3. Fallback: Parse hashtags & clean title
    cleaned = clean_title(title)
    if not cleaned:
        cleaned = "Web Tool"
    
    # Categorize based on hashtags
    category = "🛠️ Utilities & Web Tools"
    if any(k in title_lower for k in ['#ai', '#aitools', 'artificial intelligence', 'gpt', 'llm']):
        category = "🤖 AI & Machine Learning Tools"
    elif any(k in title_lower for k in ['#design', '#3d', '#graphic', '#photo', '#logo', '#art']):
        category = "🎨 Design, 3D & Graphics"
    elif any(k in title_lower for k in ['#music', '#song', '#audio', '#radio']):
        category = "🎵 Audio, Voice & Music"
    elif any(k in title_lower for k in ['#coding', '#dev', '#code', '#website']):
        category = "💻 Coding & Web Development"
    elif any(k in title_lower for k in ['#study', '#learn', '#science', '#education']):
        category = "📚 Learning, Study & Research"
    elif any(k in title_lower for k in ['#game', '#gaming', '#fun']):
        category = "🎮 Gaming & Fun Web Experiments"
    elif any(k in title_lower for k in ['#map', '#maps', '#travel']):
        category = "🗺️ Maps, 3D & Exploration"
    elif any(k in title_lower for k in ['#productivity', '#office', '#work']):
        category = "⚡ Productivity & Document Tools"

    return {
        "name": cleaned[:45],
        "url": f"https://www.google.com/search?q={re.sub(r'[^a-zA-Z0-9 ]', '', cleaned).strip().replace(' ', '+')}",
        "category": category,
        "description": f"Featured web tool in {channel_name} showcase: {cleaned}",
        "video_title": cleaned,
        "video_url": url,
        "video_id": video_id
    }

def main():
    channel_arg = sys.argv[1] if len(sys.argv) > 1 else None
    channel_name, channel_handle, raw_path = get_channel_info(channel_arg)

    print("=" * 60)
    print("🌐 YouTube Websites Extractor & Categorizer Agent")
    print(f"📺 Channel: {channel_name} ({channel_handle})")
    print(f"📂 Project Path: {PROJECT_DIR}")
    print("=" * 60)

    excel_output_path = os.path.join(PROJECT_DIR, f"📊 {channel_name}_Websites_Master_List.xlsx")
    markdown_output_path = os.path.join(PROJECT_DIR, f"📝 {channel_name}_Websites_Categorized_Report.md")
    json_output_path = os.path.join(PROJECT_DIR, f"💾 {channel_name}_websites_database.json")

    if not os.path.exists(raw_path):
        print(f"❌ Error: Raw data file not found at {raw_path}")
        return

    with open(raw_path, "r", encoding="utf-8", errors="ignore") as f:
        lines = [l.strip() for l in f if "|" in l]

    print(f"📊 Loaded {len(lines)} total video entries from channel.")

    extracted_items = []
    seen_websites = set()
    category_counts = defaultdict(int)

    for line in lines:
        parts = line.split("|")
        if len(parts) >= 3:
            raw_title = parts[0].strip()
            video_id = parts[1].strip()
            video_url = parts[2].strip()
        elif len(parts) == 2:
            raw_title = parts[0].strip()
            video_id = parts[1].strip()
            video_url = f"https://www.youtube.com/watch?v={video_id}"
        else:
            continue

        item = extract_website_info(raw_title, video_url, video_id, channel_name)
        
        # Deduplication key by website name & category
        dedup_key = (item["name"].lower(), item["category"])
        if dedup_key not in seen_websites:
            seen_websites.add(dedup_key)
            extracted_items.append(item)
            category_counts[item["category"]] += 1

    print(f"✅ Extracted {len(extracted_items)} unique categorized websites/tools exclusively for {channel_name}.")
    print("\n📈 Category Breakdown:")
    for cat, count in sorted(category_counts.items(), key=lambda x: x[1], reverse=True):
        print(f"  • {cat}: {count} tools")

    # Dedicated Channel Directory
    channel_dir = os.path.join(PROJECT_DIR, f"📁 {channel_name}")
    os.makedirs(channel_dir, exist_ok=True)

    excel_dir_path = os.path.join(channel_dir, f"📊 {channel_name}_Websites_Master_List.xlsx")
    excel_root_path = os.path.join(PROJECT_DIR, f"📊 {channel_name}_Websites_Master_List.xlsx")

    markdown_dir_path = os.path.join(channel_dir, f"📝 {channel_name}_Websites_Categorized_Report.md")
    markdown_root_path = os.path.join(PROJECT_DIR, f"📝 {channel_name}_Websites_Categorized_Report.md")

    json_dir_path = os.path.join(channel_dir, f"💾 {channel_name}_websites_database.json")
    json_root_path = os.path.join(PROJECT_DIR, f"💾 {channel_name}_websites_database.json")

    # 1. Export to Excel (.xlsx) with OpenPyXL
    print(f"\n📝 Generating Excel Master Sheet for {channel_name}...")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{channel_name[:30]}"

    # Styling definitions
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Deep Blue
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10)
    link_font = Font(name="Segoe UI", size=10, color="2563EB", underline="single")
    alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    border_side = Side(border_style="thin", color="E2E8F0")
    thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

    headers = [
        "Category",
        "Website / Tool Name",
        "Official Website URL",
        "Key Function / Description",
        "Source Video Title",
        "YouTube Video Link"
    ]
    ws.append(headers)

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[1].height = 28

    # Sort items by category then name
    sorted_items = sorted(extracted_items, key=lambda x: (x["category"], x["name"]))

    for row_idx, item in enumerate(sorted_items, 2):
        row_data = [
            item["category"],
            item["name"],
            item["url"],
            item["description"],
            item["video_title"],
            item["video_url"]
        ]
        ws.append(row_data)

        # Style row
        fill = alt_fill if row_idx % 2 == 0 else PatternFill(fill_type=None)
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_num)
            cell.font = data_font
            cell.border = thin_border
            if fill.fill_type:
                cell.fill = fill
            cell.alignment = Alignment(vertical="center")

            # Clickable Links
            if col_num in [3, 6]:
                cell.font = link_font
                cell.hyperlink = cell.value

        ws.row_dimensions[row_idx].height = 22

    # Auto-adjust column widths
    column_widths = [26, 26, 32, 50, 40, 36]
    for i, width in enumerate(column_widths, 1):
        col_letter = get_column_letter(i)
        ws.column_dimensions[col_letter].width = width

    wb.save(excel_dir_path)
    wb.save(excel_root_path)
    print(f"💾 Excel saved: {excel_root_path}")
    print(f"💾 Excel saved in channel folder: {excel_dir_path}")

    # 2. Export to JSON
    print(f"📝 Generating JSON Database for {channel_name}...")
    with open(json_dir_path, "w", encoding="utf-8") as f:
        json.dump(sorted_items, f, ensure_ascii=False, indent=2)
    with open(json_root_path, "w", encoding="utf-8") as f:
        json.dump(sorted_items, f, ensure_ascii=False, indent=2)
    print(f"💾 JSON saved: {json_root_path}")

    # 3. Export to Markdown Report
    print(f"📝 Generating Markdown Report for {channel_name}...")
    md_lines = [
        f"# 🌐 {channel_name} - Exclusive Websites & AI Tools Directory\n",
        f"> **Channel Name**: {channel_name}  ",
        f"> **Channel Handle / URL**: {channel_handle}  ",
        f"> **Total Unique Websites Identified**: {len(extracted_items)} tools  ",
        f"> **Total Videos / Shorts Scanned**: {len(lines)} videos (Exclusively from {channel_name})  ",
        f"> **Channel Isolation**: 100% Dedicated to {channel_name} (Zero content from other channels)  \n",
        "## 📊 Category Summary\n",
        "| Category | Websites Count |",
        "| :--- | :---: |"
    ]
    for cat, count in sorted(category_counts.items(), key=lambda x: x[1], reverse=True):
        md_lines.append(f"| {cat} | **{count}** |")

    md_lines.append("\n---\n")

    # Group by category
    by_category = defaultdict(list)
    for it in sorted_items:
        by_category[it["category"]].append(it)

    for cat in sorted(by_category.keys()):
        tools = by_category[cat]
        md_lines.append(f"### {cat} ({len(tools)} Tools)\n")
        md_lines.append(f"| Website / Tool | Official Link | Description ({channel_name}) | Video Reference |")
        md_lines.append("| :--- | :--- | :--- | :--- |")
        for t in tools:
            clean_desc = t["description"].replace("|", "-")
            md_lines.append(f"| **{t['name']}** | [{t['name']}]({t['url']}) | {clean_desc} | [Watch Short]({t['video_url']}) |")
        md_lines.append("\n")

    md_content = "\n".join(md_lines)
    with open(markdown_dir_path, "w", encoding="utf-8") as f:
        f.write(md_content)
    with open(markdown_root_path, "w", encoding="utf-8") as f:
        f.write(md_content)
    print(f"💾 Markdown report saved: {markdown_root_path}")
    print(f"💾 Markdown report saved in channel folder: {markdown_dir_path}")

    print("\n" + "=" * 60)
    print("🎉 All extraction & categorization tasks completed successfully!")
    print("=" * 60)

if __name__ == "__main__":
    main()
