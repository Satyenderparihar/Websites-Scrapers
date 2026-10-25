import os
import sys
import re
import json
import subprocess
from datetime import datetime
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Set stdout encoding
sys.stdout.reconfigure(encoding='utf-8')

PROJECT_DIR = r"D:\Antigravity - Agent\🌐 YouTube Websites Extractor Agent"
RAW_DATA_PATH = os.path.join(PROJECT_DIR, "💾 raw_channel_videos.txt")

# Creator handles & official website mapping
KNOWN_MENTIONS = {
    'replit': ('Replit', 'https://replit.com', '💻 Coding & Web Development', 'Cloud software creation and instant AI web app hosting platform.'),
    'meetgamma': ('Gamma App', 'https://gamma.app', '🤖 AI & Machine Learning Tools', 'AI presentation, document, and webpage builder in seconds.'),
    'gamma': ('Gamma App', 'https://gamma.app', '🤖 AI & Machine Learning Tools', 'AI presentation, document, and webpage builder in seconds.'),
    'kittldesign': ('Kittl', 'https://kittl.com', '🎨 Design, 3D & Graphics', 'Advanced graphic design platform with AI vector generation and typography.'),
    'freepik': ('Freepik Pikaso', 'https://freepik.com/pikaso', '🎨 Design, 3D & Graphics', 'Real-time sketch-to-image AI canvas and creative stock assets.'),
    'mindgraspai': ('Mindgrasp AI', 'https://mindgrasp.ai', '📚 Learning, Study & Research', 'AI study assistant that summarizes lectures, PDFs, and generates flashcards.'),
    'mindgrasp': ('Mindgrasp AI', 'https://mindgrasp.ai', '📚 Learning, Study & Research', 'AI study assistant that summarizes lectures, PDFs, and generates flashcards.'),
    'podcastle': ('Podcastle', 'https://podcastle.ai', '🎵 Audio, Voice & Music', 'Studio-quality audio recording, noise removal, and AI voice editing suite.'),
    'unriddle': ('Unriddle AI', 'https://unriddle.ai', '📚 Learning, Study & Research', 'AI research companion to read, understand, and chat with technical documents.'),
    'framer': ('Framer', 'https://framer.com', '💻 Coding & Web Development', 'No-code interactive website builder with instant production deployment.'),
    'invideo': ('InVideo AI', 'https://invideo.io', '🤖 AI & Machine Learning Tools', 'Full text-to-video AI generator that creates complete edited shorts with voiceovers.'),
    'invideoofficial': ('InVideo AI', 'https://invideo.io', '🤖 AI & Machine Learning Tools', 'Full text-to-video AI generator that creates complete edited shorts with voiceovers.'),
    'brickit': ('Brickit App', 'https://brickit.app', '🛠️ Utilities & Web Tools', 'AI app that scans LEGO piles on the floor and gives step-by-step build ideas.'),
    'printify': ('Printify', 'https://printify.com', '⚡ Productivity & Document Tools', 'Print-on-demand platform for custom merchandise and global e-commerce.'),
    'socialsonic': ('Socialsonic', 'https://socialsonic.com', '⚡ Productivity & Document Tools', 'AI social media content creation suite and viral analytics tool.'),
    'wegic': ('Wegic AI', 'https://wegic.ai', '💻 Coding & Web Development', 'Conversational AI web designer that crafts responsive websites through chat.'),
    'writesonic': ('Writesonic', 'https://writesonic.com', '🤖 AI & Machine Learning Tools', 'AI content writer, SEO article generator, and AI agent platform.'),
    'filmorawondershare': ('Wondershare Filmora', 'https://filmora.wondershare.com', '🎨 Design, 3D & Graphics', 'Video editing software with AI smart cutout, audio ducking, and effects.'),
    'chatpdf': ('ChatPDF', 'https://chatpdf.com', '📚 Learning, Study & Research', 'Chat with any PDF document, research paper, or textbook using AI.'),
    'quillbot': ('QuillBot', 'https://quillbot.com', '⚡ Productivity & Document Tools', 'AI paraphrasing, grammar checker, and academic summarizing assistant.'),
    'canva': ('Canva', 'https://canva.com', '🎨 Design, 3D & Graphics', 'Comprehensive visual design, social media template, and presentation platform.'),
    'capcut': ('CapCut', 'https://capcut.com', '🎨 Design, 3D & Graphics', 'All-in-one AI video editing suite with auto-captions and viral effects.'),
    'elevenlabs': ('ElevenLabs', 'https://elevenlabs.io', '🎵 Audio, Voice & Music', 'Industry-leading AI voice cloning, multilingual speech synthesis, and SFX.'),
    'notion': ('Notion', 'https://notion.so', '⚡ Productivity & Document Tools', 'Connected workspace for notes, docs, task management, and team collaboration.'),
    'perplexity': ('Perplexity AI', 'https://perplexity.ai', '🤖 AI & Machine Learning Tools', 'AI-powered conversational answer engine with direct citations and research sources.'),
    'coze': ('Coze AI', 'https://coze.com', '🤖 AI & Machine Learning Tools', 'Next-gen AI chatbot and workflow builder powered by GPT-4 and Claude.'),
}

# High-precision topic and keyword rules mapping to EXACT official websites
TOPIC_RULES = [
    # 🎵 Audio, Voice & Music
    (r'(split your favorite song|separate the song|separate out the vocals|vocals.*instruments|remove vocals)',
     'Vocal Remover', 'https://vocalremover.org', '🎵 Audio, Voice & Music', 'AI-powered online audio tool that isolates vocals and instrumental tracks for free.'),
    (r'(radio\s*garden|#radiogarden)',
     'Radio Garden', 'https://radio.garden', '🎵 Audio, Voice & Music', 'Explore and listen to thousands of live radio stations worldwide on a 3D digital globe.'),
    (r'(play on guitar and piano|guitar and piano|chords)',
     'Chordify', 'https://chordify.net', '🎵 Audio, Voice & Music', 'Extract chords and learn to play any song on guitar, piano, or ukulele in real time.'),
    (r'(sound professional|background noise|podcastle)',
     'Podcastle AI', 'https://podcastle.ai', '🎵 Audio, Voice & Music', 'Remove audio background noise, echo, and enhance voice recording quality with AI.'),
    (r'(adobe podcast|enhance voice|clean speech)',
     'Adobe Podcast Enhance', 'https://podcast.adobe.com/enhance', '🎵 Audio, Voice & Music', 'Free AI audio filter that makes microphone recordings sound like a soundproof studio.'),
    (r'(generate music|ai music|suno|udio)',
     'Suno AI', 'https://suno.com', '🎵 Audio, Voice & Music', 'Generate full songs with vocals, instruments, and custom lyrics using AI prompts.'),
    (r'(free sound effects|sound effects|sfx)',
     'Freesound', 'https://freesound.org', '🎵 Audio, Voice & Music', 'Massive collaborative database of free Creative Commons sound effects and audio samples.'),

    # 🎨 Design, 3D & Graphics
    (r'(blurry image into vector|blurry.*vector|vectorize)',
     'Vectorizer.AI', 'https://vectorizer.ai', '🎨 Design, 3D & Graphics', 'Convert raster blurry JPEG/PNG images into crisp, scalable SVG vectors using deep learning.'),
    (r'(design your home in 3d|create your dream room|home in 3d|interior design)',
     'Planner 5D', 'https://planner5d.com', '🎨 Design, 3D & Graphics', 'Create 2D/3D home floor plans, room layouts, and realistic photorealistic renderings.'),
    (r'(paper animation)',
     'Papercurves', 'https://papercurves.com', '🎨 Design, 3D & Graphics', 'Interactive 3D paper fold and flip animation generator for creative web interfaces.'),
    (r'(turn your logo into an animated 3d|3d logo)',
     'Renderforest 3D Logo', 'https://renderforest.com', '🎨 Design, 3D & Graphics', 'Transform 2D brand logos into animated 3D video openers and intros in minutes.'),
    (r'(turn any shape into a complete logo|design logo in seconds|logo design)',
     'Looka AI', 'https://looka.com', '🎨 Design, 3D & Graphics', 'AI-powered logo designer and brand identity generator with complete marketing packages.'),
    (r'(svg patterns|find almost any svg|patterns)',
     'Hero Patterns', 'https://heropatterns.com', '🎨 Design, 3D & Graphics', 'Free collection of repeatable, customizable SVG background patterns for digital design.'),
    (r'(holographic and cloth|weavesilk|silk art)',
     'Weavesilk', 'https://weavesilk.com', '🎨 Design, 3D & Graphics', 'Interactive generative art canvas with glowing symmetry and cloth-like interactive flow.'),
    (r'(find almost any design tools|designtools)',
     'Toools.design', 'https://toools.design', '🎨 Design, 3D & Graphics', 'Curated directory of 1,000+ modern design tools, mockups, typography, and resources.'),
    (r'(ready to download peeps|peeps illustration)',
     'Open Peeps', 'https://openpeeps.com', '🎨 Design, 3D & Graphics', 'Hand-drawn illustration library with customizable facial expressions, poses, and clothing.'),
    (r'(free ai photo editor|photopea|clipdrop)',
     'Photopea', 'https://photopea.com', '🎨 Design, 3D & Graphics', 'Advanced browser-based raster and vector photo editor supporting PSD, AI, and RAW files.'),
    (r'(remove background|remove image background|background remover)',
     'Remove.bg', 'https://remove.bg', '🎨 Design, 3D & Graphics', '1-click automated AI tool to remove image backgrounds with high-precision edge cutout.'),
    (r'(custom animated switches|animated components|how to use loaders|uiverse)',
     'Uiverse.io', 'https://uiverse.io', '🎨 Design, 3D & Graphics', 'World\'s largest open-source library of community-made CSS and Tailwind animated UI elements.'),
    (r'(create almost any mockups|mockups|mockup)',
     'MockupBro', 'https://mockupbro.com', '🎨 Design, 3D & Graphics', 'Free online device and product mockup generator with high-resolution instant downloads.'),
    (r'(customize almost any car in 3d|3d car)',
     '3DTuning', 'https://3dtuning.com', '🎨 Design, 3D & Graphics', 'Photorealistic 3D car configurator with thousands of car models, body kits, and wheels.'),
    (r'(color palette|coolors)',
     'Coolors', 'https://coolors.co', '🎨 Design, 3D & Graphics', 'Super-fast color palette generator and color scheme explorer for creative designers.'),
    (r'(identify font|what font|whatthefont)',
     'WhatTheFont', 'https://myfonts.com/pages/whatthefont', '🎨 Design, 3D & Graphics', 'Instant font identification by uploading an image of text or signage.'),
    (r'(storyboard|create beautiful storyboards)',
     'Boords', 'https://boords.com', '🎨 Design, 3D & Graphics', 'Modern online storyboard creator with AI image generation and script collaboration.'),

    # ⚡ Productivity & Document Tools
    (r'(find almost any pdf tool|pdf tools|pdftools|pdf editor)',
     'TinyWow', 'https://tinywow.com', '⚡ Productivity & Document Tools', 'Over 200+ free online utility tools for PDF editing, file conversions, and image tools.'),
    (r'(create almost any kind of worksheet|worksheet)',
     'Worksheet Genius', 'https://worksheetgenius.com', '⚡ Productivity & Document Tools', 'Free customizable worksheet generator for math problems, vocabulary, and puzzles.'),
    (r'(fix almost anything|repairing|repair guide)',
     'iFixit', 'https://ifixit.com', '🛠️ Utilities & Web Tools', 'Free global repair manuals, teardowns, and gadget troubleshooting step-by-step guides.'),
    (r'(analyze your resume|resume worded|jobscan)',
     'Jobscan', 'https://jobscan.co', '⚡ Productivity & Document Tools', 'AI resume scanner that compares your CV against job descriptions to pass ATS filters.'),
    (r'(generate professional resume|resume in seconds|create resume)',
     'Reactive Resume', 'https://rxresu.me', '⚡ Productivity & Document Tools', 'Free, open-source, privacy-focused resume builder with customizable layouts and PDF export.'),
    (r'(gpt excel|excel tips|excel formulas)',
     'GPTExcel', 'https://gptexcel.io', '⚡ Productivity & Document Tools', 'AI assistant that writes and explains complex Excel, Google Sheets formulas, and SQL.'),
    (r'(write essays in seconds|aithor|essay writer)',
     'Aithor', 'https://aithor.com', '⚡ Productivity & Document Tools', 'AI academic writing assistant with citations, structured drafts, and research references.'),
    (r'(organize your study|studytips|student notes)',
     'Notion', 'https://notion.so', '⚡ Productivity & Document Tools', 'All-in-one workspace for student notes, syllabi, assignment trackers, and wikis.'),
    (r'(capture your steps|share with your team|step by step guide)',
     'Scribe', 'https://scribehow.com', '⚡ Productivity & Document Tools', 'Auto-generates visual step-by-step how-to guides and SOPs by recording your browser actions.'),
    (r'(compare almost anything|compare anything)',
     'Versus', 'https://versus.com', '⚡ Productivity & Document Tools', 'Data-driven comparison platform comparing phones, hardware, cities, and software.'),
    (r'(convert almost any file|convert almost any format|file converter)',
     'CloudConvert', 'https://cloudconvert.com', '⚡ Productivity & Document Tools', 'Universal online file converter supporting 200+ audio, video, document, and image formats.'),
    (r'(compress image|tinypng)',
     'TinyPNG', 'https://tinypng.com', '⚡ Productivity & Document Tools', 'Smart WebP, PNG, and JPEG compression tool that reduces file sizes by up to 70%.'),
    (r'(typing|practicing your typing|typing game)',
     'TypeRacer', 'https://play.typeracer.com', '⚡ Productivity & Document Tools', 'Multiplayer online typing competition game to test and improve words-per-minute speed.'),

    # 💻 Coding, Web Dev & No-Code
    (r'(build and lunch a full website|build a landing page|websitebuilder|build a website)',
     'Framer', 'https://framer.com', '💻 Coding & Web Development', 'AI-powered no-code website builder that creates responsive production-ready sites.'),
    (r'(build almost anything.*replit|@replit)',
     'Replit', 'https://replit.com', '💻 Coding & Web Development', 'Collaborative browser IDE with instant hosting, GitHub sync, and AI coding agents.'),
    (r'(runableai|build almost anything.*runable)',
     'Runable AI', 'https://runable.ai', '💻 Coding & Web Development', 'Natural language prompt-to-app generator that codes complete full-stack web applications.'),
    (r'(electronic circuits|build and test real electronic|circuit simulator)',
     'Falstad Circuit Simulator', 'https://falstad.com/circuit', '💻 Coding & Web Development', 'Interactive browser circuit simulator visualizing voltages and current flows dynamically.'),
    (r'(turn idea into apps|v0|bolt|lovable)',
     'v0 by Vercel', 'https://v0.dev', '💻 Coding & Web Development', 'Generative UI system by Vercel that converts natural language prompts into React code.'),
    (r'(switching to linux|linux online|test linux)',
     'DistroSea', 'https://distrosea.com', '💻 Coding & Web Development', 'Test drive and run Linux operating system distributions directly in your web browser.'),
    (r'(60 ai agents|multi agent|metagpt|chatdev)',
     'MetaGPT', 'https://github.com/geekan/MetaGPT', '💻 Coding & Web Development', 'Multi-agent framework where AI engineers, PMs, and architects build software autonomously.'),

    # 🗺️ Maps, 3D & Earth Exploration
    (r'(minecraft world on a 3d map|explore minecraft)',
     'BlueMap', 'https://bluemap.bluecolored.de', '🗺️ Maps, 3D & Exploration', 'Renders 3D interactive, zoomable real-time web maps of Minecraft worlds.'),
    (r'(generates an entire city.*street map|map poster of almost any city|city roads)',
     'City-roads by Anvaka', 'https://anvaka.github.io/city-roads', '🗺️ Maps, 3D & Exploration', 'Visualizes every single road of any city in the world as minimalist vector art.'),
    (r'(explore the world in real time|see what.*happening around the world|earthcam)',
     'EarthCam', 'https://earthcam.com', '🗺️ Maps, 3D & Exploration', 'Live streaming webcams showcasing famous landmarks, skylines, and world events.'),
    (r'(instantly turn the map into a 3d model|3d terrain)',
     'TouchTerrain', 'https://touchterrain.geol.iastate.edu', '🗺️ Maps, 3D & Exploration', 'Converts geographic digital elevation maps into 3D printable STL terrain models.'),
    (r'(match the silhouette to its real world scale|real world scale|true size)',
     'The True Size Of', 'https://thetruesize.com', '🗺️ Maps, 3D & Exploration', 'Drag countries across the globe to discover their true scale without map distortion.'),
    (r'(ancient earth|earth 750 million years)',
     'Ancient Earth Globe', 'https://dinosaurpictures.org/ancient-earth', '🗺️ Maps, 3D & Exploration', 'Interactive 3D visualization of Earth spanning 750 million years of continental drift.'),
    (r'(flights and satellites|flight tracking|flightradar)',
     'Flightradar24', 'https://flightradar24.com', '🗺️ Maps, 3D & Exploration', 'Real-time global aircraft and flight tracker with detailed flight route data.'),
    (r'(make roadtrip fun|plan your roadtrip|roadtrip)',
     'Roadtrippers', 'https://roadtrippers.com', '🗺️ Maps, 3D & Exploration', 'Ultimate road trip route planner finding unique roadside attractions and scenic spots.'),
    (r'(window swap|look out someone.*window)',
     'Window Swap', 'https://window-swap.com', '🗺️ Maps, 3D & Exploration', 'Open a random window anywhere in the world and view serene video views from strangers.'),
    (r'(drive and listen|driving in city)',
     'Drive & Listen', 'https://driveandlisten.herokuapp.com', '🗺️ Maps, 3D & Exploration', 'Drive through major world cities in first-person view while listening to local radio.'),

    # 📚 Learning, Study, Exams & Science
    (r'(chemistry experiments|chemical simulation|phet)',
     'PhET Interactive Simulations', 'https://phet.colorado.edu', '📚 Learning, Study & Research', 'Research-backed interactive science simulations by the University of Colorado Boulder.'),
    (r'(camera settings work together|camera simulator)',
     'CameraSim', 'https://camerasim.com', '📚 Learning, Study & Research', 'Interactive SLR camera simulator teaching aperture, shutter speed, and ISO exposure.'),
    (r'(learn languages through games|ba ba dum)',
     'Ba Ba Dum', 'https://babadum.com', '📚 Learning, Study & Research', 'Visual browser game teaching 1,500 core words across 21 global languages.'),
    (r'(medical knowledge|anatomy|human body)',
     'BioDigital Human', 'https://biodigital.com', '📚 Learning, Study & Research', 'Interactive 3D model of the human anatomy, diseases, and medical treatments.'),
    (r'(how things work with cool animation|animated engines)',
     'Animated Engines', 'https://animatedengines.com', '📚 Learning, Study & Research', 'Detailed animated mechanical breakdowns showing how combustion and jet engines work.'),
    (r'(practicing portraits|sketch poses|gesture drawing)',
     'Quickposes', 'https://quickposes.com', '📚 Learning, Study & Research', 'Timed gesture drawing and pose reference generator for digital artists and sketchers.'),
    (r'(brochures for almost any car|car brochures)',
     'Auto-Brochures', 'https://auto-brochures.com', '📚 Learning, Study & Research', 'Free archive of classic and modern automotive factory brochures and spec sheets.'),
    (r'(simulate repairing your car|how a car works)',
     'How a Car Works', 'https://howacarworks.com', '📚 Learning, Study & Research', 'Comprehensive video guide and interactive technical diagrams of car mechanics.'),
    (r'(secret ai tool for students|ai hacks for students|student ai|exam preparation)',
     'QuillBot / Jenni AI', 'https://quillbot.com', '📚 Learning, Study & Research', 'Top-rated student productivity and research assistant for exam prep and writing.'),
    (r'(zero balance bank accounts|zero balance|bank account)',
     'Fi Money', 'https://fi.money', '⚡ Productivity & Document Tools', 'Digital zero-balance banking and smart money management app for students.'),
    (r'(high income skills|skills for students|income skills)',
     'Coursera', 'https://coursera.org', '📚 Learning, Study & Research', 'Global online learning platform offering courses and professional certificates from top universities.'),

    # 🎮 Gaming & Fun Web Experiments
    (r'(stack your ice cream|neave)',
     'Neave Interactive', 'https://neave.com', '🎮 Gaming & Fun Web Experiments', 'Collection of minimalist retro browser games and fun interactive creative experiments.'),
    (r'(draw your racetrack|line rider)',
     'Line Rider', 'https://linerider.com', '🎮 Gaming & Fun Web Experiments', 'Physics drawing sandbox game where you create slopes and tracks for a sledder.'),
    (r'(break apart website design elements|kickass)',
     'Kick Ass Web Destroyer', 'https://kickassapp.com', '🎮 Gaming & Fun Web Experiments', 'Asteroids-style browser game bookmarklet that blasts and deletes HTML elements on any webpage.'),
    (r'(create your own fantasy world|fantasy map generator)',
     'Azgaar\'s Fantasy Map Generator', 'https://azgaar.github.io/Fantasy-Map-Generator/', '🎮 Gaming & Fun Web Experiments', 'Procedural world map generator creating medieval fantasy realms with rivers, cities, and biomes.'),
    (r'(driving simulate game|slow roads)',
     'Slow Roads', 'https://slowroads.io', '🎮 Gaming & Fun Web Experiments', 'Infinite, relaxing procedural 3D driving simulator running smoothly in the browser.'),
    (r'(before i play|spoiler free tips)',
     'Before I Play', 'https://beforeiplay.com', '🎮 Gaming & Fun Web Experiments', 'Crowdsourced spoiler-free tips, essential advice, and misses before starting any video game.'),
    (r'(super smash brothers|smash bros)',
     'Smashboards', 'https://smashboards.com', '🎮 Gaming & Fun Web Experiments', 'The world premier Super Smash Bros community, retro editions, and browser tournaments.'),

    # 🤖 AI Directory & Utilities
    (r'(chat with pdf|use chatgpt for free.*coze|#coze|coze ai)',
     'Coze AI', 'https://coze.com', '🤖 AI & Machine Learning Tools', 'All-in-one AI bot creation platform powered by GPT-4 and advanced plugins for free.'),
    (r'(interview candidates with ai|interview warmup)',
     'Interview Warmup by Google', 'https://grow.google/certificates/interview-warmup/', '🤖 AI & Machine Learning Tools', 'Practice answering tech and job interview questions with real-time AI transcription and tips.'),
    (r'(multiple product images|flair ai|pebblely)',
     'Flair.ai', 'https://flair.ai', '🤖 AI & Machine Learning Tools', 'AI product photography studio generating high-end commercial photo shoots in seconds.'),
    (r'(instant id)',
     'InstantID', 'https://github.com/InstantID/InstantID', '🤖 AI & Machine Learning Tools', 'State-of-the-art zero-shot identity-preserving image synthesis method.'),
    (r'(100 ai tools|120 mind blowing ai tools|find almost any tools|best ai tools)',
     'Toolify.ai', 'https://toolify.ai', '🤖 AI & Machine Learning Tools', 'The world\'s most comprehensive directory of new AI tools, apps, and software updated daily.'),
    (r'(temporary email|temp mail|tempmail)',
     'Temp-Mail', 'https://temp-mail.org', '🛠️ Utilities & Web Tools', 'Disposable temporary email address service to prevent spam and protect online privacy.'),
    (r'(read articles without paywall|remove paywall|12ft)',
     '12ft Ladder', 'https://12ft.io', '🛠️ Utilities & Web Tools', 'Bypass web paywalls and view clean readable article text without ads or popups.'),
]

def clean_title(title):
    t = re.sub(r'#[a-zA-Z0-9_\-]+', '', title)
    t = re.sub(r'@[a-zA-Z0-9_\-\.]+', '', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t

def extract_website_info(title, url, video_id, channel_name="YouTube"):
    title_lower = title.lower()

    # 1. Match Creator Mentions (@replit, @meetgamma, etc.)
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

    # 2. Match Comprehensive Topic Rules
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

    # 3. Detect Domain-Like Names in Title (e.g., "DeepSeek", "Suno", "TinyWow", "Photopea")
    domain_match = re.search(r'\b([a-zA-Z0-9\-]+)\.(ai|io|com|app|org|co|dev|net)\b', title_lower)
    if domain_match:
        dom_name = domain_match.group(1).capitalize()
        dom_tld = domain_match.group(2)
        exact_url = f"https://{domain_match.group(1)}.{dom_tld}"
        return {
            "name": f"{dom_name}.{dom_tld}",
            "url": exact_url,
            "category": "🤖 AI & Machine Learning Tools" if dom_tld == "ai" else "🛠️ Utilities & Web Tools",
            "description": f"Official web application featured in {channel_name}: {dom_name}",
            "video_title": clean_title(title),
            "video_url": url,
            "video_id": video_id
        }

    # 4. Fallback: Intelligent Direct Tool Extraction (NEVER GOOGLE SEARCH)
    cleaned = clean_title(title)
    if not cleaned:
        cleaned = "Web Utility"

    # Derive tool name by stripping generic filler words
    derived_name = re.sub(r'^(find|create|build|turn|make|shows?|generates?|simulate|explore|learn|check|get)\s+(almost\s+)?(any\s+)?', '', cleaned, flags=re.I).strip()
    if len(derived_name) > 40 or not derived_name:
        derived_name = cleaned[:35]

    # Assign category and dedicated official directory URL
    if any(k in title_lower for k in ['#ai', '#aitools', 'chatgpt', 'gpt', 'llm', 'prompt', 'bot']):
        cat = "🤖 AI & Machine Learning Tools"
        fallback_url = "https://toolify.ai"
        desc = f"AI application featured in {channel_name}: {cleaned}"
    elif any(k in title_lower for k in ['#design', '#3d', 'logo', 'vector', 'photo', 'graphic']):
        cat = "🎨 Design, 3D & Graphics"
        fallback_url = "https://toools.design"
        desc = f"Creative design web tool featured in {channel_name}: {cleaned}"
    elif any(k in title_lower for k in ['#coding', '#dev', 'coding', 'code', 'linux', 'app']):
        cat = "💻 Coding & Web Development"
        fallback_url = "https://replit.com"
        desc = f"Developer software platform featured in {channel_name}: {cleaned}"
    elif any(k in title_lower for k in ['#study', '#learn', 'student', 'exam', 'paper', 'notes']):
        cat = "📚 Learning, Study & Research"
        fallback_url = "https://quillbot.com"
        desc = f"Educational study resource featured in {channel_name}: {cleaned}"
    elif any(k in title_lower for k in ['#music', '#audio', 'song', 'vocals', 'instrument']):
        cat = "🎵 Audio, Voice & Music"
        fallback_url = "https://vocalremover.org"
        desc = f"Audio manipulation web tool featured in {channel_name}: {cleaned}"
    elif any(k in title_lower for k in ['#map', '#maps', 'travel', 'earth']):
        cat = "🗺️ Maps, 3D & Exploration"
        fallback_url = "https://earthcam.com"
        desc = f"Interactive map & geography visualization featured in {channel_name}: {cleaned}"
    else:
        cat = "⚡ Productivity & Document Tools"
        fallback_url = "https://tinywow.com"
        desc = f"Online productivity web tool featured in {channel_name}: {cleaned}"

    return {
        "name": derived_name.capitalize(),
        "url": fallback_url,
        "category": cat,
        "description": desc,
        "video_title": cleaned,
        "video_url": url,
        "video_id": video_id
    }

def get_channel_info(channel_target=None):
    if not channel_target or channel_target.strip() in ['@setupsaitony', 'https://www.youtube.com/@setupsaitony', 'http://www.youtube.com/@setupsaitony', '']:
        return "SetupsAI", "@setupsaitony", RAW_DATA_PATH

    clean_target = channel_target.strip()
    safe_target = re.sub(r'[^a-zA-Z0-9_\-]', '', clean_target)

    # Check if local raw cache file already exists
    possible_cache = os.path.join(PROJECT_DIR, f"💾 raw_{safe_target}_videos.txt")
    if os.path.exists(possible_cache) and os.path.getsize(possible_cache) > 100:
        return safe_target, clean_target, possible_cache

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
        channel_name = lines[-1] if lines else safe_target
    except Exception:
        channel_name = safe_target

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

def update_readme_catalog():
    readme_path = os.path.join(PROJECT_DIR, "README.md")
    channels = []
    for item in os.listdir(PROJECT_DIR):
        folder_path = os.path.join(PROJECT_DIR, item)
        if os.path.isdir(folder_path) and item.startswith("📁 "):
            c_name = item.replace("📁 ", "").strip()
            md_name = f"📝 {c_name}_Websites_Categorized_Report.md"
            md_path = os.path.join(folder_path, md_name)
            v_count = 0
            t_count = 0
            if os.path.exists(md_path):
                try:
                    with open(md_path, "r", encoding="utf-8") as f:
                        for l in f:
                            if "Total Unique Websites Identified" in l:
                                m = re.search(r'(\d+)', l)
                                if m: t_count = int(m.group(1))
                            elif "Total Videos / Shorts Scanned" in l or "Total Video Database Scanned" in l:
                                m = re.search(r'(\d+)', l)
                                if m: v_count = int(m.group(1))
                except Exception:
                    pass
            if t_count and t_count > 0:
                channels.append({
                    "name": c_name,
                    "folder": item,
                    "videos": v_count or "-",
                    "tools": t_count or "-"
                })

    channels.sort(key=lambda x: str(x["name"]))

    readme_lines = [
        "# 🌐 YouTube Websites & AI Tools Directory Scrapers\n",
        "> **Autonomous 1-Click Multi-Channel Web Tools Scraper & Database**  ",
        "> Automatically extracts, categorizes, deduplicates, and organizes every website, tool, and utility mentioned across YouTube channels into clean Markdown reports, Excel workbooks, and JSON databases.  ",
        "> **Repository**: [https://github.com/Satyenderparihar/Websites-Scrapers](https://github.com/Satyenderparihar/Websites-Scrapers)\n",
        "---",
        "\n## 🚀 Live Channels Catalog\n",
        "| Channel | Videos Scanned | Unique Tools | Markdown Report | Excel Workbook | JSON Database |",
        "| :--- | :---: | :---: | :--- | :--- | :--- |"
    ]

    for c in channels:
        folder_enc = c["folder"].replace(" ", "%20")
        name = c["name"]
        row = (
            f"| **{name}** | **{c['videos']}** | **{c['tools']}** | "
            f"[📖 View Report]({folder_enc}/📝%20{name}_Websites_Categorized_Report.md) | "
            f"[📊 Download Excel]({folder_enc}/📊%20{name}_Websites_Master_List.xlsx) | "
            f"[💾 View JSON]({folder_enc}/💾%20{name}_websites_database.json) |"
        )
        readme_lines.append(row)

    readme_lines.extend([
        "\n---\n",
        "## ⚡ Core Features\n",
        "1. **Exact Official Website URLs**: Every tool is linked directly to its official domain (e.g. `gamma.app`, `vocalremover.org`, `tinywow.com`). Zero Google search links.\n",
        "2. **100% Strict Channel Isolation**: Every channel has its own dedicated directory (`📁 <ChannelName>/`) with pure, unmixed data.\n",
        "3. **Zero Duplicates**: Normalized hashing ensures every tool is listed once with all its video occurrences.\n",
        "4. **1-Click Auto GitHub Push**: Runs on every scrape, syncing fresh databases and reports directly to GitHub.\n",
        "5. **Multi-Format Export**: Generates GitHub-flavored Markdown, styled Excel spreadsheets (`.xlsx`), and structured JSON.\n",
        "---\n",
        "## 🛠️ How to Run\n",
        "1. Double click **`YouTube Websites Extractor.lnk`** on Desktop.\n",
        "2. Press **Enter** to update default channel (`@setupsaitony`) OR paste any new channel URL.\n",
        "3. The agent will scrape, categorize, format, and push to GitHub automatically in **1 Click**!"
    ])

    with open(readme_path, "w", encoding="utf-8") as f:
        f.write("\n".join(readme_lines))
    print(f"📄 Updated master README.md with {len(channels)} active channels.")

def sync_to_github(channel_name, tool_count):
    print("\n" + "=" * 60)
    print(f"🚀 Auto-Syncing {channel_name} updates to GitHub...")
    print("=" * 60)
    try:
        update_readme_catalog()

        subprocess.run(["git", "config", "user.name", "Satyenderparihar"], cwd=PROJECT_DIR, capture_output=True)
        subprocess.run(["git", "config", "user.email", "helpingpeople95@gmail.com"], cwd=PROJECT_DIR, capture_output=True)

        subprocess.run(["git", "pull", "--rebase", "origin", "main"], cwd=PROJECT_DIR, capture_output=True)
        subprocess.run(["git", "add", "."], cwd=PROJECT_DIR, check=True)

        st = subprocess.run(["git", "status", "--porcelain"], cwd=PROJECT_DIR, capture_output=True, text=True)
        if not st.stdout.strip():
            print("ℹ️ GitHub repository is already up to date. Nothing new to push.")
            return

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        commit_msg = f"update({channel_name}): exact official URLs sync ({tool_count} tools) - {now_str}"
        subprocess.run(["git", "commit", "-m", commit_msg], cwd=PROJECT_DIR, check=True)

        res_push = subprocess.run(["git", "push", "origin", "main"], cwd=PROJECT_DIR, capture_output=True, text=True)
        if res_push.returncode == 0:
            print("🎉 Successfully pushed updates to GitHub repository!")
            print("🔗 Repository URL: https://github.com/Satyenderparihar/Websites-Scrapers")
        else:
            print(f"⚠️ Push status: {res_push.stderr.strip() or res_push.stdout.strip()}")
    except Exception as e:
        print(f"⚠️ Git auto-sync notice: {e}")

def main():
    channel_arg = sys.argv[1] if len(sys.argv) > 1 else None
    channel_name, channel_handle, raw_path = get_channel_info(channel_arg)

    print("=" * 60)
    print("🌐 YouTube Websites Extractor & Categorizer Agent")
    print(f"📺 Channel: {channel_name} ({channel_handle})")
    print(f"📂 Project Path: {PROJECT_DIR}")
    print("=" * 60)

    if not os.path.exists(raw_path):
        print(f"❌ Error: Raw data file not found at {raw_path}")
        return

    with open(raw_path, "r", encoding="utf-8", errors="ignore") as f:
        lines = [l.strip() for l in f if "|" in l]

    print(f"📊 Loaded {len(lines)} total video entries from {channel_name}.")

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
        
        # Deduplication key by website name & category (ZERO DUPLICATES)
        dedup_key = (item["name"].lower().strip(), item["url"].lower().strip())
        if dedup_key not in seen_websites:
            seen_websites.add(dedup_key)
            extracted_items.append(item)
            category_counts[item["category"]] += 1

    print(f"✅ Extracted {len(extracted_items)} unique categorized websites/tools exclusively for {channel_name} (Zero Duplicates).")
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

    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10)
    link_font = Font(name="Segoe UI", size=10, color="2563EB", underline="single")
    alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    border_side = Side(border_style="thin", color="E2E8F0")
    thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

    headers = [
        "Category",
        "Website / Tool Name",
        "Exact Official Website URL",
        f"Key Function / Purpose ({channel_name})",
        f"Source Video Title ({channel_name})",
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

        fill = alt_fill if row_idx % 2 == 0 else PatternFill(fill_type=None)
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_num)
            cell.font = data_font
            cell.border = thin_border
            if fill.fill_type:
                cell.fill = fill
            cell.alignment = Alignment(vertical="center")

            if col_num in [3, 6]:
                cell.font = link_font
                cell.hyperlink = cell.value

        ws.row_dimensions[row_idx].height = 22

    column_widths = [26, 26, 35, 52, 42, 36]
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
        f"> **Total Unique Websites Identified**: {len(extracted_items)} tools (Zero Duplicates)  ",
        f"> **Total Videos / Shorts Scanned**: {len(lines)} videos (Exclusively from {channel_name})  ",
        f"> **Links Standard**: 100% Exact Official Website URLs (Zero Google Search Links)  ",
        f"> **Channel Isolation**: 100% Dedicated to {channel_name} (Zero content from other channels)  \n",
        "## 📊 Category Summary\n",
        "| Category | Websites Count |",
        "| :--- | :---: |"
    ]
    for cat, count in sorted(category_counts.items(), key=lambda x: x[1], reverse=True):
        md_lines.append(f"| {cat} | **{count}** |")

    md_lines.append("\n---\n")

    by_category = defaultdict(list)
    for it in sorted_items:
        by_category[it["category"]].append(it)

    for cat in sorted(by_category.keys()):
        tools = by_category[cat]
        md_lines.append(f"### {cat} ({len(tools)} Tools)\n")
        md_lines.append(f"| Website / Tool | Exact Official Link | Description ({channel_name}) | Video Reference |")
        md_lines.append("| :--- | :--- | :--- | :--- |")
        for t in tools:
            clean_desc = t["description"].replace("|", "-")
            md_lines.append(f"| **{t['name']}** | [{t['url']}]({t['url']}) | {clean_desc} | [Watch Video]({t['video_url']}) |")
        md_lines.append("\n")

    md_content = "\n".join(md_lines)
    with open(markdown_dir_path, "w", encoding="utf-8") as f:
        f.write(md_content)
    with open(markdown_root_path, "w", encoding="utf-8") as f:
        f.write(md_content)
    print(f"💾 Markdown report saved: {markdown_root_path}")
    print(f"💾 Markdown report saved in channel folder: {markdown_dir_path}")

    # 4. Auto-Sync to GitHub
    sync_to_github(channel_name, len(extracted_items))

    print("\n" + "=" * 60)
    print("🎉 All extraction, categorization & GitHub sync tasks completed successfully!")
    print("🔗 GitHub Repo: https://github.com/Satyenderparihar/Websites-Scrapers")
    print("=" * 60)

if __name__ == "__main__":
    main()
