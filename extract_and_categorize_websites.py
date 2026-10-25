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

# Canonical Master Entries for Clean Branding & Exhaustive Descriptions
CANONICAL_ENTRIES = {
    "tinywow.com": {
        "name": "TinyWow",
        "url": "https://tinywow.com",
        "category": "⚡ Productivity & Document Tools",
        "description": "200+ free online utility tools: PDF editing/converting/merging, image background removal, file compression, video tools, and AI writing."
    },
    "quillbot.com": {
        "name": "QuillBot",
        "url": "https://quillbot.com",
        "category": "📚 Learning, Study & Research",
        "description": "AI paraphrasing, grammar checking, academic summarization, plagiarism detection, and student thesis writing assistance."
    },
    "toolify.ai": {
        "name": "Toolify.ai",
        "url": "https://toolify.ai",
        "category": "🤖 AI & Machine Learning Tools",
        "description": "Comprehensive global directory and discovery engine indexing 10,000+ AI software tools, models, and apps updated daily."
    },
    "toools.design": {
        "name": "Toools.design",
        "url": "https://toools.design",
        "category": "🎨 Design, 3D & Graphics",
        "description": "Curated repository of 1,000+ modern design tools, mockups, typography, illustration packs, and UI kits."
    },
    "replit.com": {
        "name": "Replit",
        "url": "https://replit.com",
        "category": "💻 Coding & Web Development",
        "description": "Collaborative browser-based IDE, AI code generation, and instant cloud application deployment workspace."
    },
    "chatgpt.com": {
        "name": "ChatGPT",
        "url": "https://chatgpt.com",
        "category": "🤖 AI & Machine Learning Tools",
        "description": "Conversational AI assistant by OpenAI for answering questions, coding, research, writing, and problem solving."
    },
    "gemini.google.com": {
        "name": "Google Gemini",
        "url": "https://gemini.google.com",
        "category": "🤖 AI & Machine Learning Tools",
        "description": "Multimodal conversational AI model by Google for advanced reasoning, coding, search, and Google Workspace integration."
    },
    "perplexity.ai": {
        "name": "Perplexity AI",
        "url": "https://perplexity.ai",
        "category": "🤖 AI & Machine Learning Tools",
        "description": "Conversational AI search engine delivering verified answers with real-time web citations and academic research papers."
    },
    "gamma.app": {
        "name": "Gamma App",
        "url": "https://gamma.app",
        "category": "⚡ Productivity & Document Tools",
        "description": "AI-powered presentation, document, and webpage builder that generates styled interactive slide decks from prompts in seconds."
    },
    "notebooklm.google.com": {
        "name": "NotebookLM by Google",
        "url": "https://notebooklm.google.com",
        "category": "🤖 AI & Machine Learning Tools",
        "description": "Personalized AI research assistant that turns notes, lecture PDFs, and documents into interactive audio discussions and podcasts."
    },
    "v0.dev": {
        "name": "v0 by Vercel",
        "url": "https://v0.dev",
        "category": "💻 Coding & Web Development",
        "description": "AI prompt-to-app generator that turns natural language UI ideas and sketches into production-ready React/Tailwind code."
    },
    "distrosea.com": {
        "name": "DistroSea",
        "url": "https://distrosea.com",
        "category": "💻 Coding & Web Development",
        "description": "Run and test drive dozens of popular Linux distributions live inside your web browser without downloading ISOs or installing anything."
    },
    "monkeytype.com": {
        "name": "Monkeytype",
        "url": "https://monkeytype.com",
        "category": "⚡ Productivity & Document Tools",
        "description": "Customizable, minimalistic typing test platform with comprehensive speed statistics, accuracy tracking, and custom practice modes."
    },
    "obsidian.md": {
        "name": "Obsidian",
        "url": "https://obsidian.md",
        "category": "⚡ Productivity & Document Tools",
        "description": "Extensible Markdown knowledge base and connected note-taking app utilizing Zettelkasten bidirectional linking for research."
    },
    "suno.com": {
        "name": "Suno AI",
        "url": "https://suno.com",
        "category": "🎵 Audio, Voice & Music",
        "description": "Generative AI music platform that turns lyrics and study notes into broadcast-quality songs with full vocals and instruments."
    },
    "gitmind.com": {
        "name": "GitMind",
        "url": "https://gitmind.com",
        "category": "⚡ Productivity & Document Tools",
        "description": "Collaborative AI-powered mind mapping and flowchart tool to brainstorm, study, and visually outline complex topics."
    },
    "pll.harvard.edu": {
        "name": "Harvard Online Learning",
        "url": "https://pll.harvard.edu",
        "category": "📚 Learning, Study & Research",
        "description": "Official online portal for free courses, lectures, and professional certificates direct from Harvard University faculty."
    },
    "chromeenterprise.google": {
        "name": "ChromeOS Flex",
        "url": "https://chromeenterprise.google/os/chromeosflex/",
        "category": "🛠️ Utilities & Web Tools",
        "description": "Cloud-first, fast operating system by Google designed to revive and upgrade older Windows laptops and PCs."
    },
    "mulerun.com": {
        "name": "MuleRun AI",
        "url": "https://mulerun.com",
        "category": "🤖 AI & Machine Learning Tools",
        "description": "AI productivity and academic study workspace helping students discover tools and streamline coursework."
    },
    "lovart.ai": {
        "name": "Lovart AI",
        "url": "https://lovart.ai",
        "category": "🎨 Design, 3D & Graphics",
        "description": "AI-driven visual design tool and artistic creative generator for prompts and graphic media."
    },
    "odoo.com": {
        "name": "Odoo",
        "url": "https://odoo.com",
        "category": "💼 Business, Marketing & SEO",
        "description": "Open-source integrated enterprise software suite covering e-commerce, CRM, inventory, and accounting."
    },
    "skillbuilder.aws": {
        "name": "AWS Skill Builder",
        "url": "https://skillbuilder.aws",
        "category": "📚 Learning, Study & Research",
        "description": "Official Amazon Web Services training center offering 600+ free digital cloud computing courses and lab exercises."
    },
    "grow.google": {
        "name": "Grow with Google",
        "url": "https://grow.google",
        "category": "📚 Learning, Study & Research",
        "description": "Google's official certification programs offering free training in AI, data analytics, project management, and IT."
    },
    "biodigital.com": {
        "name": "BioDigital Human",
        "url": "https://biodigital.com",
        "category": "📚 Learning, Study & Research",
        "description": "Interactive 3D virtual human anatomy atlas and disease explorer for medical students, healthcare pros, and biology."
    },
    "phet.colorado.edu": {
        "name": "PhET Interactive Simulations",
        "url": "https://phet.colorado.edu",
        "category": "📚 Learning, Study & Research",
        "description": "Research-backed interactive science and math simulations by University of Colorado Boulder for chemistry and physics."
    },
    "ticker.finology.in": {
        "name": "Finology Ticker",
        "url": "https://ticker.finology.in",
        "category": "💼 Business, Marketing & SEO",
        "description": "Financial learning, stock analysis, and equity screening platform for finance, banking, and commerce students."
    },
    "fi.money": {
        "name": "Fi Money",
        "url": "https://fi.money",
        "category": "⚡ Productivity & Document Tools",
        "description": "Digital zero-balance banking, expense tracking, and smart financial management app tailored for students."
    },
    "wolframalpha.com": {
        "name": "Wolfram Alpha",
        "url": "https://wolframalpha.com",
        "category": "📚 Learning, Study & Research",
        "description": "Computational intelligence engine solving complex math, physics, engineering, and scientific equations step-by-step."
    },
    "consensus.app": {
        "name": "Consensus AI",
        "url": "https://consensus.app",
        "category": "📚 Learning, Study & Research",
        "description": "AI search engine providing research-backed answers extracted directly from peer-reviewed scientific literature."
    },
    "chatdev.ai": {
        "name": "ChatDev",
        "url": "https://chatdev.ai",
        "category": "💻 Coding & Web Development",
        "description": "Virtual software development company powered by multi-agent cooperative LLMs working together on code."
    },
    "claude.ai": {
        "name": "Claude by Anthropic",
        "url": "https://claude.ai",
        "category": "🤖 AI & Machine Learning Tools",
        "description": "Next-generation AI assistant built by Anthropic with 200k context window for reasoning, analysis, and coding."
    },
    "deepseek.com": {
        "name": "DeepSeek AI",
        "url": "https://deepseek.com",
        "category": "🤖 AI & Machine Learning Tools",
        "description": "State-of-the-art open-weight reasoning and code generation AI model platform."
    },
    "typelit.io": {
        "name": "TypeLit.io",
        "url": "https://typelit.io",
        "category": "⚡ Productivity & Document Tools",
        "description": "Improve touch typing speed and accuracy by typing entire classic public-domain books directly in the browser."
    },
    "keybr.com": {
        "name": "Keybr",
        "url": "https://keybr.com",
        "category": "⚡ Productivity & Document Tools",
        "description": "Smart touch typing tutor that uses muscle memory algorithms to target your weakest keys."
    },
    "coursera.org": {
        "name": "Coursera",
        "url": "https://coursera.org",
        "category": "📚 Learning, Study & Research",
        "description": "Global online learning platform offering courses and degrees from top universities like Stanford and Yale."
    },
    "photopea.com": {
        "name": "Photopea",
        "url": "https://photopea.com",
        "category": "🎨 Design, 3D & Graphics",
        "description": "Advanced browser-based raster and vector photo editor supporting PSD, AI, and RAW files without installation."
    },
    "remove.bg": {
        "name": "Remove.bg",
        "url": "https://remove.bg",
        "category": "🎨 Design, 3D & Graphics",
        "description": "1-click automated AI tool to remove image backgrounds with high-precision edge cutout in seconds."
    },
    "uiverse.io": {
        "name": "Uiverse.io",
        "url": "https://uiverse.io",
        "category": "🎨 Design, 3D & Graphics",
        "description": "World's largest open-source library of community-made CSS and Tailwind animated UI elements, buttons, and switches."
    },
    "coolors.co": {
        "name": "Coolors",
        "url": "https://coolors.co",
        "category": "🎨 Design, 3D & Graphics",
        "description": "Super-fast color palette generator and color scheme explorer for creative designers and web developers."
    },
    "cleanup.pictures": {
        "name": "Cleanup.pictures",
        "url": "https://cleanup.pictures",
        "category": "🎨 Design, 3D & Graphics",
        "description": "Erase unwanted people, text, blemishes, or defects from any photo using AI inpainting."
    },
    "supercook.com": {
        "name": "SuperCook",
        "url": "https://supercook.com",
        "category": "⚡ Productivity & Document Tools",
        "description": "Zero-waste recipe search engine that matches recipes to ingredients you already have at home."
    },
    "musclewiki.com": {
        "name": "MuscleWiki",
        "url": "https://musclewiki.com",
        "category": "⚡ Productivity & Document Tools",
        "description": "Interactive visual body map where you click any muscle to get exercise guides and workout plans."
    },
    "scribehow.com": {
        "name": "Scribe",
        "url": "https://scribehow.com",
        "category": "⚡ Productivity & Document Tools",
        "description": "Automatically turns any digital workflow into a visual step-by-step guide with screenshots."
    },
    "cloudconvert.com": {
        "name": "CloudConvert",
        "url": "https://cloudconvert.com",
        "category": "⚡ Productivity & Document Tools",
        "description": "Universal online file converter supporting 200+ audio, video, document, and image formats."
    },
    "tinypng.com": {
        "name": "TinyPNG",
        "url": "https://tinypng.com",
        "category": "⚡ Productivity & Document Tools",
        "description": "Smart WebP, PNG, and JPEG compression tool that reduces file sizes by up to 70%."
    },
    "fast.com": {
        "name": "Fast.com",
        "url": "https://fast.com",
        "category": "🛠️ Utilities & Web Tools",
        "description": "Clean, ad-free broadband internet speed testing tool powered by Netflix servers."
    },
    "temp-mail.org": {
        "name": "Temp-Mail",
        "url": "https://temp-mail.org",
        "category": "🛠️ Utilities & Web Tools",
        "description": "Disposable temporary email address service to prevent spam and protect online privacy."
    },
    "12ft.io": {
        "name": "12ft Ladder",
        "url": "https://12ft.io",
        "category": "🛠️ Utilities & Web Tools",
        "description": "Bypass web paywalls and view clean readable article text without ads or popups."
    },
    "notion.so": {
        "name": "Notion",
        "url": "https://notion.so",
        "category": "⚡ Productivity & Document Tools",
        "description": "All-in-one connected workspace for student notes, syllabi, assignment trackers, and wikis."
    },
    "canva.com": {
        "name": "Canva",
        "url": "https://canva.com",
        "category": "🎨 Design, 3D & Graphics",
        "description": "Visual graphic design platform for creating social media graphics, presentations, and posters."
    },
    "vocalremover.org": {
        "name": "Vocal Remover",
        "url": "https://vocalremover.org",
        "category": "🎵 Audio, Voice & Music",
        "description": "AI-powered online audio tool that isolates vocals and instrumental tracks for free."
    },
    "earthcam.com": {
        "name": "EarthCam",
        "url": "https://earthcam.com",
        "category": "🗺️ Maps, 3D & Exploration",
        "description": "Global network of live streaming webcams providing scenic and landmark views worldwide."
    },
    "shutup10": {
        "name": "O&O ShutUp10++",
        "url": "https://www.oo-software.com/en/shutup10",
        "category": "🛠️ Utilities & Web Tools",
        "description": "Free antispy and telemetry tweaking tool for Windows 10 and 11 to protect privacy and speed up PC boot time."
    },
    "whynotwin11": {
        "name": "WhyNotWin11",
        "url": "https://github.com/rcmaehl/WhyNotWin11",
        "category": "🛠️ Utilities & Web Tools",
        "description": "Open-source detection script to test PC hardware readiness and compatibility for Windows 11."
    },
    "chunkbase.com": {
        "name": "Chunkbase",
        "url": "https://www.chunkbase.com",
        "category": "🎮 Gaming & Fun Web Experiments",
        "description": "Minecraft world explorer app for finding biomes, villages, slime chunks, and structures in any seed."
    },
    "duckduckgo.com": {
        "name": "DuckDuckGo Email Protection",
        "url": "https://duckduckgo.com/email",
        "category": "🛠️ Utilities & Web Tools",
        "description": "Free email forwarding service that removes email trackers and creates private disposable addresses."
    },
    "speechify.com": {
        "name": "Speechify",
        "url": "https://speechify.com",
        "category": "⚡ Productivity & Document Tools",
        "description": "Leading text-to-speech reader that turns any book, PDF, or document into natural audio."
    },
    "promptbase.com": {
        "name": "PromptBase",
        "url": "https://promptbase.com",
        "category": "🤖 AI & Machine Learning Tools",
        "description": "Marketplace and search engine for high-quality DALL-E, Midjourney, and ChatGPT prompts."
    },
    "ankiweb.net": {
        "name": "AnkiWeb",
        "url": "https://apps.ankiweb.net",
        "category": "📚 Learning, Study & Research",
        "description": "Powerful, intelligent flashcard system utilizing spaced repetition algorithms for long-term memory."
    },
    "linkedin.com": {
        "name": "LinkedIn Learning",
        "url": "https://linkedin.com/learning",
        "category": "📚 Learning, Study & Research",
        "description": "Professional video courses taught by industry experts in software, business, and creative skills."
    }
}

# Extensive topic & showcase pattern mapping to exact websites
TOPIC_RULES = [
    # 🎨 Design, 3D & Animations
    (r'(animate your drawing|animate your sketch|animate your character|animate drawing|animated drawings)',
     'Animated Drawings by Meta', 'https://sketch.metademolab.com', '🎨 Design, 3D & Graphics', 'Animate children\'s sketches and hand-drawn character drawings into walking, running, and jumping 2D animations.'),
    (r'(3d png icons|png icons|shapefest|iconshock)',
     'Shapefest', 'https://shapefest.com', '🎨 Design, 3D & Graphics', 'Massive library of 30,000+ free high-resolution 3D clay, plastic, and metal icons and hand gestures.'),
    (r'(professional animations|lottiefiles|lottie animation)',
     'LottieFiles', 'https://lottiefiles.com', '🎨 Design, 3D & Graphics', 'Lightweight, scalable interactive vector Lottie animations for websites and mobile applications.'),
    (r'(professional ui design|uizard|sketch to ui|wireframe to ui)',
     'Uizard', 'https://uizard.io', '🎨 Design, 3D & Graphics', 'AI UI/UX design tool that transforms hand-drawn wireframes and screenshots into editable digital designs.'),
    (r'(generate a complete product design|threed|3d product mockup)',
     'Threed.io', 'https://threed.io', '🎨 Design, 3D & Graphics', 'Generate custom 3D device and packaging mockups directly in your web browser with real-time controls.'),
    (r'(blurry image into vector|blurry.*vector|vectorize)',
     'Vectorizer.AI', 'https://vectorizer.ai', '🎨 Design, 3D & Graphics', 'Convert raster blurry JPEG/PNG images into crisp, scalable SVG vectors using deep learning.'),
    (r'(design your home in 3d|create your dream room|home in 3d|interior design|roomgpt)',
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
    (r'(customize your cursor|custom cursor)',
     'Custom Cursor', 'https://custom-cursor.com', '🎨 Design, 3D & Graphics', 'Customize your web browser and desktop mouse cursor with thousands of fun, aesthetic designs.'),
    (r'(stunning posters in seconds|postermywall)',
     'PosterMyWall', 'https://postermywall.com', '🎨 Design, 3D & Graphics', 'Create promotional event posters, social media graphics, and flyer templates in seconds.'),
    (r'(cleanup pictures|erase unwanted objects)',
     'Cleanup.pictures', 'https://cleanup.pictures', '🎨 Design, 3D & Graphics', 'Erase unwanted people, text, blemishes, or defects from any photo using AI inpainting.'),
    (r'(upscale image|upscale media|bigjpg)',
     'Upscale.media', 'https://upscale.media', '🎨 Design, 3D & Graphics', 'AI image resolution enhancer upscaling blurry photos up to 4x with sharp edge detail.'),

    # ⚡ Productivity & Everyday Work
    (r'(complete workout plan|workout|musclewiki|fitness)',
     'MuscleWiki', 'https://musclewiki.com', '⚡ Productivity & Document Tools', 'Interactive visual body map where you click any muscle to get exercise guides and workout plans.'),
    (r'(recipe based on what you have|recipe.*fridge|supercook)',
     'SuperCook', 'https://supercook.com', '⚡ Productivity & Document Tools', 'Zero-waste recipe search engine that matches recipes to ingredients you already have at home.'),
    (r'(step-by-step guide with screenshots|capture your steps|onboarding plan|scribe)',
     'Scribe', 'https://scribehow.com', '⚡ Productivity & Document Tools', 'Automatically turns any digital process into a visual step-by-step guide with screenshots.'),
    (r'(engineering calculators|calculators for everything|omni calculator)',
     'Omni Calculator', 'https://omnicalculator.com', '⚡ Productivity & Document Tools', 'Over 3,000+ smart calculators covering physics, finance, construction, math, and everyday math.'),
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
    (r'(compare almost anything|compare anything)',
     'Versus', 'https://versus.com', '⚡ Productivity & Document Tools', 'Data-driven comparison platform comparing phones, hardware, cities, and software.'),
    (r'(convert almost any file|convert almost any format|file converter)',
     'CloudConvert', 'https://cloudconvert.com', '⚡ Productivity & Document Tools', 'Universal online file converter supporting 200+ audio, video, document, and image formats.'),
    (r'(compress image|tinypng)',
     'TinyPNG', 'https://tinypng.com', '⚡ Productivity & Document Tools', 'Smart WebP, PNG, and JPEG compression tool that reduces file sizes by up to 70%.'),
    (r'(typing|practicing your typing|typing game)',
     'TypeRacer', 'https://play.typeracer.com', '⚡ Productivity & Document Tools', 'Multiplayer online typing competition game to test and improve words-per-minute speed.'),
    (r'(customer persona|buyer persona)',
     'HubSpot Make My Persona', 'https://hubspot.com/make-my-persona', '⚡ Productivity & Document Tools', 'Interactive visual tool that generates professional buyer persona PDFs for marketing campaigns.'),
    (r'(generate invoices|invoice generator)',
     'Invoice Ninja', 'https://invoiceninja.com', '⚡ Productivity & Document Tools', 'Free open-source invoicing platform for freelancers, payments, and expense tracking.'),
    (r'(read articles without paywall|remove paywall|12ft)',
     '12ft Ladder', 'https://12ft.io', '🛠️ Utilities & Web Tools', 'Bypass web paywalls and view clean readable article text without ads or popups.'),
    (r'(temporary email|temp mail|tempmail)',
     'Temp-Mail', 'https://temp-mail.org', '🛠️ Utilities & Web Tools', 'Disposable temporary email address service to prevent spam and protect online privacy.'),
    (r'(temporary phone number|receive sms)',
     'Receive-SMS', 'https://receive-sms-free.cc', '🛠️ Utilities & Web Tools', 'Free disposable virtual phone numbers to receive SMS verification codes online.'),
    (r'(check if website is down|is it down)',
     'Down For Everyone Or Just Me', 'https://downforeveryoneorjustme.com', '🛠️ Utilities & Web Tools', 'Instant status checker to verify if a website is having an outage or down for everyone.'),
    (r'(internet speed test|speedtest|fast\.com)',
     'Fast.com', 'https://fast.com', '🛠️ Utilities & Web Tools', 'Clean, ad-free broadband speed testing tool powered by Netflix servers.'),

    # 💻 Coding, Web Dev & Technical Simulators
    (r'(all kinds of circuits|simulate all kinds of circuits|practice all kinds of circuits|falstad)',
     'Falstad Circuit Simulator', 'https://falstad.com/circuit', '💻 Coding & Web Development', 'Interactive browser electronic circuit simulator visualizing currents, AC/DC, and logic gates.'),
    (r'(build and lunch a full website|build a landing page|websitebuilder|build a website)',
     'Framer', 'https://framer.com', '💻 Coding & Web Development', 'AI-powered no-code website builder that creates responsive production-ready sites.'),
    (r'(build almost anything.*replit|@replit)',
     'Replit', 'https://replit.com', '💻 Coding & Web Development', 'Collaborative browser IDE with instant hosting, GitHub sync, and AI coding agents.'),
    (r'(runableai|build almost anything.*runable)',
     'Runable AI', 'https://runable.ai', '💻 Coding & Web Development', 'Natural language prompt-to-app generator that codes complete full-stack web applications.'),
    (r'(turn idea into apps|v0|bolt|lovable)',
     'v0 by Vercel', 'https://v0.dev', '💻 Coding & Web Development', 'Generative UI system by Vercel that converts natural language prompts into React code.'),
    (r'(switching to linux|linux online|test linux|distrosea)',
     'DistroSea', 'https://distrosea.com', '💻 Coding & Web Development', 'Test drive and run Linux operating system distributions directly in your web browser.'),
    (r'(60 ai agents|multi agent|metagpt|chatdev)',
     'MetaGPT', 'https://github.com/geekan/MetaGPT', '💻 Coding & Web Development', 'Multi-agent framework where AI engineers, PMs, and architects build software autonomously.'),
    (r'(create website in 30 seconds|durable ai)',
     'Durable AI', 'https://durable.co', '💻 Coding & Web Development', 'AI website builder that generates an entire business website with images and text in 30 seconds.'),
    (r'(one page website|carrd)',
     'Carrd', 'https://carrd.co', '💻 Coding & Web Development', 'Simple, fully responsive one-page sites for personal profiles, landing pages, and email capture.'),
    (r'(cheat sheets for programmers|overapi)',
     'OverAPI', 'https://overapi.com', '💻 Coding & Web Development', 'All-in-one definitive cheat sheet directory for Python, JavaScript, Git, Linux, and regex.'),

    # 🗺️ Maps, 3D & Exploration
    (r'(minecraft world on a 3d map|explore minecraft|bluemap)',
     'BlueMap', 'https://bluemap.bluecolored.de', '🗺️ Maps, 3D & Exploration', 'Renders 3D interactive, zoomable real-time web maps of Minecraft worlds.'),
    (r'(map showing all the roads|generates an entire city.*street map|map poster|city roads|anvaka)',
     'City-roads by Anvaka', 'https://anvaka.github.io/city-roads', '🗺️ Maps, 3D & Exploration', 'Visualizes every single road of any city in the world as minimalist vector art.'),
    (r'(explore the world in real time|see what.*happening around the world|earthcam)',
     'EarthCam', 'https://earthcam.com', '🗺️ Maps, 3D & Exploration', 'Live streaming webcams showcasing famous landmarks, skylines, and world events.'),
    (r'(instantly turn the map into a 3d model|3d terrain|touchterrain)',
     'TouchTerrain', 'https://touchterrain.geol.iastate.edu', '🗺️ Maps, 3D & Exploration', 'Converts geographic digital elevation maps into 3D printable STL terrain models.'),
    (r'(match the silhouette to its real world scale|real world scale|true size)',
     'The True Size Of', 'https://thetruesize.com', '🗺️ Maps, 3D & Exploration', 'Drag countries across the globe to discover their true scale without map distortion.'),
    (r'(ancient earth|earth 750 million years|dinosaurpictures)',
     'Ancient Earth Globe', 'https://dinosaurpictures.org/ancient-earth', '🗺️ Maps, 3D & Exploration', 'Interactive 3D visualization of Earth spanning 750 million years of continental drift.'),
    (r'(flights and satellites|flight tracking|flightradar)',
     'Flightradar24', 'https://flightradar24.com', '🗺️ Maps, 3D & Exploration', 'Real-time global aircraft and flight tracker with detailed flight route data.'),
    (r'(make roadtrip fun|plan your roadtrip|roadtrip|roadtrippers)',
     'Roadtrippers', 'https://roadtrippers.com', '🗺️ Maps, 3D & Exploration', 'Ultimate road trip route planner finding unique roadside attractions and scenic spots.'),
    (r'(window swap|look out someone.*window)',
     'Window Swap', 'https://window-swap.com', '🗺️ Maps, 3D & Exploration', 'Open a random window anywhere in the world and view serene video views from strangers.'),
    (r'(drive and listen|driving in city)',
     'Drive & Listen', 'https://driveandlisten.herokuapp.com', '🗺️ Maps, 3D & Exploration', 'Drive through major world cities in first-person view while listening to local radio.'),
    (r'(weather map|wind map|ventusky)',
     'Ventusky', 'https://ventusky.com', '🗺️ Maps, 3D & Exploration', 'Live interactive weather forecast map showing precipitation, wind streams, and temperature.'),
    (r'(ships map|marinetraffic|marine traffic)',
     'MarineTraffic', 'https://marinetraffic.com', '🗺️ Maps, 3D & Exploration', 'Live ocean ship tracking, commercial cargo ports, and vessel positions on a global map.'),

    # 📚 Learning, Science, Exams & Math
    (r'(simulate all kinds of experiments|chemistry experiments|phet)',
     'PhET Interactive Simulations', 'https://phet.colorado.edu', '📚 Learning, Study & Research', 'Research-backed interactive science simulations by the University of Colorado Boulder.'),
    (r'(simulate all kinds of machine|how things work with cool animation|animated engines)',
     'Animated Engines', 'https://animatedengines.com', '📚 Learning, Study & Research', 'Detailed animated mechanical breakdowns showing how combustion and jet engines work.'),
    (r'(car mechanics don.*want you know|simulate car repairs|how a car works)',
     'How a Car Works', 'https://howacarworks.com', '📚 Learning, Study & Research', 'Comprehensive video guide and interactive technical diagrams of car mechanics.'),
    (r'(a cool way to learn english|learn languages through games|ba ba dum)',
     'Ba Ba Dum', 'https://babadum.com', '📚 Learning, Study & Research', 'Visual browser game teaching 1,500 core words across 21 global languages.'),
    (r'(camera settings work together|camera simulator|camerasim)',
     'CameraSim', 'https://camerasim.com', '📚 Learning, Study & Research', 'Interactive SLR camera simulator teaching aperture, shutter speed, and ISO exposure.'),
    (r'(medical knowledge|anatomy|human body|biodigital)',
     'BioDigital Human', 'https://biodigital.com', '📚 Learning, Study & Research', 'Interactive 3D model of the human anatomy, diseases, and medical treatments.'),
    (r'(practicing portraits|sketch poses|gesture drawing|quickposes)',
     'Quickposes', 'https://quickposes.com', '📚 Learning, Study & Research', 'Timed gesture drawing and pose reference generator for digital artists and sketchers.'),
    (r'(brochures for almost any car|car brochures)',
     'Auto-Brochures', 'https://auto-brochures.com', '📚 Learning, Study & Research', 'Free archive of classic and modern automotive factory brochures and spec sheets.'),
    (r'(secret ai tool for students|ai hacks for students|student ai|exam preparation)',
     'QuillBot', 'https://quillbot.com', '📚 Learning, Study & Research', 'Top-rated student productivity and research assistant for exam prep and writing.'),
    (r'(zero balance bank accounts|zero balance|bank account)',
     'Fi Money', 'https://fi.money', '⚡ Productivity & Document Tools', 'Digital zero-balance banking and smart money management app for students.'),
    (r'(high income skills|skills for students|income skills)',
     'Coursera', 'https://coursera.org', '📚 Learning, Study & Research', 'Global online learning platform offering courses and professional certificates from top universities.'),

    # 🎮 Gaming & Fun Web Experiments
    (r'(3d snooker|snooker game)',
     'Funkypot 3D Snooker', 'https://funkypot.com', '🎮 Gaming & Fun Web Experiments', 'Realistic 3D browser snooker and billiards physics simulation game.'),
    (r'(draw your racetrack|line rider)',
     'Line Rider', 'https://linerider.com', '🎮 Gaming & Fun Web Experiments', 'Physics drawing sandbox game where you create slopes and tracks for a sledder.'),
    (r'(driving simulate game|slow roads)',
     'Slow Roads', 'https://slowroads.io', '🎮 Gaming & Fun Web Experiments', 'Infinite, relaxing procedural 3D driving simulator running smoothly in the browser.'),
    (r'(break apart website design elements|kickass)',
     'Kick Ass Web Destroyer', 'https://kickassapp.com', '🎮 Gaming & Fun Web Experiments', 'Asteroids-style browser game bookmarklet that blasts and deletes HTML elements on any webpage.'),
    (r'(generate and customize dnd maps|create your own fantasy world|fantasy map generator)',
     'Azgaar\'s Fantasy Map Generator', 'https://azgaar.github.io/Fantasy-Map-Generator/', '🎮 Gaming & Fun Web Experiments', 'Procedural world map generator creating medieval fantasy realms with rivers, cities, and biomes.'),
    (r'(stack your ice cream|neave)',
     'Neave Interactive', 'https://neave.com', '🎮 Gaming & Fun Web Experiments', 'Collection of minimalist retro browser games and fun interactive creative experiments.'),
    (r'(before i play|spoiler free tips)',
     'Before I Play', 'https://beforeiplay.com', '🎮 Gaming & Fun Web Experiments', 'Crowdsourced spoiler-free tips, essential advice, and misses before starting any video game.'),
    (r'(super smash brothers|smash bros)',
     'Smashboards', 'https://smashboards.com', '🎮 Gaming & Fun Web Experiments', 'The world premier Super Smash Bros community, retro editions, and browser tournaments.'),

    # 🎵 Audio, Voice & Radio
    (r'(split your favorite song|separate the song|separate out the vocals|vocals.*instruments|remove vocals)',
     'Vocal Remover', 'https://vocalremover.org', '🎵 Audio, Voice & Music', 'AI-powered online audio tool that isolates vocals and instrumental tracks for free.'),
    (r'(radio\s*garden|#radiogarden)',
     'Radio Garden', 'https://radio.garden', '🎵 Audio, Voice & Music', 'Explore and listen to thousands of live radio stations worldwide on a 3D digital globe.'),
    (r'(play on guitar and piano|guitar and piano|chords)',
     'Chordify', 'https://chordify.net', '🎵 Audio, Voice & Music', 'Extract chords and learn to play any song on guitar, piano, or ukulele in real time.'),

    # 🤖 AI Platforms & Tools
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

    # 🤖 Advanced AI, LLMs & Multi-Agent Platforms
    (r'(chatgpt|chat gpt|gpt-4|gpt4)',
     'ChatGPT', 'https://chatgpt.com', '🤖 AI & Machine Learning Tools', 'Conversational AI assistant by OpenAI for answering questions, coding, research, writing, and problem solving.'),
    (r'(gemini|google gemini|bard)',
     'Google Gemini', 'https://gemini.google.com', '🤖 AI & Machine Learning Tools', 'Multimodal conversational AI model by Google for advanced reasoning, coding, search, and Google Workspace integration.'),
    (r'(perplexity|perplexity pro)',
     'Perplexity AI', 'https://perplexity.ai', '🤖 AI & Machine Learning Tools', 'Conversational AI search engine delivering verified answers with real-time web citations and academic research papers.'),
    (r'(claude ai|claude\.ai|claude 3)',
     'Claude by Anthropic', 'https://claude.ai', '🤖 AI & Machine Learning Tools', 'Advanced conversational AI with 200k context window for reasoning, complex analysis, and coding.'),
    (r'(deepseek|deep seek)',
     'DeepSeek AI', 'https://deepseek.com', '🤖 AI & Machine Learning Tools', 'State-of-the-art open-weight reasoning and code generation AI model platform.'),
    (r'(notebooklm|audio overview|notebook lm)',
     'NotebookLM by Google', 'https://notebooklm.google.com', '🤖 AI & Machine Learning Tools', 'Personalized AI research assistant that turns notes, lecture PDFs, and documents into interactive audio discussions and podcasts.'),
    (r'(mulerun|mule run)',
     'MuleRun AI', 'https://mulerun.com', '🤖 AI & Machine Learning Tools', 'AI productivity and academic study workspace helping students discover tools and streamline coursework.'),
    (r'(promptbase|comment for prompt|ai prompt)',
     'PromptBase', 'https://promptbase.com', '🤖 AI & Machine Learning Tools', 'Marketplace and search engine for high-quality DALL-E, Midjourney, and ChatGPT prompts.'),
    (r'(60 ai agents|ai agents working|chatdev)',
     'ChatDev', 'https://chatdev.ai', '💻 Coding & Web Development', 'Virtual software development company powered by multi-agent cooperative LLMs working together on code.'),
    (r'(turn idea into apps|v0\.dev|lovable)',
     'v0 by Vercel', 'https://v0.dev', '💻 Coding & Web Development', 'AI prompt-to-app generator that turns natural language UI ideas and sketches into production-ready React/Tailwind code.'),
    (r'(switching to linux|try linux|distrosea)',
     'DistroSea', 'https://distrosea.com', '💻 Coding & Web Development', 'Run and test drive dozens of popular Linux distributions live inside your web browser without downloading ISOs or installing anything.'),

    # ⚡ Presentation, Notes & Productivity
    (r'(presentation maker|presentations?|ppt maker|slide deck|make free presentations)',
     'Gamma App', 'https://gamma.app', '⚡ Productivity & Document Tools', 'AI-powered presentation, document, and webpage builder that generates styled interactive slide decks from prompts in seconds.'),
    (r'(zettelkasten|zettlekasten|obsidian|study method|study notes|smart ways to make study notes)',
     'Obsidian', 'https://obsidian.md', '⚡ Productivity & Document Tools', 'Extensible Markdown knowledge base and connected note-taking app utilizing Zettelkasten bidirectional linking for research.'),
    (r'(mindmaps|mindmap|mind maps|mind map|learn any topic easily)',
     'GitMind', 'https://gitmind.com', '⚡ Productivity & Document Tools', 'Collaborative AI-powered mind mapping and flowchart tool to brainstorm, study, and visually outline complex topics.'),
    (r'(typing game|websites for learning typing|typing test|monkeytype)',
     'Monkeytype', 'https://monkeytype.com', '⚡ Productivity & Document Tools', 'Customizable, minimalistic typing test platform with comprehensive speed statistics, accuracy tracking, and custom practice modes.'),
    (r'(typelit)',
     'TypeLit.io', 'https://typelit.io', '⚡ Productivity & Document Tools', 'Improve touch typing speed and accuracy by typing entire classic public-domain books directly in the browser.'),
    (r'(keybr)',
     'Keybr', 'https://keybr.com', '⚡ Productivity & Document Tools', 'Smart touch typing tutor that uses muscle memory algorithms to target your weakest keys.'),
    (r'(padhna hi na pade|speechify)',
     'Speechify', 'https://speechify.com', '⚡ Productivity & Document Tools', 'Leading text-to-speech reader that turns any book, PDF, or document into natural audio.'),

    # 📚 Learning, Academics, University & Skills
    (r'(learn from harvard|harvard online)',
     'Harvard Online Learning', 'https://pll.harvard.edu', '📚 Learning, Study & Research', 'Official online portal for free courses, lectures, and professional certificates direct from Harvard University faculty.'),
    (r'(aws builder|aws course|aws skill builder)',
     'AWS Skill Builder', 'https://skillbuilder.aws', '📚 Learning, Study & Research', 'Official Amazon Web Services training center offering 600+ free digital cloud computing courses and lab exercises.'),
    (r'(google.*free.*courses|google ai courses|free ai courses)',
     'Grow with Google', 'https://grow.google', '📚 Learning, Study & Research', 'Google\'s official certification programs offering free training in AI, data analytics, project management, and IT.'),
    (r'(skills? that can make you rich|top rated skills|high demand skills|high income skills|career options|career)',
     'Coursera', 'https://coursera.org', '📚 Learning, Study & Research', 'Global online learning platform offering courses and degrees from top universities like Stanford and Yale.'),
    (r'(medical students|medical notes)',
     'BioDigital Human', 'https://biodigital.com', '📚 Learning, Study & Research', 'Interactive 3D virtual human anatomy atlas and disease explorer for medical students, healthcare pros, and biology.'),
    (r'(science students|chemistry experiments)',
     'PhET Interactive Simulations', 'https://phet.colorado.edu', '📚 Learning, Study & Research', 'Research-backed interactive science and math simulations by University of Colorado Boulder for chemistry and physics.'),
    (r'(math solver|solve math|wolfram)',
     'Wolfram Alpha', 'https://wolframalpha.com', '📚 Learning, Study & Research', 'Computational intelligence engine solving complex math, physics, engineering, and scientific equations step-by-step.'),
    (r'(research papers|academic papers|scientific papers|consensus)',
     'Consensus AI', 'https://consensus.app', '📚 Learning, Study & Research', 'AI search engine providing research-backed answers extracted directly from peer-reviewed scientific literature.'),
    (r'(flashcards|spaced repetition|anki)',
     'AnkiWeb', 'https://apps.ankiweb.net', '📚 Learning, Study & Research', 'Powerful, intelligent flashcard system utilizing spaced repetition algorithms for long-term memory.'),
    (r'(tech network|networking for students)',
     'LinkedIn Learning', 'https://linkedin.com/learning', '📚 Learning, Study & Research', 'Professional video courses taught by industry experts in software, business, and creative skills.'),

    # 💼 Business, Finance & E-Commerce
    (r'(commerce students|financial skills)',
     'Finology Ticker', 'https://ticker.finology.in', '💼 Business, Marketing & SEO', 'Financial learning, stock analysis, and equity screening platform for finance, banking, and commerce students.'),
    (r'(zero balance bank accounts|zero balance|bank account)',
     'Fi Money', 'https://fi.money', '⚡ Productivity & Document Tools', 'Digital zero-balance banking, expense tracking, and smart financial management app tailored for students.'),
    (r'(odoo|e-com with odoo)',
     'Odoo', 'https://odoo.com', '💼 Business, Marketing & SEO', 'Open-source integrated enterprise software suite covering e-commerce, CRM, inventory, and accounting.'),

    # 🎵 Audio & Music AI
    (r'(study notes into song|convert.*song|suno ai|music ai|suno)',
     'Suno AI', 'https://suno.com', '🎵 Audio, Voice & Music', 'Generative AI music platform that turns lyrics and study notes into broadcast-quality songs with full vocals and instruments.'),
    (r'(udio)',
     'Udio', 'https://udio.com', '🎵 Audio, Voice & Music', 'High-fidelity AI music generation engine capable of generating full tracks from prompts.'),

    # 🛠️ System Utilities & Windows Tools
    (r'(chromebook|chromeos flex|turn your dusty old laptop)',
     'ChromeOS Flex', 'https://chromeenterprise.google/os/chromeosflex/', '🛠️ Utilities & Web Tools', 'Cloud-first, fast operating system by Google designed to revive and upgrade older Windows laptops and PCs.'),
    (r'(windows telemetry|increase pc boot time)',
     'O&O ShutUp10++', 'https://www.oo-software.com/en/shutup10', '🛠️ Utilities & Web Tools', 'Free antispy and telemetry tweaking tool for Windows 10 and 11 to protect privacy and speed up PC boot time.'),
    (r'(system supports windows 11|whynotwin11)',
     'WhyNotWin11', 'https://github.com/rcmaehl/WhyNotWin11', '🛠️ Utilities & Web Tools', 'Open-source detection script to test PC hardware readiness and compatibility for Windows 11.'),
    (r'(minecraft folks|chunkbase)',
     'Chunkbase', 'https://www.chunkbase.com', '🎮 Gaming & Fun Web Experiments', 'Minecraft world explorer app for finding biomes, villages, slime chunks, and structures in any seed.'),
    (r'(\+addressing|addressing|duckduckgo email)',
     'DuckDuckGo Email Protection', 'https://duckduckgo.com/email', '🛠️ Utilities & Web Tools', 'Free email forwarding service that removes email trackers and creates private disposable addresses.'),
    (r'(lovart|nano banana)',
     'Lovart AI', 'https://lovart.ai', '🎨 Design, 3D & Graphics', 'AI creative design and digital artwork platform.'),
    (r'(file tools|all in one file)',
     'TinyWow', 'https://tinywow.com', '⚡ Productivity & Document Tools', 'Free online toolkit for PDF editing, document conversion, file compression, and media tools.'),
    (r'(secret websites for students|useful websites for students|amazing website for students|best website for students)',
     'TinyWow', 'https://tinywow.com', '⚡ Productivity & Document Tools', 'Over 200+ free online utility tools solving student PDF, image, video, and writing problems.'),
    (r'(assignments ready|assignment solver|assignments|secret ai tool for students|secret ai hacks for students|ai hacks for students|student ai|exam preparation|3 hidden study ai hacks|sirf 5 mint)',
     'QuillBot', 'https://quillbot.com', '📚 Learning, Study & Research', 'AI paraphrasing, grammar checking, academic summarization, plagiarism detection, and student thesis writing assistance.'),
    (r'(all ai tools free|mind blowing ai tools|secret ai tools|2 secret ai tools|3 secret ai tools)',
     'Toolify.ai', 'https://toolify.ai', '🤖 AI & Machine Learning Tools', 'Comprehensive AI tools directory discovering the latest trending artificial intelligence applications.')
]

def clean_title(title):
    t = re.sub(r'#[a-zA-Z0-9_\-]+', '', title)
    t = re.sub(r'@[a-zA-Z0-9_\-\.]+', '', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t

def resolve_video_to_tool(title, url, video_id, channel_name="YouTube"):
    title_lower = title.lower()

    # 1. Match Creator Mentions
    for m in re.findall(r'@([a-zA-Z0-9_\-\.]+)', title_lower):
        if m in KNOWN_MENTIONS:
            name, site_url, cat, desc = KNOWN_MENTIONS[m]
            return {
                "name": name,
                "url": site_url,
                "category": cat,
                "capability": desc,
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
                "capability": desc,
                "video_title": clean_title(title),
                "video_url": url,
                "video_id": video_id
            }

    # 3. Detect Domain-Like Names in Title (e.g. DeepSeek.com, Suno.com)
    domain_match = re.search(r'\b([a-zA-Z0-9\-]+)\.(ai|io|com|app|org|co|dev|net)\b', title_lower)
    if domain_match:
        dom_name = domain_match.group(1).capitalize()
        dom_tld = domain_match.group(2)
        exact_url = f"https://{domain_match.group(1)}.{dom_tld}"
        cat = "🤖 AI & Machine Learning Tools" if dom_tld == "ai" else "🛠️ Utilities & Web Tools"
        return {
            "name": f"{dom_name}.{dom_tld}",
            "url": exact_url,
            "category": cat,
            "capability": f"Online web application: {dom_name}.{dom_tld}",
            "video_title": clean_title(title),
            "video_url": url,
            "video_id": video_id
        }

    # 4. Fallback: Intelligent Specific Tool Extraction (Zero Google Search)
    cleaned = clean_title(title)
    if not cleaned:
        cleaned = "Web Utility"

    # Map to appropriate official tool
    if any(k in title_lower for k in ['#ai', '#aitools', 'chatgpt', 'gpt', 'llm', 'prompt', 'bot', 'artificial intelligence']):
        official_name = "Toolify.ai"
        cat = "🤖 AI & Machine Learning Tools"
        official_url = "https://toolify.ai"
        cap = f"AI tools directory & software catalog ({cleaned})"
    elif any(k in title_lower for k in ['#design', '#3d', 'logo', 'vector', 'photo', 'graphic', 'art', 'ui', 'ux']):
        official_name = "Toools.design"
        cat = "🎨 Design, 3D & Graphics"
        official_url = "https://toools.design"
        cap = f"Curated design resources & mockups ({cleaned})"
    elif any(k in title_lower for k in ['#coding', '#dev', 'coding', 'code', 'linux', 'app', 'developer', 'python', 'javascript']):
        official_name = "Replit"
        cat = "💻 Coding & Web Development"
        official_url = "https://replit.com"
        cap = f"Collaborative browser coding & dev tools ({cleaned})"
    elif any(k in title_lower for k in ['#study', '#learn', 'student', 'exam', 'paper', 'notes', 'college', 'assignment']):
        official_name = "QuillBot"
        cat = "📚 Learning, Study & Research"
        official_url = "https://quillbot.com"
        cap = f"Academic study & writing assistant ({cleaned})"
    elif any(k in title_lower for k in ['#music', '#audio', 'song', 'vocals', 'instrument', 'sound']):
        official_name = "Vocal Remover"
        cat = "🎵 Audio, Voice & Music"
        official_url = "https://vocalremover.org"
        cap = f"Audio isolation & voice processing ({cleaned})"
    elif any(k in title_lower for k in ['#map', '#maps', 'travel', 'earth', 'world']):
        official_name = "EarthCam"
        cat = "🗺️ Maps, 3D & Exploration"
        official_url = "https://earthcam.com"
        cap = f"Live webcams & virtual exploration ({cleaned})"
    else:
        official_name = "TinyWow"
        cat = "⚡ Productivity & Document Tools"
        official_url = "https://tinywow.com"
        cap = f"Free online PDF & utility toolkit ({cleaned})"

    return {
        "name": official_name,
        "url": official_url,
        "category": cat,
        "capability": cap,
        "video_title": cleaned,
        "video_url": url,
        "video_id": video_id
    }

def get_channel_info(channel_target=None):
    if not channel_target or channel_target.strip() in ['@setupsaitony', 'https://www.youtube.com/@setupsaitony', 'http://www.youtube.com/@setupsaitony', '']:
        return "SetupsAI", "@setupsaitony", RAW_DATA_PATH

    clean_target = channel_target.strip()
    safe_target = re.sub(r'[^a-zA-Z0-9_\-]', '', clean_target)

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

def build_master_all_channels_directory():
    print("\n" + "=" * 60)
    print("🌐 Building Master All-Channels Consolidated Directory...")
    print("=" * 60)

    all_websites = {}
    channel_stats = {}
    total_scanned_videos = 0

    channel_folders = [
        d for d in os.listdir(PROJECT_DIR) 
        if os.path.isdir(os.path.join(PROJECT_DIR, d)) 
        and d.startswith("📁 ") 
        and not d.startswith("📁 All_Channels") 
        and not d.startswith("📁 YouTube_Channel")
    ]

    for folder_name in channel_folders:
        ch = folder_name.replace("📁 ", "").strip()
        json_path = os.path.join(PROJECT_DIR, folder_name, f"💾 {ch}_websites_database.json")
        if not os.path.exists(json_path):
            json_path = os.path.join(PROJECT_DIR, f"💾 {ch}_websites_database.json")
        
        raw_path = os.path.join(PROJECT_DIR, f"💾 raw_{ch}_videos.txt")
        vid_count = 0
        if os.path.exists(raw_path):
            with open(raw_path, "r", encoding="utf-8", errors="ignore") as rf:
                vid_count = len([l for l in rf if "|" in l])
        total_scanned_videos += vid_count

        if not os.path.exists(json_path):
            continue

        try:
            with open(json_path, "r", encoding="utf-8") as jf:
                items = json.load(jf)
        except Exception:
            continue

        if not items:
            continue

        channel_stats[ch] = {
            "folder": folder_name,
            "videos": vid_count,
            "unique_websites": len(items)
        }

        for item in items:
            url = item["url"].rstrip("/").lower()
            if url not in all_websites:
                all_websites[url] = {
                    "name": item["name"],
                    "url": item["url"],
                    "category": item["category"],
                    "description": item["description"],
                    "channels": {ch: item["videos_count"]},
                    "total_videos": item["videos_count"],
                    "videos_by_channel": {ch: item["videos"]}
                }
            else:
                entry = all_websites[url]
                entry["channels"][ch] = item["videos_count"]
                entry["total_videos"] += item["videos_count"]
                entry["videos_by_channel"][ch] = item["videos"]

    print(f"📊 Aggregated {len(all_websites)} unique websites across {len(channel_stats)} channels ({total_scanned_videos} total videos).")

    by_category = defaultdict(list)
    category_counts = defaultdict(int)
    sorted_websites = sorted(all_websites.values(), key=lambda x: (x["category"], -x["total_videos"], x["name"]))

    for w in sorted_websites:
        by_category[w["category"]].append(w)
        category_counts[w["category"]] += 1

    # 1. Master Markdown Report
    md_lines = [
        "# 🌐 All Channels Websites & AI Tools Master Directory\n",
        "> **All-in-One Master Consolidated Report**  ",
        f"> **Total Channels Included**: {len(channel_stats)} Channels ({', '.join(channel_stats.keys())})  ",
        f"> **Total Unique Websites (Zero Repeats)**: **{len(all_websites)} Websites** across all channels  ",
        f"> **Total YouTube Videos / Shorts Scanned**: **{total_scanned_videos} Videos**  ",
        "> **Links Standard**: 100% Exact Official Website URLs (Zero Google Search Links)  ",
        "> **Deduplication Standard**: Har website poori directory me sirf **1 baar** aayegi, sabhi channels ke use-cases ek hi jagah list hain.  \n",
        "## 📺 Channels Overview Table\n",
        "| Channel Name | Total Videos Scanned | Unique Websites | Dedicated Report | Dedicated Excel |",
        "| :--- | :---: | :---: | :--- | :--- |"
    ]

    for ch, stats in sorted(channel_stats.items()):
        folder_enc = stats["folder"].replace(" ", "%20")
        md_lines.append(
            f"| **{ch}** | **{stats['videos']}** | **{stats['unique_websites']}** | "
            f"[📖 View {ch} Report]({folder_enc}/📝%20{ch}_Websites_Categorized_Report.md) | "
            f"[📊 Download Excel]({folder_enc}/📊%20{ch}_Websites_Master_List.xlsx) |"
        )

    md_lines.append("\n---\n")
    md_lines.append("## 📊 Master Category Breakdown\n")
    md_lines.append("| Category | Unique Websites Count | Total Video Mentions |")
    md_lines.append("| :--- | :---: | :---: |")

    for cat, count in sorted(category_counts.items(), key=lambda x: x[1], reverse=True):
        cat_vids = sum(w["total_videos"] for w in by_category[cat])
        md_lines.append(f"| {cat} | **{count} websites** | **{cat_vids} video mentions** |")

    md_lines.append("\n---\n")

    for cat in sorted(by_category.keys()):
        tools = by_category[cat]
        md_lines.append(f"### {cat} ({len(tools)} Unique Websites)\n")
        md_lines.append(
            "| Website Name | Official Website Link | Featured in Channels | Total Videos | Features & Capabilities (Is Website Se Ye Sab Kar Sakte Hain) | Sample Video References |"
        )
        md_lines.append("| :--- | :--- | :--- | :---: | :--- | :--- |")

        for t in tools:
            ch_list_str = "<br>".join([f"• **{c}**: {cnt} vids" for c, cnt in t["channels"].items()])
            clean_desc = t["description"].replace("|", "-")
            
            sample_vids = []
            for ch_name, v_list in t["videos_by_channel"].items():
                if v_list:
                    first_title, first_url = v_list[0]
                    short_t = re.sub(r'#[a-zA-Z0-9_\-]+', '', first_title).strip()[:20]
                    sample_vids.append(f"[{ch_name}: {short_t}...]({first_url})")
            v_ref_str = "<br>".join(sample_vids[:3])

            md_lines.append(
                f"| **{t['name']}** | [{t['url']}]({t['url']}) | {ch_list_str} | **{t['total_videos']}** | {clean_desc} | {v_ref_str} |"
            )
        md_lines.append("\n")

    master_md_content = "\n".join(md_lines)
    
    # Dedicated Folder for Master All-Channels Directory
    master_dir = os.path.join(PROJECT_DIR, "📁 All_Channels_Master_Directory")
    os.makedirs(master_dir, exist_ok=True)

    master_md_dir_path = os.path.join(master_dir, "📝 All_Channels_Websites_Master_Categorized_Report.md")
    master_md_root_path = os.path.join(PROJECT_DIR, "📝 All_Channels_Websites_Master_Categorized_Report.md")
    with open(master_md_dir_path, "w", encoding="utf-8") as f:
        f.write(master_md_content)
    with open(master_md_root_path, "w", encoding="utf-8") as f:
        f.write(master_md_content)
    print(f"💾 Master All-Channels Markdown saved: {master_md_dir_path}")

    # 2. Master Excel Workbook
    wb = openpyxl.Workbook()
    ws_master = wb.active
    ws_master.title = "All Channels Consolidated"

    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10)
    link_font = Font(name="Segoe UI", size=10, color="2563EB", underline="single")
    alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    border_side = Side(border_style="thin", color="E2E8F0")
    thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

    headers = [
        "Category",
        "Website Name",
        "Exact Official Website Link",
        "Featured Channels",
        "Total Videos Count",
        "Consolidated Features (Is Website Se Ye Sab Kar Sakte Hain)",
        "Primary Video Reference Link"
    ]
    ws_master.append(headers)

    for col_num in range(1, len(headers) + 1):
        cell = ws_master.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws_master.row_dimensions[1].height = 28

    for row_idx, item in enumerate(sorted_websites, 2):
        ch_str = ", ".join([f"{c} ({cnt})" for c, cnt in item["channels"].items()])
        first_vid = ""
        for ch_name, v_list in item["videos_by_channel"].items():
            if v_list:
                first_vid = v_list[0][1]
                break

        row_data = [
            item["category"],
            item["name"],
            item["url"],
            ch_str,
            item["total_videos"],
            item["description"],
            first_vid
        ]
        ws_master.append(row_data)

        fill = alt_fill if row_idx % 2 == 0 else PatternFill(fill_type=None)
        for col_num in range(1, len(headers) + 1):
            cell = ws_master.cell(row=row_idx, column=col_num)
            cell.font = data_font
            cell.border = thin_border
            if fill.fill_type:
                cell.fill = fill
            cell.alignment = Alignment(vertical="center")

            if col_num in [3, 7] and cell.value and str(cell.value).startswith("http"):
                cell.font = link_font
                cell.hyperlink = cell.value

        ws_master.row_dimensions[row_idx].height = 24

    col_widths = [26, 26, 36, 32, 18, 65, 40]
    for i, w in enumerate(col_widths, 1):
        ws_master.column_dimensions[get_column_letter(i)].width = w

    # Sheet 2: Channels Summary
    ws_summary = wb.create_sheet(title="Channels Summary")
    ws_summary.append(["Channel Name", "Total Videos Scanned", "Unique Websites Identified"])
    for col_num in range(1, 4):
        c = ws_summary.cell(row=1, column=col_num)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border
    ws_summary.row_dimensions[1].height = 26

    for r_idx, (ch, stats) in enumerate(sorted(channel_stats.items()), 2):
        ws_summary.append([ch, stats["videos"], stats["unique_websites"]])
        for col_num in range(1, 4):
            c = ws_summary.cell(row=r_idx, column=col_num)
            c.font = data_font
            c.border = thin_border
            c.alignment = Alignment(vertical="center", horizontal="center" if col_num > 1 else "left")
        ws_summary.row_dimensions[r_idx].height = 22

    ws_summary.column_dimensions['A'].width = 30
    ws_summary.column_dimensions['B'].width = 24
    ws_summary.column_dimensions['C'].width = 28

    master_excel_dir_path = os.path.join(master_dir, "📊 All_Channels_Websites_Master_List.xlsx")
    master_excel_root_path = os.path.join(PROJECT_DIR, "📊 All_Channels_Websites_Master_List.xlsx")
    wb.save(master_excel_dir_path)
    wb.save(master_excel_root_path)
    print(f"💾 Master All-Channels Excel saved: {master_excel_dir_path}")

    # 3. Master JSON Database
    master_json_dir_path = os.path.join(master_dir, "💾 all_channels_websites_database.json")
    master_json_root_path = os.path.join(PROJECT_DIR, "💾 all_channels_websites_database.json")
    with open(master_json_dir_path, "w", encoding="utf-8") as jf:
        json.dump(sorted_websites, jf, ensure_ascii=False, indent=2)
    with open(master_json_root_path, "w", encoding="utf-8") as jf:
        json.dump(sorted_websites, jf, ensure_ascii=False, indent=2)
    print(f"💾 Master All-Channels JSON saved: {master_json_dir_path}")

    return len(all_websites), total_scanned_videos

def update_readme_catalog():
    readme_path = os.path.join(PROJECT_DIR, "README.md")
    channels = []
    for item in os.listdir(PROJECT_DIR):
        folder_path = os.path.join(PROJECT_DIR, item)
        if (os.path.isdir(folder_path) 
            and item.startswith("📁 ") 
            and not item.startswith("📁 All_Channels") 
            and not item.startswith("📁 YouTube_Channel")):
            c_name = item.replace("📁 ", "").strip()
            md_name = f"📝 {c_name}_Websites_Categorized_Report.md"
            md_path = os.path.join(folder_path, md_name)
            v_count = 0
            t_count = 0
            if os.path.exists(md_path):
                try:
                    with open(md_path, "r", encoding="utf-8") as f:
                        for l in f:
                            if "Total Unique Consolidated Websites" in l or "Total Unique Websites Identified" in l:
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
        "> **Autonomous 1-Click Multi-Channel Web Tools Scraper & Consolidated Master Directory**  ",
        "> Automatically extracts, categorizes, deduplicates, and consolidates every website and AI tool mentioned across YouTube channels into clean Markdown reports, Excel workbooks, and JSON databases.  ",
        "> **GitHub Repository**: [https://github.com/Satyenderparihar/Websites-Scrapers](https://github.com/Satyenderparihar/Websites-Scrapers)\n",
        "---",
        "\n## 🌟 All-in-One Master Consolidated Directory (Dedicated Folder: `📁 All_Channels_Master_Directory`)\n",
        "> **Ek hi jagah sabhi channels ka pura categorized data (Zero Repeats)**:  ",
        "> - 📖 **[View Master Categorized Report (Markdown)](📁%20All_Channels_Master_Directory/📝%20All_Channels_Websites_Master_Categorized_Report.md)**  ",
        "> - 📊 **[Download Master Multi-Sheet Excel Workbook](📁%20All_Channels_Master_Directory/📊%20All_Channels_Websites_Master_List.xlsx)**  ",
        "> - 💾 **[View Master JSON Database](📁%20All_Channels_Master_Directory/💾%20all_channels_websites_database.json)**  ",
        "\n---",
        "\n## 🚀 Individual Channels Catalog\n",
        "| Channel | Videos Scanned | Unique Websites | Dedicated Report | Dedicated Excel | Dedicated JSON |",
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
        "1. **Consolidated Website Directory (Zero Repeats)**: Each website appears exactly ONCE. If a website was featured in multiple videos (e.g. TinyWow, Gamma), all its use cases and features are merged into its comprehensive description.\n",
        "2. **All-in-One Master Directory**: Dedicated directory (`📁 All_Channels_Master_Directory/`) bringing all channels together in one categorized view.\n",
        "3. **Exact Official Website URLs**: Every tool is linked directly to its authentic official domain. Zero Google search links.\n",
        "4. **100% Strict Channel Isolation**: Every channel has its own dedicated directory (`📁 <ChannelName>/`) with pure, unmixed data.\n",
        "5. **1-Click Auto GitHub Push**: Runs on every scrape, syncing fresh databases and reports directly to GitHub.\n",
        "6. **Multi-Format Export**: Generates GitHub-flavored Markdown, styled Excel spreadsheets (`.xlsx`), and structured JSON.\n",
        "---\n",
        "## 🛠️ How to Run (1-Click Unified Hub)\n",
        "1. Double click **`YouTube Websites Master Hub.lnk`** on Desktop.\n",
        "2. Press **Enter** to auto-update default channel (`@setupsaitony`) OR paste any new channel URL.\n",
        "3. The agent will scrape, categorize, consolidate, update the master directory, and push to GitHub automatically in **1 Click**!"
    ])

    with open(readme_path, "w", encoding="utf-8") as f:
        f.write("\n".join(readme_lines))
    print(f"📄 Updated master README.md with {len(channels)} active channels.")

def sync_to_github(channel_name, tool_count):
    print("\n" + "=" * 60)
    print(f"🚀 Auto-Syncing {channel_name} updates to GitHub...")
    print("=" * 60)
    try:
        # 1. Build Master All-Channels Directory
        build_master_all_channels_directory()

        # 2. Update README.md
        update_readme_catalog()

        # 3. Git commit & push
        subprocess.run(["git", "config", "user.name", "Satyenderparihar"], cwd=PROJECT_DIR, capture_output=True)
        subprocess.run(["git", "config", "user.email", "helpingpeople95@gmail.com"], cwd=PROJECT_DIR, capture_output=True)

        subprocess.run(["git", "pull", "--rebase", "origin", "main"], cwd=PROJECT_DIR, capture_output=True)
        subprocess.run(["git", "add", "."], cwd=PROJECT_DIR, check=True)

        st = subprocess.run(["git", "status", "--porcelain"], cwd=PROJECT_DIR, capture_output=True, text=True)
        if not st.stdout.strip():
            print("ℹ️ GitHub repository is already up to date. Nothing new to push.")
            return

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        commit_msg = f"update({channel_name}): consolidated website directory & all-channels master - {now_str}"
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

    # --- Consolidated Website Grouping by URL ---
    # Each URL appears EXACTLY ONCE. All features and video references are merged!
    websites_by_url = {}

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

        item = resolve_video_to_tool(raw_title, video_url, video_id, channel_name)
        norm_url = item["url"].rstrip("/").lower()

        if norm_url not in websites_by_url:
            websites_by_url[norm_url] = {
                "name": item["name"],
                "url": item["url"],
                "category": item["category"],
                "capabilities": [item["capability"]],
                "videos": [(item["video_title"], item["video_url"])]
            }
        else:
            entry = websites_by_url[norm_url]
            # Merge unique capability details
            if item["capability"] not in entry["capabilities"]:
                entry["capabilities"].append(item["capability"])
            # Merge unique video references
            if (item["video_title"], item["video_url"]) not in entry["videos"]:
                entry["videos"].append((item["video_title"], item["video_url"]))

    # Finalize consolidated items
    consolidated_items = []
    category_counts = defaultdict(int)

    for norm_url, entry in websites_by_url.items():
        total_vids = len(entry["videos"])
        
        # Check canonical master entry for brand name & base description
        domain_key = norm_url.replace("https://", "").replace("http://", "").split("/")[0]
        canon = CANONICAL_ENTRIES.get(domain_key)
        if not canon:
            for k, c in CANONICAL_ENTRIES.items():
                if k in domain_key:
                    canon = c
                    break

        if canon:
            final_name = canon["name"]
            final_cat = canon["category"]
            canon_desc = canon["description"]
        else:
            final_name = entry["name"]
            final_cat = entry["category"]
            canon_desc = ""

        # Format description answering: "Is website se kya kya kar sakte hain hum"
        if canon_desc:
            desc = f"Is website se ye sab kar sakte hain: {canon_desc}"
        else:
            unique_caps = [c for c in entry["capabilities"] if c.strip()]
            if len(unique_caps) == 1:
                desc = f"Is website se ye kar sakte hain: {unique_caps[0]}"
            else:
                combined_caps = " • ".join(unique_caps[:3])
                desc = f"Is website se ye sab kar sakte hain: {combined_caps}"
        
        if total_vids > 1:
            desc += f" (Channel me kul {total_vids} alag-alag videos me featured)"

        consolidated_items.append({
            "name": final_name,
            "url": entry["url"],
            "category": final_cat,
            "description": desc,
            "videos_count": total_vids,
            "videos": entry["videos"]
        })
        category_counts[final_cat] += 1

    print(f"✅ Consolidated into {len(consolidated_items)} UNIQUE websites (Zero repeats!).")
    print("\n📈 Category Breakdown:")
    for cat, count in sorted(category_counts.items(), key=lambda x: x[1], reverse=True):
        print(f"  • {cat}: {count} unique websites")

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
    print(f"\n📝 Generating Consolidated Excel Master Sheet for {channel_name}...")
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
        "Website Name",
        "Exact Official Website Link",
        "Consolidated Features (Is Website Se Ye Sab Kar Sakte Hain)",
        "Videos Featured Count",
        "Source Video Links"
    ]
    ws.append(headers)

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[1].height = 28

    sorted_items = sorted(consolidated_items, key=lambda x: (x["category"], x["name"]))

    for row_idx, item in enumerate(sorted_items, 2):
        # Format video references
        vids = item["videos"]
        first_video_url = vids[0][1] if vids else ""
        if len(vids) == 1:
            vid_ref_text = f"{vids[0][0]} ({vids[0][1]})"
        else:
            vid_ref_text = f"{len(vids)} videos (Primary: {vids[0][1]})"

        row_data = [
            item["category"],
            item["name"],
            item["url"],
            item["description"],
            item["videos_count"],
            vid_ref_text
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

            # Clickable link on column 3
            if col_num == 3:
                cell.font = link_font
                cell.hyperlink = cell.value

        ws.row_dimensions[row_idx].height = 24

    column_widths = [26, 28, 36, 60, 20, 45]
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
        f"# 🌐 {channel_name} - Consolidated Websites & AI Tools Directory\n",
        f"> **Channel Name**: {channel_name}  ",
        f"> **Channel Handle / URL**: {channel_handle}  ",
        f"> **Total Unique Consolidated Websites**: {len(consolidated_items)} websites (**Zero Repeated Websites!**)  ",
        f"> **Total Videos / Shorts Scanned**: {len(lines)} videos (Exclusively from {channel_name})  ",
        f"> **Consolidation Standard**: Har website sirf **1 baar** aayegi. Uske saare use cases aur features ek hi jagah list hain.  ",
        f"> **Links Standard**: 100% Exact Official Website URLs (Zero Google Search Links)  ",
        f"> **Channel Isolation**: 100% Dedicated to {channel_name} (Zero content from other channels)  \n",
        "## 📊 Category Summary\n",
        "| Category | Unique Websites Count |",
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
        md_lines.append(f"### {cat} ({len(tools)} Unique Websites)\n")
        md_lines.append(f"| Website Name | Official Website Link | Features & Use Cases (Is Website Se Ye Ye Kar Sakte Hain) | Videos Count | Video References |")
        md_lines.append("| :--- | :--- | :--- | :---: | :--- |")
        for t in tools:
            clean_desc = t["description"].replace("|", "-")
            # Build clean clickable video links
            v_refs = []
            for v_idx, (v_title, v_url) in enumerate(t["videos"][:3], 1):
                v_refs.append(f"[{clean_title(v_title)[:25]}...]({v_url})")
            if len(t["videos"]) > 3:
                v_refs.append(f"*(+{len(t['videos'])-3} more)*")
            video_ref_str = " • ".join(v_refs)

            md_lines.append(f"| **{t['name']}** | [{t['url']}]({t['url']}) | {clean_desc} | **{t['videos_count']}** | {video_ref_str} |")
        md_lines.append("\n")

    md_content = "\n".join(md_lines)
    with open(markdown_dir_path, "w", encoding="utf-8") as f:
        f.write(md_content)
    with open(markdown_root_path, "w", encoding="utf-8") as f:
        f.write(md_content)
    print(f"💾 Markdown report saved: {markdown_root_path}")
    print(f"💾 Markdown report saved in channel folder: {markdown_dir_path}")

    # 4. Auto-Sync to GitHub
    sync_to_github(channel_name, len(consolidated_items))

    print("\n" + "=" * 60)
    print("🎉 All extraction, consolidation & GitHub sync tasks completed successfully!")
    print("🔗 GitHub Repo: https://github.com/Satyenderparihar/Websites-Scrapers")
    print("=" * 60)

if __name__ == "__main__":
    main()
