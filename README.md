# Word Image Fixer Pro

![Word Image Fixer Pro](https://img.shields.io/badge/version-1.0.0-blue)
![Python](https://img.shields.io/badge/Python-3.9+-green)
![FastAPI](https://img.shields.io/badge/FastAPI-0.104.1-009688)
![License](https://img.shields.io/badge/license-MIT-orange)

**Free online tool to fix Word document image issues instantly.**

![Screenshot](https://via.placeholder.com/800x400/3b82f6/ffffff?text=Word+Image+Fixer+Pro)

## 🚀 Features

- ✅ **Fix Cut-off Images** - Auto-correct line spacing to show full images
- ✅ **Smart Resizing** - Scale images to fit page or table width
- ✅ **Fix Position** - Convert floating images to inline
- ✅ **Marker Insertion** - Insert images using filename markers
- ✅ **Auto Page Detection** - Automatically detects document page width
- ✅ **Multi-language** - English, 中文, 한국어, 日本語, Français
- ✅ **Privacy First** - Files deleted immediately after processing

## 📦 Quick Deploy

### Deploy to Vercel (Free - Easiest)

1. Fork this repository to your GitHub
2. Go to [vercel.com](https://vercel.com)
3. Click "Import Project" → Select your forked repo
4. Click "Deploy" - Done!

[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https://github.com/yourusername/word-fixer-pro)

### Deploy to Render (Free)

1. Fork this repository
2. Go to [render.com](https://render.com)
3. Click "New Web Service" → Connect GitHub
4. Select repo and deploy

### Local Development

```bash
# Clone the repository
git clone https://github.com/yourusername/word-fixer-pro.git
cd word-fixer-pro

# Install dependencies
pip install -r requirements.txt

# Run the server
uvicorn main:app --reload

# Open browser
open http://localhost:8000