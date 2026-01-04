# API Key Setup

## 🔑 Grok API Key Configuration

The AI parsing feature requires a Grok API key. For security, the key is **not** included in this repository.

## Setup Instructions

### Option 1: Parent Folder (Recommended)

1. Navigate to the parent folder (one level up from rafas_folder)
2. Create file: `grok_config.txt`
3. Add your API key as the only content (one line)
4. Save

**Example:**
```
If rafas_folder is at: c:\dev\canva\rafas_folder\
Create key file at:    c:\dev\canva\grok_config.txt
```

**Advantages:**
- Shared across all subfolders
- Easy to find and update
- Outside the main package folder

### Option 2: Local Config

1. Copy `grok_config.txt.sample` to `grok_config.txt` (in rafas_folder)
2. Open `grok_config.txt`
3. Replace `YOUR-GROK-API-KEY-HERE` with your actual key
4. Save

**Advantages:**
- Self-contained
- Portable with this folder only

## Getting Your API Key

1. Visit: https://console.x.ai/
2. Sign in / Create account
3. Navigate to API Keys section
4. Create a new API key
5. Copy the key (starts with `xai-`)

## File Format

The config file should contain **only** your API key:

```
xai-YOUR-ACTUAL-API-KEY-HERE
```

**No extra text, no quotes, just the key!**

## Security Notes

- ✅ `grok_config.txt` is in `.gitignore` (won't be pushed to GitHub)
- ✅ `C:\source\` is outside the project (safe)
- ⚠️ Never share your API key
- ⚠️ Never commit your API key to git
- ⚠️ Regenerate if exposed

## Troubleshooting

### "API Key Missing" Error

**Problem:** App can't find your API key

**Solutions:**
1. Check file exists: `C:\source\grok_config.txt` OR `grok_config.txt`
2. Verify file contains your key (no extra text)
3. Make sure no `.sample` extension
4. Restart the application

### API Key Not Working

**Problem:** Key loads but API returns errors

**Solutions:**
1. Verify key is valid at https://console.x.ai/
2. Check your API usage/quota
3. Ensure key starts with `xai-`
4. Try regenerating the key

## Features Requiring API Key

- 🤖 **AI Order Parser** - Parse raw order text automatically

## Features NOT Requiring API Key

- ✅ Manual order entry
- ✅ CSV file loading
- ✅ Image processing
- ✅ PDF generation
- ✅ Everything except AI parsing!

## Development Notes

If you're forking/cloning this repo:

1. Create your own API key
2. Set up your config file
3. Never commit `grok_config.txt`
4. Use `grok_config.txt.sample` as template

## Questions?

- Grok API Docs: https://docs.x.ai/
- Grok Console: https://console.x.ai/

