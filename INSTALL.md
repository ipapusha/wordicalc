# WordiCalc - Installation Guide

## Quick Installation (3 minutes)

### Step 1: Download the Required Files
Download these three files from this repository:
- `WordiCalc.bas` - Main WordiCalc functions
- `JsonConverter.bas` - JSON parsing library
- `Dictionary.cls` - Cross-platform dictionary (Mac compatibility)

### Step 2: Open Excel VBA Editor
- **Windows**: Press `Alt + F11`
- **Mac**: Press `Opt + F11` (or `Fn + Opt + F11`)

### Step 3: Import All Files
**Important**: Import all three files for full functionality.

1. In the VBA Editor, right-click on **VBAProject** in the Project Explorer
2. Select **Import File...**
3. Browse to and select `Dictionary.cls` first
4. Repeat: Import `JsonConverter.bas` 
5. Repeat: Import `WordiCalc.bas`

**Alternative method (copy/paste):**
1. Right-click VBAProject → Insert → **Class Module** (rename to "Dictionary")
2. Copy contents of `Dictionary.cls` and paste
3. Right-click VBAProject → Insert → **Module** (for JsonConverter)
4. Copy contents of `JsonConverter.bas` and paste  
5. Right-click VBAProject → Insert → **Module** (for WordiCalc)
6. Copy contents of `WordiCalc.bas` and paste

### Step 4: Save Your Workbook
1. Save your workbook as a **macro-enabled file** (`.xlsm` format)
2. When prompted about macros, click **Enable Macros**

### Step 5: Configure API Access
In any Excel cell, enter:
```excel
=LLMConfig("set", "openai_api_key", "your-api-key-here")
```

### Step 6: Test the Installation
```excel
=LLM("Hello, are you working?")
```

**Success!** You should see a response from the AI.

---

## Detailed Instructions

### Getting an OpenAI API Key
1. Go to [platform.openai.com/api-keys](https://platform.openai.com/api-keys)
2. Sign up or log in to your account
3. Click **"Create new secret key"**
4. Copy the key (starts with `sk-`)
5. Store it securely - you won't see it again

### Alternative: Using Local AI (Ollama, LocalAI)
```excel
=LLMConfig("set", "openai_api_endpoint", "http://localhost:11434/v1/chat/completions")
=LLMConfig("set", "openai_api_key", "dummy-not-needed")
=LLMConfig("set", "openai_model", "llama3")
```

### Enabling Macros Permanently
1. Go to **File** → **Options** → **Trust Center**
2. Click **Trust Center Settings**
3. Select **Macro Settings**
4. Choose **"Enable all macros"** or **"Disable all macros with notification"**

---

## Troubleshooting

### "Function not recognized"
- Ensure all three files are imported correctly:
  - `Dictionary.cls` (as Class Module)
  - `JsonConverter.bas` (as Module) 
  - `WordiCalc.bas` (as Module)
- Save workbook as `.xlsm` format
- Enable macros when prompted

### "Compile error: User-defined type not defined"
- Ensure `Dictionary.cls` is imported as a **Class Module**
- Restart Excel and reopen the workbook
- This error typically means the Dictionary class wasn't imported correctly

### "API key not configured"
- Run: `=LLMConfig("set", "openai_api_key", "your-key")`
- Check your key is valid at OpenAI platform

### "HTTP Request Failed"
- Check internet connection
- Verify API endpoint URL
- Ensure API key has credits/usage remaining

### "VBA Editor won't open"
- Enable Developer tab: **File** → **Options** → **Customize Ribbon** → Check **Developer**
- Try **Developer** → **Visual Basic** from ribbon

---

## Installation Options

### Option A: Single Workbook
- Import all three files (`Dictionary.cls`, `JsonConverter.bas`, `WordiCalc.bas`) into your specific workbook
- Functions only available in that workbook

### Option B: Excel Add-in (All Workbooks)
1. Create new workbook and import all three files
2. Save as **Excel Add-in** (`.xlam`)
3. Go to **File** → **Options** → **Add-ins**
4. Click **Manage Excel Add-ins** → **Go**
5. Browse and select your `.xlam` file
6. Check the box to enable it

The add-in approach makes LLM functions available in all Excel workbooks.

## Platform Compatibility

### Windows Excel
- ✅ Works with Excel 2010, 2013, 2016, 2019, 2021, 365
- ✅ No external references required
- ✅ Automatic fallback to native `Scripting.Dictionary` when available

### Mac Excel  
- ✅ Works with Excel 2016, 2019, 2021, 365 for Mac
- ✅ Full compatibility without Microsoft Scripting Runtime
- ✅ Uses cross-platform `Dictionary.cls` implementation

### Dependencies Explained
**Why three files?**
1. **`WordiCalc.bas`**: Your main LLM functions
2. **`JsonConverter.bas`**: Professional JSON parsing (handles complex API responses) 
3. **`Dictionary.cls`**: Mac-compatible dictionary (replaces Windows-only `Scripting.Dictionary`)

These dependencies ensure WordiCalc works reliably across all platforms and handles real-world JSON responses from AI APIs.

---

## Verification

Check your installation with:
```excel
=LLMStatus()
```

Should show:
```
WordiCalc v1.1
API Key: Configured
Endpoint: https://api.openai.com/v1/chat/completions (default)
Model: gpt-3.5-turbo (default)
```

---

## Next Steps

Once installed, see the main [README.md](README.md) for:
- Function usage examples
- Configuration options
- Advanced features
- Troubleshooting tips