# WordiCalc - Installation Guide

## Quick Installation (3 minutes)

### Step 1: Download the Extension
Download the `WordiCalc.bas` file from this repository.

### Step 2: Open Excel VBA Editor
- **Windows**: Press `Alt + F11`
- **Mac**: Press `Opt + F11` (or `Fn + Opt + F11`)

### Step 3: Import the Module
1. In the VBA Editor, right-click on **VBAProject** in the Project Explorer
2. Select **Insert** → **Module** (or right-click and choose **Import File...**)
3. If using Import File, browse to and select `WordiCalc.bas`
4. If using Insert Module, copy the entire contents of `WordiCalc.bas` and paste it

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
- Ensure the WordiCalc module is imported correctly
- Save workbook as `.xlsm` format
- Enable macros when prompted

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
- Import `WordiCalc.bas` into your specific workbook
- Functions only available in that workbook

### Option B: Excel Add-in (All Workbooks)
1. Create new workbook with the module
2. Save as **Excel Add-in** (`.xlam`)
3. Go to **File** → **Options** → **Add-ins**
4. Click **Manage Excel Add-ins** → **Go**
5. Browse and select your `.xlam` file
6. Check the box to enable it

The add-in approach makes LLM functions available in all Excel workbooks.

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