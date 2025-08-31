# WordiCalc Installation Guide

## Quick Setup

### 1. Download Files
Get these three files from this repository:
- `WordiCalc.bas` - Main LLM functions
- `LibJSON.bas` - JSON parsing library
- `Dictionary.cls` - Cross-platform dictionary

### 2. Import into Excel
1. Open Excel and press Alt+F11 (Windows) or Opt+F11 (Mac) for VBA Editor
2. Right-click on VBAProject in Project Explorer
3. Select "Import File..." and import all three files:
   - `Dictionary.cls` (will appear as Class Module)
   - `LibJSON.bas` (will appear as Module)
   - `WordiCalc.bas` (will appear as Module)

### 3. Save and Enable
1. Save workbook as .xlsm format
2. Enable macros when prompted

### 4. Configure API Key
In any Excel cell:
```
=LLMConfig("set", "openai_api_key", "your-api-key-here")
```

### 5. Test
```
=LLM("Hello, are you working?")
```

## Alternative Import Method
If file import doesn't work:
1. Create Class Module named "Dictionary", paste Dictionary.cls content
2. Create Module for LibJSON, paste LibJSON.bas content
3. Create Module for WordiCalc, paste WordiCalc.bas content

## API Configuration

### OpenAI (Default)
```
=LLMConfig("set", "openai_api_key", "sk-your-key")
=LLMConfig("set", "openai_model", "gpt-4")
```

### Local AI (Ollama)
```
=LLMConfig("set", "openai_api_endpoint", "http://localhost:11434/v1/chat/completions")
=LLMConfig("set", "openai_api_key", "dummy")
=LLMConfig("set", "openai_model", "llama3")
```

### Azure OpenAI
```
=LLMConfig("set", "openai_api_endpoint", "https://your-resource.openai.azure.com/...")
=LLMConfig("set", "openai_api_key", "your-azure-key")
```

## Troubleshooting

**Function not recognized**
- Ensure all three files are imported
- Save as .xlsm format
- Enable macros

**Compile error: User-defined type not defined**
- Import Dictionary.cls as Class Module (not regular Module)
- Restart Excel

**API key not configured**
- Run: `=LLMConfig("set", "openai_api_key", "your-key")`

**HTTP Request Failed**
- Check internet connection
- Verify API endpoint and key
- For local APIs, ensure server is running

## Installation Options

**Single Workbook**: Import files into specific workbook (functions only available there)

**Excel Add-in**: Create new workbook, import files, save as .xlam, enable in Excel Add-ins (functions available in all workbooks)

## Verification
Check installation:
```
=LLMStatus()
```

Should show configured API key, endpoint, and model.