# WordiCalc - AI-Powered Excel Extension

A VBA Excel extension that adds an `=LLM(...)` function for integrating OpenAI and compatible AI models directly into spreadsheets.

## Quick Start

1. **Install**: Download `WordiCalc.bas`, `LibJSON.bas`, `Dictionary.cls` and import into Excel VBA
2. **Configure**: `=LLMConfig("set", "openai_api_key", "your-key")` (optionally set endpoint: `=LLMConfig("set", "openai_api_endpoint", "custom-url")`)
3. **Use**: `=LLM("What is 2+2?")`

See [INSTALL.md](INSTALL.md) for detailed setup.

## Usage

```excel
=LLM(prompt, [sys], [schema], [values], [useJson])
```

**Parameters:**
- `prompt` (required): Question or instruction for AI
- `sys` (optional): Context for AI behavior (default: "You are a helpful assistant.")
- `schema` (optional): "string", "integer", "float", or "choice" (default: "string")
- `values` (optional): Comma-separated values for "choice" schema
- `useJson` (optional): Use structured JSON output (default: FALSE)

**Examples:**
```excel
=LLM("What is the capital of France?")
=LLM("How many days in February 2024?", , "integer")
=LLM("Is this positive: 'Great work!'", , "choice", "positive,negative,neutral")
=LLM("Analyze trend: 100, 150, 200", "Business analyst")
```

**Configuration:**
```excel
=LLMConfig("set", "openai_api_key", "sk-your-key")
=LLMConfig("set", "openai_model", "gpt-4")
=LLMConfig("set", "openai_api_endpoint", "custom-url")  # For Ollama, Azure, etc.
```

## Architecture

```
Excel Cell: =LLM("What is 2+2?")
     |
     v
WordiCalc.bas:LLM() Function
     |
     v
ValidateParameters() -> Check prompt, schema, values
     |
     v
CallAPI()
     |
     v
BuildRequestBody() -> Build HTTP request
     |
     |--- useJson=FALSE: Enhance system prompt 
     |    ("Respond with only a single integer.")
     |
     |--- useJson=TRUE: Add JSON schema to request (for supported models)
     |    ({"type":"object","properties":{"value":{"type":"integer"}}})
     |
     v
HttpRequest() -> Send POST to OpenAI API
     |
     v
API Response -> JSON with choices[0].message.content
     |
     v
ExtractContent() -> Parse API response
     |
     v
ConvertOutput() -> Convert to Excel data type
     |
     v
Return value to Excel cell
```

## Components

- **WordiCalc.bas**: Main functions and API integration
- **LibJSON.bas**: JSON parser ([VBA-FastJSON](https://github.com/cristianbuse/VBA-FastJSON))
- **Dictionary.cls**: Cross-platform dictionary ([VBA-FastDictionary](https://github.com/cristianbuse/VBA-FastDictionary))

## Inspiration

The name *WordiCalc* comes from [VisiCalc](https://en.wikipedia.org/wiki/VisiCalc), the original killer spreadsheet app for the Apple II that served as inspiration for Lotus 1-2-3, Multiplan, and ultimately Microsoft Excel. See [history](https://en.wikipedia.org/wiki/Microsoft_Excel#Early_history).

* [sheets-llm](https://github.com/nicucalcea/sheets-llm)
* [cellm](https://github.com/getcellm/cellm) 
* [otto](https://ottogrid.ai/)
* Excel Copilot

## Troubleshooting

**Function not recognized**: Import all three files and save as .xlsm  
**Compile error**: Import Dictionary.cls as Class Module (not regular Module)  
**API errors**: Check internet connection and API key validity  

**Security**: API keys are stored in workbook properties. Do not share workbooks containing API keys.
