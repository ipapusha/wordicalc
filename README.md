# WordiCalc - AI-Powered Excel Extension

A VBA Excel extension that adds an `=LLM(...)` function for integrating OpenAI and compatible AI models directly into spreadsheets. 


## Features

- OpenAI integration with ChatGPT models and compatible APIs
- Flexible output types: strings, integers, floats, or predefined choices
- Custom system messages for AI context
- Support for local AI (Ollama, LocalAI) and cloud providers
- Secure API key storage in workbook properties
- Cross-platform compatibility (Windows/Mac)

## Quick Start

1. **Install**: Download `WordiCalc.bas`, `LibJSON.bas`, `Dictionary.cls` and import into Excel VBA
2. **Configure**: `=LLMConfig("set", "openai_api_key", "your-key")`
3. **Use**: `=LLM("What is 2+2?")`

See [INSTALL.md](INSTALL.md) for detailed setup instructions.

## Function Syntax

```excel
=LLM(prompt, [system_message], [output_schema], [allowed_values], [use_json_schema])
```

**Parameters:**
- `prompt` (required): Question or instruction for AI
- `system_message` (optional): Context for AI behavior
- `output_schema` (optional): "string", "integer", "float", or "choice"
- `allowed_values` (optional): Comma-separated values for "choice" schema
- `use_json_schema` (optional): Use structured JSON output (default: FALSE)

## Usage Examples

### Basic Usage
```excel
=LLM("What is the capital of France?")
=LLM("Analyze this data trend", "You are a business analyst")
```

### Typed Outputs
```excel
=LLM("How many days in February 2024?", , "integer")
=LLM("What's pi to 3 decimal places?", , "float")
=LLM("Is this positive: 'Great work!'", , "choice", "positive,negative,neutral")
```

### JSON Schema Mode (Advanced)
```excel
=LLM("Count words: hello world", , "integer", , TRUE)
=LLM("Classify sentiment", , "choice", "positive,negative,neutral", TRUE)
```

## Configuration Functions

```excel
=LLMConfig("set", "openai_api_key", "sk-your-key")
=LLMConfig("set", "openai_model", "gpt-4")
=LLMConfig("set", "openai_api_endpoint", "custom-url")
=LLMConfig("get", "openai_model")
=LLMConfig("list")
=LLMStatus()
=LLMModels()
```

## Supported APIs

### OpenAI
```excel
=LLMConfig("set", "openai_api_key", "sk-...")
=LLMConfig("set", "openai_model", "gpt-4")
```

### Ollama (Local)
```excel
=LLMConfig("set", "openai_api_endpoint", "http://localhost:11434/v1/chat/completions")
=LLMConfig("set", "openai_api_key", "dummy")
=LLMConfig("set", "openai_model", "llama3")
```

### Azure OpenAI
```excel
=LLMConfig("set", "openai_api_endpoint", "https://your-resource.openai.azure.com/...")
=LLMConfig("set", "openai_api_key", "your-azure-key")
```

## Use Cases

- **Data Analysis**: `=LLM("Summarize trend: 100, 150, 200", "Business analyst")`
- **Classification**: `=LLM("Categorize: customer complaint", , "choice", "billing,shipping,product")`
- **Text Processing**: `=LLM("Extract email from: Contact john@example.com")`
- **Calculations**: `=LLM("$1000 at 5% for 3 years compound?", , "float")`

## File Structure

The extension consists of three VBA files:

- **WordiCalc.bas**: Main functions and API integration
- **LibJSON.bas**: High-performance JSON parser ([VBA-FastJSON](https://github.com/cristianbuse/VBA-FastJSON))
- **Dictionary.cls**: Cross-platform dictionary ([VBA-FastDictionary](https://github.com/cristianbuse/VBA-FastDictionary))

Both libraries by Ion Cristian Buse, licensed under MIT.

## Inspiration

* [sheets-llm](https://github.com/nicucalcea/sheets-llm)
* [cellm](https://github.com/getcellm/cellm)
* [Excel Copilot](https://support.microsoft.com/en-us/copilot-excel)
* [otto](https://ottogrid.ai/)

and others

## Common Issues

**Function not recognized**: Import all three files and save as .xlsm

**Compile error**: Import Dictionary.cls as Class Module (not regular Module)

**API errors**: Check internet connection and API key validity

**Performance**: Use specific prompts and faster models like gpt-3.5-turbo for simple tasks

## Security

API keys are stored securely in workbook custom properties, not visible in normal Excel interface. Don't share workbooks containing API keys.

## License

Provided as-is for educational and commercial use. Comply with your API provider's terms of service.
