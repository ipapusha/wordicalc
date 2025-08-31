# WordiCalc - AI-Powered Excel Extension

A pure VBA Excel extension that adds an `=LLM(...)` function for integrating OpenAI models directly into Excel spreadsheets. Everything is contained in one file for easy installation!

## Features

- **OpenAI Integration**: Supports OpenAI (ChatGPT) models and OpenAI-compatible APIs
- **Flexible Output Schemas**: Return strings, integers, floats, or predefined choices
- **Customizable System Messages**: Define context and instructions for the AI
- **Custom API Endpoints**: Use OpenAI or compatible APIs (Ollama, LocalAI, etc.)
- **Secure Configuration**: API keys stored safely in workbook properties
- **Error Handling**: Comprehensive error messages and validation
- **Single File**: One `.bas` file contains everything!

## Function Signature

```excel
=LLM(prompt, [system_message], [output_schema], [allowed_values], [use_json_schema])
```

### Parameters

- **prompt** (required): The question or instruction for the AI
- **system_message** (optional): Context/instructions for the AI (default: "You are a helpful assistant.")
- **output_schema** (optional): Response format - "string", "integer", "float", or "choice" (default: "string")
- **allowed_values** (optional): Comma-separated list of valid responses for "choice" schema
- **use_json_schema** (optional): Use JSON schema instead of system prompt instructions (default: FALSE)

## Quick Installation

**3-minute setup!** See [INSTALL.md](INSTALL.md) for detailed instructions.

1. **Download**: Get the required files:
   - `WordiCalc.bas` - Main LLM functions
   - `JsonConverter.bas` - JSON parsing library  
   - `Dictionary.cls` - Cross-platform dictionary (Mac compatibility)
2. **Import**: Open Excel VBA Editor (`Alt+F11`), right-click VBAProject → Import File (import all 3 files)
3. **Save**: Save workbook as `.xlsm` and enable macros
4. **Configure**: `=LLMConfig("set", "openai_api_key", "your-key")`
5. **Test**: `=LLM("Hello!")`

**That's it!** For detailed steps, troubleshooting, and advanced options, see the complete [Installation Guide](INSTALL.md).

## Configuration

### Set OpenAI API Key
```excel
=LLMConfig("set", "openai_api_key", "sk-your-key-here")
```

### Set Model (Optional)
```excel
=LLMConfig("set", "openai_model", "gpt-4")
```

### Set Custom API Endpoint (Optional)
For OpenAI-compatible APIs like Ollama, LocalAI, or other providers:
```excel
=LLMConfig("set", "openai_api_endpoint", "http://localhost:11434/v1/chat/completions")
```

### Check Configuration
```excel
=LLMConfig("get", "openai_api_key")
=LLMConfig("list")
=LLMStatus()
```

### Discover Available Models
```excel
=LLMModels()
```
Lists all models available at your configured endpoint.

## Usage Examples

### Basic Usage
```excel
=LLM("What is the capital of France?")
```

### Integer Output
```excel
=LLM("How many days are in February 2024?", , "integer")
```

### Choice Output
```excel
=LLM("Is this positive or negative: 'Great job!'", , "choice", "positive,negative,neutral")
```

### Using JSON Schema (Advanced)
For models that support structured output (like GPT-4), you can use JSON schemas instead of system prompt instructions:
```excel
=LLM("Classify: 'Great job!'", , "choice", "positive,negative,neutral", TRUE)
=LLM("Count words: hello world", , "integer", , TRUE)
```

**Difference between approaches:**
- **System Prompt (default)**: Adds instructions to your system message asking the AI to respond in the format
- **JSON Schema**: Uses the API's `response_format` with strict JSON schema validation

JSON Schema is more reliable for supported models but may not work with all APIs.

### Using Different Models
First discover what models are available, then configure and use:
```excel
=LLMModels()
=LLMConfig("set", "openai_model", "gpt-4")
=LLM("Write a haiku about Excel")
```

### Custom System Message
```excel
=LLM("Analyze Q1 sales: $100k", "You are a financial analyst expert")
```

## Compatible APIs

Works with any OpenAI-compatible API:

### OpenAI (Default)
```excel
=LLMConfig("set", "openai_api_key", "sk-...")
=LLMModels()
=LLMConfig("set", "openai_model", "gpt-4")
```

### Ollama (Local)
```excel
=LLMConfig("set", "openai_api_endpoint", "http://localhost:11434/v1/chat/completions")
=LLMConfig("set", "openai_api_key", "dummy-key-not-needed")
=LLMModels()
=LLMConfig("set", "openai_model", "llama3")
```

### LocalAI
```excel
=LLMConfig("set", "openai_api_endpoint", "http://localhost:8080/v1/chat/completions")
=LLMConfig("set", "openai_api_key", "your-local-api-key")
=LLMConfig("set", "openai_model", "your-model-name")
```

### Azure OpenAI
```excel
=LLMConfig("set", "openai_api_endpoint", "https://your-resource.openai.azure.com/openai/deployments/your-model/chat/completions?api-version=2024-02-15-preview")
=LLMConfig("set", "openai_api_key", "your-azure-api-key")
=LLMConfig("set", "openai_model", "gpt-35-turbo")
```

## Use Cases

### Data Analysis
```excel
=LLM("Summarize trend: sales 100, 150, 200", "You analyze business data")
```

### Classification
```excel
=LLM("Customer email: 'Delayed delivery!'", , "choice", "billing,shipping,product,service")
```

### Text Processing
```excel
=LLM("Extract email from: Contact john@example.com")
```

### Calculations
```excel
=LLM("$1000 at 5% for 3 years compound interest?", , "float")
```

## Supported Models

### OpenAI Models
- `gpt-3.5-turbo` (default)
- `gpt-4`
- `gpt-4-turbo`
- `gpt-4o`
- `gpt-4o-mini`

### Local Models (via Ollama/LocalAI)
- `llama3`
- `codellama`
- `mistral`
- `phi3`
- Any model supported by your local API

## Troubleshooting

### Common Issues

1. **"Error: OpenAI API key not configured"**
   - Set your API key: `=LLMConfig("set", "openai_api_key", "your-key")`

2. **"HTTP Request Failed"**
   - Check internet connection
   - Verify API endpoint is correct
   - For local APIs, ensure server is running

3. **Function not recognized**
   - Ensure `ExcelLLM.bas` is imported
   - Save workbook as `.xlsm` and enable macros

### Performance Tips

- Use specific, concise prompts
- Cache responses by copying values if static
- Use faster models like `gpt-3.5-turbo` for simple tasks
- Use `gpt-4` for complex reasoning

## Security Notes

- API keys stored securely in workbook custom properties
- Keys not visible in normal Excel interface
- Don't share workbooks containing API keys

## Advanced Configuration

### Check All Settings
```excel
=LLMStatus()
```

### Discover Available Models
```excel
=LLMModels()
```

### Clear Configuration
```excel
=LLMConfig("clear", "openai_api_key")
=LLMConfig("clear", "openai_api_endpoint")
```

### Example Configurations

**Using with Ollama locally:**
```excel
=LLMConfig("set", "openai_api_endpoint", "http://localhost:11434/v1/chat/completions")
=LLMConfig("set", "openai_api_key", "ollama")
=LLMConfig("set", "openai_model", "llama3")
=LLM("Hello")
```

**Using different OpenAI-compatible service:**
```excel
=LLMConfig("set", "openai_api_endpoint", "https://api.your-service.com/v1/chat/completions")
=LLMConfig("set", "openai_api_key", "your-service-key")
=LLMConfig("set", "openai_model", "your-model-name")
```

## File Structure

WordiCalc consists of three VBA files:

### `WordiCalc.bas` (Main Module)
- Main `LLM()` function and variants (`LLMConfig`, `LLMStatus`, `LLMModels`)
- Configuration management (API key, endpoint, model settings)
- HTTP client for API communication
- OpenAI API integration
- Function registration for IntelliSense tooltips
- All core functionality

### `JsonConverter.bas` (JSON Library)
- Professional JSON parsing and serialization
- Cross-platform compatibility (Windows/Mac)
- Robust handling of nested objects and arrays
- **Source**: [VBA-tools/VBA-JSON](https://github.com/VBA-tools/VBA-JSON) by Tim Hall
- **License**: MIT

### `Dictionary.cls` (Cross-Platform Dictionary)
- Drop-in replacement for `Scripting.Dictionary`
- Mac compatibility (no Microsoft Scripting Runtime dependency)
- Automatic fallback to native `Scripting.Dictionary` on Windows when available
- **Source**: [VBA-tools/VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary) by Tim Hall  
- **License**: MIT

## Dependencies & Credits

WordiCalc builds upon excellent open-source VBA libraries:

- **JSON Processing**: [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) - Professional JSON converter for VBA
- **Cross-Platform Dictionary**: [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary) - Mac-compatible dictionary implementation

Both libraries are created by Tim Hall and licensed under MIT. WordiCalc integrates these libraries to provide a robust, cross-platform AI integration solution for Excel.

### Why These Dependencies?

1. **Robust JSON Parsing**: VBA's built-in JSON support is limited. VBA-JSON handles complex nested structures reliably.
2. **Mac Compatibility**: `Scripting.Dictionary` isn't available on Mac. VBA-Dictionary provides seamless cross-platform support.
3. **Professional Quality**: These are battle-tested libraries used by thousands of VBA developers.
4. **MIT Licensed**: Compatible with both personal and commercial use.

## License

This project is provided as-is for educational and commercial use. Ensure you comply with your API provider's terms of service.