# WordiCalc Examples

## Setup

| Example | Formula | Description | Expected Result |
|---------|---------|-------------|-----------------|
| Configure Endpoint | `=LLMConfig("set", "openai_api_endpoint", "https://api.openai.com/v1/chat/completions")` | Set API endpoint (optional) | Configuration set |
| Configure API Key | `=LLMConfig("set", "openai_api_key", "sk-your-key-here")` | Set up OpenAI API key | Configuration set |
| List Models | `=LLMModels()` | See available models | List of available models |
| Set Model | `=LLMConfig("set", "openai_model", "gpt-4")` | Configure specific model | Configuration set |
| Check Status | `=LLMStatus()` | View current configuration | Shows version and settings |

## Examples

| Example | Formula | Description | Expected Result |
|---------|---------|-------------|-----------------|
| Simple Question | `=LLM("What is the capital of France?")` | Basic factual query | Paris |
| Creative Writing | `=LLM("Write a haiku about Excel")` | Creative text generation | Excel haiku poem |
| Custom System Message | `=LLM("Explain gravity", "You are a physics teacher")` | Query with context | Simple physics explanation |
| Text Summary | `=LLM("Sales grew 25% in Q2 due to marketing", "Summarize in one sentence")` | Text summarization | One sentence summary |
| Integer Result | `=LLM("What is 15 * 23?", , "integer")` | Math with integer output | 345 |
| Float Result | `=LLM("Product costs $29.99 plus tax", , "float")` | Extract price as decimal | 29.99 |
| Rating Extraction | `=LLM("5 stars! Excellent service", , "integer")` | Extract numeric rating | 5 |
| Word Count | `=LLM("Count words: hello world test", , "integer")` | Count words in text | 3 |
| Sentiment Analysis | `=LLM("I love this product!", , "choice", "positive,negative,neutral")` | Classify text sentiment | positive |
| Customer Issues | `=LLM("Billing error complaint", , "choice", "billing,shipping,product,service")` | Categorize support ticket | billing |
| Language Detection | `=LLM("Bonjour comment allez-vous?", , "choice", "english,french,spanish,german")` | Detect language | french |
| Yes/No Decision | `=LLM("Should I invest at age 25?", "You are a financial advisor", "choice", "yes,no")` | Advisory with limited options | yes |
| Email Domain | `=LLM("Extract domain: john@example.com")` | Extract specific data | example.com |
| Data Classification | `=LLM("VIP customer complaint", , "choice", "standard,priority,urgent")` | Prioritize items | urgent |
| Format Conversion | `=LLM("Convert to title case: hello world")` | Text formatting | Hello World |
| Date Extraction | `=LLM("Meeting on March 15th, 2024")` | Extract dates from text | March 15th, 2024 |

## Advanced JSON Schema Examples

| Example | Formula | Description | Expected Result |
|---------|---------|-------------|-----------------|
| JSON Choice | `=LLM("Great product!", , "choice", "positive,negative,neutral", TRUE)` | Sentiment with JSON schema | positive |
| JSON Integer | `=LLM("Count: one two three four", , "integer", , TRUE)` | Number extraction with JSON | 4 |
| JSON Float | `=LLM("Pi to 2 decimals", , "float", , TRUE)` | Precise number with JSON | 3.14 |
| JSON Priority | `=LLM("Server is down!", , "choice", "low,medium,high,critical", TRUE)` | Issue priority with JSON schema | critical |
| JSON Word Count | `=LLM("The quick brown fox jumps", , "integer", , TRUE)` | Count words with JSON schema | 5 |
| JSON Temperature | `=LLM("Room feels warm, about 75 degrees", , "integer", , TRUE)` | Extract temperature with JSON | 75 |
| JSON Category | `=LLM("New laptop purchase", , "choice", "hardware,software,service,training", TRUE)` | Expense categorization with JSON | hardware |
| JSON Score | `=LLM("Customer rated us 4.5 out of 5 stars", , "float", , TRUE)` | Extract rating with JSON schema | 4.5 |
| JSON Status | `=LLM("Order shipped successfully", , "choice", "pending,processing,shipped,delivered", TRUE)` | Order status with JSON schema | shipped |
| JSON Age | `=LLM("25-year-old customer", , "integer", , TRUE)` | Extract age with JSON schema | 25 |

## Tips for Using Examples

1. **Default Arguments**: Use empty parameters (just commas) when you want default values
2. **String Schema**: Default output type, no need to specify unless you want other types
3. **Choice Values**: Always provide comma-separated options for choice schema
4. **JSON Schema**: Set last parameter to TRUE for structured output (works with compatible models)
5. **System Messages**: Use the second parameter to provide context or role instructions

## Common Patterns

- Basic question: `=LLM("your question")`
- With context: `=LLM("your question", "you are an expert in...")`
- Integer output: `=LLM("your question", , "integer")`
- Multiple choice: `=LLM("classify this", , "choice", "option1,option2,option3")`
- JSON mode: `=LLM("your question", , "integer", , TRUE)`