# Prompt Management System

This module provides tools for managing, versioning, and testing AI prompts used throughout the Outlook Extract application.

## Features

- **Prompt Versioning**: Track changes to prompts over time
- **A/B Testing**: Compare different prompt variants
- **Metadata Tracking**: Store additional context with each prompt
- **Integration**: Seamlessly integrate with existing code using decorators or direct API calls

## Installation

1. Install the required dependencies:

```bash
pip install prompt-keeper promptpilot
```

2. Add the prompts directory to your Python path or install the package in development mode:

```bash
pip install -e .
```

## Usage

### Basic Usage

```python
from outlook_extractor.prompts.utils import track_prompt

@track_prompt(
    prompt_id="email_summarization",
    model="gpt-4",
    parameters={"temperature": 0.7},
    tags=["email", "summarization"]
)
def summarize_email(email_content: str) -> str:
    """
    Summarize the given email content.
    
    Email content:
    {email_content}
    """
    # Your implementation here
    return "Summary..."
```

### Using PromptTemplate

```python
from outlook_extractor.prompts.utils import PromptTemplate

# Create a prompt template
prompt = PromptTemplate(
    prompt_id="email_classification",
    template="""
    Classify this email: {email_content}
    Categories: IMPORTANT, ACTION_REQUIRED, REFERENCE, PROMOTIONAL, SPAM
    """,
    model="gpt-4",
    tags=["email", "classification"]
)

# Use the template
classification = prompt.format(email_content="Hello, please review this document...")
```

### Running A/B Tests

```python
from outlook_extractor.prompts import prompt_manager

# Define prompt variants
variants = [
    {
        'name': 'direct_style',
        'prompt': 'Extract: {email_content}',
        'parameters': {'temperature': 0.3}
    },
    {
        'name': 'polite_style',
        'prompt': 'Please extract: {email_content}',
        'parameters': {'temperature': 0.3}
    }
]

# Run the test
results = prompt_manager.run_ab_test(
    test_name="email_extraction_style_test",
    prompt_variants=variants,
    test_inputs=[{"email_content": "Test email content"}],
    evaluation_metrics={"length": lambda x: len(x)},
    num_runs=2
)
```

## Configuration

Configuration is handled through the main application config. You can override default settings in your config file:

```yaml
prompts:
  base_dir: ~/.outlook_extract/prompts
  enable_tracking: true
  default_model: gpt-4
  default_parameters:
    temperature: 0.7
    max_tokens: 1000
```

## Best Practices

1. **Use Descriptive IDs**: Choose IDs that clearly describe the prompt's purpose
2. **Version Control**: Commit prompt templates to version control
3. **Documentation**: Include examples and expected outputs in docstrings
4. **Testing**: Create unit tests for your prompts
5. **Monitoring**: Regularly review prompt performance and update as needed

## Examples

See the `examples/` directory for complete usage examples, including:

- Email summarization
- Email classification
- Response drafting
- A/B testing

## License

This project is licensed under the MIT License - see the [LICENSE](../LICENSE) file for details.
