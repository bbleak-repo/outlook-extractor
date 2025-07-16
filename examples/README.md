# Prompt Management Examples

This directory contains example scripts demonstrating how to use the prompt management system in the Outlook Extract application.

## Available Examples

### 1. `prompt_management_demo.py`

A comprehensive demo that shows:
- Tracking prompts with the `@track_prompt` decorator
- Creating and using `PromptTemplate` objects
- Running A/B tests on different prompt variants
- Evaluating prompt performance

#### How to Run

```bash
# Install required dependencies
pip install prompt-keeper promptpilot

# Run the demo
python examples/prompt_management_demo.py
```

## Integration with Outlook Extract

To use the prompt management system in your own code:

```python
from outlook_extractor.prompts import (
    PromptTemplate,
    track_prompt,
    prompt_manager,
    get_prompt_config
)

# Track a prompt with a decorator
@track_prompt(
    prompt_id="your_prompt_id",
    model="gpt-4",
    parameters={"temperature": 0.7},
    tags=["your_tag"]
)
def your_function():
    """Your prompt goes here."""
    pass

# Or use PromptTemplate
template = PromptTemplate(
    prompt_id="your_template_id",
    template="Your template with {variables}",
    model="gpt-4"
)

# Run A/B tests
results = prompt_manager.run_ab_test(
    test_name="your_test",
    prompt_variants=[...],
    test_inputs=[...],
    evaluation_metrics={...}
)
```

## Best Practices

1. **Prompt IDs**: Use descriptive, unique IDs for your prompts
2. **Versioning**: The system automatically tracks prompt versions
3. **Testing**: Use A/B testing to compare different prompt variants
4. **Documentation**: Include clear docstrings in your prompt functions
5. **Organization**: Use tags to organize related prompts

## Troubleshooting

If you encounter any issues:
1. Ensure all dependencies are installed
2. Check that the prompt database is accessible
3. Verify your configuration in `config.yaml`
4. Check the application logs for detailed error messages
