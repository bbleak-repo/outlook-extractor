"""
Prompt Management Example for Outlook Extract

This script demonstrates how to use the prompt management system in the
Outlook Extract application to track, version, and test prompts.
"""

import os
import sys
from pathlib import Path
from datetime import datetime

# Add the project root to the Python path
project_root = Path(__file__).parent.parent
sys.path.append(str(project_root))

from outlook_extractor.prompts import (
    PromptTemplate,
    track_prompt,
    prompt_manager,
    get_prompt_config
)

def main():
    """Run the prompt management example."""
    print("=== Outlook Extract - Prompt Management Example ===\n")
    
    # Show the current prompt configuration
    config = get_prompt_config()
    print("Current prompt configuration:")
    print(f"Base directory: {config['prompts']['base_dir']}")
    print(f"Default model: {config['prompts']['default_model']}")
    print(f"Default parameters: {config['prompts']['default_parameters']}")
    
    # Example 1: Using the @track_prompt decorator
    print("\n=== Example 1: Tracking a prompt with @track_prompt decorator ===")
    
    @track_prompt(
        prompt_id="email_summarization",
        model="gpt-4",
        parameters={"temperature": 0.7, "max_tokens": 500},
        tags=["email", "summarization"]
    )
    def summarize_email(email_content: str) -> str:
        """
        Summarize the given email content.
        
        The email may contain multiple paragraphs and sections. 
        Please provide a concise summary that captures the main points, 
        key details, and any action items or requests mentioned in the email.
        
        Email content:
        {email_content}
        """
        # In a real implementation, this would call an AI model
        return f"Summary for email with {len(email_content)} characters"
    
    # Call the decorated function
    test_email = """
    Subject: Project Update: Q3 Marketing Campaign
    From: marketing@company.com
    
    Hi Team,
    
    I wanted to provide an update on our Q3 marketing campaign. We've seen 
    a 15% increase in engagement compared to last quarter, which is great news! 
    However, we need to address the bounce rate, which has increased by 5%.
    
    Action items:
    1. Review the email content for the next campaign (due Friday)
    2. Schedule a brainstorming session for new ad creatives
    3. Prepare a report on the performance metrics for the leadership team
    
    Let me know if you have any questions.
    
    Best,
    Alex
    """
    
    summary = summarize_email(test_email)
    print(f"\nGenerated summary: {summary}")
    
    # Example 2: Using PromptTemplate
    print("\n=== Example 2: Using PromptTemplate ===")
    
    # Create a prompt template
    classification_template = PromptTemplate(
        prompt_id="email_classification",
        template="""
        Classify the following email into one of these categories:
        - IMPORTANT: Urgent or high-priority emails
        - ACTION_REQUIRED: Emails that require a response or action
        - REFERENCE: Informational emails for future reference
        - PROMOTIONAL: Marketing or promotional content
        - SPAM: Unwanted or unsolicited emails
        
        Email subject: {subject}
        Sender: {sender}
        First 200 chars: {preview}
        
        Return only the category name, nothing else.
        """,
        model="gpt-4",
        tags=["email", "classification"]
    )
    
    # Format the prompt with test data
    formatted_prompt = classification_template.format(
        subject="Urgent: Server Maintenance Tonight",
        sender="it@company.com",
        preview="This is to inform you about scheduled maintenance tonight from 10 PM to 2 AM..."
    )
    
    print("\nFormatted prompt:")
    print("-" * 80)
    print(formatted_prompt.strip())
    print("-" * 80)
    
    # Track the prompt
    prompt_manager.track_prompt(
        prompt_id=classification_template.prompt_id,
        prompt_text=classification_template.template,
        model=classification_template.model,
        parameters=classification_template.parameters,
        source_file=__file__,
        line_number=inspect.currentframe().f_lineno,
        tags=classification_template.tags,
        metadata={
            'description': 'Template for classifying emails',
            'version': classification_template.version
        }
    )
    
    # Example 3: Running an A/B test
    print("\n=== Example 3: Running an A/B test ===")
    
    # Define prompt variants
    variants = [
        {
            'name': 'direct_style',
            'prompt': """
            Extract the following information from this email:
            - Sender name
            - Company name
            - Request type (e.g., information, meeting, support)
            - Urgency level (low, medium, high)
            
            Email: {email_content}
            """,
            'parameters': {'temperature': 0.3}
        },
        {
            'name': 'polite_style',
            'prompt': """
            Could you please help extract the following details from this email?
            
            We're looking for:
            1. The name of the sender
            2. Their company name
            3. What they're requesting (information, meeting, support, etc.)
            4. How urgent their request seems (low, medium, high)
            
            Here's the email content:
            {email_content}
            
            Thank you for your help!
            """,
            'parameters': {'temperature': 0.3}
        }
    ]
    
    # Define test inputs
    test_emails = [
        {
            'email_content': """
            Subject: Urgent: Need access to the quarterly report
            From: John Smith <john.smith@example.com>
            
            Hi team,
            
            I hope this email finds you well. I'm reaching out because I need access to 
            the Q2 financial report as soon as possible. Our board meeting has been moved 
            up to tomorrow morning, and I need to prepare my presentation.
            
            Could you please grant me access to the report at your earliest convenience?
            
            Best regards,
            John
            """
        },
        {
            'email_content': """
            Subject: Follow-up on our meeting
            From: Sarah Johnson <sarah.j@acmecorp.com>
            
            Hello,
            
            I'm just following up on our discussion last week about potential 
            collaboration opportunities between our companies. I'd love to schedule 
            another call to explore this further.
            
            Please let me know your availability for next week.
            
            Warm regards,
            Sarah
            """
        }
    ]
    
    # Define evaluation metrics
    def completeness_score(response: str) -> int:
        """Score 1-4 based on how many fields were extracted."""
        fields = ['sender name', 'company', 'request type', 'urgency']
        return sum(1 for field in fields if field.lower() in response.lower())
    
    # Run the A/B test
    print("\nRunning A/B test with 2 prompt variants and 2 test emails...")
    results = prompt_manager.run_ab_test(
        test_name="email_extraction_style_test",
        prompt_variants=variants,
        test_inputs=test_emails,
        evaluation_metrics={"completeness": completeness_score},
        num_runs=1
    )
    
    # Print the results
    print("\nA/B Test Results:")
    print(f"Test name: {results['test_name']}")
    print(f"Variants: {', '.join(results['variants'].keys())}")
    print("\nMetrics:")
    for metric, variant_scores in results['metrics'].items():
        print(f"\n{metric}:")
        for variant, scores in variant_scores.items():
            print(f"  {variant}: {scores}")
    
    print("\n=== Prompt Management Example Complete ===")

if __name__ == "__main__":
    import inspect
    main()
