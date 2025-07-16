"""
Example module demonstrating how to use the prompt management system
for email processing in Outlook Extract.
"""
import os
import sys
from pathlib import Path
from typing import Dict, List, Any, Optional

# Add the project root to the Python path
project_root = Path(__file__).parent.parent.parent.parent
sys.path.append(str(project_root))

from outlook_extractor.prompts.utils import PromptTemplate, track_prompt
from outlook_extractor.prompts import prompt_manager

# Example 1: Using the @track_prompt decorator
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

# Example 2: Using the PromptTemplate class
class EmailProcessor:
    """
    Example class demonstrating prompt management for email processing.
    """
    
    def __init__(self):
        # Initialize prompt templates
        self.classification_prompt = PromptTemplate(
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
        
        self.response_draft_prompt = PromptTemplate(
            prompt_id="email_response_draft",
            template="""
            Draft a professional email response based on the following context:
            
            Original email from {sender}:
            {email_content}
            
            Response requirements:
            - Tone: {tone}
            - Key points to include: {key_points}
            - Length: Approximately {word_count} words
            
            Please structure the response with appropriate greeting, body, and closing.
            """,
            model="gpt-4",
            tags=["email", "response", "draft"]
        )
    
    def classify_email(self, subject: str, sender: str, preview: str) -> str:
        """
        Classify an email based on its metadata and preview.
        """
        # Format the prompt with the provided variables
        prompt = self.classification_prompt.format(
            subject=subject,
            sender=sender,
            preview=preview[:200]  # First 200 characters
        )
        
        # In a real implementation, you would call an AI model here
        # For example: response = ai_client.complete(prompt=prompt, ...)
        print(f"\n[DEBUG] Classification prompt:\n{prompt}\n")
        
        # Simulate a response
        return "IMPORTANT"
    
    def draft_response(
        self,
        sender: str,
        email_content: str,
        tone: str = "professional",
        key_points: List[str] = None,
        word_count: int = 150
    ) -> str:
        """
        Draft a response to an email.
        """
        # Format the prompt with the provided variables
        prompt = self.response_draft_prompt.format(
            sender=sender,
            email_content=email_content,
            tone=tone,
            key_points=", ".join(key_points) if key_points else "None provided",
            word_count=word_count
        )
        
        # In a real implementation, you would call an AI model here
        print(f"\n[DEBUG] Response draft prompt:\n{prompt}\n")
        
        # Simulate a response
        return f"Draft response to email from {sender} with {len(email_content)} characters"

def example_ab_testing():
    """
    Example of running an A/B test between different prompt variants.
    """
    # Define the prompt variants to test
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
    
    # Test inputs
    test_emails = [
        """
        Subject: Urgent: Need access to the quarterly report
        From: John Smith <john.smith@example.com>
        
        Hi team,
        
        I hope this email finds you well. I'm reaching out because I need access to 
        the Q2 financial report as soon as possible. Our board meeting has been moved 
        up to tomorrow morning, and I need to prepare my presentation.
        
        Could you please grant me access to the report at your earliest convenience?
        
        Best regards,
        John
        """,
        """
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
    ]
    
    # Evaluation metrics
    def completeness_score(response: str) -> int:
        """Score 1-5 based on how many fields were extracted."""
        fields = ['sender name', 'company', 'request type', 'urgency']
        return sum(1 for field in fields if field.lower() in response.lower())
    
    def clarity_score(response: str) -> int:
        """Score 1-5 based on clarity of the response."""
        # Simple heuristic: longer responses with good structure tend to be clearer
        lines = [line.strip() for line in response.split('\n') if line.strip()]
        return min(5, len(lines) // 2)  # Cap at 5
    
    # Run the A/B test
    results = prompt_manager.run_ab_test(
        test_name="email_extraction_style_test",
        prompt_variants=variants,
        test_inputs=[{"email_content": email} for email in test_emails],
        evaluation_metrics={
            "completeness": completeness_score,
            "clarity": clarity_score
        },
        num_runs=2
    )
    
    print("\nA/B Test Results:")
    print(json.dumps(results, indent=2))
    return results

if __name__ == "__main__":
    import json
    
    print("=== Example 1: Email Summarization ===")
    email_content = """
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
    
    summary = summarize_email(email_content)
    print(f"Summary: {summary}")
    
    print("\n=== Example 2: Email Classification ===")
    processor = EmailProcessor()
    
    subject = "Urgent: Server Downtime Tonight"
    sender = "it@company.com"
    preview = "This is to inform you about scheduled maintenance tonight from 10 PM to 2 AM..."
    
    category = processor.classify_email(subject, sender, preview)
    print(f"Email category: {category}")
    
    print("\n=== Example 3: A/B Testing ===")
    example_ab_testing()
