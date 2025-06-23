import requests
import json
import time
from typing import Tuple, Optional, List
import logging

class LLMMapper:
    """
    LLM-based mapping component for financial statement line items.
    Uses Ollama with Mistral for free local inference.
    """
    
    def __init__(self, ollama_url: str = "http://localhost:11434"):
        self.ollama_url = ollama_url
        self.model_name = "mistral"
        self.setup_logging()
        
    def setup_logging(self):
        """Setup logging for LLM operations"""
        self.logger = logging.getLogger("llm_mapper")
        self.logger.setLevel(logging.INFO)
        
    def check_ollama_available(self) -> bool:
        """Check if Ollama is running and Mistral model is available"""
        try:
            # Check if Ollama is running
            response = requests.get(f"{self.ollama_url}/api/tags", timeout=5)
            if response.status_code != 200:
                return False
            
            # Check if Mistral model is available
            models = response.json().get("models", [])
            return any(model["name"].startswith(self.model_name) for model in models)
            
        except Exception as e:
            self.logger.warning(f"Ollama not available: {e}")
            return False
    
    def call_ollama(self, prompt: str) -> Optional[str]:
        """
        Call Ollama API with the given prompt.
        Returns the generated response or None if failed.
        """
        try:
            payload = {
                "model": self.model_name,
                "prompt": prompt,
                "stream": False,
                "options": {
                    "temperature": 0.1,  # Low temperature for consistent results
                    "top_p": 0.9,
                    "max_tokens": 512
                }
            }
            
            response = requests.post(
                f"{self.ollama_url}/api/generate",
                json=payload,
                timeout=30
            )
            
            if response.status_code == 200:
                result = response.json()
                return result.get("response", "").strip()
            else:
                self.logger.error(f"Ollama API error: {response.status_code}")
                return None
                
        except Exception as e:
            self.logger.error(f"Error calling Ollama: {e}")
            return None
    
    def create_mapping_prompt(self, description: str, template_items: List[str], 
                            section_context: str, statement_type: str) -> str:
        """
        Create a structured prompt for mapping financial line items.
        """
        prompt = f"""You are a financial statement mapping expert. Your task is to map a financial line item to the most appropriate template item.

STATEMENT TYPE: {statement_type.upper()}
SECTION: {section_context.upper()}

FINANCIAL LINE ITEM TO MAP: "{description}"

AVAILABLE TEMPLATE ITEMS:
{chr(10).join(f"- {item}" for item in template_items)}

INSTRUCTIONS:
1. Analyze the financial line item and find the best match from the template items
2. Consider synonyms, abbreviations, and common variations
3. Pay attention to the section context (e.g., assets vs liabilities)
4. If no good match exists, return "Other"
5. Provide a confidence score from 0.0 to 1.0

RESPONSE FORMAT (exact format required):
[template_item_name, confidence_score, reasoning]

Examples:
- "Accounts Receivable" → "Accounts Receivable" (high confidence)
- "Trade Receivables" → "Accounts Receivable" (high confidence) 
- "Net Receivables" → "Accounts Receivable" (medium confidence)
- "Unusual Item" → "Other" (low confidence)

Your response:"""

        return prompt
    
    def parse_llm_response(self, response: str) -> Tuple[Optional[str], float, str]:
        """
        Parse the LLM response to extract template item, confidence, and reasoning.
        Returns (template_item, confidence_score, reasoning)
        """
        try:
            # Clean up the response
            response = response.strip()
            
            # Look for the expected format: [item, confidence, reasoning]
            if response.startswith('[') and response.endswith(']'):
                # Extract content between brackets
                content = response[1:-1].strip()
                
                # Split by commas, but be careful with commas in the reasoning
                parts = content.split(',')
                if len(parts) >= 2:
                    template_item = parts[0].strip().strip('"\'')
                    confidence_str = parts[1].strip()
                    
                    # Extract confidence score
                    try:
                        confidence = float(confidence_str)
                        confidence = max(0.0, min(1.0, confidence))  # Clamp between 0 and 1
                    except ValueError:
                        confidence = 0.5  # Default confidence if parsing fails
                    
                    # Extract reasoning (everything after the second comma)
                    reasoning = ','.join(parts[2:]).strip().strip('"\'') if len(parts) > 2 else "No reasoning provided"
                    
                    return template_item, confidence, reasoning
            
            # Fallback parsing for different response formats
            lines = response.split('\n')
            for line in lines:
                line = line.strip()
                if '→' in line or '->' in line:
                    # Extract template item from arrow notation
                    parts = line.replace('→', '->').split('->')
                    if len(parts) >= 2:
                        template_item = parts[1].strip().strip('"\'')
                        return template_item, 0.7, "Parsed from arrow notation"
            
            # If all else fails, look for template items in the response
            for item in ["Other", "Cash and equivalents", "Accounts Receivable", "Inventory"]:
                if item.lower() in response.lower():
                    return item, 0.5, "Found in response text"
            
            return None, 0.0, "Could not parse response"
            
        except Exception as e:
            self.logger.error(f"Error parsing LLM response: {e}")
            return None, 0.0, f"Parsing error: {e}"
    
    def map_with_llm(self, description: str, template_items: List[str], 
                    section_context: str, statement_type: str) -> Tuple[Optional[str], float, str]:
        """
        Map a financial line item using LLM.
        Returns (template_item, confidence_score, reasoning)
        """
        if not self.check_ollama_available():
            return None, 0.0, "Ollama not available"
        
        prompt = self.create_mapping_prompt(description, template_items, section_context, statement_type)
        
        # Call LLM
        response = self.call_ollama(prompt)
        if not response:
            return None, 0.0, "LLM call failed"
        
        # Parse response
        template_item, confidence, reasoning = self.parse_llm_response(response)
        
        self.logger.info(f"LLM mapping: '{description}' -> '{template_item}' (confidence: {confidence:.2f})")
        
        return template_item, confidence, reasoning
    
    def get_llm_suggestions(self, unmapped_items: List[dict], template_items: List[str], 
                          section_context: str, statement_type: str) -> List[dict]:
        """
        Get LLM suggestions for unmapped items.
        Returns list of suggestions with template_item, confidence, and reasoning.
        """
        suggestions = []
        
        for item in unmapped_items:
            description = item.get('description', '')
            if not description:
                continue
            
            template_item, confidence, reasoning = self.map_with_llm(
                description, template_items, section_context, statement_type
            )
            
            if template_item and confidence > 0.3:  # Only include reasonable suggestions
                suggestions.append({
                    'description': description,
                    'value': item.get('value'),
                    'suggested_template_item': template_item,
                    'confidence': confidence,
                    'reasoning': reasoning
                })
            
            # Small delay to avoid overwhelming the LLM
            time.sleep(0.1)
        
        return suggestions

def main():
    """Test the LLM mapper"""
    mapper = LLMMapper()
    
    if not mapper.check_ollama_available():
        print("❌ Ollama not available. Please install and run Ollama with Mistral model.")
        print("Install: https://ollama.ai/")
        print("Run: ollama pull mistral && ollama run mistral")
        return
    
    print("✅ Ollama with Mistral is available!")
    
    # Test mapping
    test_cases = [
        ("Cash and cash equivalents", ["Cash and equivalents", "Other"], "current_assets", "balance_sheet"),
        ("Trade receivables", ["Accounts Receivable", "Other"], "current_assets", "balance_sheet"),
        ("Property, plant and equipment", ["Net PPE", "Other"], "noncurrent_assets", "balance_sheet"),
    ]
    
    for description, template_items, section, stmt_type in test_cases:
        print(f"\n--- Testing: {description} ---")
        result = mapper.map_with_llm(description, template_items, section, stmt_type)
        print(f"Result: {result}")

if __name__ == "__main__":
    main() 