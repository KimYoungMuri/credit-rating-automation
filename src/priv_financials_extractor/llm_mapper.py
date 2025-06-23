import requests
import json
import time
from typing import Tuple, Optional, List, Dict
import logging
import re

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
    
    def call_ollama(self, prompt: str, timeout: int = 15) -> Optional[str]:
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
                    "max_tokens": 256  # Reduced from 512 to speed up
                }
            }
            
            response = requests.post(
                f"{self.ollama_url}/api/generate",
                json=payload,
                timeout=timeout  # Reduced from 30 to 15 seconds
            )
            
            if response.status_code == 200:
                result = response.json()
                return result.get("response", "").strip()
            else:
                self.logger.error(f"Ollama API error: {response.status_code}")
                return None
                
        except requests.exceptions.Timeout:
            self.logger.warning(f"Ollama API timeout for prompt: {prompt[:50]}...")
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

    def create_batch_section_assignment_prompt(self, descriptions: List[str]) -> str:
        """Creates a prompt to ask the LLM to assign sections to a batch of items."""
        items_formatted = "\n".join([f'- "{desc}"' for desc in descriptions])
        
        prompt = f"""You are an expert financial analyst. Your task is to categorize a batch of financial line items into the correct section of a Balance Sheet.

Line Items:
{items_formatted}

Available Sections:
- current_assets
- noncurrent_assets
- current_liabilities
- noncurrent_liabilities
- equity

IMPORTANT INSTRUCTIONS:
1. Analyze each line item carefully. Pay attention to keywords and context.
2. Skip any total/subtotal lines (lines containing "Total", "Sum", "Subtotal", "Net" when referring to totals).
3. For ambiguous items, consider the financial context and typical placement.

SECTION GUIDELINES:
- current_assets: Items expected to be converted to cash within one year (cash, receivables, inventory, prepaid expenses, short-term investments)
- noncurrent_assets: Long-term assets (property/equipment, goodwill, intangibles, long-term investments, deferred assets)
- current_liabilities: Obligations due within one year (accounts payable, short-term debt, accrued expenses, current portions of long-term debt)
- noncurrent_liabilities: Long-term obligations (long-term debt, deferred taxes, pension obligations, long-term leases)
- equity: Ownership interests (common stock, retained earnings, paid-in capital, treasury stock, noncontrolling interests)

SPECIFIC EXAMPLES:
- "Notes receivable—current portion" → current_assets (it's a receivable, even if "current portion")
- "GOODWILL—Net" → noncurrent_assets (goodwill is always a long-term asset)
- "Long-term debt—current portion" → current_liabilities (current portion of long-term debt)
- "LONG-TERM INCENTIVE" → current_liabilities or noncurrent_liabilities (compensation obligation, not equity)
- "SUBCHAPTER S INCOME TAX DEPOSIT" → current_liabilities (tax obligation)
- "NONCONTROLLING INTERESTS" → equity (ownership interest)

Return a single JSON object where keys are the line item descriptions and values are the corresponding section from the list above.
IMPORTANT: Your entire response must be ONLY the JSON object, with no other text before or after it.

Example Response Format:
{{
  "Cash and cash equivalents": "current_assets",
  "GOODWILL—Net": "noncurrent_assets",
  "Long-term debt—current portion": "current_liabilities",
  "Notes receivable—current portion": "current_assets"
}}

Your response:"""
        return prompt

    def assign_sections_batch_with_llm(self, descriptions: List[str]) -> Optional[Dict[str, str]]:
        """Assigns sections to a batch of line items using a single LLM call."""
        if not self.check_ollama_available():
            self.logger.warning("Cannot assign sections with LLM, Ollama not available.")
            return None

        prompt = self.create_batch_section_assignment_prompt(descriptions)
        print(f"\n[DEBUG] Sending batch section assignment prompt to Ollama:")
        print(f"[DEBUG] Prompt length: {len(prompt)} characters")
        print(f"[DEBUG] First 500 chars of prompt: {prompt[:500]}...")
        
        try:
            response = requests.post(
                f"{self.ollama_url}/api/generate",
                json={
                    "model": self.model_name,
                    "prompt": prompt,
                    "stream": False,
                    "options": {
                        "temperature": 0.1,
                        "top_p": 0.9,
                        "max_tokens": 512
                    }
                },
                timeout=30  # Reduced from 120 to 30 seconds
            )
            print(f"[DEBUG] Ollama HTTP status: {response.status_code}")
            print(f"[DEBUG] Ollama raw response length: {len(response.text)} characters")
            print(f"[DEBUG] Ollama raw response: {response.text}")

            if response.status_code != 200:
                self.logger.error(f"Ollama API error: {response.status_code} - {response.text}")
                return None

            result = response.json()
            response_text = result.get("response", "").strip()
            print(f"[DEBUG] Ollama 'response' field length: {len(response_text)} characters")
            print(f"[DEBUG] Ollama 'response' field: {repr(response_text)}")

            # Attempt to parse the JSON response with multiple strategies
            assignments = None
            
            # Strategy 1: Look for JSON object with regex
            try:
                json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
                if json_match:
                    json_str = json_match.group(0)
                    print(f"[DEBUG] Found JSON match: {json_str}")
                    assignments = json.loads(json_str)
                    if isinstance(assignments, dict):
                        print(f"[DEBUG] Successfully parsed JSON dict with {len(assignments)} items")
                        return assignments
                    else:
                        print(f"[DEBUG] LLM returned valid JSON but not a dictionary: {type(assignments)}")
            except json.JSONDecodeError as e:
                print(f"[DEBUG] Strategy 1 failed: {e}")
            
            # Strategy 2: Try to clean up the response and parse as JSON
            try:
                # Remove any text before the first {
                cleaned_text = response_text
                brace_start = cleaned_text.find('{')
                if brace_start != -1:
                    cleaned_text = cleaned_text[brace_start:]
                
                # Remove any text after the last }
                brace_end = cleaned_text.rfind('}')
                if brace_end != -1:
                    cleaned_text = cleaned_text[:brace_end+1]
                
                print(f"[DEBUG] Strategy 2 cleaned text: {cleaned_text}")
                assignments = json.loads(cleaned_text)
                if isinstance(assignments, dict):
                    print(f"[DEBUG] Strategy 2 succeeded with {len(assignments)} items")
                    return assignments
            except json.JSONDecodeError as e:
                print(f"[DEBUG] Strategy 2 failed: {e}")
            
            # Strategy 3: Try to extract key-value pairs manually
            try:
                print(f"[DEBUG] Attempting manual parsing...")
                manual_assignments = {}
                # Look for patterns like "description": "section"
                pattern = r'"([^"]+)"\s*:\s*"([^"]+)"'
                matches = re.findall(pattern, response_text)
                print(f"[DEBUG] Found {len(matches)} key-value matches: {matches}")
                
                for key, value in matches:
                    if value in ['current_assets', 'noncurrent_assets', 'current_liabilities', 'noncurrent_liabilities', 'equity']:
                        manual_assignments[key] = value
                
                if manual_assignments:
                    print(f"[DEBUG] Manual parsing succeeded with {len(manual_assignments)} items")
                    return manual_assignments
            except Exception as e:
                print(f"[DEBUG] Strategy 3 failed: {e}")
            
            # If all strategies fail
            self.logger.error(f"All JSON parsing strategies failed for LLM response: {response_text}")
            return None
                
        except requests.exceptions.Timeout:
            self.logger.warning(f"Ollama API timeout for batch section assignment. Falling back to rule-based.")
            return None
        except requests.exceptions.ConnectionError:
            self.logger.warning(f"Ollama connection error. Falling back to rule-based.")
            return None
        except Exception as e:
            self.logger.error(f"Unexpected error calling Ollama: {e}")
            return None

    def create_section_assignment_prompt(self, description: str, statement_type: str) -> str:
        """Creates a prompt to ask the LLM to assign a section."""
        if statement_type == 'balance_sheet':
            sections = ['current_assets', 'noncurrent_assets', 'current_liabilities', 'noncurrent_liabilities', 'equity']
            prompt = f"""You are an expert financial analyst. Your task is to categorize a financial line item into the correct section of a Balance Sheet.

Line Item: "{description}"

Available Sections:
- current_assets
- noncurrent_assets
- current_liabilities
- noncurrent_liabilities
- equity

Instructions:
1. Analyze the line item carefully. Pay attention to keywords like "current", "noncurrent", "asset", "liability", "debt", "receivable", "payable", "equity", etc.
2. Return *only* the single most appropriate section name from the list above. Do not add any explanation.

Example 1:
Line Item: "Accounts receivable—net"
Response: current_assets

Example 2:
Line Item: "Long-term debt—current portion"
Response: current_liabilities

Example 3:
Line Item: "GOODWILL—Net"
Response: noncurrent_assets

Your response:"""
        else:
            # Placeholder for other statement types
            return ""
        return prompt

    def assign_section_with_llm(self, description: str, statement_type: str) -> Optional[str]:
        """Assigns a section to a line item using an LLM."""
        if not self.check_ollama_available():
            self.logger.warning("Cannot assign section with LLM, Ollama not available.")
            return None

        prompt = self.create_section_assignment_prompt(description, statement_type)
        if not prompt:
            return None

        response = self.call_ollama(prompt)
        if not response:
            return None

        # The response should be just the section name
        valid_sections = ['current_assets', 'noncurrent_assets', 'current_liabilities', 'noncurrent_liabilities', 'equity']
        cleaned_response = response.strip().lower().replace("response:", "").strip()
        
        # Take the first word of the response in case of extra text
        first_word = cleaned_response.split()[0]
        
        if first_word in valid_sections:
            return first_word
        else:
            self.logger.warning(f"LLM returned an invalid section '{cleaned_response}' for item '{description}'.")
            # Try to find a valid section in the response as a fallback
            for section in valid_sections:
                if section in cleaned_response:
                    return section
            return None

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