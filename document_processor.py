import pandas as pd
import docx
from openai import OpenAI
import os
from openai import OpenAIError
import tiktoken
from pptx import Presentation
from config import OPENAI_API_KEY

def validate_api_key(api_key):
    try:
        client = OpenAI(api_key=api_key)
        client.models.list()
        return True, client
    except OpenAIError as e:
        print(f"API Key validation failed: {str(e)}")
        return False, None

def read_docx(file_path):
    doc = docx.Document(file_path)
    text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
    return text

def read_excel(file_path):
    df = pd.read_excel(file_path)
    return df.to_string()

def get_token_count(text, model="gpt-3.5-turbo"):
    """Count tokens in a text string."""
    encoding = tiktoken.encoding_for_model(model)
    return len(encoding.encode(text))

def chunk_content(text, max_tokens=14000):
    """Split content into chunks that fit within token limits."""
    chunks = []
    current_chunk = ""
    current_tokens = 0
    
    # Split text into paragraphs
    paragraphs = text.split('\n')
    
    for paragraph in paragraphs:
        paragraph_tokens = get_token_count(paragraph)
        
        # If single paragraph exceeds limit, split it into smaller pieces
        if paragraph_tokens > max_tokens:
            words = paragraph.split()
            temp_chunk = ""
            for word in words:
                if get_token_count(temp_chunk + " " + word) < max_tokens:
                    temp_chunk += " " + word
                else:
                    chunks.append(temp_chunk.strip())
                    temp_chunk = word
            if temp_chunk:
                chunks.append(temp_chunk.strip())
            continue
            
        # Check if adding paragraph exceeds limit
        if current_tokens + paragraph_tokens < max_tokens:
            current_chunk += "\n" + paragraph
            current_tokens += paragraph_tokens
        else:
            chunks.append(current_chunk.strip())
            current_chunk = paragraph
            current_tokens = paragraph_tokens
    
    if current_chunk:
        chunks.append(current_chunk.strip())
    
    return chunks

def analyze_chunk(client, chunk, template_info=None):
    """Analyze content with reference to template structure."""
    try:
        system_prompt = """You are a professional business analyst. Provide a comprehensive analysis with the following structure, ensuring all numerical data is clearly formatted:

1. Executive Summary
2. Key Metrics (each on new line, strictly in format "MetricName: Number" or "Category: XX%"):
   - Revenue: 1234567
   - Growth Rate: 25%
   - Market Share: 45%
   - Customer Count: 5000
3. Trend Analysis (each on new line, in format "Trend: Number"):
   - Q1 Growth Trend: 15
   - Q2 Growth Trend: 25
   - Q3 Growth Trend: 35
4. Segment Analysis (each on new line, in format "Segment: Number"):
   - Enterprise Segment: 45
   - SMB Segment: 30
   - Consumer Segment: 25
5. Performance Metrics (each on new line, in format "Metric: Number"):
   - Sales Performance: 85
   - Customer Satisfaction: 92
   - Market Penetration: 78
6. Recommendations (each with impact percentage):
   - Recommendation 1 (Impact: 30%)
   - Recommendation 2 (Impact: 25%)

Ensure EVERY numerical value is presented in the exact format specified above for proper chart generation."""

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": chunk}
            ],
            temperature=0.7,
            max_tokens=2000
        )
        return response.choices[0].message.content
    except OpenAIError as e:
        return f"Error analyzing chunk: {str(e)}"

def process_document(file_path):
    try:
        # Use static API key
        is_valid, client = validate_api_key(OPENAI_API_KEY)
        if not is_valid:
            return "Error: Invalid OpenAI API key. Please check the configuration."

        # Debug information
        file_extension = os.path.splitext(file_path)[1].lower()
        print(f"Processing file with extension: {file_extension}")
        
        if file_extension in ['.xlsx', '.csv']:
            try:
                if file_extension == '.xlsx':
                    content = pd.read_excel(file_path, engine='openpyxl').to_string()
                else:
                    content = pd.read_csv(file_path).to_string()
                print("Successfully read Excel/CSV file")
            except Exception as e:
                print(f"Error reading Excel/CSV file: {str(e)}")
                return f"Error processing file: {str(e)}"
        elif file_extension == '.docx':
            try:
                content = read_docx(file_path)
                print("Successfully read DOCX file")
            except Exception as e:
                print(f"Error reading DOCX file: {str(e)}")
                return f"Error processing file: {str(e)}"
        else:
            print(f"Unsupported file format: {file_extension}")
            return f"Unsupported file format: {file_extension}"

        # Process content in chunks
        chunks = chunk_content(content)
        print(f"Content split into {len(chunks)} chunks")

        all_analyses = []
        for i, chunk in enumerate(chunks, 1):
            print(f"Processing chunk {i} of {len(chunks)}...")
            chunk_analysis = analyze_chunk(client, chunk)
            all_analyses.append(chunk_analysis)

        # Combine all analyses
        if len(all_analyses) == 1:
            return all_analyses[0]
        
        # If multiple chunks, summarize them
        combined_analysis = "\n\n=== Combined Analysis ===\n\n"
        for i, analysis in enumerate(all_analyses, 1):
            combined_analysis += f"\nSection {i}:\n{analysis}\n"
            
        # Generate final summary
        try:
            summary_response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "Summarize the following analyses into a concise, coherent summary:"},
                    {"role": "user", "content": combined_analysis}
                ]
            )
            return summary_response.choices[0].message.content
        except OpenAIError as api_error:
            return combined_analysis  # Fallback to combined analysis if summary fails
                
    except Exception as e:
        print(f"Error in process_document: {str(e)}")
        return f"Error processing document: {str(e)}"
