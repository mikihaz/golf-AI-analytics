import pandas as pd
import docx
from openai import OpenAI
import os
from openai import OpenAIError
import tiktoken
from pptx import Presentation
from template_analyzer import TemplateAnalyzer
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

def analyze_reference_ppt(file_path):
    """Extract structure and formatting from reference PPT."""
    prs = Presentation(file_path)
    template_info = {
        'slides': [],
        'charts': [],
        'metrics': set()
    }
    
    for slide in prs.slides:
        slide_info = {'type': 'content', 'elements': []}
        
        for shape in slide.shapes:
            if hasattr(shape, "chart"):
                slide_info['type'] = 'chart'
                chart_type = str(shape.chart.chart_type)
                chart_data = {
                    'type': chart_type,
                    'has_legend': shape.chart.has_legend,
                    'categories': []
                }
                template_info['charts'].append(chart_type)
                
            if hasattr(shape, "text"):
                # Extract potential metric keywords
                text = shape.text.lower()
                for keyword in ['performance', 'metrics', 'kpi', 'growth', 'rate', 
                              'revenue', 'sales', 'profit', 'percentage']:
                    if keyword in text:
                        template_info['metrics'].add(keyword)
                
                slide_info['elements'].append({
                    'type': 'text',
                    'content': shape.text
                })
        
        template_info['slides'].append(slide_info)
    
    return template_info

def analyze_chunk(client, chunk, template_info=None):
    """Analyze content with reference to template structure."""
    try:
        system_prompt = """You are a professional business analyst. Provide a comprehensive analysis with the following structure:

1. Executive Summary (with key highlights)
2. Detailed Analysis:
   - Key Performance Metrics (with specific numbers and percentages)
   - Growth Trends (provide % changes)
   - Performance Comparisons (with numerical data)
   - Risk Areas (quantified if possible)
3. Segment Analysis:
   - Break down performance by different categories
   - Market share percentages
   - Year-over-year comparisons
4. Recommendations:
   - Actionable insights with expected impact (in %)
   - Priority areas (with scoring 1-10)
   - Growth opportunities (with potential % gains)

Format all numerical data clearly as "Category: Number" or "Metric: XX%" for easy extraction.
Include at least 8-10 different metrics and percentages."""

        if template_info:
            metrics = ', '.join(template_info['metrics'])
            charts = ', '.join(template_info['charts'])
            system_prompt += f"\n\nFocus on these metrics: {metrics}"
            system_prompt += f"\nFormat data suitable for these chart types: {charts}"

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": chunk}
            ],
            temperature=0.7,  # Add some creativity while maintaining accuracy
            max_tokens=2000   # Allow for longer, more detailed responses
        )
        return response.choices[0].message.content
    except OpenAIError as e:
        return f"Error analyzing chunk: {str(e)}"

def process_document(file_path, reference_path=None):
    try:
        # Initialize template analyzer
        template_analyzer = TemplateAnalyzer()
        
        # Use static API key
        is_valid, client = validate_api_key(OPENAI_API_KEY)
        if not is_valid:
            return "Error: Invalid OpenAI API key. Please check the configuration."

        # Analyze reference template if provided
        template_data = None
        if reference_path:
            print("Learning from reference template...")
            template_data = template_analyzer.learn_from_template(reference_path)
            
            # Save patterns for future use
            patterns_file = "template_patterns.json"
            template_analyzer.save_patterns(patterns_file)

        # Analyze reference template if provided
        template_info = None
        if reference_path:
            print("Analyzing reference template...")
            template_info = analyze_reference_ppt(reference_path)

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

        # Process content in chunks with template info
        chunks = chunk_content(content)
        print(f"Content split into {len(chunks)} chunks")
        
        def analyze_with_template(client, chunk):
            """Analyze content using template-based prompt"""
            system_prompt = template_analyzer.generate_prompt(template_data) if template_data else \
                          "Analyze the content and provide structured insights with numerical data."
            
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": chunk}
                ]
            )
            return response.choices[0].message.content

        all_analyses = []
        for i, chunk in enumerate(chunks, 1):
            print(f"Processing chunk {i} of {len(chunks)}...")
            chunk_analysis = analyze_with_template(client, chunk)
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
