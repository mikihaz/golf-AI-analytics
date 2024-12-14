import pandas as pd
import docx
from openai import OpenAI
import os
from openai import OpenAIError
import tiktoken
from pptx import Presentation
from config import OPENAI_API_KEY
import json  # Add this import at the top with other imports

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

def get_players_list(file_path):
    """Extract list of players from the Excel/CSV file."""
    try:
        if file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        else:
            df = pd.read_csv(file_path)
        
        # Assuming player names are in a column named 'Player' or similar
        player_column = next(col for col in df.columns if 'player' in col.lower())
        players = df[player_column].unique().tolist()
        return players
    except Exception as e:
        print(f"Error getting players list: {str(e)}")
        return []

def get_player_column(df):
    """Identify the player column name in the dataframe."""
    possible_names = ['Player', 'player', 'Name', 'name', 'PLAYER', 'PlayerName', 'Player Name']
    
    for name in possible_names:
        if name in df.columns:
            return name
            
    # If no exact match, try partial matches
    for col in df.columns:
        if any(name.lower() in col.lower() for name in ['player', 'name']):
            return col
            
    raise ValueError("Could not find player column in the data")

def analyze_player_performance(client, df, selected_player):
    """Analyze specific player's performance compared to others."""
    try:
        # Get the correct player column name
        player_column = next(col for col in df.columns if 'player' in col.lower())
        
        # Verify player exists in the data
        if selected_player not in df[player_column].values:
            return f"Error: Player '{selected_player}' not found in the data"
        
        # Filter numeric columns only and remove columns with all NaN values
        numeric_columns = df.select_dtypes(include=['int64', 'float64']).columns.tolist()
        numeric_columns = [col for col in numeric_columns if not df[col].isna().all()]
        
        # Calculate player stats
        player_data = df[df[player_column] == selected_player]
        
        # Handle NaN values in player stats
        player_stats = {}
        for col in numeric_columns:
            val = player_data[col].iloc[0]
            player_stats[col] = round(float(val), 2) if pd.notna(val) else None
        
        # Remove None/NaN values from stats
        player_stats = {k: v for k, v in player_stats.items() if v is not None}
        
        # Calculate overall stats handling NaN values
        all_players_stats = {
            'mean': df[numeric_columns].mean().apply(lambda x: round(float(x), 2) if pd.notna(x) else None).to_dict(),
            'max': df[numeric_columns].max().apply(lambda x: round(float(x), 2) if pd.notna(x) else None).to_dict(),
            'min': df[numeric_columns].min().apply(lambda x: round(float(x), 2) if pd.notna(x) else None).to_dict()
        }
        
        # Calculate rankings and percentiles only for non-NaN values
        rankings = {}
        percentiles = {}
        for col in numeric_columns:
            if pd.notna(player_data[col].iloc[0]):
                # Filter out NaN values for ranking calculations
                valid_data = df[df[col].notna()]
                if not valid_data.empty:
                    rankings[col] = int(valid_data[col].rank(ascending=False)[player_data.index[0]])
                    percentiles[col] = int(100 * (len(valid_data[valid_data[col] <= player_data[col].iloc[0]]) / len(valid_data)))
        
        # Clean up stats dictionaries to remove any remaining None values
        all_players_stats = {k: {k2: v2 for k2, v2 in v.items() if v2 is not None} 
                           for k, v in all_players_stats.items()}
        
        # Create analysis prompt
        analysis_prompt = f"""Analyze the following golf player statistics and provide a detailed comparison:

Player: {selected_player}

Player's Current Statistics:
{json.dumps(player_stats, indent=2)}

Rankings (out of {len(df)} players):
{json.dumps(rankings, indent=2)}

Percentile Rankings:
{json.dumps(percentiles, indent=2)}

Statistical Context:
- Mean: {json.dumps(all_players_stats['mean'], indent=2)}
- Best: {json.dumps(all_players_stats['max'], indent=2)}
- Worst: {json.dumps(all_players_stats['min'], indent=2)}

Provide a detailed analysis with the following structure:
1. Overall Performance Summary
2. Key Strengths (top 25% percentile metrics)
3. Areas for Improvement (bottom 25% percentile metrics)
4. Comparative Analysis with Average Players
5. Statistical Highlights
6. Specific Recommendations for Improvement

Use exact numbers and percentages in your analysis.
Format all metrics as "Metric Name: Value" for proper chart generation.
Include percentile rankings in the analysis."""

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a professional golf analyst. Provide detailed insights comparing the player's performance with others."},
                {"role": "user", "content": analysis_prompt}
            ],
            temperature=0.7,
            max_tokens=2000
        )
        
        return response.choices[0].message.content
    except Exception as e:
        return f"Error analyzing player performance: {str(e)}\nDataframe columns: {', '.join(df.columns.tolist())}"

def process_document(file_path, selected_player=None):
    try:
        is_valid, client = validate_api_key(OPENAI_API_KEY)
        if not is_valid:
            return "Error: Invalid OpenAI API key. Please check the configuration."

        file_extension = os.path.splitext(file_path)[1].lower()
        
        if file_extension in ['.xlsx', '.csv']:
            try:
                if file_extension == '.xlsx':
                    df = pd.read_excel(file_path)
                else:
                    df = pd.read_csv(file_path)
                
                if selected_player:
                    # Perform player-specific analysis
                    return analyze_player_performance(client, df, selected_player)
                else:
                    # Return list of players
                    return get_players_list(file_path)
                    
            except Exception as e:
                return f"Error processing file: {str(e)}"
        else:
            return f"Unsupported file format: {file_extension}"
            
    except Exception as e:
        return f"Error processing document: {str(e)}"
