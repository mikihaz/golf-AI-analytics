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
    try:
        player_column = 'Player_Name'
        player_data = df[df[player_column] == selected_player]
        
        # Get handicap group statistics
        player_handicap = float(player_data['Handicap'].iloc[0])
        handicap_group = df[
            (df['Handicap'] >= player_handicap - 2) & 
            (df['Handicap'] <= player_handicap + 2)
        ]
        
        # Calculate handicap group stats for each hole
        handicap_group_stats = {
            col: handicap_group[col].mean() 
            for col in df.columns 
            if col.startswith('H_') and col.endswith('_GS')
        }
        
        # Add time of day analysis
        try:
            df['Hour'] = pd.to_datetime(df['Date'] + ' ' + df['Tee_Off']).dt.hour
            morning_rounds = player_data[df['Hour'] < 12]
            afternoon_rounds = player_data[df['Hour'] >= 12]
            
            time_analysis = {
                'morning_stats': {
                    'avg_score': round(float(morning_rounds['Tot_par'].mean()), 2) if not morning_rounds.empty else None,
                    'rounds_played': len(morning_rounds),
                    'best_score': round(float(morning_rounds['Tot_par'].min()), 2) if not morning_rounds.empty else None
                },
                'afternoon_stats': {
                    'avg_score': round(float(afternoon_rounds['Tot_par'].mean()), 2) if not afternoon_rounds.empty else None,
                    'rounds_played': len(afternoon_rounds),
                    'best_score': round(float(afternoon_rounds['Tot_par'].min()), 2) if not afternoon_rounds.empty else None
                }
            }
        except Exception as e:
            time_analysis = None

        metrics_data = {
            'holes': {col: {
                'score': player_data[col].iloc[0],
                'field_avg': df[col].mean(),
                'handicap_group_avg': handicap_group_stats.get(col)
            } for col in df.columns if col.startswith('H_') and col.endswith('_GS')},
            'handicap': player_handicap,
            'total_par': float(player_data['Tot_par'].iloc[0]),
            'handicap_group': {
                'size': len(handicap_group),
                'avg_total_par': float(handicap_group['Tot_par'].mean()),
                'range': f"{player_handicap-2} to {player_handicap+2}",
                'best_total_par': float(handicap_group['Tot_par'].min()),
                'worst_total_par': float(handicap_group['Tot_par'].max())
            },
            'time_analysis': time_analysis
        }

        analysis_prompt = f"""You are a professional golf analyst. Analyze this player's performance data:

Player: {selected_player}

Available Data:
{json.dumps(metrics_data, indent=2)}

For time analysis, Games with tee time before 10 am are morning games. Rest are noon games. In noon game — since sun is up and since it’s winter time the night dew goes away. This makes a ball roll extra distance.

Provide a detailed analysis following this structure:
1. Overall Performance Summary:
   - FORMAT AS: "Total Par: [value]"
   - FORMAT AS: "Handicap: [value]"
   - FORMAT AS: "vs Handicap Group Avg: [value]"

2. Time of Day Analysis (IMPORTANT):
   - FORMAT AS: "Morning Average: [value]"
   - FORMAT AS: "Afternoon Average: [value]"
   - FORMAT AS: "Preferred Time: [Morning/Afternoon]"
   - FORMAT AS: "Time Performance Gap: [value]"
   - Analyze scoring patterns by time of day
   - Identify optimal tee time preference
   - Compare morning vs afternoon consistency

3. Handicap Group Comparison:
   - FORMAT AS: "Handicap Group Range: [range]"
   - FORMAT AS: "Group Size: [number] players"
   - FORMAT AS: "Player vs Group Avg: [difference]"
   - Analyze performance relative to handicap group
   - Identify key performance gaps

4. Hole-by-Hole Analysis:
   - FORMAT AS: "Hole X: [player score] (Handicap Group Avg: [avg], Field Avg: [avg])"
   - Show score differences
   - Identify patterns

5. Key Findings:
   - FORMAT AS: "Finding: [description]"
   - Include time of day insights
   - Compare with handicap group
   - Highlight strongest periods of play

6. Recommendations:
   - FORMAT AS: "Recommendation: [action] (Impact: XX%)"
   - Suggest optimal tee times
   - Provide specific practice focus areas
   - Include time-based strategies

Requirements:
- EMPHASIZE both time of day and handicap analysis
- All numerical data MUST be formatted as "Metric: Value"
- Include specific morning vs afternoon comparisons
- Show percentage differences
- Provide clear tee time preferences

Focus Areas:
- Morning vs Afternoon performance
- Handicap group comparison
- Time-based scoring patterns
- Optimal playing conditions
- Strategic tee time selection"""

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a professional golf analyst specializing in statistical analysis and performance improvement recommendations."},
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
