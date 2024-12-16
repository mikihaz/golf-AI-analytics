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
        player_column = 'Player_Name'  # Match your data structure
        
        # Verify player exists in the data
        if selected_player not in df[player_column].values:
            return f"Error: Player '{selected_player}' not found in the data"
        
        # Calculate player stats
        player_data = df[df[player_column] == selected_player]
        
        # Add hole-by-hole analysis
        hole_columns = sorted([col for col in df.columns if col.startswith('H_') and col.endswith('_GS')])
        hole_stats = {}
        for hole in hole_columns:
            hole_num = int(hole.split('_')[1])
            player_score = player_data[hole].mean()
            field_average = df[hole].mean()
            hole_stats[f"Hole {hole_num}"] = {
                'player_score': round(float(player_score), 2),
                'field_average': round(float(field_average), 2),
                'difference': round(float(player_score - field_average), 2)
            }
        
        # Add time of day analysis (modified to handle time strings)
        try:
            df['Hour'] = pd.to_datetime(df['Date'] + ' ' + df['Tee_Off']).dt.hour
            morning_rounds = player_data[df['Hour'] < 12]
            afternoon_rounds = player_data[df['Hour'] >= 12]
            
            time_analysis = {
                'morning_avg': round(float(morning_rounds['Tot_par'].mean()), 2) if not morning_rounds.empty else 0,
                'afternoon_avg': round(float(afternoon_rounds['Tot_par'].mean()), 2) if not afternoon_rounds.empty else 0,
                'morning_rounds': len(morning_rounds),
                'afternoon_rounds': len(afternoon_rounds)
            }
        except Exception as e:
            # Fallback if time analysis fails
            time_analysis = {
                'morning_avg': 0,
                'afternoon_avg': 0,
                'morning_rounds': 0,
                'afternoon_rounds': 0
            }
        
        # Add handicap comparison
        player_handicap = float(player_data['Handicap'].iloc[0])
        similar_handicaps = df[
            (df['Handicap'] >= player_handicap - 2) & 
            (df['Handicap'] <= player_handicap + 2) & 
            (df['Player_Name'] != selected_player)
        ]
        
        handicap_analysis = {
            'player_avg': round(float(player_data['Tot_par'].mean()), 2),
            'similar_handicap_avg': round(float(similar_handicaps['Tot_par'].mean()), 2),
            'handicap_range': f"{player_handicap-2} to {player_handicap+2}",
            'similar_players_count': len(similar_handicaps['Player_Name'].unique())
        }

        # Update analysis prompt
        analysis_prompt = f"""Analyze the following golf player statistics and provide a detailed comparison:

Player: {selected_player}

Hole-by-Hole Analysis:
{json.dumps(hole_stats, indent=2)}

Handicap Comparison:
{json.dumps(handicap_analysis, indent=2)}

Player's Current Statistics:
- Total Par: {player_data['Tot_par'].iloc[0]}
- Handicap: {player_handicap}

Hole Performance Metrics:
{"".join(f"Hole {i}: {hole_stats[f'Hole {i}']['player_score']} (Field Avg: {hole_stats[f'Hole {i}']['field_average']})\n" for i in range(1, 19))}

Provide a detailed analysis with the following structure:
1. Overall Performance Summary
2. Hole-by-Hole Performance:
   - List each hole's average score compared to field average
   - Identify strongest and weakest holes
   - Highlight notable patterns
3. Handicap Group Comparison:
   - Compare performance against similar handicap players
   - Identify areas where player outperforms or underperforms their handicap group
4. Key Strengths
5. Areas for Improvement
6. Specific Recommendations

Format all numerical data as "Metric: Value" for proper chart generation.
Include specific hole scores and handicap comparisons."""

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
