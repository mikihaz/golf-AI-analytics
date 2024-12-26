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
        player_column = next(
            col for col in df.columns if 'player' in col.lower())
        players = df[player_column].unique().tolist()
        return players
    except Exception as e:
        print(f"Error getting players list: {str(e)}")
        return []


def get_player_column(df):
    """Identify the player column name in the dataframe."""
    possible_names = ['Player', 'player', 'Name',
                      'name', 'PLAYER', 'PlayerName', 'Player Name']

    for name in possible_names:
        if name in df.columns:
            return name

    # If no exact match, try partial matches
    for col in df.columns:
        if any(name.lower() in col.lower() for name in ['player', 'name']):
            return col

    raise ValueError("Could not find player column in the data")


def analyze_player_performance(client, df, selected_player):
    hole_index_data = [
        {"hole_no": 1, "par": 4, "stroke_index": 11},
        {"hole_no": 2, "par": 3, "stroke_index": 17},
        {"hole_no": 3, "par": 4, "stroke_index": 3},
        {"hole_no": 4, "par": 5, "stroke_index": 13},
        {"hole_no": 5, "par": 4, "stroke_index": 9},
        {"hole_no": 6, "par": 4, "stroke_index": 5},
        {"hole_no": 7, "par": 4, "stroke_index": 1},
        {"hole_no": 8, "par": 4, "stroke_index": 15},
        {"hole_no": 9, "par": 4, "stroke_index": 7},
        {"hole_no": 10, "par": 4, "stroke_index": 2},
        {"hole_no": 11, "par": 4, "stroke_index": 8},
        {"hole_no": 12, "par": 4, "stroke_index": 14},
        {"hole_no": 13, "par": 3, "stroke_index": 16},
        {"hole_no": 14, "par": 4, "stroke_index": 4},
        {"hole_no": 15, "par": 5, "stroke_index": 12},
        {"hole_no": 16, "par": 4, "stroke_index": 18},
        {"hole_no": 17, "par": 4, "stroke_index": 10},
        {"hole_no": 18, "par": 4, "stroke_index": 6}
    ]
    try:
        analysis_prompt = f"""You are an elite PGA-level golf performance analyst specializing in statistical analysis and player improvement. Analyze this player's detailed performance data:

Provided All Score Data:
////START////
{df.to_string()}
////END////

Provided Golf index Data:
////START////
{json.dumps(hole_index_data)}
////END////

Selected Player: {selected_player}

Course Conditions Context:
- Morning Rounds (Before 10 AM): Typically slower greens due to morning dew, cooler temperatures affecting ball flight
- Afternoon Rounds: Faster green speeds, wind patterns more variable, drier conditions affecting roll distance
- Winter Conditions: Afternoon sun removes dew effect, potentially increasing roll distance by 10-15%

Provide a comprehensive technical analysis following this structure:

1. Overall Performance Metrics:
   - FORMAT AS: "Total Par: [value]" (Include + or - relative to course par)
   - FORMAT AS: "Handicap Index: [value]" (Include trend direction)
   - FORMAT AS: "Strokes Gained vs Handicap Group: [value]"
   - FORMAT AS: "Scoring Average: [value]"
   - FORMAT AS: "GIR Percentage: [value]%"

2. Time of Day Performance Analysis (CRITICAL):
   - FORMAT AS: "AM Scoring Average: [value]"
   - FORMAT AS: "PM Scoring Average: [value]"
   - FORMAT AS: "Optimal Playing Window: [Morning/Afternoon]"
   - FORMAT AS: "Performance Delta: [value] strokes"
   - Detailed green speed impact analysis
   - Wind pattern adaptation metrics
   - Temperature impact on distance control

3. Technical Handicap Analysis:
   - FORMAT AS: "Handicap Peer Group: [range]"
   - FORMAT AS: "Statistical Peer Group Size: [number] players"
   - FORMAT AS: "Strokes Gained/Lost vs Peer Group: [value]"
   - Comparative shot pattern analysis
   - Scoring distribution on par 3s/4s/5s
   - Course management efficiency rating

4. Hole-by-Hole Statistical Breakdown:
   - FORMAT AS: "Hole [X] ([Par]): [player score] (Peer Avg: [avg], Field Avg: [avg])"
   - FORMAT AS: "Total Pars: [number]"
   - FORMAT AS: "Double Bogeys or Worse: [number]"
   - Time of Day Analysis:
      * FORMAT AS: "Morning Stats - Hole [X]:"
        - Average Score: [value]
        - Total Pars: [number]
        - Double Bogeys or Worse: [number]
      * FORMAT AS: "Afternoon Stats - Hole [X]:"
        - Average Score: [value]
        - Total Pars: [number]
        - Double Bogeys or Worse: [number]
   - Risk/reward decision points
   - Shot distribution patterns
   - Critical scoring opportunities
   - Recovery shot efficiency

5. Round-by-Round Performance:
   - FORMAT AS: "Game [Date] [Time]:"
     * Gross Score: [value]
     * Total Pars: [number]
     * Total Bogeys: [number]
     * Total Double Bogeys or Worse: [number]
   - Include morning/afternoon split
   - Trend analysis
   - Pattern identification

6. Key Performance Insights:
   - FORMAT AS: "Strategic Finding: [detailed description]"
   - Time-based performance variations including:
     * Morning vs Afternoon par conversion rates
     * Time-specific double bogey patterns
     * Scoring distribution by time of day
   - Course management decisions
   - Scoring pattern anomalies
   - Statistical strengths/weaknesses
   - Weather impact correlations

7. Professional Development Recommendations:
   - FORMAT AS: "Technical Recommendation: [specific action] (Projected Impact: XX%)"
   - Optimal tee time strategy
   - Shot selection modifications
   - Practice priority areas
   - Course management adjustments
   - Environmental adaptation strategies

Technical Analysis Requirements:
- Emphasize strokes gained/lost metrics
- Include detailed morning vs afternoon statistical comparisons
- Analyze scoring patterns relative to playing conditions
- Evaluate decision-making efficiency
- Quantify performance under varying conditions
- Provide specific practice protocols

Key Analysis Points:
- Shot pattern distribution
- Scoring efficiency by hole type
- Time-of-day performance correlation
- Weather impact assessment
- Course management decision points
- Statistical trend analysis
- Performance optimization opportunities
- Risk/reward efficiency metrics

Ensure all numerical data follows strict formatting for accurate statistical tracking and trend analysis.

Additional Analysis Requirements:
- Calculate and highlight total pars for each hole and overall rounds
- Track double bogey or worse frequency by hole and time of day
- Compare morning vs afternoon performance for each statistical category
- Identify patterns in scoring distribution across different rounds
- Analyze par conversion rates by time of day
- Track progression of double bogey avoidance"""

        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a professional golf analyst specializing in statistical analysis and performance improvement recommendations."},
                {"role": "user", "content": analysis_prompt}
            ],
            temperature=0.7,
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
