# Golf Player Performance Analyzer 🏌️‍♂️

An AI-powered analytics tool providing detailed golf performance analysis using OpenAI's GPT model.

## Features 🌟
- Data Analysis from Excel/CSV files
- AI-Powered Insights via OpenAI GPT
- Performance Comparisons
- Visual Analytics & Charts
- Export to PowerPoint
- Streamlit Interface

## Installation 🛠️

```bash
# Clone repository
git clone https://github.com/yourusername/golf-AI-analytics.git
cd golf-AI-analytics

# Setup environment
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Configure API
export OPENAI_API_KEY=your_api_key_here
```

## Usage 📊
1. Run: `streamlit run app.py`
2. Upload stats file (Excel/CSV)
3. Select analysis options
4. View/download results

## Requirements 📋
- Python 3.8+
- OpenAI API key
- Packages in `requirements.txt`

## Data Format 📈
- Player names column
- Numeric statistics columns
- One row per player

## Analysis Features 🎯
- Performance Summary
- Comparative Analysis
- Detailed Metrics
- Custom Recommendations

## Security 🔒
- Local API key storage
- No persistent data
- In-memory analysis

## Contributing 🤝
1. Fork repository
2. Create feature branch
3. Submit pull request

## Support 💬
- Open issues
- Contact maintainers
- Check documentation

## License 📝
MIT License

