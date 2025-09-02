# Semantic Spreadsheet Search Engine

A revolutionary search engine that understands the meaning behind spreadsheet data rather than just its structure. This system bridges the gap between human semantic thinking and spreadsheet structural reality.

# Features
1. Semantic Understanding: Recognizes business concepts, synonyms, and relationships

2. Intent Detection: Automatically classifies queries into LOOKUP, CALCULATION, EXPLANATION, or COMPARISON

3. Natural Language Processing: Understands queries like "find profitability metrics" or "show me sales trends"

4. Multi-Sheet Analysis: Works across multiple spreadsheet tabs with contextual understanding

5. Formula Intelligence: Understands and explains spreadsheet formulas and dependencies
   

# Installation

Prerequisites:
1. Python 3.8 or higher

2. pip (Python package manager)

Step-by-Step Setup:
1. Clone the repository <br><br>
     git clone https://github.com/Adarsh-k0802/semantic_search_script.git <br>
     cd semantic_search_script
   
2. Create a virtual environment<br><br>
   ->On Windows<br><br>
    python -m venv venv<br>
    venv\Scripts\activate<br>

    ->On macOS/Linux<br><br>
    python3 -m venv venv<br>
    source venv/bin/activate

3. Install dependencies<br><br>
    pip install -r requirements.txt

4. Set up environment variables<br><br>
     ->Create a .env file in the project root<br>
      echo "GOOGLE_API_KEY=your_google_api_key_here" > .env <br><br>

      Note: You'll need to obtain a Google API key for Gemini model access.

5. Run the application<br><br>
   python main.py


# Example Queries

1. Lookup: "find january target revenue", "what was Q3 sales"

2. Calculation: "total revenue", "average monthly sales"

3. Explanation: "how is profit calculated", "explain growth formula"

4. Comparison: "compare Q1 vs Q2 performance", "budget vs actual analysis"


# How It Works

1. **Spreadsheet Parsing**: Intelligently detects tables, headers, and cell relationships

2. **Semantic Chunking**: Creates meaningful data chunks with business context

3. **Vector Embedding**: Transforms content into numerical representations

4. **Intent Classification**: Determines user goals from natural language queries

5. **Semantic Search**: Finds conceptually related content, not just keyword matches

6. **Contextual Response**: Generanswers answers based on detected intent


# Supported Spreadsheet Formats

1. Microsoft Excel (.xlsx, .xls)

# Acknowledgments

1. Built with Python and powerful NLP libraries

2. Uses Google's Gemini model for advanced language understanding

3. FAISS for efficient vector similarity search

4. OpenPyXL for spreadsheet manipulation

