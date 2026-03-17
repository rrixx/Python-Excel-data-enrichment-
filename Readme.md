The project is split into two specialized modules:

### 1. Contact & Lead Enrichment

- **Automated Website Discovery**: Finds the official website of the agency via Google Maps.
    
- **Intelligent Email Scraping**: Navigates discovered websites and extracts valid business emails using regex (while bypassing common garbage data like `.png` or `.jpg` links).
    
- **LinkedIn Matching**: Specifically targets LinkedIn profiles for intermediaries to facilitate professional networking.
    
- **Lead Info**: Retrieves official phone numbers directly from Google Business profiles.
    

### 2. Precise Geocoding

- **Coordinate Extraction**: Fetches exact **Latitude** and **Longitude** for every entry.
    
- **Address Normalization**: Cleans corporate suffixes (e.g., _SRL, SPA, SNC_) to ensure higher match accuracy during Google Maps queries.
    

### ⚡ Performance & Efficiency

- **Smart Caching**: Implements a local cache based on the RUI registration number to prevent redundant API calls for recurring entries, significantly reducing costs and execution time.
    
- **SSL Resilience**: Built-in support for scraping websites with outdated or invalid SSL certificates.
    
- **Bulk Processing**: Native support for multi-sheet Excel files.


## Tech Stack

- **Language**: Python 3.x
    
- **APIs**: [Serper.dev](https://serper.dev/) (Google Search & Places)
    
- **Libraries**:
    
    - `openpyxl`: For advanced Excel I/O.
        
    - `BeautifulSoup4`: For HTML parsing and email extraction.
        
    - `requests`: For API interaction and web navigation.
        
    - `re` & `json`: For data processing.

## Getting Started

### Prerequisites

1. A **Serper.dev API Key**.
    
2. An input Excel file named `rui_intermediari_A_B_E_collaborazioni.xlsx` containing the columns: `Nome Intermediario`, `Città`, and `Numero Iscrizione RUI`.
    

### Installation

1. Clone the repository:
    
    Bash
    
    ```
    git clone https://github.com/your-username/rui-data-enricher.git
    cd rui-data-enricher
    ```
    
2. Install dependencies:
    
    Bash
    
    ```
    pip install requests openpyxl beautifulsoup4 urllib3
    ```
    

### Configuration

Update the `SERPER_API_KEY` variable in the scripts:

Python

```
SERPER_API_KEY = "YOUR_API_KEY_HERE"
```