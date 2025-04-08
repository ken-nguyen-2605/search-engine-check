# Virtual Treasury Web Scraper

A tool that scrapes Virtual Treasury search results and checks for entity mentions in the content.

## Installation

1. Clone or download this repository to your local machine.

2. Create a virtual environment (recommended):

    ```
    python -m venv venv
    ```

3. Activate the virtual environment:

    - On Windows:
        ```
        venv\Scripts\activate
        ```
    - On macOS/Linux:
        ```
        source venv/bin/activate
        ```

4. Install required dependencies:

    ```
    pip install -r requirements.txt
    ```

5. Install a compatible ChromeDriver:
    - Download the ChromeDriver that matches your Chrome browser version from [https://sites.google.com/chromium.org/driver/](https://sites.google.com/chromium.org/driver/)
    - Extract the executable and place it in your PATH or in the project directory

## Configuring Questions Data

The search queries and entities to find are defined in the `questions_data.py` file.

1. Open `questions_data.py` in a text editor
2. Modify the `question_tuple` list with your search terms and entities
3. Each item should be a tuple with two elements:
    - First element: The search phrase to query
    - Second element: The entity to look for in the results

Example:

```python
question_tuple = [
    ("Edward I's involvement in Irish governance", "Edward I"),
    ("Medieval castle construction techniques", "stone masonry"),
    ("History of Dublin parliament", "Irish parliament"),
    # Add more search items here...
]
```

You should add at least 25 tuples to fully utilize the tool. Each tuple will generate a search on Virtual Treasury, and the script will check if the specified entity appears in the search results.

## Running the Tool

To run the web scraper:

1. Make sure your virtual environment is activated
2. Execute the script with Python:
    ```
    python web_scrape.py
    ```

The script will:

1. Open a Chrome browser window (make sure you have Chrome installed)
2. For each search term in `question_tuple`:
    - Search on virtualtreasury.ie
    - Process up to 10 results per search
    - Check if the entity is present in specific sections
3. Save results to an Excel file named `virtual_treasury_search_results.xlsx`

## Understanding the Results

The Excel file contains:

-   One row per search query
-   Columns for each search result (up to 10 per query)
-   For each result:
    -   "y" if the entity was found exactly as specified
    -   Otherwise, a reason explaining why it wasn't an exact match
    -   The URL of the result page

## Troubleshooting

-   **Browser crashes or doesn't open**: Make sure you have Chrome installed and the correct ChromeDriver version.
-   **Script runs slowly**: The script includes delays to avoid overloading the website. You can adjust the `time.sleep()` values in the code.
-   **No results found**: Check your search queries and make sure they're likely to return results on the Virtual Treasury website.
-   **Selenium errors**: Update your ChromeDriver and Selenium package to compatible versions.

## Notes

-   The script opens a visible Chrome window by default. To run in headless mode (no visible window), uncomment the headless option in the code.
-   The tool assumes the Virtual Treasury website structure remains unchanged. Website updates may require code modifications.
