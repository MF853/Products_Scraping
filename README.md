# Products_Scraping

## Overview

**Products_Scraping** is a Python application that scrapes product prices from various online stores and exports the data into an Excel file. This tool allows users to automate price monitoring across multiple websites, gathering up-to-date pricing information for comparison.

## Features

- **Scrapes product prices** from different e-commerce websites.
- **Exports data to Excel** using the `openpyxl` library.
- **Supports multiple stores**: easily adaptable to different online store structures.
- **Easy to extend**: Add new store URLs and product selectors as needed.

## Technologies Used

- **Python 3.8+**
- **BeautifulSoup** for parsing HTML
- **Selenium** for browser automation
- **Requests** for making HTTP requests
- **Openpyxl** for Excel file manipulation

## Requirements

Install the dependencies listed in the `requirements.txt` file using pip:

```bash
pip install -r requirements.txt
```

### Major Dependencies:

- `beautifulsoup4`: Used to parse and extract data from HTML pages.
- `selenium`: Automates browser actions to handle dynamic websites.
- `openpyxl`: Manages Excel files to store the scraped product data.
- `requests`: Sends HTTP requests to fetch page contents.
- `pandas`: For data manipulation before exporting to Excel.

## Setup

### 1. Install Python

Ensure you have Python 3.8 or higher installed on your system. You can download Python [here](https://www.python.org/downloads/).

### 2. Install Dependencies

Run the following command to install all the required libraries:

```bash
pip install -r requirements.txt
```

### 3. Configure Selenium

You'll need a WebDriver (e.g., ChromeDriver or GeckoDriver for Firefox) to use Selenium. You can install the necessary driver using the `webdriver-manager` library, which is already included in the requirements file.

## Usage

1. **Run the Scraper**:

   Execute the main script to start scraping:

   ```bash
   python script.py
   ```

2. **Output**:

   The scraped data is exported to an Excel file, which contains:
   
   - Store name
   - Product name
   - Price
