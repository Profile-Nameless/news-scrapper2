# Google News Scraper and PDF Generator

This Python script is designed to scrape news articles related to a specific query from Google News, store the data in an Excel file, and convert each article to a PDF.

## Features

1. **News Search**: Searches Google News for articles related to the query and parses each article. The parsed data includes the URL, title, text, authors, and publishing date of each article.

2. **Excel File Creation**: Stores the parsed data in an Excel file named 'newslinks.xlsx'. Each query has its own worksheet in the workbook, and the worksheets are named after the query.

3. **PDF Generation**: Converts each article to a PDF. The PDFs are stored in a directory named after the query under the 'Google_News_PDFs' directory. If there are already PDFs in the directory, the function will start creating new PDFs from the next number.

## Libraries Used

- googlesearch
- openpyxl
- newspaper
- os
- pdfkit
- glob

## Usage

To use this script, simply run it and enter your search query when prompted. The script will then scrape the news articles, store the data in an Excel file, and convert each article to a PDF.

## Note

This script uses the `wkhtmltopdf` tool to convert web pages to PDF. You need to have this tool installed on your system and provide the path to the executable in the script.
