URL - Title - SEO Title Optimizer
===============================

This program is a web scraper that uses Selenium to automate the process of exporting CSVs from Ahrefs for a list of URLs.

The program takes a list of URLs from a text file, logs into Ahrefs, and downloads the CSVs for the top 10 organic keywords for each URL. It then merges all the CSVs into one file and cleans it up before exporting it to an Excel file.

Installation
------------

To use this program, you need to install the following Python packages:

*   `selenium`
*   `webdriver_manager`
*   `pandas`
*   `openpyxl`

You can install them using pip:

Copy code

`pip install selenium webdriver_manager pandas openpyxl`

You also need to download the appropriate version of the ChromeDriver executable for your system and put it in your PATH.

Usage
-----

1.  Create a file named `urls.txt` in the same directory as the Python script.
2.  Add the URLs you want to scrape to `urls.txt`, with each URL on a separate line.
3.  Run the Python script in your terminal with `python ahrefs_scraper.py`.
4.  Wait for the program to complete its run, then check for the output Excel file named with the top keyword.

Program Flow
------------

The program works as follows:

1.  Reads the `urls.txt` file and converts it into a dictionary of URLs.
2.  Launches the Chrome browser using Selenium.
3.  Logs into Ahrefs using your credentials.
4.  For each URL in the dictionary:
    1.  Enters the URL into the search box.
    2.  Downloads the CSVs for the top 10 organic keywords for that URL.
    3.  Merges the downloaded CSVs into one file.
    4.  Cleans up the merged file and exports it to an Excel file.
5.  Closes the browser.

The program uses various functions and libraries to achieve the above functionality. These include `webdriver`, `time`, `pandas`, `os`, `re`, `openpyxl`, and `selenium`.

Disclaimer
----------

This script is for educational purposes only. Use it at your own risk. The author does not guarantee the accuracy, reliability, or suitability of this script for any purpose. The author is not responsible for any damage or loss caused by the use of this script.
