# scrape-items-price
### Description
- Scrape top 3 defined item's price in list of "search term / item name", aggregated by top sales.

## How it works for a single item
![How it works for a single item](/docs-image/scrapeeshopee.gif)

1. List down "search term / item name" under Search Terms column
2. Save and close the excel file
3. Run the python script
4. Check back the excel file (expected to have multiple new sheet correspond to the number of item name.

## Modules required
beautifulsoup4==4.10.0 <br>
openpyxl==3.0.7 <br>
pandas==1.3.5 <br>
selenium==4.1.0 <br>
webdriver_manager==3.5.2 <br><br>

**You can also use requirements.txt to install the packages. How? Follow this [link].**

## Locate excel.file (xlsx) to change pathline
### Line 21 (to read)
![line21](/docs-image/line21.png)

### Line 79 (to write)
![line79](/docs-image/line79.png)

[link]: https://stackoverflow.com/questions/7225900/how-to-install-packages-using-pip-according-to-the-requirements-txt-file-from-a
