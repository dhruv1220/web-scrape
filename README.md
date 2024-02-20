# Web Scrape Research Articles
A python script that utilizes Selenium and Beautiful Soup to scrape research papers off the web and put them in an excel sheet.  
The website scraped was: **The Institute for Operations Research and the Management Sciences** (https://pubsonline.informs.org/)  
The research articles from the year 2011 onwards are scraped.

## Pre-requisites
1. Python
2. Google chromedriver (preferably a stable version)
3. Selenium and BeautifulSoup libraries on Python
  
**Pro Tip:** Put the chromedriver executable in `/usr/local/bin` file path.

## How to Run
Run the Python file on terminal using the following command:
```code
python3 web_scrape.py
```

## Output Format
The output is an excel sheet by the name of `research_articles.xlsx`. I have posted my output for reference too in a file called `research_articles_final.xlsx`.  

It contains the following columns:-
- _Volume:_ Volume of the research article. Usually a year is associated with a particular volume. For eg: 2024 is Volume 70, 2023 is Volume 69, and so on.
- _Issue:_ The issue in that particular volume. Usually a month is associated with a particular issue. For eg: January is Issue 1, February is Issue 2, and so on.
- _Title:_ Title of the research article
- _Author:_ The list of authors of the research article
- _Abstract:_ The complete abstract text of the research article
- _Accepted_By:_ The name of the person who accepted the research article
- _Accepted_Dept:_ The department of the person who accepted the research article
- _URL:_ The link that was scraped to get all the research article information

**Sample Output**  
![image](https://github.com/dhruv1220/web-scrape/assets/35871415/36ca364e-1332-47a7-9c8f-7d6f47105d46)
