import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from random import uniform
from selenium.webdriver.chrome.options import Options
from datetime import datetime
import pandas as pd

# Initialize the driver
# driver = webdriver.Chrome()

def get_volume_issue_list(decade, num):
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(options=options)

    # Navigate to the webpage
    driver.get('https://pubsonline.informs.org/loi/mnsc/group/d' + str(decade) + '.y' + str(num))

    # Get the page source and create a BeautifulSoup object
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # Find all links that contain volume and issue information
    volume_issue_links = soup.find_all('a', class_='issue-info__vol-issue')

    # List to hold the extracted volume and issue numbers
    volume_issue_list = []

    for link in volume_issue_links:
        text = link.get_text(strip=True)  # Get the text content of the link
        # Use string methods or a regular expression to extract numbers
        # Assuming the format is always "Volume XX, Issue Y"
        parts = text.replace('Volume', '').replace('Issue', '').split(',')
        if len(parts) == 2:
            volume = parts[0].strip()
            issue = parts[1].strip()
            volume_issue_list.append((volume, issue))

    # Close the browser
    driver.quit()

    return volume_issue_list


# Get the current year
current_year = datetime.now().year

volume_issue_list_complete = []
for decade in [2020, 2010]:
    diff = min(current_year + 1 - decade, 10)
    for year in range(diff):
        if decade + year > 2010:
            volume_issue_list_complete.append(get_volume_issue_list(decade, decade + year))

print(volume_issue_list_complete)


def get_research_article_links(v, s):
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(options=options)

    # Navigate to the webpage
    driver.get('https://pubsonline.informs.org/toc/mnsc/' + v + '/' + s)

    # Get the page source and create a BeautifulSoup object
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    research_articles_heading = soup.find(lambda tag: tag.name == "h2" and "Research Articles" in tag.text)
    research_articles_heading2 = soup.find(lambda tag: tag.name == "h2" and "Research Article" in tag.text)

    yearly_links = []

    # If the articles are contained within a sibling or child of the heading's parent container
    if research_articles_heading:
        articles_container = research_articles_heading.find_next_sibling('div')
        while articles_container and articles_container.name == 'div':
            article_links = articles_container.find_all('a', href=True)
            links = [link['href'] for link in article_links if 'doi.org' in link['href']]
            yearly_links.append(links[0])

            articles_container = articles_container.find_next_sibling()

    elif research_articles_heading2:
        articles_container = research_articles_heading2.find_next_sibling('div')
        while articles_container and articles_container.name == 'div':
            article_links = articles_container.find_all('a', href=True)
            links = [link['href'] for link in article_links if 'doi.org' in link['href']]
            yearly_links.append(links[0])

            articles_container = articles_container.find_next_sibling()

    else:
        print("Research Articles section not found")

    # Close the browser
    driver.quit()

    return yearly_links


#### Test Line - Remove ######
# volume_issue_list_complete = [[('66', '12'), ('66', '11'), ('66', '10'), ('66', '9'), ('66', '8'), ('66', '7'), ('66', '6'), ('66', '5'), ('66', '4'), ('66', '3'), ('66', '2'), ('66', '1')], [('67', '12'), ('67', '11'), ('67', '10'), ('67', '9'), ('67', '8'), ('67', '7'), ('67', '6'), ('67', '5'), ('67', '4'), ('67', '3'), ('67', '2'), ('67', '1')], [('68', '12'), ('68', '11'), ('68', '10'), ('68', '9'), ('68', '8'), ('68', '7'), ('68', '6'), ('68', '5'), ('68', '4'), ('68', '3'), ('68', '2'), ('68', '1')], [('69', '12'), ('69', '11'), ('69', '10'), ('69', '9'), ('69', '8'), ('69', '7'), ('69', '6'), ('69', '5'), ('69', '4'), ('69', '3'), ('69', '2'), ('69', '1')], [('70', '1')], [('57', '12'), ('57', '11'), ('57', '10'), ('57', '9'), ('57', '8'), ('57', '7'), ('57', '6'), ('57', '5'), ('57', '4'), ('57', '3'), ('57', '2'), ('57', '1')], [('58', '12'), ('58', '11'), ('58', '10'), ('58', '9'), ('58', '8'), ('58', '7'), ('58', '6'), ('58', '5'), ('58', '4'), ('58', '3'), ('58', '2'), ('58', '1')], [('59', '12'), ('59', '11'), ('59', '10'), ('59', '9'), ('59', '8'), ('59', '7'), ('59', '6'), ('59', '5'), ('59', '4'), ('59', '3'), ('59', '2'), ('59', '1')], [('60', '12'), ('60', '11'), ('60', '10'), ('60', '9'), ('60', '8'), ('60', '7'), ('60', '6'), ('60', '5'), ('60', '4'), ('60', '3'), ('60', '2'), ('60', '1')], [('61', '12'), ('61', '11'), ('61', '10'), ('61', '9'), ('61', '8'), ('61', '7'), ('61', '6'), ('61', '5'), ('61', '4'), ('61', '3'), ('61', '2'), ('61', '1')], [('62', '12'), ('62', '11'), ('62', '10'), ('62', '9'), ('62', '8'), ('62', '7'), ('62', '6'), ('62', '5'), ('62', '4'), ('62', '3'), ('62', '2'), ('62', '1')], [('63', '12'), ('63', '11'), ('63', '10'), ('63', '9'), ('63', '8'), ('63', '7'), ('63', '6'), ('63', '5'), ('63', '4'), ('63', '3'), ('63', '2'), ('63', '1')], [('64', '12'), ('64', '11'), ('64', '10'), ('64', '9'), ('64', '8'), ('64', '7'), ('64', '6'), ('64', '5'), ('64', '4'), ('64', '3'), ('64', '2'), ('64', '1')], [('65', '12'), ('65', '11'), ('65', '10'), ('65', '9'), ('65', '8'), ('65', '7'), ('65', '6'), ('65', '5'), ('65', '4'), ('65', '3'), ('65', '2'), ('65', '1')]]

all_yearly_links = {}
for volume in volume_issue_list_complete:
    for issue in volume:
        v, s = issue
        all_yearly_links[(v,s)] = get_research_article_links(v, s)

print(all_yearly_links)


def get_article_metadata(v, s, url):
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(options=options)

    # Navigate to the webpage
    driver.get(url)

    # Get the page source and create a BeautifulSoup object
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # Get Citation Title
    title_tag = soup.find('h1', class_='citation__title')
    title_text = title_tag.get_text().strip() if title_tag else 'Title not found'
    # print(title_text)

    # If Link doesn't exist
    if title_text == 'Title not found':
        print(url + ' does not exist!')
        return -1

    # Get Citation Author
    citations = soup.find('div', class_='citation')
    authors = []
    for citation in citations:
        author_tags = citation.find_all('a', class_='entryAuthor')
        for author_tag in author_tags:
            authors.append(author_tag.get_text().strip())

    authors = list(set(authors))
    # print(authors)


    # Get Abstract Text
    complete_abstract_tag = soup.find('div', class_='abstractSection abstractInFull')
    complete_abstract_text = complete_abstract_tag.get_text() if complete_abstract_tag else 'Abstract not found'
    if 'This paper was accepted by' in complete_abstract_text:
        abstract_text = complete_abstract_text.split('This paper was accepted by ')[0].strip()
    elif 'This paper has been accepted by' in complete_abstract_text:
        abstract_text = complete_abstract_text.split('This paper has been accepted by ')[0].strip()
    else:
        abstract_text = complete_abstract_text

    try:
        if 'This paper was accepted by' in complete_abstract_text:
            complete_abstract_text = complete_abstract_text.split('This paper was accepted by ')[1]
        elif 'This paper has been accepted by' in complete_abstract_text:
            complete_abstract_text = complete_abstract_text.split('This paper has been accepted by ')[1]

        accepted_by = complete_abstract_text.split(',')[0].strip()
        # print(accepted_by)

        accepted_dept = complete_abstract_text.split(',')[1].strip()[:-1]
        # print(accepted_dept)
    except:
        print(url)
        accepted_by = 'Not Found'
        accepted_dept = 'Not Found'

    # Close the browser
    driver.quit()

    return {
        'Volume': v,
        'Issue': s,
        'Title': title_text,
        'Author': authors,
        'Abstract': abstract_text,
        'Accepted_By': accepted_by,
        'Accepted_Dept': accepted_dept,
        'URL': url
    }

output_list = []
for key, value in all_yearly_links.items():
    v, s = key
    for link in value:
        temp = get_article_metadata(v, s, link)
        if temp != -1:
            output_list.append(temp)
    
    # # Uncomment following lines if all the links not getting processed at once
    # df_temp = pd.DataFrame(output_list)
    # df_temp.to_excel('Desktop/research_articles_temp.xlsx', index=False)

print(output_list)

# Convert the list of dictionaries to a DataFrame
df = pd.DataFrame(output_list)
print(df)

# Specify the filename and path for your Excel file
filename = 'Desktop/research_articles.xlsx'

# Export the DataFrame to an Excel file
df.to_excel(filename, index=False)