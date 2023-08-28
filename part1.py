'''
pip install requests
pip install pandas
pip install beautifulsoup4
pip install python-pptx
'''


import requests
from bs4 import BeautifulSoup
import pandas as pd

# Number of pages to fetch
num_pages = 3

# Initialize lists to store article details
counter = 1
titles = []
categories = []
dates = []
authors = []
descriptions = []
urls = []

# Set to keep track of visited URLs
visited_urls = []

# Iterate over the pages
for page in range(1, num_pages + 1):
    # Send a GET request to the webpage
    url = f"https://powerbi.microsoft.com/en-us/blog/?page={page}"
    response = requests.get(url)

    # Create BeautifulSoup object to parse the HTML content
    soup = BeautifulSoup(response.content, "html.parser")

    # Find all article elements
    articles = soup.find_all("article", class_="post")

    # Extract information from each article
    for article in articles:
        # Extract URL
        base_url = "https://powerbi.microsoft.com"
        link = article.find("a", string="» Read more") or article.find("a", class_="button-fabricfeatured")
        article_url = base_url + link["href"] if link else ""

        # Skip if the URL is already visited
        if article_url in visited_urls:
            continue

        # Mark URL as visited
        visited_urls.append(article_url)

        # Extract title
        title = article.find("a")["title"]
        titles.append(title.strip() if title else "")

        # Extract categories
        category_links = article.find_all("a", class_="tag")
        category_list = [link.text.strip() for link in category_links]
        categories.append(", ".join(category_list))

        # Extract date
        date = article.find("span", class_="icon-fabricpalette02").find_next_sibling("span")
        dates.append(date.text.strip() if date else "")

        # Extract author
        author = article.find("a", href=lambda href: href and "/blog/author/" in href)
        authors.append(author.text.strip() if author else "")

        # Extract description
        metadata = article.find("div", class_="metadata")
        metadata.extract()  # Remove the metadata from the article
        description = article.find("p", dir="ltr")
        descriptions.append(description.text.strip() if description else "")

        # Store the counter value
        urls.append(counter)
        counter += 1

# Create a DataFrame with the extracted article details
data = {
    "Blogpost#": urls,
    "Title": titles,
    "Categories": categories,
    "Date": dates,
    "Author": authors,
    "Description": descriptions,
    "URL": visited_urls
}
df = pd.DataFrame(data)

# Save the DataFrame to an Excel file
output_file = "powerbi_blog_articles.xlsx"
df.to_excel(output_file, index=False)

print("Excel file created successfully.")


