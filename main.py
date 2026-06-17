from fastapi import FastAPI, Query
import requests
from bs4 import BeautifulSoup

app = FastAPI(
    title="Market Intelligence API", 
    description="A production-grade RESTful API that handles dynamic pagination."
)

@app.get("/api/v1/books")
def get_scraped_books(pages: int = Query(default=1, ge=1, le=10, description="The number of pages to scrape")):
    books_list = []
    
    for page in range(1, pages + 1):
        url = f"https://books.toscrape.com/catalogue/page-{page}.html"
        response = requests.get(url)
        
        if response.status_code != 200:
            continue
            
        soup = BeautifulSoup(response.text, "html.parser")
        articles = soup.find_all("article", class_="product_pod")
        
        for article in articles:
            title = article.h3.a["title"]
            price = article.find("p", class_="price_color").text
            
            books_list.append({
                "title": title,
                "price": price.replace("Â", "").strip()
            })
            
    return {
        "status": "success", 
        "pages_scraped": pages,
        "count": len(books_list), 
        "data": books_list
    }