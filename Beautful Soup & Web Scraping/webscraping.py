import requests
from bs4 import BeautifulSoup as bs

r = requests.get("https://keithgalli.github.io/web-scraping/example.html")

soup = bs(r.content)

first_header = soup.find("h2")
print(first_header)

headers = soup.find_all("h2")
print (headers)
#print(soup.prettify())


list_header=soup.find(["h2h","h1"]) # lists
list_headers=soup.find_all(["h2h","h1"]) # lists

paragraps=soup.find_all("p", attrs={"id": "paragraph-id"}) #can pass in the attribute of the p to filter it
print(paragraps)


