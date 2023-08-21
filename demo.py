import requests
import openpyxl
from bs4 import BeautifulSoup

excel = openpyxl.Workbook()
page = excel.active
page.append(("S.no", "Name", "Year", "Rating"))
try:
    response = requests.get(
        "https://www.imdb.com/search/title/?genres=Adventure&sort=user_rating,desc&title_type=feature&num_votes=25000,&ref_=chttp_gnr_1")
    soup = BeautifulSoup(response.text, "html.parser")
    # print(soup)
    movies = soup.find("div", class_="lister-list").find_all("div", class_="lister-item")
    for movie in movies:
        # print(movie)
        index = movie.find("h3").find("span", class_="lister-item-index").text.split('.')
        name = movie.find("h3").a.text
        year = movie.find("h3").find("span", class_="lister-item-year").text.replace('(', "").replace(')', "")
        rate = movie.find("div", class_="ratings-imdb-rating").strong.text
        print(index[0],name,year,rate)
        # page.append((index[0], name, year, rate))


except Exception as e:
    print(e)

excel.save(filename="imdbscrap.xlsx")
