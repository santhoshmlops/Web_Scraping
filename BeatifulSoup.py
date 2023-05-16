from bs4 import BeautifulSoup
import requests,openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title="Movie List"
sheet.append(['Rank','Movie','Rating','Year'])

try:
    response = requests.get("https://www.imdb.com/chart/top/")
    soup=BeautifulSoup(response.text,'html.parser')
    movies = soup.find('tbody',class_="lister-list").find_all("tr")

    for movie in movies :
        rank = movie.find('td',class_="titleColumn").get_text(strip=True).split('.')[0]
        movie_name = movie.find('td',class_="titleColumn").a.text
        rating = movie.find('td',class_="ratingColumn imdbRating").strong.text
        year = movie.find('td',class_="titleColumn").span.text.replace('(',"")
        year = year.replace(')',"")
        # print(rank,movie_name,rating,year)
        sheet.append([rank,movie_name,rating,year])     
 
except Exception as e:
    print(e)

excel.save("Movies.xlsx")