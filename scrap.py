from bs4 import BeautifulSoup
import requests, openpyxl
excel = openpyxl.Workbook()
sheet = excel.active
sheet.title='Top Rated Movies'
sheet.append(['Rank','Movie Name','Year of Release','Ratings'])


try:
    source = requests.get("https://www.imdb.com/chart/top/")
    source.raise_for_status()
    soup = BeautifulSoup(source.text,'html.parser')

    movies = soup.find('tbody',class_=("lister-list")).find_all('tr')

    for movie in movies:
        name = movie.find('td',class_="titleColumn").a.text
        
        rank = movie.find('td',class_="titleColumn").get_text(strip=True).split('.')[0]
        
        year = movie.find('td', class_="titleColumn").span.text.strip('()')
        
        rating = movie.find('td',class_="ratingColumn imdbRating").strong.text
        sheet.append([rank,name,year,rating])

except Exception as e:
    print(e)    

excel.save("IMBD Rating.xlsx")
