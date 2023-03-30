from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()

sheet = excel.active
sheet.title = "Indian IMBD Rating"

sheet.append(['Rank','Movie Name','Year of Release','Rating'])


try:
    source = requests.get("https://www.imdb.com/india/top-rated-indian-movies/")
    source.raise_for_status()
    soup = BeautifulSoup(source.text,'html.parser')
    
    movies = soup.find('tbody',class_=("lister-list")).find_all('tr')
    
    for movie in movies:

        name = movie.find('td',class_="titleColumn").a.text
        
        rank = movie.find('td',class_="titleColumn").get_text(strip=True).split('.')[0]
        
        year = movie.find('td', class_="titleColumn").span.text.strip('()')
        
        rate = movie.find('td',class_="ratingColumn imdbRating").strong.text
        sheet.append([rank,name,year,rate])
except Exception as e:
    print(e)    

excel.save("IMBD Rating of India.xlsx")