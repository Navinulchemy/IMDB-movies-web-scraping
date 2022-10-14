import requests,openpyxl
from bs4 import BeautifulSoup
excel=openpyxl.Workbook()
sheet=excel.active
sheet.title="Top Rated Movies"
sheet.append(["Rank","Title","Year","Rating"])
# print(excel.sheetnames)


try:
    source=requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()

    parse=BeautifulSoup(source.text,'html.parser')
    
    movies=parse.find('tbody',class_="lister-list").find_all('tr')
    
    for movie in movies:
        rank=movie.find('td',class_="titleColumn").get_text(strip=True).split(".")[0]
        name=movie.find('td',class_="titleColumn").a.text
        year=movie.find('td',class_="titleColumn").span.text.strip("()")
        rating=movie.find('td',class_="ratingColumn imdbRating").strong.text

        print(rank,name,year,rating)
        sheet.append([rank,name,year,rating])
      
except Exception as e:
    print(e)
excel.save("Idmb Ratings.xlsx")