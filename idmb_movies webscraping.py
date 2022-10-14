#importing the necessary lib
import requests,openpyxl
from bs4 import BeautifulSoup
#creating a empty excel file
excel=openpyxl.Workbook()
#identifying the active sheet
sheet=excel.active
#setting uo the title for the sheet
sheet.title="Top Rated Movies"
#setting up the column headers
sheet.append(["Rank","Title","Year","Rating"])
# print(excel.sheetnames)


try:
    #requesting the webpage for the html source
    source=requests.get('https://www.imdb.com/chart/top/')
    #will indicate if any error
    source.raise_for_status()
     #scrapes the entire html context
    parse=BeautifulSoup(source.text,'html.parser')
    
    
    # fetching the necessary data needed for our process from their respective individual tags in html content
    
    movies=parse.find('tbody',class_="lister-list").find_all('tr')
    
    for movie in movies:
        rank=movie.find('td',class_="titleColumn").get_text(strip=True).split(".")[0]
        name=movie.find('td',class_="titleColumn").a.text
        year=movie.find('td',class_="titleColumn").span.text.strip("()")
        rating=movie.find('td',class_="ratingColumn imdbRating").strong.text

        print(rank,name,year,rating)
        
        #appending the each individual info to a excel file 
        sheet.append([rank,name,year,rating])
      
except Exception as e:  #indicates if any error occurs
    print(e)
excel.save("Idmb Ratings.xlsx")  #saving locally the entire weather status as a excel file
