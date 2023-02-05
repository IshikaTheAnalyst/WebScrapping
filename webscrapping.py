from bs4 import BeautifulSoup
import requests, openpyxl
excel=openpyxl.Workbook()
sheet=excel.active
sheet.title='Top Rated Movies'
sheet.append(['Movie Rank','Movie Name','Year of Release','IMDB Rating'])

#use request module to access this website 
#this 'source' will have html code for this website 
#requests module will not be able to check if this link is correct or not so to ensure this link is fine we will use raise to status
try: 
    source=requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()  #this will throw an error if the link is not correct 
    soup=BeautifulSoup(source.text,'html.parser')  #BS will take html content of this website and then it is going to parse it using parser 
    movies=soup.find('tbody',class_="lister-list").find_all('tr') #used findall because there are many tr tags
    for movie in movies:
        name= movie.find('td', class_="titleColumn").a.text
        rank= movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0] # this will remove all white spaces, split will seperate rank and name 
        year= movie.find('td', class_="titleColumn").span.text.strip('()') # this will strip out the brackets 
        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text
        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating])
       
  

except Exception as e:
    print(e)

excel.save('IMDB Movie rating.xlsx')