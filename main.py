from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
# print(excel.sheetnames)

sheet = excel.active
sheet.title = 'Top rated movies'
# print(excel.sheetnames)

sheet.append(['Movie Rank','Movie Nmae', 'Movie Year', 'Movie Rating'])

url = 'https://www.imdb.com/chart/top/'

try:
    movie_data = requests.get(url)
    movie_data.raise_for_status()

    soup = BeautifulSoup(movie_data.text, 'html.parser')
    movies = soup.find('tbody', class_="lister-list").find_all('tr')
    
    for movie in movies:
        movie_name = movie.find('td', class_="titleColumn").a.text
        movie_rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]
        movie_year = movie.find('td', class_="titleColumn").span.text.strip('()')
        movie_rating = movie.find('td', class_="ratingColumn imdbRating").strong.text

        # print(movie_rank, movie_name,movie_year,movie_rating)
        sheet.append([movie_rank, movie_name,movie_year,movie_rating])
except Exception as e:
    print(e)

excel.save('IMDB Movie Rating.xlsx')