import pandas as pd

from imdb import IMDb

# create an instance of the IMDb class
imdb = IMDb()

input_films = pd.read_csv('input_films.csv')

all_results = []

print("Fetching data for films:")

for index, row in input_films.iterrows():
    
    print(str(index+1) + ". " + row['title'])
    
    film_search = imdb.search_movie(row['title'])
    
    film_id = film_search[0].movieID
    
    film = imdb.get_movie(film_id)
    
    top_cast = []
    for cast_member in film['cast'][:20]:
        top_cast.append(cast_member['name'])
    
    result = {
        "title": film['title'],
        "rating": film['rating'],
        "genre": film['genres'],
        "year": film['year'],
        "director": film['directors'][0]['name'],
        "runtime": film['runtimes'][0],
        "cast": top_cast
    }
    
    all_results.append(result)
    
all_results_df = pd.DataFrame.from_dict(all_results)

headers = ["title","year","rating","genre","runtime","director","cast"]
final_results_df = all_results_df[headers]

top_actors_df = final_results_df['cast'].apply(lambda x: pd.Series(x).value_counts()).sum().to_frame("counts").sort_values("counts", ascending=False)
top_actors_df.index.name = "Actor"

top_genres_df = final_results_df['genre'].apply(lambda x: pd.Series(x).value_counts()).sum().to_frame("counts").sort_values("counts", ascending=False)
top_genres_df.index.name = "Genre"

top_directors_df = final_results_df['director'].value_counts().to_frame("counts").sort_values("counts", ascending=False)
top_directors_df.index.name = "Director"

with pd.ExcelWriter('output.xlsx') as writer:
    final_results_df.to_excel(writer, sheet_name='Films', index=False)
    top_actors_df.to_excel(writer, sheet_name='Actors')
    top_genres_df.to_excel(writer, sheet_name='Genres')
    top_directors_df.to_excel(writer, sheet_name='Directors')