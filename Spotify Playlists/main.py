import requests as rq 
from bs4 import BeautifulSoup
import spotipy
from spotipy.oauth2 import SpotifyOAuth


CLIENT_ID = "702319fc01e7427e88001256cae7caf9"
CLIENT_SECRET = "1cbd4ee19b6a432da2229f2f233a161d"
USERNAME= "lf0mm26e72iud172g753eld8s"
REDIRECT_URI = "https://example.com/callback"

date = input("Which year do you want to travel to? Type the date in this format YYYY-MM-DD: ")

url = f"https://www.billboard.com/charts/hot-100/{date}/"

response = rq.get(url)

billboard_webpage = response.text

soup = BeautifulSoup(billboard_webpage, "html.parser")

# print(soup.prettify())

song_position =[position.get_text().strip() for position in soup.find_all(name="span", class_="a-font-primary-bold-l")]
song_title = [title.get_text().strip() for title in soup.find_all(name="h3", class_="a-no-trucate")]
song_artist = [artist.get_text().strip() for artist in soup.find_all(name="span", class_="a-no-trucate")]


sp = spotipy.Spotify(
    auth_manager=SpotifyOAuth(
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET,
        redirect_uri=REDIRECT_URI,
        scope="playlist-read-private playlist-modify-private playlist-modify-public",
        show_dialog=True,
        cache_path=r"100 Days of coding\Intermediate +\Day 46 - Spotify Playlists\token.txt"
    )
)

song_uris = []
year = date.split("-")[0]
for song in song_title:
    result = sp.search(q=f"track:{song} year:{year}", type="track")
    # print(result)
    try:
        uri = result["tracks"]["items"][0]["uri"]
        song_uris.append(uri)
    except IndexError:
        print(f"{song} doesn't exist in Spotify. Skipped.")

#Creating playlist
playlist = sp.user_playlist_create(user="lf0mm26e72iud172g753eld8s", name=f"{date} Billboard 100", public=True, description=f"The best hits from {year}.")
# print(playlist)

try:
    sp.playlist_add_items(playlist_id=playlist["id"],items=song_uris)
    print("List created!")
except spotipy.exceptions.SpotifyException:
    print("List not created :(")

# # Get the current user ID
# user_id = sp.current_user()["id"]
# print(f"User ID: {user_id}")






