# ------
# Goals
#   - Scrape website to find games
#   - Populate excel sheet with games
#   - Populate excel sheet with formulas
#   - This is purely to develop the skeleton document, not to update the scores
# -------


# Get all bowl games
#   Scrape URL for bowl games
#   Get bowl game name, team, and date
#   INPUT: URL
#   OUTPUT: Bowl Name, Date, Home Team, Away Team
#       output could be a class, list of lists, dict, etc.
#       should sort by date
#       List of lists would be very simple, class would be inclusive

import requests
from bs4 import BeautifulSoup
# -------------------------------------------------------------------------------------------------------------------
# Basic download of page
# -------------------------------------------------------------------------------------------------------------------

url = "https://www.ncaa.com/scoreboard/football/fbs/2019/14"
page = requests.get(url)
soup = BeautifulSoup(page.content, 'html.parser')

# -------------------------------------------------------------------------------------------------------------------
# Get all containers, they contain the date
# -------------------------------------------------------------------------------------------------------------------

container = soup.find_all(class_="gamePod_content-division")
# date = container[0].find('h6')

# -------------------------------------------------------------------------------------------------------------------
# Get all the games
# -------------------------------------------------------------------------------------------------------------------

gamesList = []
for pod in container:
    date = pod.find('h6').get_text()
    teamNames = [team.get_text() for team in pod.find_all(class_="gamePod-game-team-name")]
    # teamScores = [score.get_text() for score in pod.find_all(class_="gamePod-game-team-score")]
    for index in range(0,len(teamNames), 2):
        homeTeam = teamNames[index]
        awayTeam = teamNames[index + 1]
        gameIndex = [date, 'BOWL NAME' + str(index), homeTeam, awayTeam]
        gamesList.append(gameIndex)

# -------------------------------------------------------------------------------------------------------------------
# Create the worksheet
# -------------------------------------------------------------------------------------------------------------------

from openpyxl import Workbook

players = ["Brad", "Burton", "Chris", "Martin"]
playerEntry = []
for user in players:
    playerEntry += [user + 'Home', user + 'Away', user + 'Score']

wb = Workbook()
destFile = "sampleSkeleton.xlsx"

ws1 = wb.active
ws1.title = "CoverPage"

wb.create_sheet("BowlGames")
wb.active = wb["BowlGames"]

# Create the Header
ws = wb.active
appendHeader = ['DATE', 'BOWL NAME', 'HOME TEAM', 'AWAY TEAM'] + playerEntry + ['ACTUAL HOME', 'ACTUAL AWAY']
ws.append(appendHeader)

# Fill in the games
for game in gamesList:
    ws.append(game)

wb.save(filename = destFile)