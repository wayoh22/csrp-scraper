import time
import requests
import xlsxwriter
from login_credentials import username, password
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def login_to_schedule():
    link = driver.find_element_by_link_text("Logon").click()
    element = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "username")))
    username_entry = driver.find_element_by_name("username").send_keys(username)
    password_entry = driver.find_element_by_name("password").send_keys(password + Keys.RETURN)
    main = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CLASS_NAME, "ui-link-inherit"))
    )
    referee_schedule_link = driver.find_element_by_link_text("Referee Schedule").click()

PATH = "/Users/rtamburro/Documents/chromedriver"
driver = webdriver.Chrome(PATH)
driver.get("https://csrp.ctreferee.net/")

login_to_schedule()

main = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.ID, "mobile_referee_schedule_form"))
    )

dates = []
leagues = []
home_teams = []
away_teams = []
venues = []

games = main.find_elements_by_xpath("/html/body/div[1]/div[2]/form/div[3]/div")

def sort_game_information(date, teams):
    date_of_game = date.rsplit(None, 3)[0]
    team_names = teams.partition("vs")
    home_team = team_names[0].title().strip()
    away_team = team_names[2].title().strip()
    dates.append(date_of_game)
    home_teams.append(home_team)
    away_teams.append(away_team)

for game in games:
    game_information = game.find_element_by_class_name("text_simple").text.split("\n")
    date = game_information[0]
    league = game_information[1]
    teams = game_information[2]
    venue = game_information[3]

    sort_game_information(date, teams)

    leagues.append(league)
    venues.append(venue)

date_range_selector = driver.find_element_by_name("game_selector").click()
quarter_selector = driver.find_element_by_xpath("/html/body/div[1]/div[2]/form/div[2]/div/select/option[2]").click()
second = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.ID, "mobile_referee_schedule_form"))
    )
second_games = second.find_elements_by_xpath("/html/body/div[1]/div[2]/form/div[3]/div")

for game in second_games:
    game_information = game.find_element_by_class_name("text_simple").text.split("\n")
    date = game_information[0]
    league = game_information[1]
    teams = game_information[2]
    venue = game_information[3]

    sort_game_information(date, teams)

    leagues.append(league)
    venues.append(venue)

workbook = xlsxwriter.Workbook("referee-schedule-2.xlsx")
worksheet = workbook.add_worksheet("Fall 2021")

def write_headers_to_worksheet(worksheet):
    headers = ["Date", "Home Team", "Away Team", "Position", "Fee Paid", "Cash", "Check", "DD", "Venmo", "League", "Gender", "Venue"]
    row = 0
    col = 0
    for header in headers:
        worksheet.write_string(row, col, header)
        col += 1

def write_dates_to_worksheet(worksheet, dates):
    row = 1
    col = 0
    for x in dates:
        worksheet.write_string(row, col, x)
        row +=1

def write_home_teams_to_worksheet(worksheet, home_teams):
    row = 1
    col = 1
    for x in home_teams:
        worksheet.write_string(row, col, x)
        row += 1

def write_away_teams_to_worksheet(worksheet, away_teams):
    row = 1
    col = 2
    for x in away_teams:
        worksheet.write_string(row, col, x)
        row += 1

def write_leagues_to_worksheet(worksheet, leagues):
    row = 1
    col = 9
    for x in leagues:
        worksheet.write_string(row, col, x)
        row += 1

def write_venues_to_worksheet(worksheet, venues):
    row = 1
    col = 11
    for x in venues:
        worksheet.write_string(row, col, x)
        row += 1

write_headers_to_worksheet(worksheet)
write_dates_to_worksheet(worksheet, dates)
write_home_teams_to_worksheet(worksheet, home_teams)
write_away_teams_to_worksheet(worksheet, away_teams)
write_leagues_to_worksheet(worksheet, leagues)
write_venues_to_worksheet(worksheet, venues)

workbook.close()



