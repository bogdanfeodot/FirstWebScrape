from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

import pandas as pd
import pyodbc

import time
import datetime

import smtplib
from email.mime.multipart import MIMEMultipart


browser = webdriver.Chrome(executable_path='C:/Python projects/Web Scraping Airplanes/chromedriver')

#Next, I will quickly go to Expedia https://www.expedia.com/ to check the interface and the options available to choose from.
#I click right click + inspect on the ticket type (roundtrip, one way, etc.) to see the tags related to it.
#As we can see below it has a 'label' tag with 'id = flight-type-roundtrip-label-hp-flight'.

#Setting ticket types paths
return_ticket = "//label[@id='flight-type-roundtrip-label-hp-flight']"
one_way_ticket = "//label[@id='flight-type-one-way-label-hp-flight']"
multi_ticket = "//label[@id='flight-type-multi-dest-label-hp-flight']"

#Then I define a function to choose a ticket type:
def ticket_chooser(ticket):
    try:
        ticket_type = browser.find_element_by_xpath(ticket)
        ticket_type.click()
    except Exception as e:
        pass

def dep_country_chooser(dep_country):
    fly_from = browser.find_element_by_xpath("//input[@id='flight-origin-hp-flight']")
    time.sleep(1)
    fly_from.clear()
    time.sleep(1.5)
    fly_from.send_keys('  ' + dep_country)
    time.sleep(1.5)
    first_item = browser.find_element_by_xpath("//a[@id='aria-option-0']")
    time.sleep(1.5)
    first_item.click()

def arrival_country_chooser(arrival_country):
    fly_to = browser.find_element_by_xpath("//input[@id='flight-destination-hp-flight']")
    time.sleep(1)
    fly_to.clear()
    time.sleep(1.5)
    fly_to.send_keys('  ' + arrival_country)
    time.sleep(1.5)
    first_item = browser.find_element_by_xpath("//a[@id='aria-option-0']")
    time.sleep(1.5)
    first_item.click()

def dep_date_chooser(month, day, year):
    dep_date_button = browser.find_element_by_xpath("//input[@id='flight-departing-hp-flight']")
    dep_date_button.clear()
    dep_date_button.send_keys(month + '/' + day + '/' + year)


def return_date_chooser(month, day, year):
    return_date_button = browser.find_element_by_xpath("//input[@id='flight-returning-hp-flight']")
    for i in range(11):
        return_date_button.send_keys(Keys.BACKSPACE)
    return_date_button.send_keys(month + '/' + day + '/' + year)

def search():
    search = browser.find_element_by_xpath("//button[@class='btn-primary btn-action gcw-submit']")
    search.click()
    time.sleep(15)
    print('Results ready!')

df = pd.DataFrame()
def compile_data():
    global df
    global dep_times_list
    global arr_times_list
    global airlines_list
    global price_list
    global durations_list
    global stops_list
    global layovers_list
    #departure times
    dep_times = browser.find_elements_by_xpath("//span[@data-test-id='departure-time']")
    dep_times_list = [value.text for value in dep_times]
    #arrival times
    arr_times = browser.find_elements_by_xpath("//span[@data-test-id='arrival-time']")
    arr_times_list = [value.text for value in arr_times]
    #airline name
    airlines = browser.find_elements_by_xpath("//span[@data-test-id='airline-name']")
    airlines_list = [value.text for value in airlines]
    #prices
    prices = browser.find_elements_by_xpath("//span[@data-test-id='listing-price-dollars']")
    price_list = [value.text.replace(",","").split('$')[1] for value in prices]
    #durations
    durations = browser.find_elements_by_xpath("//span[@data-test-id='duration']")
    durations_list = [value.text for value in durations]
    #stops
    stops = browser.find_elements_by_xpath("//span[@class='number-stops']")
    stops_list = [value.text for value in stops]
    #layovers
    layovers = browser.find_elements_by_xpath("//span[@data-test-id='layover-airport-stops']")
    layovers_list = [value.text for value in layovers]
    now = datetime.datetime.now()
    current_date = (str(now.year) + '-' + str(now.month) + '-' + str(now.day))
    current_time = (str(now.hour) + ':' + str(now.minute))
    current_price = 'price' #+ '(' + current_date + '---' + current_time + ')'
    for i in range(len(dep_times_list)):
        try:
            df.loc[i, 'departure_time'] = dep_times_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'arrival_time'] = arr_times_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'airline'] = airlines_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'duration'] = durations_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'stops'] = stops_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'layovers'] = layovers_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, str(current_price)] = price_list[i]
        except Exception as e:
            pass
    print('Excel Sheet Created!')

#email credentials
username = 'feodotbogdan@hotmail.com'
password = 'Parolaesecret4'

def connect_mail(username, password):
    global server
    server = smtplib.SMTP('smtp-mail.outlook.com', 587)
    #server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(username, password)


#Create message template for email
def create_msg():
    global msg
    msg = '\nCurrent Cheapest flight:\n\nDeparture time: {}\nArrival time: {}\nAirline: {}\nFlight duration: {}\nNo. of stops: {}\nPrice: {}\n'.format(cheapest_dep_time,
                       cheapest_arrival_time,
                       cheapest_airline,
                       cheapest_duration,
                       cheapest_stops,
                       cheapest_price)

def send_email(msg):
    print('Sending mail ...')
    global message
    message = MIMEMultipart()
    message['Subject'] = 'Current Best flight'
    message['From'] = 'feodotbogdan@hotmail.com'
    message['to'] = 'feodotbogdan@gmail.com'
    server.sendmail('feodotbogdan@hotmail.com', 'feodotbogdan@gmail.com', msg)
    server.close()


for i in range(24):    
    link = 'https://www.expedia.com/'
    browser.get(link)
    time.sleep(5)
    #choose flights only
    flights_only = browser.find_element_by_xpath("//button[@id='tab-flight-tab-hp']")
    flights_only.click()
    ticket_chooser(return_ticket)
    dep_country_chooser('Bucharest')
    arrival_country_chooser('Bali')
    dep_date_chooser('10', '01', '2019')
    return_date_chooser('10', '10', '2019')
    search()
    compile_data()

    #save values for email
    current_values = df.iloc[0]
    cheapest_dep_time = current_values[0]
    cheapest_arrival_time = current_values[1]
    cheapest_airline = current_values[2]
    cheapest_duration = current_values[3]
    cheapest_stops = current_values[4]
    cheapest_price = current_values[-1]
    print('run {} completed!'.format(i))
    
    create_msg()
    try:
        connect_mail(username,password)
    except:
        print ('Something went wrong...')

    send_email(msg)
    print('Email sent!')
    
    df.to_excel('flights.xlsx')
    
    connStr = pyodbc.connect('DRIVER={ODBC Driver 13 for SQL Server};SERVER=DESKTOP-6PNM517;DATABASE=PythonDB;Trusted_Connection=yes')

    cursor = connStr.cursor()
    
    for index,row in df.iterrows():
        cursor.execute("INSERT INTO dbo.Flights([departure_time],[arrival_time],[airline],[duration],[stops],[layovers],[price]) values (?, ?, ?, ?, ?, ?, ?)"
                    , row['departure_time'], row['arrival_time'], row['airline'], row['duration'], row['stops'], row['layovers'], row['price']) 
        connStr.commit()
    cursor.close()
    connStr.close()

    time.sleep(3600)