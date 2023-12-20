
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time
# Load the survey data from Excel
data = pd.read_excel(
    '/Users/sorendahlmann/Desktop/Uni/digitalfront/Canteen_Survey_AU_Herning (2).xlsx')
# print(data)
print(data.columns)


def fill_form(response):

    # Open the MS Form webpage
    driver.get('https://forms.office.com/pages/responsepage.aspx?id=Nh39Ycv-yke319DfA3ChmEtvcebBEIVHjs1szO8lEXhUMDBRNFoyOTdINlhQTEpQWEUzQThBRTgyUi4u')

    # Wait for the page to load
    time.sleep(2)

    # Selecting the gender
    gender = response['1. What is your gender?']
    if gender == 'Woman':
        driver.find_element(By.CSS_SELECTOR, '[aria-label="Woman"]').click()
    elif gender == 'Man':
        driver.find_element(By.CSS_SELECTOR, '[aria-label="Man"]').click()
    elif gender == 'Non-binary':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Non-binary"]').click()
    elif gender == 'Prefer not to say':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Prefer not to say"]').click()

    time.sleep(0.5)

    # Selecting the canteen usage frequency
    usage = response['2. How often do you use the university canteen?']
    if usage == 'Daily':
        driver.find_element(By.CSS_SELECTOR, '[aria-label="Daily."]').click()
    elif usage == 'Twice a week':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Twice a week."]').click()
    elif usage == 'Three times a week':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Three times a week."]').click()
    elif usage == 'Four times a week':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Four times a week."]').click()
    elif usage == 'Five times a week':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Five times a week."]').click()
    time.sleep(0.5)

    satisfaction_rating = response['3. How satisfied are you with the quality of food offered in the canteen?']
    if satisfaction_rating == 1:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="1 Bad"]').click()
    elif satisfaction_rating == 2:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="2"]').click()
    elif satisfaction_rating == 3:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="3"]').click()
    elif satisfaction_rating == 4:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="4"]').click()
    elif satisfaction_rating == 5:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="5"]').click()
    elif satisfaction_rating == 6:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="6"]').click()
    elif satisfaction_rating == 7:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="7"]').click()
    elif satisfaction_rating == 8:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="8"]').click()
    elif satisfaction_rating == 9:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="9"]').click()
    elif satisfaction_rating == 10:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="10 Good"]').click()
    time.sleep(0.5)

    def handle_question(question_label, response_value):
        if response_value == 'Poor':
            driver.find_element(
                By.CSS_SELECTOR, f'[aria-label="{question_label}, Poor"]').click()
        elif response_value == 'Below average':
            driver.find_element(
                By.CSS_SELECTOR, f'[aria-label="{question_label}, Below average"]').click()
        elif response_value == 'Average':
            driver.find_element(
                By.CSS_SELECTOR, f'[aria-label="{question_label}, Average"]').click()

        elif response_value == 'Good':
            driver.find_element(
                By.CSS_SELECTOR, f'[aria-label="{question_label}, Good"]').click()
        elif response_value == 'Above good':
            driver.find_element(
                By.CSS_SELECTOR, f'[aria-label="{question_label}, Above good"]').click()
    time.sleep(0.5)

 # Handling new columns
    handle_question("Vegetarian", response['Vegetarian'])
    handle_question("Vegan", response['Vegan'])
    handle_question("Gluten free", response['Gluten free'])
    handle_question("Halal and kosher", response['Halal and kosher'])

    affordability = response['5. How would you rate the affordability of the food in the canteen?']
    if affordability == 'Very affordable':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Very affordable"]').click()
    elif affordability == 'Affordable':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Affordable"]').click()
    elif affordability == 'Neutral':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Neutral"]').click()
    elif affordability == 'Expensive':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Expensive"]').click()
    elif affordability == 'Very expensive':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Very expensive"]').click()
    time.sleep(0.5)

    environment_rating = response['6. How would you rate the overall environment of the canteen?']
    if environment_rating == 'Very good':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Very good"]').click()
    elif environment_rating == 'Good':
        driver.find_element(By.CSS_SELECTOR, '[aria-label="Good"]').click()
    elif environment_rating == 'Average':
        driver.find_element(By.CSS_SELECTOR, '[aria-label="Average"]').click()
    elif environment_rating == 'Poor':
        driver.find_element(By.CSS_SELECTOR, '[aria-label="Poor"]').click()
    elif environment_rating == 'Very poor':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Very poor"]').click()
    time.sleep(0.5)

    waiting_time_satisfaction = response['7. How satisfied are you with the waiting time in the canteen during peak hours?']
    waiting_time_options = {
        'Very satisfied': '[aria-label="Very satisfied"]',
        'Satisfied': '[aria-label="Satisfied"]',
        'Neutral': '[aria-label="Neutral"]',
        'Dissatisfied': '[aria-label="Dissatisfied"]',
        'Very dissatisfied': '[aria-label="Very dissatisfied"]'
    }
    question_7_group = driver.find_element(
        By.CSS_SELECTOR, '[aria-labelledby="QuestionId_r70df4624aab8415dbc1ddd8e2809fc9b QuestionInfo_r70df4624aab8415dbc1ddd8e2809fc9b"]')

    question_7_group.find_element(
        By.CSS_SELECTOR, waiting_time_options[waiting_time_satisfaction]).click()

    time.sleep(0.5)

    # Handling question 8
    staff_interaction = response['8. How would you rate your interactions with the canteen staff in terms of friendliness and efficiency?']
    staff_interaction_options = {
        'Very satisfied': '[aria-label="Very satisfied"]',
        'Satisfied': '[aria-label="Satisfied"]',
        'Neutral': '[aria-label="Neutral"]',
        'Dissatisfied': '[aria-label="Dissatisfied"]',
        'Very dissatisfied': '[aria-label="Very dissatisfied"]'
    }

    question_8_group = driver.find_element(
        By.CSS_SELECTOR, '[aria-labelledby="QuestionId_r9aef720467fe47698e1cabe397846d37 QuestionInfo_r9aef720467fe47698e1cabe397846d37"]')

    question_8_group.find_element(
        By.CSS_SELECTOR, staff_interaction_options[staff_interaction]).click()

    time.sleep(0.5)

    seating_satisfaction = response['9. Are you satisfied with the seating arrangements and space management in the canteen, especially during peak hours?']
    if seating_satisfaction == 'Very satisfied':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Very satisfied"]').click()
    elif seating_satisfaction == 'Satisfied':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Satisfied"]').click()
    elif seating_satisfaction == 'Neutral':
        driver.find_element(By.CSS_SELECTOR, '[aria-label="Neutral"]').click()
    elif seating_satisfaction == 'Dissatisfied':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Dissatisfied"]').click()
    elif seating_satisfaction == 'Very dissatisfied':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Very dissatisfied"]').click()
    time.sleep(0.5)

    cleanliness_rating = response['10. How would you rate the cleanliness and hygiene standards in the canteen?']
    if cleanliness_rating == 'Excellent':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Excellent"]').click()
    elif cleanliness_rating == 'Good':
        driver.find_element(By.CSS_SELECTOR, '[aria-label="Good"]').click()
    elif cleanliness_rating == 'Average':
        driver.find_element(By.CSS_SELECTOR, '[aria-label="Average"]').click()
    elif cleanliness_rating == 'Below average':
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="Below average"]').click()
    elif cleanliness_rating == 'Poor':
        driver.find_element(By.CSS_SELECTOR, '[aria-label="Poor"]').click()
    time.sleep(0.5)

    recommendation_rating = response['11. How likely is it that you would recommend the canteen to a friend or colleague?']
    if recommendation_rating == 1:
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="1 Not likely at all"]').click()
    elif recommendation_rating == 2:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="2"]').click()
    elif recommendation_rating == 3:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="3"]').click()
    elif recommendation_rating == 4:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="4"]').click()
    elif recommendation_rating == 5:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="5"]').click()
    elif recommendation_rating == 6:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="6"]').click()
    elif recommendation_rating == 7:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="7"]').click()
    elif recommendation_rating == 8:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="8"]').click()
    elif recommendation_rating == 9:
        driver.find_element(By.CSS_SELECTOR, '[aria-label="9"]').click()
    elif recommendation_rating == 10:
        driver.find_element(
            By.CSS_SELECTOR, '[aria-label="10 Extremely likely"]').click()

    # ... [rest of the script] ...

    # Wait a bit before filling the next form
    time.sleep(15)


# Iterate over each row (respondent) in the DataFrame and fill the form
for index, row in data.iterrows():
    driver = webdriver.Chrome()

    fill_form(row)
    driver.quit()


# Close the browser when done
