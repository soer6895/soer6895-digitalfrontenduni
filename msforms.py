
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException

import pandas as pd
import time
data = pd.read_excel(
    '/Users/sorendahlmann/Desktop/Uni/digitalfront/Final_Generated_Canteen_Survey_Data.xlsx')
# print(data)
# print(data.columns)

sleepp = 0.02
failed_entries = []

driver = webdriver.Chrome()
time.sleep(1.5)


def fill_form(response):
    try:

        # open the MS Form webpage
        driver.get('https://forms.office.com/pages/responsepage.aspx?id=Nh39Ycv-yke319DfA3ChmEtvcebBEIVHjs1szO8lEXhUMDBRNFoyOTdINlhQTEpQWEUzQThBRTgyUi4u')

        # Wait for the page to load
        time.sleep(0.5)

        # selecting the gender
        gender = response['1. What is your gender?']
        if gender == 'Woman':
            driver.find_element(
                By.CSS_SELECTOR, '[aria-label="Woman"]').click()
        elif gender == 'Man':
            driver.find_element(By.CSS_SELECTOR, '[aria-label="Man"]').click()
        elif gender == 'Non-binary':
            driver.find_element(
                By.CSS_SELECTOR, '[aria-label="Non-binary"]').click()
        elif gender == 'Prefer not to say':
            driver.find_element(
                By.CSS_SELECTOR, '[aria-label="Prefer not to say"]').click()
        else:
            print(f"Invalid gender: {gender}")
            return False

        time.sleep(sleepp)

        # Handling the age question
        age_response = response['What is your age?']
        age_options = {
            'Under 18': '[aria-label="Under 18"]',
            '18-24': '[aria-label="18-24"]',
            '25-34': '[aria-label="25-34"]',
            '35-44': '[aria-label="35-44"]',
            '45-54': '[aria-label="45-54"]'
        }

        # find the radio group for the age question using its aria-labelledby attribute
        age_group = driver.find_element(
            By.CSS_SELECTOR, '[aria-labelledby="QuestionId_r43607cce379e48a2b7e2d46a2df0d11c QuestionInfo_r43607cce379e48a2b7e2d46a2df0d11c"]')

        # click the appropriate radio button
        age_group.find_element(
            By.CSS_SELECTOR, age_options[age_response]).click()

        time.sleep(sleepp)

        # selecting the canteen usage frequency
        usage = response['2. How often do you use the university canteen?']
        if usage == 'Daily':
            driver.find_element(
                By.CSS_SELECTOR, '[aria-label="Daily."]').click()
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
        else:
            print(f"Invalid gender: {usage}")
            return False
        time.sleep(sleepp)

        # Handling question 3
        satisfaction_rating = response['3. How satisfied are you with the quality of food offered in the canteen?']
        satisfaction_options = {
            1: '[aria-label="1 Bad"]',
            2: '[aria-label="2"]',
            3: '[aria-label="3"]',
            4: '[aria-label="4"]',
            5: '[aria-label="5"]',
            6: '[aria-label="6"]',
            7: '[aria-label="7"]',
            8: '[aria-label="8"]',
            9: '[aria-label="9"]',
            10: '[aria-label="10 Good"]'
        }

        # find the radio group for question 3 using its aria-labelledby attribute
        question_3_group = driver.find_element(
            By.CSS_SELECTOR, '[aria-labelledby="QuestionId_r491d56e985ef4692a196e50b5e8dd07e QuestionInfo_r491d56e985ef4692a196e50b5e8dd07e ll_QuestionId_r491d56e985ef4692a196e50b5e8dd07e rl_QuestionId_r491d56e985ef4692a196e50b5e8dd07e"]')

        # click the appropriate radio button
        question_3_group.find_element(
            By.CSS_SELECTOR, satisfaction_options[satisfaction_rating]).click()

        time.sleep(sleepp)

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

    # handling new columns
        handle_question("Vegetarian", response['Vegetarian'])
        time.sleep(0.1)

        handle_question("Vegan", response['Vegan'])
        time.sleep(0.1)

        handle_question("Gluten free", response['Gluten free'])
        time.sleep(0.1)

        handle_question("Halal and kosher", response['Halal and kosher'])

        # Handling question 5
        affordability = response['5. How would you rate the affordability of the food in the canteen?']
        affordability_options = {
            'Very affordable': '[aria-label="Very affordable"]',
            'Affordable': '[aria-label="Affordable"]',
            'Neutral': '[aria-label="Neutral"]',
            'Expensive': '[aria-label="Expensive"]',
            'Very expensive': '[aria-label="Very expensive"]'
        }
        time.sleep(sleepp)

        # Find the radio group for question 5 using its aria-labelledby attribute
        # Make sure to replace the values in the selector below with the correct ones for question 5
        question_5_group = driver.find_element(
            By.CSS_SELECTOR, '[aria-labelledby="QuestionId_r9f04d901963f476390b0053a9fad5536 QuestionInfo_r9f04d901963f476390b0053a9fad5536"]')

        # Click the appropriate radio button
        question_5_group.find_element(
            By.CSS_SELECTOR, affordability_options[affordability]).click()

        time.sleep(sleepp)

        # Handling question 6
        environment_rating = response['6. How would you rate the overall environment of the canteen?']
        environment_options = {
            'Very good': '[aria-label="Very good"]',
            'Good': '[aria-label="Good"]',
            'Average': '[aria-label="Average"]',
            'Poor': '[aria-label="Poor"]',
            'Very poor': '[aria-label="Very poor"]'
        }

        # Find the radio group for question 6 using its aria-labelledby attribute
        question_6_group = driver.find_element(
            By.CSS_SELECTOR, '[aria-labelledby="QuestionId_rac998bf9695f4e4fb1ac170c12e7869d QuestionInfo_rac998bf9695f4e4fb1ac170c12e7869d"]')

        # Click the appropriate radio button
        question_6_group.find_element(
            By.CSS_SELECTOR, environment_options[environment_rating]).click()

        time.sleep(sleepp)

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

        time.sleep(sleepp)

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

        time.sleep(sleepp)

        # Handling question 9
        seating_satisfaction = response['9. Are you satisfied with the seating arrangements and space management in the canteen, especially during peak hours?']
        seating_satisfaction_options = {
            'Very satisfied': '[aria-label="Very satisfied"]',
            'Satisfied': '[aria-label="Satisfied"]',
            'Neutral': '[aria-label="Neutral"]',
            'Dissatisfied': '[aria-label="Dissatisfied"]',
            'Very dissatisfied': '[aria-label="Very dissatisfied"]'
        }

        question_9_group = driver.find_element(
            By.CSS_SELECTOR, '[aria-labelledby="QuestionId_re201bb5fef574143846bfb0d157347c1 QuestionInfo_re201bb5fef574143846bfb0d157347c1"]')

        question_9_group.find_element(
            By.CSS_SELECTOR, seating_satisfaction_options[seating_satisfaction]).click()

        time.sleep(sleepp)

        # Handling question 10
        cleanliness_rating = response['10. How would you rate the cleanliness and hygiene standards in the canteen?']
        cleanliness_options = {
            'Excellent': '[aria-label="Excellent"]',
            'Good': '[aria-label="Good"]',
            'Average': '[aria-label="Average"]',
            'Below average': '[aria-label="Below average"]',
            'Poor': '[aria-label="Poor"]'
        }

        question_10_group = driver.find_element(
            By.CSS_SELECTOR, '[aria-labelledby="QuestionId_r8830aa5d0ff74165a31cd23e366fd4c1 QuestionInfo_r8830aa5d0ff74165a31cd23e366fd4c1"]')

        question_10_group.find_element(
            By.CSS_SELECTOR, cleanliness_options[cleanliness_rating]).click()

        time.sleep(sleepp)

        # Handling question 11
        recommendation_rating = response['11. How likely is it that you would recommend the canteen to a friend or colleague?']
        recommendation_options = {
            1: '[aria-label="1 Not likely at all"]',
            2: '[aria-label="2"]',
            3: '[aria-label="3"]',
            4: '[aria-label="4"]',
            5: '[aria-label="5"]',
            6: '[aria-label="6"]',
            7: '[aria-label="7"]',
            8: '[aria-label="8"]',
            9: '[aria-label="9"]',
            10: '[aria-label="10 Extremely likely"]'
        }

        # Find the radio group for question 11 using its aria-labelledby attribute
        question_11_group = driver.find_element(
            By.CSS_SELECTOR, '[aria-labelledby="QuestionId_rc9e7b087e0144b40a41479406f531da9 QuestionInfo_rc9e7b087e0144b40a41479406f531da9 ll_QuestionId_rc9e7b087e0144b40a41479406f531da9 rl_QuestionId_rc9e7b087e0144b40a41479406f531da9"]')

        if recommendation_rating in recommendation_options:
            # Click the button
            question_11_group.find_element(
                By.CSS_SELECTOR, recommendation_options[recommendation_rating]).click()
        else:
            print(
                f"Invalid recommendation rating at row {index}: {recommendation_rating}")
            return False

        time.sleep(sleepp)

        submit_button = driver.find_element(
            By.CSS_SELECTOR, '[data-automation-id="submitButton"]')
        submit_button.click()

        time.sleep(1)
    except WebDriverException as e:
        print(f"Error occurred: {e}")
        driver.refresh()
        return False


# Iterate over each row (respondent) in the DataFrame and fill the form
total_rows = len(data)
for index, row in data.iterrows():
    print(f"Processing row {index}")
    if total_rows > 40:
        break
    if not fill_form(row):
        failed_entries.append(row)

    if index == total_rows - 1:
        print("Processed all rows.")
        break

# Retry for failed entries
for row in failed_entries:
    fill_form(row)
