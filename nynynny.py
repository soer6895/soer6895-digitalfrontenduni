import pandas as pd
import numpy as np

# Survey question options
gender_options = ['Man', 'Woman', 'Non-binary', 'Prefer not to say']
frequency_options = ['Daily', 'Twice a week',
                     'Three times a week', 'Four times a week', 'Five times a week']
food_quality_options = list(range(1, 11))  # 1 being Bad, 10 being Good
Vegetarian_options = ['Poor', 'Below average', 'Average', 'Good', 'Above good']
Vegan_options = ['Poor', 'Below average', 'Average', 'Good', 'Above good']
Gluten_free_options = ['Poor', 'Below average',
                       'Average', 'Good', 'Above good']
Halal_and_kosher_options = [
    'Poor', 'Below average', 'Average', 'Good', 'Above good']
affordability_options = ['Very affordable',
                         'Affordable', 'Neutral', 'Expensive', 'Very expensive']
environment_options = ['Very good', 'Good', 'Average', 'Poor', 'Very poor']
waiting_time_options = ['Very satisfied', 'Satisfied',
                        'Neutral', 'Dissatisfied', 'Very dissatisfied']
staff_interaction_options = [
    'Very satisfied', 'Satisfied', 'Neutral', 'Dissatisfied', 'Very dissatisfied']
seating_options = ['Very satisfied', 'Satisfied',
                   'Neutral', 'Dissatisfied', 'Very dissatisfied']
cleanliness_options = ['Excellent', 'Good', 'Average', 'Below average', 'Poor']
# 1 being Not likely at all, 10 being Extremely likely
recommend_options = list(range(1, 11))
age_options = ['Under 18', '18-24', '25-34', '45-54']

# Survey question headings
survey_headings = [
    "1. What is your gender?",
    "2. How often do you use the university canteen?",
    "3. How satisfied are you with the quality of food offered in the canteen?",
    "Vegetarian", "Vegan", "Gluten free", "Halal and kosher",
    "5. How would you rate the affordability of the food in the canteen?",
    "6. How would you rate the overall environment of the canteen?",
    "7. How satisfied are you with the waiting time in the canteen during peak hours?",
    "8. How would you rate your interactions with the canteen staff in terms of friendliness and efficiency?",
    "9. Are you satisfied with the seating arrangements and space management in the canteen, especially during peak hours?",
    "10. How would you rate the cleanliness and hygiene standards in the canteen?",
    "11. How likely is it that you would recommend the canteen to a friend or colleague?",
    "What is your age?"
]

# Number of responses to generate (can be adjusted as needed)
num_responses = 20

# Adjusting the gender distribution to have more men than women
gender_distribution = np.random.choice(
    gender_options, num_responses, p=[0.55, 0.35, 0.05, 0.05])

# Food quality ratings with a tendency to be rated below average (1 to 10 scale)
food_quality_ratings = np.random.choice(food_quality_options, num_responses, p=[
                                        0.15, 0.15, 0.15, 0.15, 0.10, 0.10, 0.05, 0.05, 0.05, 0.05])

# Staff interaction satisfaction (1 to 5 scale)
# Higher probability for Neutral and Dissatisfied
staff_interaction_probabilities = [0.10, 0.20, 0.40, 0.20, 0.10]
staff_interaction_ratings = np.random.choice(
    staff_interaction_options, num_responses, p=staff_interaction_probabilities)

# Overall environment rating (1 to 5 scale), positively correlated with staff interaction
environment_ratings = np.array([environment_options.index(
    rating) + 1 for rating in staff_interaction_ratings]) + np.random.randint(-1, 2, num_responses)
environment_ratings = np.clip(environment_ratings, 1, 5)
environment_ratings = [environment_options[rating - 1]
                       for rating in environment_ratings]

# Likelihood of recommending the canteen (1 to 10 scale), negatively correlated with food quality
recommendation_likelihood = 11 - food_quality_ratings + \
    np.random.randint(-1, 2, num_responses)
recommendation_likelihood = np.clip(recommendation_likelihood, 1, 10)

# Age distribution focusing mostly on 18-24 and 25-34
age_distribution = np.random.choice(
    ['18-24', '25-34'], num_responses, p=[0.7, 0.3])

# Generating random responses for the remaining questions
generated_data_with_factors = pd.DataFrame({
    survey_headings[0]: gender_distribution,
    survey_headings[1]: np.random.choice(frequency_options, num_responses),
    survey_headings[2]: food_quality_ratings,
    survey_headings[3]: np.random.choice(Vegetarian_options, num_responses),
    survey_headings[4]: np.random.choice(Vegan_options, num_responses),
    survey_headings[5]: np.random.choice(Gluten_free_options, num_responses),
    survey_headings[6]: np.random.choice(Halal_and_kosher_options, num_responses),
    survey_headings[7]: np.random.choice(affordability_options, num_responses),
    survey_headings[8]: environment_ratings,
    survey_headings[9]: np.random.choice(waiting_time_options, num_responses),
    survey_headings[10]: staff_interaction_ratings,
    survey_headings[11]: np.random.choice(seating_options, num_responses),
    survey_headings[12]: np.random.choice(cleanliness_options, num_responses),
    survey_headings[13]: recommendation_likelihood,
    survey_headings[14]: age_distribution
})

# Saving the generated dataset to an Excel file
output_file_path = 'Generated_Canteen_Survey_Data.xlsx'
generated_survey_data.to_excel(output_file_path, index=False)
