import numpy as np
import pandas as pd
from scipy.stats import norm
import random

# function to convert normal distribution values to categories


def normal_dist_to_categories(n, categories, skewness_range=(-0.5, 0.5)):
    # randomly deciding the skewness (left or right)
    skewness = random.uniform(*skewness_range)
    # generating normal distribution values with random skewness
    values = norm.rvs(size=n, loc=skewness, scale=1)
    # normalizing the values to the range [0, 1]
    values_normalized = (values - values.min()) / (values.max() - values.min())
    # Converting normalized values to categories
    categories_idx = (values_normalized * len(categories)).astype(int)
    categories_idx = np.clip(categories_idx, 0, len(categories) - 1)
    return [categories[i] for i in categories_idx]


# Survey question options and headings
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

# number of responses to generate
num_responses = 46

# adjusting the gender distribution
gender_distribution = np.random.choice(
    gender_options, num_responses, p=[0.55, 0.35, 0.05, 0.05])

# food quality ratings with a tendency to be below average
food_quality_ratings = np.random.choice(food_quality_options, num_responses, p=[
                                        0.05, 0.10, 0.20, 0.25, 0.10, 0.10, 0.05, 0.05, 0.05, 0.05])

# higher probability for Neutral and Dissatisfied
staff_interaction_probabilities = [0.10, 0.20, 0.20, 0.40, 0.10]
staff_interaction_ratings = np.random.choice(
    staff_interaction_options, num_responses, p=staff_interaction_probabilities)


# generating the likelihood of recommending the canteen independently
recommendation_likelihood = np.random.choice(food_quality_options, num_responses, p=[
    0.05, 0.10, 0.15, 0.25, 0.15, 0.10, 0.05, 0.05, 0.05, 0.05])


# higher probability for 18-24 and 25-34 age groups
age_probabilities = [0.05, 0.45, 0.45, 0.05]
age_distribution = np.random.choice(
    age_options, num_responses, p=age_probabilities)


# generating the dataset with normal distribution for specific categories
generated_data_with_all_factors = pd.DataFrame({
    survey_headings[0]: gender_distribution,
    survey_headings[1]: normal_dist_to_categories(num_responses, frequency_options),
    survey_headings[2]: food_quality_ratings,
    survey_headings[3]: normal_dist_to_categories(num_responses, Vegetarian_options),
    survey_headings[4]: normal_dist_to_categories(num_responses, Vegan_options),
    survey_headings[5]: normal_dist_to_categories(num_responses, Gluten_free_options),
    survey_headings[6]: normal_dist_to_categories(num_responses, Halal_and_kosher_options),
    survey_headings[7]: normal_dist_to_categories(num_responses, affordability_options),
    survey_headings[8]: normal_dist_to_categories(num_responses, environment_options),
    survey_headings[9]: normal_dist_to_categories(num_responses, waiting_time_options),
    survey_headings[10]: staff_interaction_ratings,
    survey_headings[11]: normal_dist_to_categories(num_responses, seating_options),
    survey_headings[12]: normal_dist_to_categories(num_responses, cleanliness_options),
    survey_headings[13]: recommendation_likelihood,
    survey_headings[14]: age_distribution
})

final_output_file_path = 'Final_Generated_Canteen_Survey_Data123.xlsx'
generated_data_with_all_factors.to_excel(final_output_file_path, index=False)
