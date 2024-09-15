import matplotlib.pyplot as plt
import numpy as np
import os

def plot_and_save_multi(your_scores, drexel_mba_scores, switch_categories=False):
    # Define categories with the switch
    if switch_categories:
        categories = ['Higher Level Contacts', 'External Contacts', 'Cross-Functional Contacts']
        # Convert scores to percentages
        your_scores = [score * 100 for score in your_scores]
        drexel_mba_scores = [score * 100 for score in drexel_mba_scores]
        x_limit = 100  # Set x-axis limit for percentages
        x_ticks = np.arange(0, 101, 20)  # Set x-axis ticks for percentages
        score_format = '{:.0f}%'  # Format labels as percentages
    else:
        categories = ['Strong', 'Weak', 'Size']
        # Cast scores to integers
        your_scores = [int(score) for score in your_scores]
        drexel_mba_scores = [int(score) for score in drexel_mba_scores]
        x_limit = 20  # Set x-axis limit for regular scale
        x_ticks = np.arange(0, 21, 5)  # Set x-axis ticks for regular scale
        score_format = '{:.0f}'  # Format labels as integers

    # Adjusted bar positions with added gaps
    y_pos = np.array([2, 1, 0]) * 1.2  # Adding gaps by multiplying y positions

    # Colors (Pantone 294C and Pantone 7548C)
    pantone_294C = '#002F6C'  # Navy Blue for Drexel MBA Average
    pantone_7548C = '#FFC72C'  # Yellow for Your Score

    # Create horizontal bar chart
    plt.figure(figsize=(8, 4))

    # Plotting the bars for Drexel MBA Average and Your Score with spacing
    plt.barh(y_pos - 0.2, drexel_mba_scores, height=0.4, color=pantone_294C, label='Drexel MBA Average')
    plt.barh(y_pos + 0.2, your_scores, height=0.4, color=pantone_7548C, label='Your Score')

    # Label the bars with formatted numbers
    for i in range(len(categories)):
        plt.text(drexel_mba_scores[i] + 0.5, y_pos[i] - 0.2, score_format.format(drexel_mba_scores[i]), 
                 va='center', fontsize=12, color='black')
        plt.text(your_scores[i] + 0.5, y_pos[i] + 0.2, score_format.format(your_scores[i]), 
                 va='center', fontsize=12, color='black')

    # Y-axis labels
    plt.yticks(y_pos, categories)

    # X-axis limits and ticks
    plt.xlim(0, x_limit)
    plt.xticks(x_ticks, fontsize=14)

    # Ensure the y-axis line is visible
    plt.axvline(x=0, color='black', linewidth=1)  # Y-axis

    # Save the plot as an image file with the appropriate name
    plt.tight_layout()

    if not os.path.exists('CHART_IMAGES'):
        os.makedirs('CHART_IMAGES')
    if switch_categories:
        plt.savefig('CHART_IMAGES/breadth.png')
    else:
        plt.savefig('CHART_IMAGES/strength.png')
    
    plt.close()
