import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import os

def plot_and_save_single(attribute, your_score, drexel_mba_avg, standard_deviation, switch_scale=False):
    # Bar positions (Reversed to show yellow bar above blue bar)
    categories = ['Drexel MBA Average', 'Your Score']
    scores = [drexel_mba_avg, your_score]
    y_pos = np.arange(len(categories))

    # Colors (Pantone 294C and Pantone 7548C)
    pantone_294C = '#002F6C'  # Navy Blue
    pantone_7548C = '#FFC72C'  # Yellow

    # Create horizontal bar chart
    plt.figure(figsize=(8, 4))

    # Add light gray grid lines for every 2nd tick
    if attribute == 'Total_Size':
        scale_max = 20
        x_limit = max(scale_max, max(scores))  # Add a buffer to the x-limit
        for tick in np.arange(0, x_limit + 1, 5):
            plt.axvline(x=tick, color='lightgray', linewidth=0.5, zorder=1)
        plt.xticks(np.arange(0, x_limit + 1, 5), fontsize=14)
    else:
        scale_max = 5 if switch_scale else 7  # Set max scale based on switch
        x_limit = max(scale_max, max(scores))  # Add a buffer to the x-limit
        for tick in np.arange(0, x_limit + 1, 2):
            plt.axvline(x=tick, color='lightgray', linewidth=0.5, zorder=1)
        plt.xticks(np.arange(0, x_limit + 1, 1), fontsize=14)

    # Create slimmer bars closer together
    bars = plt.barh(y_pos, scores, height=0.7, left=0.01, color=[pantone_294C, pantone_7548C], zorder=3)

    # Adjust the y-axis limits to create space between bars and borders
    plt.ylim(-0.6, len(categories) - 0.4)

    # Remove y-ticks
    plt.yticks([])

    # Set x-axis limit to start slightly before 1 for better visibility
    plt.xlim(0.95, x_limit)  # Start x-axis at 0.5 for a bit of scale before 1

    # Add labels within the bars, formatting to omit trailing zeros
    for i, bar in enumerate(bars):
        label = f'{scores[i]:.2f}'.rstrip('0').rstrip('.')
        label_position = max(1, bar.get_width() - 0.2)  # Ensure label position is not too close to the start
        plt.text(label_position, bar.get_y() + bar.get_height() / 2,
                 label, ha='right', va='center', color='white', fontsize=18)

    # Add standard deviation next to the Drexel MBA Average bar
    sd_label = f'SD: {standard_deviation:.2f}'.rstrip('0').rstrip('.')
    sd_position = drexel_mba_avg + 0.5  # Add some spacing for readability
    plt.text(sd_position, y_pos[0], sd_label, ha='left', va='center', color='black', fontsize=18)

    # Save the plot as an image file named after the attribute
    plt.tight_layout()
    if not os.path.exists('CHART_IMAGES'):
        os.makedirs('CHART_IMAGES')
    plt.savefig(f'CHART_IMAGES/{attribute}.png')
    plt.close()



# # Test case setup for the plot_and_save_single function
# attributes = [
#     ('Total3_Size', 7, 3.21, 1.46, False) # Low score with scale switch to 5
# ]

# # Run the function with different test cases
# for attribute, your_score, drexel_mba_avg, standard_deviation, switch_scale in attributes:
#     plot_and_save_single(attribute, your_score, drexel_mba_avg, standard_deviation, switch_scale)

# print("Test plots saved to the 'CHART_IMAGES' directory.")
