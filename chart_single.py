import matplotlib.pyplot as plt
import numpy as np

def plot_and_save_score(attribute, your_score, drexel_mba_avg, standard_deviation):
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
    for tick in np.arange(0, 8, 2):
        plt.axvline(x=tick + 1, color='lightgray', linewidth=0.5, zorder=1)

    # Set x-ticks with increased font size
    plt.xticks(np.arange(0, 8, 1), fontsize=14)

    # Create slimmer bars closer together
    bars = plt.barh(y_pos, scores, height=0.7, left=0.01, color=[pantone_294C, pantone_7548C], zorder=3)

    # Adjust the y-axis limits to create space between bars and borders
    plt.ylim(-0.6, len(categories) - 0.4)

    # Remove y-ticks
    plt.yticks([])

    # Set x-axis limit (Assuming a 1-7 scale)
    plt.xlim(1, 7)

    # Add labels within the bars, formatting to omit trailing zeros
    for i, bar in enumerate(bars):
        label = f'{scores[i]:.2f}'.rstrip('0').rstrip('.')
        plt.text(bar.get_width() - 0.2, bar.get_y() + bar.get_height() / 2,
                 label, ha='right', va='center', color='white', fontsize=18)

    # Add standard deviation next to the Drexel MBA Average bar
    sd_label = f'SD: {standard_deviation:.2f}'.rstrip('0').rstrip('.')
    plt.text(drexel_mba_avg + 0.2, y_pos[0], sd_label, ha='left', va='center', color='black', fontsize=18)

    # Save the plot as an image file named after the attribute
    plt.tight_layout()
    plt.savefig(f'CHART_IMAGES/{attribute}.png')
    plt.close()