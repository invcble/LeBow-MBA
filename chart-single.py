import matplotlib.pyplot as plt
import numpy as np

# Dummy data
your_score = 6.0
drexel_mba_avg = 5.53
standard_deviation = 1.5

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
    plt.axvline(x=tick+1, color='lightgray', linewidth=0.5, zorder=1)

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
    plt.text(bar.get_width() - 0.2, bar.get_y() + bar.get_height()/2, 
             label, ha='right', va='center', color='white', fontsize=18)

# Add standard deviation next to the Drexel MBA Average bar
sd_label = f'SD: {standard_deviation:.2f}'.rstrip('0').rstrip('.')
plt.text(drexel_mba_avg + 0.2, y_pos[0], sd_label, ha='left', va='center', color='black', fontsize=18)

# Show the plot
plt.tight_layout()
plt.show()




# import matplotlib.pyplot as plt

# # Colors (Pantone 294C and Pantone 7548C)
# pantone_294C = '#002F6C'  # Navy Blue
# pantone_7548C = '#FFC72C'  # Yellow

# # Create figure for legend
# plt.figure(figsize=(6, 2))

# # Plot dummy points to create the legend
# plt.plot([], [], color=pantone_7548C, label='Your Score', marker='s', markersize=15, linestyle='None')
# plt.plot([], [], color=pantone_294C, label='Drexel MBA Average', marker='s', markersize=15, linestyle='None')

# # Display the legend with square icons and a different font
# legend = plt.legend(loc='center', fontsize=12, frameon=False, ncol=2, handletextpad=1, columnspacing=3,
#                     prop={'family': 'Verdana'})  # Change 'Verdana' to your preferred font

# # Set square icons in legend
# for handle in legend.legendHandles:
#     handle.set_marker('s')

# # Hide axes
# plt.axis('off')

# # Show the legend-only image
# plt.tight_layout()
# plt.show()

