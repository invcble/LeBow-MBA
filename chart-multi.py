import matplotlib.pyplot as plt
import numpy as np

# Data
categories = ['Strong', 'Weak', 'Size']
drexel_mba_scores = [3, 3, 6]
your_scores = [3, 1, 4]

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

# Label the bars with numbers (without unnecessary decimals)
for i in range(len(categories)):
    plt.text(drexel_mba_scores[i] + 0.2, y_pos[i] - 0.2, str(drexel_mba_scores[i]), va='center', fontsize=12, color='black')
    plt.text(your_scores[i] + 0.2, y_pos[i] + 0.2, str(your_scores[i]), va='center', fontsize=12, color='black')

# Y-axis labels
plt.yticks(y_pos, categories)

# X-axis limits and ticks (Set scale from 0 to 20 with increments of 5)
plt.xlim(0, 20)
plt.xticks(np.arange(0, 21, 5), fontsize=14)

# Ensure the y-axis line is visible
plt.axvline(x=0, color='black', linewidth=1)  # Y-axis

# Show the plot
plt.tight_layout()
plt.show()
