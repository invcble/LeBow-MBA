

import matplotlib.pyplot as plt

# Colors (Pantone 294C and Pantone 7548C)
pantone_294C = '#002F6C'  # Navy Blue
pantone_7548C = '#FFC72C'  # Yellow

# Create figure for legend
plt.figure(figsize=(6, 2))

# Plot dummy points to create the legend
plt.plot([], [], color=pantone_7548C, label='Your Score', marker='s', markersize=15, linestyle='None')
plt.plot([], [], color=pantone_294C, label='Drexel MBA Average', marker='s', markersize=15, linestyle='None')

# Display the legend with square icons and a different font
legend = plt.legend(loc='center', fontsize=12, frameon=False, ncol=2, handletextpad=1, columnspacing=3,
                    prop={'family': 'Verdana'})  # Change 'Verdana' to your preferred font

# Set square icons in legend
for handle in legend.legendHandles:
    handle.set_marker('s')

# Hide axes
plt.axis('off')

# Show the legend-only image
plt.tight_layout()
plt.show()

