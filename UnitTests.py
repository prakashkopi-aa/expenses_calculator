# This just displays a pie chart with dummy data
import matplotlib.pyplot as plt

# Data to plot
labels = ['A', 'B', 'C', 'D']
amounts = [100, 200, 300, 50]  # dollar amounts
colors = ['gold', 'yellowgreen', 'lightcoral', 'lightskyblue']

# Calculate total amount
total_amount = sum(amounts)

# Define a function to format the label
def autopct_label(pct):
    absolute = int(pct/100.*total_amount)
    return '${:,.0f} ({:.1f}%)'.format(absolute, pct)

# Plotting the pie chart
plt.pie(amounts, labels=labels, colors=colors, autopct=autopct_label, startangle=140)

# Equal aspect ratio ensures that pie is drawn as a circle.
plt.axis('equal')

# Show the plot
plt.show()