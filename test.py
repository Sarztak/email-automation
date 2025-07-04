import datetime

# Get the current time
current_time = datetime.datetime.now()

# Subtract two days from the current time
past_time = current_time - datetime.timedelta(hours=2)

# Format the date and time
formatted_date = past_time.strftime('%m/%d/%Y %H:%M %p')

# Print the formatted date
print("Current Time:", current_time.strftime('%m/%d/%Y %H:%M %p'))
print("Time Two Days Ago:", formatted_date)
