import telebot
import pandas as pd
import os
import datetime

# Getting today's date
current_date = datetime.date.today()

# Setting options for DataFrame display
pd.set_option('display.max_colwidth', 50)  # maximum column width
pd.set_option('display.max_rows', None)  # displaying all DataFrame rows
pd.set_option('display.unicode.east_asian_width', True)  # handling character width for adequate table display

# Bot token
TOKEN = ' '

# Initializing the bot
bot = telebot.TeleBot(TOKEN)

# Loading data from the pivot table file
pivot = pd.read_excel("Pivot.xlsx")

# Converting the "Quantity" column to numeric format and filling missing values with zeros
pivot["Quantity"] = pd.to_numeric(pivot["Quantity"]).fillna(0).astype(int)


def format_form(form):
    """Function to format data from the form."""
    # Converting the "Quantity" column to numeric format, replacing missing values with zeros, and converting to integers
    form["Quantity"] = pd.to_numeric(form["Quantity"], errors='coerce').fillna(0).astype(int)
    # Selecting the necessary columns and removing the first row, resetting indexes
    form = form[['Item', 'Quantity']].iloc[1:109].reset_index(drop=True)
    return form


# Handler for the /start command
@bot.message_handler(commands=['start'])
def send_welcome(message):
    """Handler for the /start command."""
    with open("form.xlsx", 'rb') as form:
        # Sending a welcome message
        bot.reply_to(message,
                     "Good day! Welcome to PontBakery. To place an order, fill out this form and send it in a reply message. Change the file name to the name of your point.")
        # Sending the file to the user
        bot.send_document(message.chat.id, form)

# Handler for receiving filled forms
@bot.message_handler(content_types=['document'])
def handle_document(message):
    """Handler for filled forms."""

    # Downloading the file
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    # Saving the received file
    with open(message.document.file_name, 'wb') as new_file:
        new_file.write(downloaded_file)

    # Reading and formatting the client's Excel file
    df = pd.read_excel(message.document.file_name, header=2)
    df = format_form(df)

    # Updating data in the pivot table
    pivot["Quantity"] = pivot["Quantity"] + df["Quantity"]
    pivot.to_excel("pivots/"+str(current_date) + ".xlsx", index=False)

    # Displaying the order contents with non-zero quantities
    print(df[df["Quantity"] > 0].to_string(index=False))

    # Formatting the table for sending to the chat
    formatted_table = "Quantity   Item      \n"
    for index, row in df[df["Quantity"] > 0].iterrows():
        name = row['Item']
        quantity = str(row['Quantity'])
        number_of_digits = int(len(str(quantity)))
        formatted_table += " " * (6 - number_of_digits) + quantity + " " * (16 - number_of_digits) + name + "\n"

    response = "Your order: \n" + formatted_table + "\n Order accepted"

    # Sending the response to the user
    bot.reply_to(message, response)

    # Deleting the file after use
    os.remove(message.document.file_name)


# Running the bot
bot.polling()
