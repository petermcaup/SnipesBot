HELLO
This is a personal project to automate the data collection for a game of snipes using a discord bot

#### Bot Commands: ####
- /snipe : takes in a number (1 or 2), a username of the person you sniped, and requires that you attach an image when running the command, for proof
  - This info (along with a timestamp) is entered into the excel sheet which has a macro to handle new data (refreshed pivot tables, sorts charts, etc.)
  - Once the data is properly recorded in the spreadsheet, the bot sends a message confirming this along with the proof image and tags the snipee so they are aware they've been shot
  - If the excel sheet is open when the bot tries to record data, it handles the error gracefully and sends a message in discord tagging the owner of the bot, telling them to close the sheet before the sniper resubmits

If you'd like to run the bot yourself you can totally do that! My plan is to run the bot either from my desktop PC that I'll just leave on all the time or figure out how to run the bot from my raspberry pi which is currently unused.
#### Installation instructions: ####
Setup virtual env (which will end up in a .venv folder)
From within the venv:
```bash
# Install Requirements
pip install -r requirements.txt

# Make any changes to the bot by editing snipes_bot.py

# Run the bot
python snipes_bot.py
```

Future features / To-Do list:
- /stats or /leaderboard for users to quickly get info about the standings of the game
- /appeal or /remove for users to request a change or deletion of a snipe (this would remove the need for an admin to go into the spreadsheet and delete or change data)
- A check for what channel the command is being run from or have the bot send the snipe confirmation message in the #ssnipes chat no matter where the user runs the command from
- Make the excel macro run upon opening the spreadsheet to ensure clean and sorted data whenever viewing