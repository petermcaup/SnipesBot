import discord
from discord import app_commands
from discord.ext import commands
import openpyxl
import os
from datetime import datetime
# Ensure private.py contains TOKEN and OWNER_ID
from private import private

# --- CONFIGURATION ---
TOKEN = private.token
OWNER_ID = private.owner_id
EXCEL_FILE = 'SNIPESSTATS.xlsm' 
CURRENT_SEASON = 'SPRING2026'
# ---------------------

intents = discord.Intents.default()
bot = commands.Bot(command_prefix="!", intents=intents)

def save_to_excel(sniper_name, sniper_id, number, snipee_name, snipee_id, proof_url):
    """Adds all relevant data to excel sheet."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Load or Create Workbook and Sheet
    if not os.path.exists(EXCEL_FILE):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = CURRENT_SEASON
        sheet.append(["Sniper", "Points", "Snipee", "Timestamp", "Proof Link", "Sniper ID", "Snipee ID"])
    else:
        workbook = openpyxl.load_workbook(EXCEL_FILE, keep_vba=True)
        if CURRENT_SEASON in workbook.sheetnames:
            sheet = workbook[CURRENT_SEASON]
        else:
            sheet = workbook.create_sheet(CURRENT_SEASON)
            sheet.append(["Sniper", "Points", "Snipee", "Timestamp", "Proof Link", "Sniper ID", "Snipee ID"])

    # Find next empty row in Column A
    next_row = 1
    while sheet.cell(row=next_row, column=1).value is not None:
        next_row += 1

    # Data Entry - IDs included for formula use
    sheet.cell(row=next_row, column=1).value = sniper_name
    sheet.cell(row=next_row, column=2).value = number
    sheet.cell(row=next_row, column=3).value = snipee_name
    sheet.cell(row=next_row, column=4).value = timestamp
    sheet.cell(row=next_row, column=5).value = proof_url
    sheet.cell(row=next_row, column=6).value = str(sniper_id) 
    sheet.cell(row=next_row, column=7).value = str(snipee_id)
    
    workbook.save(EXCEL_FILE)

@bot.event
async def on_ready():
    print(f'Logged in as {bot.user}!')
    try:
        synced = await bot.tree.sync()
        print(f"Synced {len(synced)} command(s).")
    except Exception as e:
        print(e)

@bot.tree.command(name="snipe", description="Add a Snipe to the Excel Sheet")
@app_commands.describe(number="Points value", user="Who did you snipe? (Leave blank for Alumni)", proof="Photo proof")
@app_commands.choices(number=[
    app_commands.Choice(name="1", value=1),
    app_commands.Choice(name="2", value=2),
    app_commands.Choice(name="Alumni Snipe", value=5)
])
# Set user: discord.User = None to make it optional in Discord UI
async def snipe(interaction: discord.Interaction, number: int, proof: discord.Attachment, user: discord.User = None):
    # 1. Defer to avoid timeout
    await interaction.response.defer()
    
    sniper_name = interaction.user.name
    sniper_id = interaction.user.id

    # 2. Check for Alumni Logic
    if user is None:
        if number == 5:
            snipee_name = "Alumni"
            snipee_id = "0000" # Placeholder ID for your formulas
            display_message = f"**{sniper_name} got an Alumni Snipe for 5 points!**"
        else:
            # Error case: User skipped 'user' but didn't pick Alumni Snipe
            await interaction.followup.send("❌ You must select a user unless it is an Alumni Snipe (5 pts).")
            return
    else:
        snipee_name = user.name
        snipee_id = user.id
        display_message = f"**<@{snipee_id}> got shot by {sniper_name} for {number} points**"
    
    # 3. Save and Respond
    try:
        save_to_excel(sniper_name, sniper_id, number, snipee_name, snipee_id, proof.url)
        
        await interaction.followup.send(f"{display_message}\n{proof.url}")

    except PermissionError:
        await interaction.followup.send(
            f"**ERROR LOGGING SNIPE**\n"
            f"⚠️ <@{OWNER_ID}> **NEEDS TO CLOSE THE EXCEL SHEET**"
        )

bot.run(TOKEN)