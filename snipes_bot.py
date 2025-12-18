import discord
from discord import app_commands
from discord.ext import commands
import openpyxl
import os
from datetime import datetime
from . import private

# --- CONFIGURATION ---
TOKEN = private.token
OWNER_ID = private.owner_id
EXCEL_FILE = 'SNIPESSTATS.xlsm' 
CURRENT_SEASON = 'FALL2025'
# ---------------------

intents = discord.Intents.default()
bot = commands.Bot(command_prefix="!", intents=intents)

def save_to_excel(sniper_name, sniper_id, number, snipee_name, snipee_id, proof_url):
    """Saves names for display and IDs for formula stability."""
    
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not os.path.exists(EXCEL_FILE):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = CURRENT_SEASON
        # Added IDs to columns F and G
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

    # Data Entry
    sheet.cell(row=next_row, column=1).value = sniper_name
    sheet.cell(row=next_row, column=2).value = number
    sheet.cell(row=next_row, column=3).value = snipee_name
    sheet.cell(row=next_row, column=4).value = timestamp
    sheet.cell(row=next_row, column=5).value = proof_url
    # These are your "Anchor" columns for formulas
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
@app_commands.describe(number="Points value", user="Who did you snipe?", proof="Photo proof")
@app_commands.choices(number=[
    app_commands.Choice(name="1", value=1),
    app_commands.Choice(name="2", value=2)
])
async def snipe(interaction: discord.Interaction, number: int, user: discord.User, proof: discord.Attachment):
    # Capture both Name (for display) and ID (for formulas)
    sniper_name = interaction.user.name
    sniper_id = interaction.user.id
    snipee_name = user.name
    snipee_id = user.id
    
    try:
        save_to_excel(sniper_name, sniper_id, number, snipee_name, snipee_id, proof.url)
        
        await interaction.response.send_message(
            f"**<@{snipee_id}> got shot by {sniper_name} for {number} points**\n"
            f"{proof.url}"
        )

    except PermissionError:
        await interaction.response.send_message(
            f"⚠️ <@{OWNER_ID}> **NEEDS TO CLOSE THE EXCEL SHEET** snipes cannot be logged while the file is open."
        )

bot.run(TOKEN)