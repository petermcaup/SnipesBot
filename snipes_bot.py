import discord
from discord import app_commands
from discord.ext import commands
import openpyxl
import json
from datetime import datetime
from private import private
import sys
import os

# --- DYNAMIC PATHING ---
if getattr(sys, 'frozen', False):
    # Path of the .exe inside the 'dist' folder
    EXE_LOCATION = os.path.dirname(sys.executable)
    # Move UP one level to the main 'SnipesBot' folder
    BASE_DIR = os.path.dirname(EXE_LOCATION)
else:
    # If running as a .py script, assume it's already in the main folder
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --- CONFIGURATION ---
TOKEN = private.token
OWNER_ID = int(private.owner_id) 

# These will now point to SnipesBot\SNIPESSTATS.xlsm instead of SnipesBot\dist\SNIPESSTATS.xlsm
EXCEL_FILE = os.path.join(BASE_DIR, 'SNIPESSTATS.xlsm') 
REG_FILE = os.path.join(BASE_DIR, 'private', 'registrations.json')

# Ensure the private directory exists in the main folder
PRIVATE_DIR = os.path.join(BASE_DIR, 'private')
if not os.path.exists(PRIVATE_DIR):
    os.makedirs(PRIVATE_DIR)

print(f"Bot starting... Working Directory: {BASE_DIR}")

# --- DATA PERSISTENCE HELPERS ---

def load_data():
    """Loads season and registration data from JSON."""
    if not os.path.exists(REG_FILE):
        # Default state if no file exists
        return {"season": "SPRING2026", "registrations": {}}
    with open(REG_FILE, 'r') as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return {"season": "SPRING2026", "registrations": {}}

def save_data(season, registrations):
    """Saves season and registration data to JSON."""
    data = {
        "season": season,
        "registrations": registrations
    }
    with open(REG_FILE, 'w') as f:
        json.dump(data, f, indent=4)

# Initialize current season from the saved file
_initial_data = load_data()
CURRENT_SEASON = _initial_data.get("season", "SPRING2026")

def get_display_name(user_id, default_name):
    """Returns registered name or discord username."""
    data = load_data()
    regs = data.get("registrations", {})
    return regs.get(str(user_id), default_name)

# --- EXCEL LOGIC ---

def save_to_excel(sniper_name, sniper_id, number, snipee_name, snipee_id, proof_url):
    """Saves snipe data to the specific season tab in Excel."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not os.path.exists(EXCEL_FILE):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = CURRENT_SEASON
        sheet.append(["Sniper", "Points", "Snipee", "Timestamp", "Proof Link", "Sniper ID", "Snipee ID"])
    else:
        workbook = openpyxl.load_workbook(EXCEL_FILE, keep_vba=True)
        # Check if season tab exists, else create it
        if CURRENT_SEASON in workbook.sheetnames:
            sheet = workbook[CURRENT_SEASON]
        else:
            sheet = workbook.create_sheet(CURRENT_SEASON)
            sheet.append(["Sniper", "Points", "Snipee", "Timestamp", "Proof Link", "Sniper ID", "Snipee ID"])

    # Find next empty row in Column A (avoids overwriting charts/pivot tables elsewhere)
    next_row = 1
    while sheet.cell(row=next_row, column=1).value is not None:
        next_row += 1

    sheet.cell(row=next_row, column=1).value = sniper_name
    sheet.cell(row=next_row, column=2).value = number
    sheet.cell(row=next_row, column=3).value = snipee_name
    sheet.cell(row=next_row, column=4).value = timestamp
    sheet.cell(row=next_row, column=5).value = proof_url
    sheet.cell(row=next_row, column=6).value = str(sniper_id) 
    sheet.cell(row=next_row, column=7).value = str(snipee_id)
    
    workbook.save(EXCEL_FILE)

# --- BOT SETUP ---

intents = discord.Intents.default()
bot = commands.Bot(command_prefix="!", intents=intents)

@bot.event
async def on_ready():
    print(f'Logged in as {bot.user}!')
    try:
        synced = await bot.tree.sync()
        print(f"Synced {len(synced)} command(s).")
    except Exception as e:
        print(e)

# --- ADMIN COMMANDS ---

@bot.tree.command(name="change_season", description="Update the active Excel tab name (Owner Only)")
@app_commands.describe(new_season="The new season name (e.g., FALL2026)")
async def change_season(interaction: discord.Interaction, new_season: str):
    if interaction.user.id != OWNER_ID:
        await interaction.response.send_message("You don't have permission for this.", ephemeral=True)
        return

    global CURRENT_SEASON
    old_season = CURRENT_SEASON
    CURRENT_SEASON = new_season.upper()
    
    # Persist the change
    data = load_data()
    save_data(CURRENT_SEASON, data.get("registrations", {}))

    await interaction.response.send_message(
        f"✅ **Season Updated!**\nOld: `{old_season}`\nNew: `{CURRENT_SEASON}`\n"
        f"Data will now be logged in the `{CURRENT_SEASON}` tab.", 
        ephemeral=True
    )

@bot.tree.command(name="register", description="Assign a custom name to a Discord user (Owner Only)")
@app_commands.describe(user="The Discord user", name="The real name to use in Excel")
async def register(interaction: discord.Interaction, user: discord.User, name: str):
    if interaction.user.id != OWNER_ID:
        await interaction.response.send_message("You don't have permission for this.", ephemeral=True)
        return

    data = load_data()
    regs = data.get("registrations", {})
    regs[str(user.id)] = name
    save_data(CURRENT_SEASON, regs)
    
    await interaction.response.send_message(f"✅ Registered **{user.name}** as **{name}**.", ephemeral=True)

@bot.tree.command(name="deregister", description="Remove a custom name registration (Owner Only)")
@app_commands.describe(name="The registered name to remove")
async def deregister(interaction: discord.Interaction, name: str):
    if interaction.user.id != OWNER_ID:
        await interaction.response.send_message("You don't have permission for this.", ephemeral=True)
        return

    data = load_data()
    regs = data.get("registrations", {})
    
    user_id_to_remove = next((uid for uid, n in regs.items() if n == name), None)
    
    if user_id_to_remove:
        del regs[user_id_to_remove]
        save_data(CURRENT_SEASON, regs)
        await interaction.response.send_message(f"🗑️ Removed registration for **{name}**.", ephemeral=True)
    else:
        await interaction.response.send_message(f"❌ No registration found for **{name}**.", ephemeral=True)

@deregister.autocomplete('name')
async def deregister_autocomplete(interaction: discord.Interaction, current: str):
    data = load_data()
    names = list(data.get("registrations", {}).values())
    return [
        app_commands.Choice(name=n, value=n)
        for n in names if current.lower() in n.lower()
    ][:25]

# --- MAIN GAME COMMAND ---

@bot.tree.command(name="snipe", description="Add a Snipe to the Excel Sheet")
@app_commands.describe(number="Points value", user="Who did you snipe? (Leave blank for Alumni)", proof="Photo proof")
@app_commands.choices(number=[
    app_commands.Choice(name="1", value=1),
    app_commands.Choice(name="2", value=2),
    app_commands.Choice(name="Alumni Snipe", value=5)
])
async def snipe(interaction: discord.Interaction, number: int, proof: discord.Attachment, user: discord.User = None):
    # Defer immediately to prevent timeout errors
    await interaction.response.defer()
    
    sniper_display = get_display_name(interaction.user.id, interaction.user.name)
    sniper_id = interaction.user.id

    # Handle Alumni logic vs Standard Snipe
    if user is None:
        if number == 5:
            snipee_display = "Alumni"
            snipee_id = "0000"
            display_message = f"**{sniper_display} got an Alumni Snipe for 5 points!**"
        else:
            await interaction.followup.send("❌ You must select a user unless it is an Alumni Snipe (5 pts).", ephemeral=True)
            return
    else:
        snipee_display = get_display_name(user.id, user.name)
        snipee_id = user.id
        display_message = f"**<@{snipee_id}> ({snipee_display}) got shot by {sniper_display} for {number} points**"
    
    try:
        save_to_excel(sniper_display, sniper_id, number, snipee_display, snipee_id, proof.url)
        await interaction.followup.send(f"{display_message}\n{proof.url}")
    except PermissionError:
        await interaction.followup.send(f"⚠️ <@{OWNER_ID}> **CLOSE THE EXCEL SHEET**")
    except Exception as e:
        # THIS IS YOUR DEBUGGER: It will print the exact error to Discord
        await interaction.followup.send(f"❌ **TECHNICAL ERROR:** `{str(e)}`")
        print(f"Error details: {e}")

bot.run(TOKEN)