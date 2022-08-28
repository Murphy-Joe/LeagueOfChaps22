import random
import time
import xlwings as xw
wb = xw.Book('draft_slots.xlsx')
sheet = wb.sheets['Sheet1']

def reset():
    for i in range(2, 14):
        sheet.range('B' + str(i)).value = ""
        sheet.range('C' + str(i)).value = sheet.range('C' + str(i)).value.strip()

reset()

def set_correct_draft_slot(cell, team):
    sheet.range('B' + cell).value = team + " "
    time.sleep(1)
    sheet.range('B' + cell).value = team

def set_incorrect_draft_slot(cell, team):
    sheet.range('B' + cell).value = team + " "
    time.sleep(0.25)
    sheet.range('B' + cell).value = ""

def remove_highlight_selected_team(cell):
    sheet.range('C' + cell).value = sheet.range('C' + cell).value.strip()

def remove_highlight_prev_team(cell):
    sheet.range('B' + cell).value = sheet.range('B' + cell).value.strip()

def highlight_selecting_team(cell, team):
    sheet.range('C' + cell).value = team + " "

teams = {
    "Joe": 8,
    "TJ": 6,
    "Andrew": 2,
    "Joey": 1,
    "Matt": 3,
    "Garen": 4,
    "Daniel": 10,
    "Tony": 7,
    "Melinda": 5,
    "Nate": 9,
    "Kevin": 12,
    "Todd": 11
}

slots_left = sorted(list(teams.values()))

previous_slot_cell = 1

for i, (team, slot) in enumerate(teams.items()):
    slot_cell = str(slot + 1)
    previous_list_cell = str(i+1)
    current_list_cell = str(i+2)

    remove_highlight_selected_team(previous_list_cell)
    # remove_highlight_prev_team(str(previous_slot_cell))
    highlight_selecting_team(current_list_cell, team)

    # highlight through slots left
    for slot_left in reversed(slots_left):
        sheet.range('B' + str(slot_left+1)).value = " "
        time.sleep(0.1)
        sheet.range('B' + str(slot_left+1)).value = ""

    slots_left.remove(slot)
    set_correct_draft_slot(slot_cell, team)

    previous_slot_cell = slot+1
    
sheet.range('C13').value = sheet.range('C13').value.strip()

for x in range(6):
    for i in range(2,14):
        sheet.range('B' + str(i)).value = sheet.range('B' + str(i)).value + "!"
    time.sleep(0.1)
    for i in range(2,14): 
        sheet.range('B' + str(i)).value = sheet.range('B' + str(i)).value.replace("!", "")
