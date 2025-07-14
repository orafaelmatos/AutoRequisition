import pyautogui as p
from time import sleep
import pandas as pd
import subprocess


# Hotkey function
def hotkeys(x, z=''):
    y = p.hotkey(x, z)
    sleep(0.5)
    return y

################################## OPENING HPRO AND LOGIN ##################################
def open_hpro():
    # Launch the HPRO application
    subprocess.Popen("C:\\ClientTeste\\PH2.EXE")
    sleep(15)  # Wait for HPRO to fully load

def login_user():
    # Enter username and password, then press Enter
    p.moveTo(866, 559)
    p.click()
    p.typewrite("matos")
    hotkeys('tab')
    p.typewrite("123456")
    hotkeys('enter')
    sleep(2)

################################## STOCK REQUEST ##################################
def request_stock_item():
    # Navigate to stock requisition screen
    p.moveTo(223, 33); p.click()  # Stock menu
    p.moveTo(295, 62); p.click()  # Stock requisition
    p.moveTo(566, 56); p.click()  # Maintenance

    # Read Excel data
    df = pd.read_excel('data/request_stock.xlsx')
    codes = df['CODIGO'].astype(str).tolist()
    quantities = df['QUANTIDADE'].astype(str).tolist()

    for code, qty in zip(codes, quantities):
        hotkeys('i')  # Include new item
        p.typewrite("Rafael Matos")
        hotkeys('tab')
        p.typewrite("3,0110")
        hotkeys('tab')
        p.typewrite(code, interval=0.2)

        p.moveTo(842, 473); p.doubleClick()
        p.typewrite(qty, interval=0.2)

        p.moveTo(912, 593); p.click()
        p.typewrite('ITEM ESTA NO KANBAN', interval=0.2)

        hotkeys('alt', 'g')
        sleep(0.5)

################################## PURCHASE REQUEST ##################################
def request_purchase_item():
    # Navigate to purchase requisition screen
    p.moveTo(303, 33); p.click()  # Supplies menu
    p.moveTo(353, 154); p.click()  # Requisition
    p.moveTo(608, 151); p.click()  # Typing screen

    df = pd.read_excel('data/request_purchase.xlsx')
    necessity_dates = df['DATA NECESSIDADE'].astype(str).tolist()
    codes = df['CODIGO'].astype(str).tolist()
    quantities = df['QUANTIDADE'].astype(str).tolist()
    projects = df['PROJETOS'].astype(str).tolist()
    observations = df['OBSERVACAO'].astype(str).tolist()

    for date, code, qty, project, note in zip(necessity_dates, codes, quantities, projects, observations):
        hotkeys('i')  # Include new item

        p.typewrite(date)  # Necessity Date
        hotkeys('tab')
        p.typewrite('Rafael Matos')  # Requester
        hotkeys('tab')
        p.typewrite('3,0110')  # Department
        hotkeys('tab')
        p.typewrite(code)  # Product code
        hotkeys('tab')
        p.typewrite(qty)  # Quantity

        hotkeys('alt', 'i')  # Include Project
        p.typewrite(project)
        hotkeys('alt', 'g')  # Save Project

        p.moveTo(1065, 721); p.click()  # Move to observation field
        p.typewrite(note)  # Write observation

        hotkeys('alt', 'g')  # Save requisition
