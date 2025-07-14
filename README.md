# Automation Bot for HPRO System
This project automates repetitive daily tasks using Python. It functions as a virtual assistant designed specifically for the HPRO system used in my workplace.

## Overview
The automation process reads data from Excel files using pandas and interacts with the system interface through pyautogui, simulating human actions such as mouse clicks and keyboard typing.

The workflow is simple:

I populate an Excel file with the required information to request items (e.g., product code, quantity, department, requester, etc.).

The bot automatically navigates through the HPRO system and submits each item request on my behalf.

## Why It Matters
Manually submitting a single item request in the HPRO system takes approximately 1 minute and 40 seconds.

With this automation:

A single item is requested in ~4 seconds

50 items can be requested in ~20 minutes

Compared to ~1 hour and 20 minutes manually

This saves over 1 hour per day on repetitive work, allowing me to focus on more important tasks.

## Technologies Used
Python
  - pyautogui — to automate mouse and keyboard input
  
  - pandas — to process Excel spreadsheets
  
  - openpyxl — for Excel file manipulation
  
  - subprocess — to launch the HPRO application
