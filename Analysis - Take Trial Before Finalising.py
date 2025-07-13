import time
import json
from datetime import datetime
from datetime import time as tm
from types import SimpleNamespace
import pyotp
import sys

import requests
import pandas as pd
import numpy as np
import openpyxl
import xlwings as xw
from threading import Thread

from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtGui import QFont
from PyQt5.QtGui import QColor
from sklearn.metrics import r2_score

import pyqtgraph as pg
from pyqtgraph import TextItem, mkPen, QtCore
from pyqtgraph.exporters import ImageExporter

import os
import ctypes

import webbrowser

def enable_ansi_support():
    if os.name == 'nt':  # Check if the OS is Windows
        kernel32 = ctypes.windll.kernel32
        hStdOut = kernel32.GetStdHandle(-11)  # Get handle to standard output
        mode = ctypes.c_uint32()
        kernel32.GetConsoleMode(hStdOut, ctypes.byref(mode))
        mode.value |= 0x0004  # Enable virtual terminal processing
        kernel32.SetConsoleMode(hStdOut, mode)

enable_ansi_support()

tdate = datetime.now().date()
code = None
access = None

def time_fun():
    ttime = datetime.now().time().replace(microsecond=0)
    ttime = ttime.strftime("%H:%M:%S")
    return ttime

def show_totp(secret):
    totp = pyotp.TOTP(secret)
    while code is None:
        otp = totp.now()
        time_left = totp.interval - (int(time.time()) % totp.interval)

        sys.stdout.write('\033[s')           # Save cursor position
        sys.stdout.write('\033[3F')          # Move up 3 lines
        sys.stdout.write('\r\033[K\n')       # Clear 1st line (blank)
        sys.stdout.write('\r\033[K')         # Clear 2nd line (TOTP)
        sys.stdout.write(f"TOTP: {otp} | Expires in: {time_left:2d} sec\n")
        sys.stdout.write('\r\033[K\n')       # Clear 3rd line (blank)
        sys.stdout.write('\033[u')           # Restore cursor to input
        sys.stdout.flush()
        time.sleep(1)

if not os.path.exists('Credentials/login_details.json'):
    print("User Details not found. First Create a User Base & Retry. Exiting program.")
    sys.exit()

with open('Credentials/login_details.json', 'r') as file_read:
    users_data = json.load(file_read)

allowed_namess = users_data.keys()
allowed_names = [name.lower() for name in allowed_namess]

name_dict = {}

for i in range(len(allowed_names)):
    name_dict[f'{allowed_names[i]}'] = f'{tdate}_access_code_{allowed_names[i]}.json'

name_list = name_dict.values()

file_list = os.listdir(f'Credentials/Data')

for name in name_list:
    if name in file_list:
        with open(f'Credentials/Data/{name}', 'r') as file_read:
            access = json.load(file_read)
            acc_name = name[23:][:-5]

if not access:

    while True:
        acc_name = input(f'\nEnter Name of Account Holder to Login From {list(allowed_namess)} : ').lower()
        if acc_name in allowed_names:
            break
        else:
            print(f"\nInvalid User. Please Enter Registered User Name {list(allowed_namess)}'.")

    try:
        with open(f'Credentials/Data/{tdate}_access_code_{acc_name}.json', 'r') as file_read:
            access = json.load(file_read)

    except:

        with open('Credentials/login_details.json', 'r') as file_read:
            login_details = json.load(file_read)

        api_key = login_details[f'{acc_name.capitalize()}']['api_key']
        api_secret = login_details[f'{acc_name.capitalize()}']['api_secret']
        api_auth = login_details[f'{acc_name.capitalize()}']['api_auth']
        api_pin = login_details[f'{acc_name.capitalize()}']['pin']
        mobile_no = login_details[f'{acc_name.capitalize()}']['Mob No.']
        hold_name = login_details[f'{acc_name.capitalize()}']['full_name']

        print(f'\nTrying to Login from below details :')
        print(f'Account Holder: {hold_name}')
        print(f'Mobile No.: {mobile_no}')
        print(f'Pin: {api_pin}')
        print('Goto below URL and Enter the Required Details to Proceed')

        uri = 'https://www.google.com/'
        url1 = f'https://api.upstox.com/v2/login/authorization/dialog?response_type=code&client_id={api_key}&redirect_uri={uri}\n'
        print(f'\n{url1}\n\n')
        webbrowser.open(url1)

        thread = Thread(target=show_totp, args=(api_auth,), daemon=True)
        thread.start()

        code = input('Enter the Code : ')
        url = 'https://api.upstox.com/v2/login/authorization/token'
        headers = {
            'accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded',
        }

        data = {
            'code': code,
            'client_id': api_key,
            'client_secret': api_secret,
            'redirect_uri': uri,
            'grant_type': 'authorization_code',
        }

        response = requests.post(url, headers=headers, data=data)
        access = response.json()['access_token']
        print(f'\nLogin Successful, Status Code : {response.status_code}')
        print(f'User Name : {response.json()['user_name']}\nEmail ID : {response.json()['email']}')

        with open(f'Credentials/Data/{tdate}_access_code_{acc_name}.json', 'w') as file_write:
            json.dump(access, file_write)

print(f'\nLogin Successful from Account : {acc_name.capitalize()}')


#############################################################################

def instrument():
    inst_url = 'https://assets.upstox.com/market-quote/instruments/exchange/complete.csv.gz'
    instrument = pd.read_csv(inst_url)
    instrument.to_csv('Credentials/instrument.csv')

while True:
    yn = input('\nDo you Want to Update Instrument : 0 / 1 : ')
    if yn == '1' or yn == '0':
        break
    else:
        print("\nInvalid Selection. Please enter either '0' or '1'.")

if yn=='1':
    instrument()
    print("Instrument Data Updated Successfully")
try:
    df = pd.read_csv('Credentials/instrument.csv', index_col=0)
except:
    instrument()
    print("Can't find 'Instrument.csv' file, Latest Instrument Data Downloaded Successfully")
    df = pd.read_csv('Credentials/instrument.csv', index_col=0)

df_niftyoptions = df[(df['exchange'] == 'NSE_FO') & (df['instrument_type'] == 'OPTIDX') & (df['name'] == 'NIFTY')]
expiry_list_nifty = df_niftyoptions['expiry'].unique().tolist()
expiry_list_nifty.sort()

df_bnf = df[(df['exchange'] == 'NSE_FO') & (df['instrument_type'] == 'OPTIDX') & (df['name'] == 'BANKNIFTY')]
expiry_list_bnf = df_bnf['expiry'].unique().tolist()
expiry_list_bnf.sort()

df_sensex = df[(df['exchange'] == 'BSE_FO') & (df['instrument_type'] == 'OPTIDX') & (df['name'] == 'SENSEX')]
expiry_list_sensex = df_sensex['expiry'].unique().tolist()
expiry_list_sensex.sort()

app_analysis = xw.App(visible=True, add_book=False)
app_analysis.display_alerts = False
    
wb = app_analysis.books.open('Credentials/Analysis.xlsm')

summary = wb.sheets['summary']
nifty_0 = wb.sheets['nifty_0']
nifty_1 = wb.sheets['nifty_1']
nifty_3 = wb.sheets['nifty_3']
bnf_0 = wb.sheets['bnf_0']
sensex_0 = wb.sheets['sensex_0']

instrument_key_nifty = 'NSE_INDEX|Nifty 50'
instrument_key_bnf = 'NSE_INDEX|Nifty Bank'
instrument_key_sensex = 'BSE_INDEX|SENSEX'

structure_initial = {}
structure_current = {}
past_data={}
past_data_obv={}

a=b=c=d=e=f=1
initialize=1

t_date = datetime.now().date()

t_time = datetime.now().time().replace(microsecond=0)
start_time = tm(9,15,0,0)
end_time = tm(15,30,0,0)
print()
while t_time < start_time:
    t_time = datetime.now().time().replace(microsecond=0)
    print(f'\rCurrent Time : {t_time} | Market Will Start at {start_time}', end='', flush=True)
    time.sleep(1)

wb.macro("StartMonitoring")()

if t_time < end_time:
    print(f'\n\nProgram Started at {t_time}\n')

################################################################

# Function to convert time strings to timestamps
def format_time_ticks(values, scale, spacing):
    """Convert timestamp values to formatted time strings"""
    result = []
    for val in values:
        try:
            if val > 0:  # Make sure it's a positive timestamp
                dt = datetime.fromtimestamp(val)
                result.append(dt.strftime('%H:%M:%S'))
            else:
                result.append('')
        except (ValueError, OSError, OverflowError) as e:
            # print(f"Error with timestamp {val}: {e}")
            result.append('')
    return result

# Function to format timestamps as time strings for the axis
def time_string_to_timestamp(time_series):
    """
    Accept a pandas Series (or any iterable) of time strings and return a list of Unix timestamps.
    """
    today = datetime.now().date()
    def to_ts(time_str):
        try:
            time_obj = datetime.strptime(time_str, '%H:%M:%S').time()
            dt = datetime.combine(today, time_obj)
            ts = time.mktime(dt.timetuple())
            return ts
        except (ValueError, TypeError):
            return time.time()  # fallback to current time's timestamp

    ts_list = [to_ts(t) for t in time_series]

    return ts_list

# Start the application
app = QApplication(sys.argv)
# Initialize the main Qt application (app)
# This is required before creating any GUI windows or widgets
# sys.argv passes command-line arguments to the application (needed for some Qt features)

screens = app.screens() # Get the List of all connected screens/monitors

# ---------- Window 1 ----------
main1 = QMainWindow() # Create a new instance of Main Window (main1) — Creates a window
main1.setWindowTitle("Window 1")
win = pg.GraphicsLayoutWidget() # Create a new GraphicsLayoutWidget (win) for displaying multiple plots/subplots | # Creates a widget (not a window) — Most widgets are not windows; they are components (buttons, plots, etc.) that you put inside a window.
main1.setCentralWidget(win) # Set the GraphicsLayoutWidget (win) inside the Main Window (main1) as the central widget — puts the widget inside the window
# main1.resize(900, 1000)
main1.showMaximized()

# ---------- Window 2 ----------
main2 = QMainWindow() # Create a new instance of Main Window (main2) — Creates another window
main2.setWindowTitle("Window 2")
win2 = pg.GraphicsLayoutWidget() # Create another GraphicsLayoutWidget (win2) for displaying multiple plots/subplots | # Creates a widget (not a window) — Most widgets are not windows; they are components (buttons, plots, etc.) that you put inside a window.
main2.setCentralWidget(win2) # Set the GraphicsLayoutWidget (win2) inside the Main Window (main2) as the central widget — puts the widget inside the window
# main2.resize(900, 1000)
main2.showMaximized()

if len(screens) >= 2:
    if len(screens) == 2:
        screen1_geo = screens[0].geometry()
        screen2_geo = screens[1].geometry()
    if len(screens) == 3:
        screen1_geo = screens[0].geometry()
        screen2_geo = screens[1].geometry()
        screen3_geo = screens[2].geometry()

screen_geo = screens[0].geometry()
width = screen_geo.width()
height = screen_geo.height()

# Example: place side by side
main1.move(0, 0)
main2.move(width // 2, 0)

def screen():
    global screens, main1, main2
    
    disp2 = summary.range('B40').value
    disp1 = summary.range('C40').value
    disp3 = summary.range('D40').value

    full2 = summary.range('B41').value
    full1 = summary.range('C41').value
    full3 = summary.range('D41').value

    if disp1 == 1:
        main1.move(screen1_geo.x(), screen1_geo.y())
        main1.showMaximized()
        if full1 == 'F':
            main1.showFullScreen()
    if disp1 == 2:
        main2.move(screen1_geo.x(), screen1_geo.y())
        main2.showMaximized()
        if full1 == 'F':
            main2.showFullScreen()

    if disp2 == 1:
        main1.move(screen2_geo.x(), screen2_geo.y())
        main1.showMaximized()
        if full2 == 'F':
            main1.showFullScreen()
    if disp2 == 2:
        main2.move(screen2_geo.x(), screen2_geo.y())
        main2.showMaximized()
        if full2 == 'F':
            main2.showFullScreen()

    if disp3 == 1:
        main1.move(screen3_geo.x(), screen3_geo.y())
        main1.showMaximized()
        if full3 == 'F':
            main1.showFullScreen()
    if disp3 == 2:
        main2.move(screen3_geo.x(), screen3_geo.y())
        main2.showMaximized()
        if full3 == 'F':
            main2.showFullScreen()


# else:
#     # Single display: place side by side or stacked
#     screen_geo = screens[0].geometry()
#     width = screen_geo.width()
#     height = screen_geo.height()
    
#     # Example: place side by side
#     main1.move(0, 0)
#     main2.move(width // 2, 0)

main1.show()
main2.show()

# Initialize data for each expiry's time axis
x1, x2, x3 = [], [], []  # Time arrays for each expiry
y0_1, y1_1, y6_1, y7_1, y8_1, y9_1, y10_1, y11_1, y12_1, y13_1, y14_1, y15_1, y16_1, y17_1, y18_1, y19_1 = [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], []
y0_2, y1_2, y6_2, y7_2, y8_2, y9_2, y10_2, y11_2, y12_2, y13_2, y15_2, y16_2, y17_2, y18_2, y19_2 = [], [], [], [], [], [], [], [], [], [], [], [], [], [], []
y0_3, y1_3, y6_3, y7_3, y8_3, y9_3, y10_3, y11_3, y12_3, y13_3, y15_3, y16_3, y17_3, y18_3, y19_3 = [], [], [], [], [], [], [], [], [], [], [], [], [], [], []
y0_4, y1_4, y6_4, y7_4, y8_4, y9_4, y10_4, y11_4, y12_4, y13_4, y15_4, y16_4, y17_4, y18_4, y19_4 = [], [], [], [], [], [], [], [], [], [], [], [], [], [], []

##########################################################

# Add this after creating the win object
def keyPressEvent(event):
    if event.key() == Qt.Key_F11:  # F11 is commonly used for toggling fullscreen
        if win.isFullScreen():
            win.showMaximized()
        else:
            win.showFullScreen()
    elif event.key() == Qt.Key_Escape and win.isFullScreen():  # ESC to exit fullscreen
        win.showMaximized()

fullscreen_active1 = False  # Global variable to track state
fullscreen_active2 = False  # Global variable to track state
# fs=None
def check_excel_for_full_screen():
    global fullscreen_active1, fullscreen_active2 
    fsb = summary.range('B19').value
    fsc = summary.range('C19').value

    if fsb == 'F' and not fullscreen_active1:
        if not main1.isFullScreen():     # Only switch if not already fullscreen
            main1.showFullScreen()
        fullscreen_active1 = True
    elif fsb != 'F' and fullscreen_active1:
        if main1.isFullScreen():         # Only switch if fullscreen now
            main1.showMaximized()
        fullscreen_active1 = False

    if fsc == 'F' and not fullscreen_active2:
        if not main2.isFullScreen():     # Only switch if not already fullscreen
            main2.showFullScreen()
        fullscreen_active2 = True
    elif fsb != 'F' and fullscreen_active2:
        if main2.isFullScreen():         # Only switch if fullscreen now
            main2.showMaximized()
        fullscreen_active2 = False

##########################################################
dte_decay = {30:'2-4 %', 29:'3-5 %', 28:'3-5 %', 27:'3-5 %',26:'3-5 %',25:'3-5 %',24:'4-6 %',23:'4-6 %',21:'4-6 %',20:'4-6 %',19:'6-10 %',18:'6-10 %',17:'6-10 %',16:'6-10 %',15:'6-10 %',14:'8-12 %',13:'8-12 %',12:'8-12 %',11:'8-12 %',10:'8-12 %',9:'10-16 %',8:'10-16 %',7:'10-16 %',6:'12-20 %',5:'12-20 %',4:'15-25 %',3:'20-30 %',2:'25-40 %',1:'35-55 %',0:'100 %'}
lin_data_val = None
quad_data_val = None
def update_regression(y, straddle_curve, linear_curve, quad_curve, linear_eqn=None, quad_eqn=None):
    global lin_data_val, quad_data_val
    y_np = np.array(y[-200:])
    x_np = np.arange(len(y_np))  # auto-generated x

    straddle_curve.setData(x_np, y_np)

    lin_coeffs = None
    quad_coeffs = None

    # ----- Linear Regression -----
    if len(x_np) >= lin_data_val:
        x_lin = x_np[-lin_data_val:] - np.mean(x_np[-lin_data_val:])
        y_lin = y_np[-lin_data_val:]
        lin_coeffs = np.polyfit(x_lin, y_lin, 1)
        y_fit = np.polyval(lin_coeffs, x_lin)
        linear_curve.setData(x_np[-lin_data_val:], y_fit)

        # R² for Linear
        r2_lin = r2_score(y_lin, y_fit)
    else:
        linear_curve.clear()
        r2_lin = None

    # ----- Quadratic Regression -----
    if len(x_np) >= quad_data_val:
        x_quad = x_np[-quad_data_val:] - np.mean(x_np[-quad_data_val:])
        y_quad = y_np[-quad_data_val:]
        quad_coeffs = np.polyfit(x_quad, y_quad, 2)
        y_fit = np.polyval(quad_coeffs, x_quad)
        quad_curve.setData(x_np[-quad_data_val:], y_fit)

        # R² for Quadratic
        r2_quad = r2_score(y_quad, y_fit)
    else:
        quad_curve.clear()
        r2_quad = None

    # ----- Display Equations at Top Center -----
    if linear_eqn and lin_coeffs is not None:
        m, c = lin_coeffs
        text = f"y = {m:.2f}x + {c:.2f}"
        if r2_lin is not None:
            text += f"   R² = {r2_lin:.3f}"
        linear_eqn.setText(text)
        linear_eqn.setPos(x_np.mean(), min(y_np))

    if quad_eqn and quad_coeffs is not None:
        a, b, c2 = quad_coeffs
        text = f"y = {a:.3f}x² + {b:.2f}x + {c2:.2f}"
        if r2_quad is not None:
            text += f"   R² = {r2_quad:.3f}"
        quad_eqn.setText(text)
        quad_eqn.setPos(x_np.mean(), max(y_np) - ((max(y_np) - min(y_np)) * 0.05))



def one_time(expiry_names):
    global dte_decay
    label_dict = {}
    today = expiry_names[0][8:18]

    ##################################### 2nd Window - Column 1 #######################################

    #Plot00 - 2nd Window
    #####################
    win2_plot00 = win2.addPlot(row=0, col=0, title='Nifty 50 "NEXT" : CE/PE OTMs')
    win2_plot00.addLegend()
    line_ce = win2_plot00.plot([], pen=mkPen('g', width=3), name='CE OTMs')
    line_pe = win2_plot00.plot([], pen=mkPen('y', width=3), name='PE OTMs')
    win2_plot00.addLine(y=0, pen=pg.mkPen(color='w', width=1.5))

    # Set the x-axis formatting to use time strings
    win2_plot00.getAxis('bottom').setTickStrings = lambda values, scale, spacing: format_time_ticks(values) # Set custom x-axis time labels

    # 1) Create a Viewbox
    # 2) Create Line/Plot to be plotted in view box
    # 3) Add/Connect that Plot/Line to the View box
    # 4) Link the main plot's Right axis to the View Box as the reference axis (This is the 1st point where we are connecting the View Box to the Main Plot via main plot's Right Axis)
    # 5) Then we Sync X-Axis but De-Sync the y axis
    # 6) Then i ask to show (make visible) the Right Axis (without this it wont make right axis visible), but this command make the axis and data range both visible.
    # 7) Then i command to not to show the data but only the Y - axis
    # 8) Then i ask to add the legend
    # 9) Then finally give command to show the view box graph on main plot's scene with the above attributes. Main Plot's Scene is where the actual visualization of view box will be done Finally
    # I think this is the proper line of action for remebering what we are doing inside the code and whats happening behind the scene.
    
    # ViewBox : Plot00 - 2nd Window
    ################################

    win2_plot00_viewbox = pg.ViewBox() # Creates a new ViewBox

    plot00_viewbox_ce_oi = pg.PlotDataItem(pen=mkPen(color='g', width=3, style=QtCore.Qt.DotLine), name='CE_OI') # Create PlotDataItems to plot
    win2_plot00_viewbox.addItem(plot00_viewbox_ce_oi) #Add PlotDataItems to the ViewBox

    #Same Here - Repeat the Process
    plot00_viewbox_pe_oi = pg.PlotDataItem(pen=mkPen(color='y', width=3, style=QtCore.Qt.DotLine), name='PE_OI')
    win2_plot00_viewbox.addItem(plot00_viewbox_pe_oi)

    win2_plot00.getAxis('right').linkToView(win2_plot00_viewbox) #Link the Main Plot's Right Axis to the ViewBox

    # Synchronize the X-axis, but not the Y-axis
    win2_plot00_viewbox.setXLink(win2_plot00)
    win2_plot00_viewbox.setYLink(None)

    win2_plot00.showAxis('right') #Make the Right Axis Visible
    win2_plot00.getAxis('right').setStyle(showValues=False) # Hide Right Axis Values (if not needed)

    # Add Legend to Main Plot
    win2_plot00.addLegend()
    win2_plot00.legend.addItem(plot00_viewbox_ce_oi, 'CE_OI')
    win2_plot00.legend.addItem(plot00_viewbox_pe_oi, 'PE_OI')

    # Finally, Add ViewBox to Main Plot's Scene
    win2_plot00.scene().addItem(win2_plot00_viewbox)

    #These two functions ensure that the custom right-side ViewBox remains geometrically aligned with the main plot and dynamically adjusts its Y-axis range based on the data being displayed.
    # For win2_plot00
    def updateViews_win2_00():
        """Ensure the right ViewBox is geometrically aligned with the left ViewBox."""
        win2_plot00_viewbox.setGeometry(win2_plot00.vb.sceneBoundingRect())

    def adjust_right_view_win2_00():
        """Adjust the right Y-axis range of win2_plot00 based on y8_4 and y9_4."""
        combined = []
        if len(y8_4) > 0:
            combined.extend(y8_4)
        if len(y9_4) > 0:
            combined.extend(y9_4)

        if combined:
            min_val = min(combined) * 0.99  # 1% padding below
            max_val = max(combined) * 1.01  # 1% padding above
            win2_plot00_viewbox.setYRange(min_val, max_val)

    # Connect resize event to keep ViewBoxes aligned
    win2_plot00.vb.sigResized.connect(updateViews_win2_00)



    # Plot10 - 2nd Window
    ######################
    
    win2_plot10 = win2.addPlot(row=1, col=0, title='plot 2')

    # Plot20 - 2nd Window
    ######################
    win2_plot20 = win2.addPlot(row=2, col=0, title='Nifty "NEAR" Expiry : Straddle with Regression')
    win2_plot20.addLegend(offset=(10, 10))
    straddle_1 = win2_plot20.plot([], pen=mkPen('w', width=3), name='ATM Straddle')
    linear_1 = win2_plot20.plot([], pen=mkPen('r', width=3), name='Linear Regression')
    quad_1 = win2_plot20.plot([], pen=mkPen('y', width=3), name='Quadratic Regression')

    # Defining Text Object
    linear_eqn_1 = pg.TextItem('', color='r', anchor=(0.5, 1))  # top-center anchor
    linear_eqn_1.setFont(QFont("Arial", 10))
    win2_plot20.addItem(linear_eqn_1)

    quad_eqn_1 = pg.TextItem('', color='y', anchor=(0.5, 1))    # top-center anchor
    quad_eqn_1.setFont(QFont("Arial", 10))
    win2_plot20.addItem(quad_eqn_1)    

    ##################################### 2nd Window - Column 2 #######################################

    expiry_nifty = expiry_names[1][-10:]
    date1 = datetime.strptime(today, "%Y-%m-%d")
    date2 = datetime.strptime(expiry_nifty, "%Y-%m-%d")
    dte = (date2 - date1).days
    decay_text = dte_decay.get(dte, "N/A")

    # Plot01 - 2nd Window
    ######################
    win2_plot01 = win2.addPlot(row=0, col=1, title=f"T: {expiry_names[1][8:18]} | E: {expiry_names[1][-10:]} | {dte} DTE ({decay_text})")
    win2_plot01.addLegend()
    line_ce_atm = win2_plot01.plot([], pen=mkPen('g', width=3), name='CE ATM')
    line_pe_atm = win2_plot01.plot([], pen=mkPen('y', width=3), name='PE ATM')
    line_straddle = win2_plot01.plot([], pen=mkPen('m', width=3), name='Straddle')
    win2_plot01.addLine(y=0, pen=pg.mkPen(color='w', width=1.5))

    # Set the x-axis formatting to use time strings
    win2_plot01.getAxis('bottom').setTickStrings = lambda values, scale, spacing: format_time_ticks(values) # Set custom x-axis time labels

    # ViewBox : Plot01 - 2nd Window
    ################################

    win2_plot01_viewbox = pg.ViewBox()

    plot01_viewbox_abs_straddle = pg.PlotDataItem(pen=mkPen('w', width=3), name='Abs Straddle')
    plot01_viewbox_vwap = pg.PlotDataItem(pen=mkPen('r', width=3), name='VWAP')
    win2_plot01_viewbox.addItem(plot01_viewbox_abs_straddle)
    win2_plot01_viewbox.addItem(plot01_viewbox_vwap)

    win2_plot01.getAxis('right').linkToView(win2_plot01_viewbox)

    win2_plot01_viewbox.setXLink(win2_plot01)
    win2_plot01_viewbox.setYLink(None)

    win2_plot01.showAxis('right')
    win2_plot01.getAxis('right').setStyle(showValues=False)

    win2_plot01.addLegend()
    win2_plot01.legend.addItem(plot01_viewbox_abs_straddle, 'Abs Straddle')
    win2_plot01.legend.addItem(plot01_viewbox_vwap, 'VWAP')

    win2_plot01.scene().addItem(win2_plot01_viewbox)

    # For win2_plot01
    def updateViews_win2_01():
        """Ensure the right ViewBox is geometrically aligned with the left ViewBox."""
        win2_plot01_viewbox.setGeometry(win2_plot01.vb.sceneBoundingRect())

    def adjust_right_view_win2_01():
        """Adjust the right Y-axis range of win2_plot01 based on y16_4 and y17_4."""
        combined = []
        if len(y16_4) > 0:
            combined.extend(y16_4)
        if len(y17_4) > 0:
            combined.extend(y17_4)

        if combined:
            min_val = min(combined) * 0.99  # 1% padding below
            max_val = max(combined) * 1.01  # 1% padding above
            win2_plot01_viewbox.setYRange(min_val, max_val)

    # Connect resize event to keep ViewBoxes aligned
    win2_plot01.vb.sigResized.connect(updateViews_win2_01)

    # Plot11 - 2nd Window
    ######################

    win2_plot11 = win2.addPlot(row=1, col=1, title='Nifty 50 "NEXT" Expiry : Straddle with Regression')
    win2_plot11.addLegend(offset=(10, 10))
    straddle_4 = win2_plot11.plot([], pen=mkPen('w', width=3), name='ATM Straddle')
    linear_4 = win2_plot11.plot([], pen=mkPen('r', width=3), name='Linear Regression')
    quad_4 = win2_plot11.plot([], pen=mkPen('y', width=3), name='Quadratic Regression')

    # Defining Text Object
    linear_eqn_4 = pg.TextItem('', color='r', anchor=(0.5, 1))  # top-center anchor
    linear_eqn_4.setFont(QFont("Arial", 10))
    win2_plot11.addItem(linear_eqn_4)

    quad_eqn_4 = pg.TextItem('', color='y', anchor=(0.5, 1))    # top-center anchor
    quad_eqn_4.setFont(QFont("Arial", 10))
    win2_plot11.addItem(quad_eqn_4)

    # Plot21 - 2nd Window
    ######################

    win2_plot21 = win2.addPlot(row=2, col=1, title='Bank Nifty Near Expiry : Straddle with Regression')
    win2_plot21.addLegend(offset=(10, 10))
    straddle_2 = win2_plot21.plot([], pen=mkPen('w', width=3), name='ATM Straddle')
    linear_2 = win2_plot21.plot([], pen=mkPen('r', width=3), name='Linear Regression')
    quad_2 = win2_plot21.plot([], pen=mkPen('y', width=3), name='Quadratic Regression')

    # Defining Text Object
    linear_eqn_2 = pg.TextItem('', color='r', anchor=(0.5, 1))  # top-center anchor
    linear_eqn_2.setFont(QFont("Arial", 10))
    win2_plot21.addItem(linear_eqn_2)

    quad_eqn_2 = pg.TextItem('', color='y', anchor=(0.5, 1))    # top-center anchor
    quad_eqn_2.setFont(QFont("Arial", 10))
    win2_plot21.addItem(quad_eqn_2)

    ##################################### 2nd Window - Column 3 #######################################

    # Plot 02 - 2nd Window
    #######################
    win2_plot02 = win2.addPlot(row=0, col=2, title='Nifty 50 : CE/PE OTMs Implied Volatility')
    win2_plot02.addLegend()
    line_ce_iv = win2_plot02.plot([], pen=mkPen('g', width=3), name='CE IV')
    line_pe_iv = win2_plot02.plot([], pen=mkPen('y', width=3), name='PE IV')
    win2_plot02.addLine(y=0, pen=pg.mkPen(color='w', width=1.5))

    # Set the x-axis formatting to use time strings
    win2_plot02.getAxis('bottom').setTickStrings = lambda values, scale, spacing: format_time_ticks(values) # Set custom x-axis time labels

    # ViewBox : Plot02 - 2nd Window
    ################################

    win2_plot02_viewbox = pg.ViewBox()  # Creates a new ViewBox for right-side plotting

    plot02_viewbox_1 = pg.PlotDataItem(pen=mkPen(color='g', width=3, style=QtCore.Qt.DotLine), name='CE OBV')
    plot02_viewbox_2 = pg.PlotDataItem(pen=mkPen(color='y', width=3, style=QtCore.Qt.DotLine), name='PE OBV')
    win2_plot02_viewbox.addItem(plot02_viewbox_1)
    win2_plot02_viewbox.addItem(plot02_viewbox_2)

    win2_plot02.getAxis('right').linkToView(win2_plot02_viewbox)

    win2_plot02_viewbox.setXLink(win2_plot02)
    win2_plot02_viewbox.setYLink(None)

    win2_plot02.showAxis('right')
    win2_plot02.getAxis('right').setStyle(showValues=False)

    win2_plot02.addLegend()
    win2_plot02.legend.addItem(plot02_viewbox_1, 'CE OBV')
    win2_plot02.legend.addItem(plot02_viewbox_2, 'PE OBV')

    win2_plot02.scene().addItem(win2_plot02_viewbox)

    # Floating Labels (for CE/PE OBV)
    font_ce = QFont('Arial', 12)
    font_ce.setBold(True)
    label_plot02_ce = TextItem(anchor=(1, 0.5))
    label_plot02_ce.setFont(font_ce)
    label_plot02_ce.setColor(QColor(255, 0, 0))  # Bright red
    win2_plot02_viewbox.addItem(label_plot02_ce)
    label_dict['label_plot02_ce'] = label_plot02_ce

    font_pe = QFont('Arial', 12)
    font_pe.setBold(True)
    label_plot02_pe = TextItem(anchor=(1, 0.5))
    label_plot02_pe.setFont(font_pe)
    label_plot02_pe.setColor(QColor(255, 255, 255))  # Bright white
    win2_plot02_viewbox.addItem(label_plot02_pe)
    label_dict['label_plot02_pe'] = label_plot02_pe

    # For win2_plot02
    def updateViews_win2_02():
        """Ensure the right ViewBox is geometrically aligned with the left ViewBox."""
        win2_plot02_viewbox.setGeometry(win2_plot02.vb.sceneBoundingRect())

    def adjust_right_view_win2_02():
        """Adjust the right Y-axis range of win2_plot02 based on y18_4 and y19_4."""
        combined = []
        if len(y18_4) > 0:
            combined.extend(y18_4)
        if len(y19_4) > 0:
            combined.extend(y19_4)

        if combined:
            min_val = min(combined) * 0.99  # 1% padding below
            max_val = max(combined) * 1.01  # 1% padding above
            win2_plot02_viewbox.setYRange(min_val, max_val)

    # Connect resize event to keep ViewBoxes aligned
    win2_plot02.vb.sigResized.connect(updateViews_win2_02)

    # Plot 12 - 2nd Window
    #######################
    win2_plot12 = win2.addPlot(row=1, col=2, title='plot 2')
    
    # Plot 22 - 2nd Window
    #######################
    win2_plot22 = win2.addPlot(row=2, col=2, title='Sensex "NEAR" Expiry : Straddle with Regression')
    win2_plot22.addLegend(offset=(10, 10))
    straddle_3 = win2_plot22.plot([], pen=mkPen('w', width=3), name='ATM Straddle')
    linear_3 = win2_plot22.plot([], pen=mkPen('r', width=3), name='Linear Regression')
    quad_3 = win2_plot22.plot([], pen=mkPen('y', width=3), name='Quadratic Regression')

    # Defining Text Object
    linear_eqn_3 = pg.TextItem('', color='r', anchor=(0.5, 1))  # top-center anchor
    linear_eqn_3.setFont(QFont("Arial", 10))
    win2_plot22.addItem(linear_eqn_3)

    quad_eqn_3 = pg.TextItem('', color='y', anchor=(0.5, 1))    # top-center anchor
    quad_eqn_3.setFont(QFont("Arial", 10))
    win2_plot22.addItem(quad_eqn_3)

    ##########################################################################################################################

    #*************************Column 1*************************************#
    expiry_nifty = expiry_names[0][-10:]
    date1 = datetime.strptime(today, "%Y-%m-%d")
    date2 = datetime.strptime(expiry_nifty, "%Y-%m-%d")
    dte = (date2 - date1).days
    decay_text = dte_decay.get(dte, "N/A")

    #Plot 00
    ######################################################################
    plot00 = win.addPlot(row=0, col=0, title=f"T: {expiry_names[0][8:18]} | Nifty - OTMs | E: {expiry_names[0][-10:]} | {dte} DTE ({decay_text})") # Add a plot at (0,0) with dynamic title
    plot00.addLegend() # Add a legend to the plot
    line_0_1 = plot00.plot([], pen=mkPen('g', width=3), name='CE OTMs') # Add green line for CE OTMs (empty data)
    line_1_1 = plot00.plot([], pen=mkPen('y', width=3), name='PE OTMs') # Add yellow line for PE OTMs (empty data)

    plot00.addLine(y=0, pen=pg.mkPen(color='w', width=1.5)) # Add horizontal line at y=0 (white, 1.5px)

    # plot00.showGrid(x=True, y=True, alpha=0.3)

    # Set the x-axis formatting to use time strings
    plot00.getAxis('bottom').setTickStrings = lambda values, scale, spacing: format_time_ticks(values) # Set custom x-axis time labels

    # Create a Text Object inside the Plot - latter can be updated with time in loop
    time_label = pg.TextItem("Default_Text", color='g', anchor=(0.5, 0.5))
    time_label.setFont(QFont('Arial', 11))
    time_label.setPos(0.25, 30) 
    plot00.addItem(time_label)

    # Plot 00 : Floating Labels
    label_0 = pg.TextItem(anchor=(1, 0.5))
    label_0.setFont(QFont('Arial', 12))
    plot00.addItem(label_0)
    label_dict['label_0'] = label_0

    label_1 = pg.TextItem(anchor=(1, 0.5))
    label_1.setFont(QFont('Arial', 12))
    plot00.addItem(label_1)
    label_dict['label_1'] = label_1

    ######################################################################

    #Plot 10
    #######################################################################
    plot10 = win.addPlot(row=1, col=0, title="Nifty - CE-PE ATM, Straddle")
    plot10.addLegend()
    line_10_1 = plot10.plot([], pen=mkPen('g', width=3), name='CE ATM')
    line_11_1 = plot10.plot([], pen=mkPen('y', width=3), name='PE ATM')
    line_12_1 = plot10.plot([], pen=mkPen('m', width=3), name='ATM Straddle')

    plot10.addLine(y=0, pen=pg.mkPen(color='w', width=1.5))

    # plot10.showGrid(x=True, y=True, alpha=0.3)
    plot10.getAxis('bottom').setTickStrings = lambda values, scale, spacing: format_time_ticks(values)

    # Create a Text Object inside the Plot - latter can be updated with time in loop
    straddle_label_0 = pg.TextItem("Default_Text", color='lightgreen', anchor=(0.5, 0.5))
    straddle_label_0.setFont(QFont('Arial', 11))
    straddle_label_0.setPos(0.25, 30)
    plot10.addItem(straddle_label_0)

    # For plot10 : Floating Labels
    label_10 = pg.TextItem(anchor=(1, 0.5))
    label_10.setFont(QFont('Arial', 12))
    plot10.addItem(label_10)
    label_dict['label_10'] = label_10

    label_11 = pg.TextItem(anchor=(1, 0.5))
    label_11.setFont(QFont('Arial', 12))
    plot10.addItem(label_11)
    label_dict['label_11'] = label_11

    label_12 = pg.TextItem(anchor=(1, 0.5))
    label_12.setFont(QFont('Arial', 12))
    plot10.addItem(label_12)
    label_dict['label_12'] = label_12

    #######################################################################

    #Plot 20
    #######################################################################
    plot20 = win.addPlot(row=2, col=0, title="Nifty - CE/PE OTMs Implied Volatility")
    plot20.addLegend()
    line_6_1 = plot20.plot([], pen=mkPen('g', width=3), name='CE IV')
    line_7_1 = plot20.plot([], pen=mkPen('y', width=3), name='PE IV')
    line_13_1 = plot20.plot([], pen=mkPen('r', width=3), name='SPOT')

    plot20.addLine(y=0, pen=pg.mkPen(color='w', width=1.5))

    # plot20.showGrid(x=True, y=True, alpha=0.3)

    plot20.getAxis('bottom').setTickStrings = lambda values, scale, spacing: format_time_ticks(values)

    # Create a Text Object inside the Plot - latter can be updated with time in loop
    spot_label_0 = pg.TextItem("Default_Text", color='r', anchor=(0.5, 0.5))
    spot_label_0.setFont(QFont('Arial', 11))
    spot_label_0.setPos(0.25, 30)
    plot20.addItem(spot_label_0)

    # For plot20 : Floating Labels
    label_6 = pg.TextItem(anchor=(1, 0.5))
    label_6.setFont(QFont('Arial', 12))
    plot20.addItem(label_6)
    label_dict['label_6'] = label_6

    label_7 = pg.TextItem(anchor=(1, 0.5))
    label_7.setFont(QFont('Arial', 12))
    plot20.addItem(label_7)
    label_dict['label_7'] = label_7

    label_13 = pg.TextItem(anchor=(1, 0.5))
    label_13.setFont(QFont('Arial', 12))
    plot20.addItem(label_13)
    label_dict['label_13'] = label_13
    ##################################################################################


    #*************************Column 2*************************************#
    expiry_bnf = expiry_names[3][-10:]
    date1 = datetime.strptime(today, "%Y-%m-%d")
    date2 = datetime.strptime(expiry_bnf, "%Y-%m-%d")
    dte = (date2 - date1).days
    decay_text = dte_decay.get(dte, "N/A")

    #Plot 01
    ###################################################################################
    plot01 = win.addPlot(row=0, col=1, title=f"T: {expiry_names[0][8:18]} | Bank Nifty - OTMs | E: {expiry_names[3][-10:]} | {dte} DTE ({decay_text})")
    plot01.addLegend()
    line_0_2 = plot01.plot([], pen=mkPen('g', width=3), name='CE OTMs')
    line_1_2 = plot01.plot([], pen=mkPen('y', width=3), name='PE OTMs')

    plot01.addLine(y=0, pen=pg.mkPen(color='w', width=1.5))

    # plot01.showGrid(x=True, y=True, alpha=0.3)

    plot01.getAxis('bottom').setTickStrings = lambda values, scale, spacing: format_time_ticks(values)

    # For plot01 : Floating Labels
    label_0_2 = pg.TextItem(anchor=(1, 0.5))
    label_0_2.setFont(QFont('Arial', 12))
    plot01.addItem(label_0_2)
    label_dict['label_0_2'] = label_0_2

    label_1_2 = pg.TextItem(anchor=(1, 0.5))
    label_1_2.setFont(QFont('Arial', 12))
    plot01.addItem(label_1_2)
    label_dict['label_1_2'] = label_1_2
    ###################################################################################

    #Plot 11
    ###################################################################################
    plot11 = win.addPlot(row=1, col=1, title="Bank Nifty - CE-PE ATM, Straddle")
    plot11.addLegend()
    line_10_2 = plot11.plot([], pen=mkPen('g', width=3), name='CE ATM')
    line_11_2 = plot11.plot([], pen=mkPen('y', width=3), name='PE ATM')
    line_12_2 = plot11.plot([], pen=mkPen('m', width=3), name='ATM Straddle')

    plot11.addLine(y=0, pen=pg.mkPen(color='w', width=1.5))

    # plot11.showGrid(x=True, y=True, alpha=0.3)

    plot11.getAxis('bottom').setTickStrings = lambda values, scale, spacing: format_time_ticks(values)

    # Create a Text Object inside the Plot - latter can be updated with time in loop
    straddle_label_1 = pg.TextItem('Default_Text', color='lightgreen', anchor=(0.5, 0.5))
    straddle_label_1.setFont(QFont('Arial', 11))
    straddle_label_1.setPos(0.25, 30)
    plot11.addItem(straddle_label_1)

    # For plot11 : Floating Labels
    label_10_2 = pg.TextItem(anchor=(1, 0.5))
    label_10_2.setFont(QFont('Arial', 12))
    plot11.addItem(label_10_2)
    label_dict['label_10_2'] = label_10_2

    label_11_2 = pg.TextItem(anchor=(1, 0.5))
    label_11_2.setFont(QFont('Arial', 12))
    plot11.addItem(label_11_2)
    label_dict['label_11_2'] = label_11_2

    label_12_2 = pg.TextItem(anchor=(1, 0.5))
    label_12_2.setFont(QFont('Arial', 12))
    plot11.addItem(label_12_2)
    label_dict['label_12_2'] = label_12_2
    ####################################################################################

    #Plot 21
    ####################################################################################
    plot21 = win.addPlot(row=2, col=1, title="Bank Nifty - CE/PE OTMs Implied Volatility")
    plot21.addLegend()
    line_6_2 = plot21.plot([], pen=mkPen('g', width=3), name='CE IV')
    line_7_2 = plot21.plot([], pen=mkPen('y', width=3), name='PE IV')
    line_13_2 = plot21.plot([], pen=mkPen('r', width=3), name='SPOT')

    plot21.addLine(y=0, pen=pg.mkPen(color='w', width=1.5))

    # plot21.showGrid(x=True, y=True, alpha=0.3)

    plot21.getAxis('bottom').setTickStrings = lambda values, scale, spacing: format_time_ticks(values)

    # Create a Text Object inside the Plot - latter can be updated with time in loop
    spot_label_1 = pg.TextItem('Default_Text', color='r', anchor=(0.5, 0.5))
    spot_label_1.setFont(QFont('Arial', 11))  # Optional: Set font and size
    spot_label_1.setPos(0.25, 30)  # Initial position
    plot21.addItem(spot_label_1)

    # For plot21 : Floating Labels
    label_6_2 = pg.TextItem(anchor=(1, 0.5))
    label_6_2.setFont(QFont('Arial', 12))
    plot21.addItem(label_6_2)
    label_dict['label_6_2'] = label_6_2

    label_7_2 = pg.TextItem(anchor=(1, 0.5))
    label_7_2.setFont(QFont('Arial', 12))
    plot21.addItem(label_7_2)
    label_dict['label_7_2'] = label_7_2

    label_13_2 = pg.TextItem(anchor=(1, 0.5))
    label_13_2.setFont(QFont('Arial', 12))
    plot21.addItem(label_13_2)
    label_dict['label_13_2'] = label_13_2
    ######################################################################################

    #*************************Column 3*************************************#
    expiry_sensex = expiry_names[4][-10:]
    date1 = datetime.strptime(today, "%Y-%m-%d")
    date2 = datetime.strptime(expiry_sensex, "%Y-%m-%d")
    dte = (date2 - date1).days
    decay_text = dte_decay.get(dte, "N/A")

    #Plot 02
    #######################################################################################
    plot02 = win.addPlot(row=0, col=2, title=f"T: {expiry_names[0][8:18]} | Sensex - OTMs | E: {expiry_names[4][-10:]} | {dte} DTE ({decay_text})")
    plot02.addLegend()
    line_0_3 = plot02.plot([], pen=mkPen('g', width=3), name='CE OTMs')
    line_1_3 = plot02.plot([], pen=mkPen('y', width=3), name='PE OTMs')

    plot02.addLine(y=0, pen=pg.mkPen(color='w', width=1.5))

    # plot02.showGrid(x=True, y=True, alpha=0.3)

    plot02.getAxis('bottom').setTickStrings = lambda values, scale, spacing: format_time_ticks(values)

    # For plot02 : Floating Labels
    label_0_3 = pg.TextItem(anchor=(1, 0.5))
    label_0_3.setFont(QFont('Arial', 12))
    plot02.addItem(label_0_3)
    label_dict['label_0_3'] = label_0_3

    label_1_3 = pg.TextItem(anchor=(1, 0.5))
    label_1_3.setFont(QFont('Arial', 12))
    plot02.addItem(label_1_3)
    label_dict['label_1_3'] = label_1_3
    ########################################################################################

    #Plot 12
    ########################################################################################
    plot12 = win.addPlot(row=1, col=2, title="Sensex - CE-PE ATM, Straddle")
    plot12.addLegend()
    line_10_3 = plot12.plot([], pen=mkPen('g', width=3), name='CE ATM')
    line_11_3 = plot12.plot([], pen=mkPen('y', width=3), name='PE ATM')
    line_12_3 = plot12.plot([], pen=mkPen('m', width=3), name='ATM Straddle')

    plot12.addLine(y=0, pen=pg.mkPen(color='w', width=1.5))

    # plot12.showGrid(x=True, y=True, alpha=0.3)

    plot12.getAxis('bottom').setTickStrings = lambda values, scale, spacing: format_time_ticks(values)

    # Create a Text Object inside the Plot - latter can be updated with time in loop
    straddle_label_2 = pg.TextItem('Default_Text' ,color='lightgreen', anchor=(0.5, 0.5))
    straddle_label_2.setFont(QFont('Arial', 11))  # Optional: Set font and size
    straddle_label_2.setPos(0.25, 30)  # Initial position
    plot12.addItem(straddle_label_2)

    # For plot12 : Floating Labels
    label_10_3 = pg.TextItem(anchor=(1, 0.5))
    label_10_3.setFont(QFont('Arial', 12))
    plot12.addItem(label_10_3)
    label_dict['label_10_3'] = label_10_3

    label_11_3 = pg.TextItem(anchor=(1, 0.5))
    label_11_3.setFont(QFont('Arial', 12))
    plot12.addItem(label_11_3)
    label_dict['label_11_3'] = label_11_3

    label_12_3 = pg.TextItem(anchor=(1, 0.5))
    label_12_3.setFont(QFont('Arial', 12))
    plot12.addItem(label_12_3)
    label_dict['label_12_3'] = label_12_3
    ##########################################################################################

    #Plot 22
    ##########################################################################################
    plot22 = win.addPlot(row=2, col=2, title="Sensex - CE/PE OTMs Implied Volatility")
    plot22.addLegend()
    line_6_3 = plot22.plot([], pen=mkPen('g', width=3), name='CE IV')
    line_7_3 = plot22.plot([], pen=mkPen('y', width=3), name='PE IV')
    line_13_3 = plot22.plot([], pen=mkPen('r', width=3), name='SPOT')

    plot22.addLine(y=0, pen=pg.mkPen(color='w', width=1.5))

    # plot22.showGrid(x=True, y=True, alpha=0.3)

    plot22.getAxis('bottom').setTickStrings = lambda values, scale, spacing: format_time_ticks(values)

    # Create a Text Object inside the Plot - latter can be updated with time in loop
    spot_label_2 = pg.TextItem('Default_Text', color='r', anchor=(0.5, 0.5))
    spot_label_2.setFont(QFont('Arial', 11))  # Optional: Set font and size
    spot_label_2.setPos(0.25, 30)  # Initial position
    plot22.addItem(spot_label_2)

    # For plot22 : Floating Labels
    label_6_3 = pg.TextItem(anchor=(1, 0.5))
    label_6_3.setFont(QFont('Arial', 12))
    plot22.addItem(label_6_3)
    label_dict['label_6_3'] = label_6_3

    label_7_3 = pg.TextItem(anchor=(1, 0.5))
    label_7_3.setFont(QFont('Arial', 12))
    plot22.addItem(label_7_3)
    label_dict['label_7_3'] = label_7_3

    label_13_3 = pg.TextItem(anchor=(1, 0.5))
    label_13_3.setFont(QFont('Arial', 12))
    plot22.addItem(label_13_3)
    label_dict['label_13_3'] = label_13_3
    ##########################################################################################

    for plot in [plot00, plot10, plot20, plot01, plot11, plot21, plot02, plot12, plot22, win2_plot00, win2_plot01, win2_plot02]:
        plot.getAxis('bottom').tickStrings = format_time_ticks

    #################################################################################

    # ViewBox : plot00
    ###################################################################################

    vb_right_00 = pg.ViewBox()
    
    line_00_1_right = pg.PlotDataItem(pen=mkPen('g', width=3, style=QtCore.Qt.DotLine), name='CE_OI')
    line_00_2_right = pg.PlotDataItem(pen=mkPen('y', width=3, style=QtCore.Qt.DotLine), name='PE_OI')
    vb_right_00.addItem(line_00_1_right)
    vb_right_00.addItem(line_00_2_right)

    plot00.getAxis('right').linkToView(vb_right_00)

    vb_right_00.setXLink(plot00) # Keeps the X-axes synchronized — so both ViewBoxes scroll/zoom in sync on the time/index axis.
    vb_right_00.setYLink(None) # Makes sure Y-axis is not linked to the main plot’s Y-axis. That’s important because CE_OI/PE_OI might have a totally different scale from price/indicator.

    plot00.showAxis('right') # Makes the right Y-axis visible on the plot. Without this, you wouldn't see the axis lines or ticks, even if linked.
    plot00.getAxis('right').setStyle(showValues=False) # Hides numeric tick labels on the right Y-axis (optional). You still see the axis line, but not the numbers — cleaner UI if values aren't useful.

    plot00.addLegend()
    plot00.legend.addItem(line_00_1_right, 'CE_OI')
    plot00.legend.addItem(line_00_2_right, 'PE_OI')

    plot00.scene().addItem(vb_right_00) # Adds your custom ViewBox to the graphics scene — mandatory so it’s visible and rendered. This “attaches” it to the visual hierarchy
 
    def updateViews_00():
        """Ensure the right ViewBox is geometrically aligned with the left ViewBox."""
        vb_right_00.setGeometry(plot00.vb.sceneBoundingRect())

    def adjust_right_view_00():
        """Adjust the right Y-axis range of plot00 based on y8_1 and y9_1."""
        combined = []
        if y8_1:
            combined.extend(y8_1)
        if y9_1:
            combined.extend(y9_1)

        if combined:
            min_val = min(combined) * 0.99  # 1% padding below
            max_val = max(combined) * 1.01  # 1% padding above
            vb_right_00.setYRange(min_val, max_val)

    # Connect the left ViewBox resize signal to keep the right ViewBox aligned
    plot00.vb.sigResized.connect(updateViews_00)


    #################################################################################

    # ViewBox : plot10
    ###################################################################################

    vb_right_10 = pg.ViewBox()  # Creates a custom ViewBox for the right Y-axis

    line_10_1_right = pg.PlotDataItem(pen=mkPen('w', width=3), name='Straddle')
    line_10_2_right = pg.PlotDataItem(pen=mkPen('r', width=3), name='VWAP')
    vb_right_10.addItem(line_10_1_right)
    vb_right_10.addItem(line_10_2_right)

    plot10.getAxis('right').linkToView(vb_right_10)  # Link the right Y-axis to vb_right_10

    vb_right_10.setXLink(plot10)  # Synchronize X-axis zoom/pan with the main plot
    vb_right_10.setYLink(None)  # Make sure the Y-axis is independent

    plot10.showAxis('right')  # Show the right axis (makes the axis line visible)
    plot10.getAxis('right').setStyle(showValues=False)  # Optionally hide numeric tick labels on right Y-axis

    plot10.addLegend()
    plot10.legend.addItem(line_10_1_right, 'Abs Straddle')
    plot10.legend.addItem(line_10_2_right, 'VWAP')

    plot10.scene().addItem(vb_right_10)

    def updateViews_10():
        """Ensure the right ViewBox is geometrically aligned with the left ViewBox."""
        vb_right_10.setGeometry(plot10.vb.sceneBoundingRect())

    def adjust_right_view_10():
        """Adjust the right Y-axis range of plot10 based on y16_1 and y17_1."""
        if len(y16_1) > 0 or len(y17_1) > 0:
            combined = []
            if len(y16_1) > 0:
                combined.extend(y16_1)
            if len(y17_1) > 0:
                combined.extend(y17_1)

            min_val = min(combined) - abs(min(combined) * 0.01)  # 1% padding below
            max_val = max(combined) + abs(max(combined) * 0.01)  # 1% padding above
            vb_right_10.setYRange(min_val, max_val)

    # Keep the ViewBoxes aligned during window/plot resize
    plot10.vb.sigResized.connect(updateViews_10)
    #############################################################

    # ViewBox: plot20
    #################################################################################

    vb_right_20 = pg.ViewBox()  # Creates a new ViewBox for the right side of plot20

    line_13_1_right = pg.PlotDataItem(pen=mkPen(color='g', width=3, style=QtCore.Qt.DotLine), name='CE OBV')
    line_13_22_right = pg.PlotDataItem(pen=mkPen(color='y', width=3, style=QtCore.Qt.DotLine), name='PE OBV')
    vb_right_20.addItem(line_13_1_right)  # Add the data item to the right ViewBox
    vb_right_20.addItem(line_13_22_right)  # Add the data item to the right ViewBox

    plot20.getAxis('right').linkToView(vb_right_20)  # Link right axis to right ViewBox

    vb_right_20.setXLink(plot20)  # Keep X-axis linked for alignment
    vb_right_20.setYLink(None)  # Explicitly unlink Y-axis

    plot20.showAxis('right')
    plot20.getAxis('right').setStyle(showValues=False) # Disable the display of numbers (data) on the right Y-axis

    plot20.addLegend()
    plot20.legend.addItem(line_13_1_right, 'CE OBV')
    plot20.legend.addItem(line_13_22_right, 'PE OBV')

    plot20.scene().addItem(vb_right_20)

    # ViewBox Plot 20 : Floating Labels
    font = QFont('Arial', 12)
    font.setBold(True)
    label_13_1_right = TextItem(anchor=(1, 0.5))
    label_13_1_right.setFont(font)
    label_13_1_right.setColor(QColor(255, 0, 0))  # Bright red
    vb_right_20.addItem(label_13_1_right)
    label_dict['label_13_1_right'] = label_13_1_right

    font2 = QFont('Arial', 12)
    font2.setBold(True)
    label_13_2_right = TextItem(anchor=(1, 0.5))
    label_13_2_right.setFont(font2)
    label_13_2_right.setColor(QColor(255, 255, 255))  # Bright white
    vb_right_20.addItem(label_13_2_right)
    label_dict['label_13_2_right'] = label_13_2_right

    def updateViews_20():
        """Ensure the right ViewBox is aligned with the left ViewBox geometrically."""
        vb_right_20.setGeometry(plot20.vb.sceneBoundingRect())

    def adjust_right_view_20():
        """Dynamically adjust the range of the right Y-axis for plot20 based on y18_1 and y19_1."""
        if len(y18_1) > 0 or len(y19_1) > 0:
            combined = []
            if len(y18_1) > 0:
                combined.extend(y18_1)
            if len(y19_1) > 0:
                combined.extend(y19_1)

            min_val = min(combined) - abs(min(combined) * 0.01)  # 0.1% padding below
            max_val = max(combined) + abs(max(combined) * 0.01)  # 0.1% padding above
            vb_right_20.setYRange(min_val, max_val)


    plot20.vb.sigResized.connect(updateViews_20)


    ###############################################################################

    # ViewBox : plot01
    ###################################################################################

    vb_right_01 = pg.ViewBox()

    line_01_1_right = pg.PlotDataItem(pen=mkPen(color='g', width=3, style=QtCore.Qt.DotLine), name='CE_OI')
    line_01_2_right = pg.PlotDataItem(pen=mkPen(color='y', width=3, style=QtCore.Qt.DotLine), name='PE_OI')
    vb_right_01.addItem(line_01_1_right)
    vb_right_01.addItem(line_01_2_right)

    plot01.getAxis('right').linkToView(vb_right_01)

    vb_right_01.setXLink(plot01)
    vb_right_01.setYLink(None)

    plot01.showAxis('right')
    plot01.getAxis('right').setStyle(showValues=False)

    plot01.addLegend()
    plot01.legend.addItem(line_01_1_right, 'CE_OI')
    plot01.legend.addItem(line_01_2_right, 'PE_OI')

    plot01.scene().addItem(vb_right_01)

    def updateViews_01():
        """Ensure the right ViewBox is geometrically aligned with the left ViewBox."""
        vb_right_01.setGeometry(plot01.vb.sceneBoundingRect())

    def adjust_right_view_01():
        """Dynamically adjust the range of the right Y-axis for plot01 based on y8_2 and y9_2."""
        combined = []
        if y8_2:
            combined.extend(y8_2)
        if y9_2:
            combined.extend(y9_2)

        if combined:
            min_val = min(combined) * 0.99  # 1% padding below
            max_val = max(combined) * 1.01  # 1% padding above
            vb_right_01.setYRange(min_val, max_val)

    plot01.vb.sigResized.connect(updateViews_01)

    # ViewBox : plot11
    ###################################################################################

    vb_right_11 = pg.ViewBox()

    line_11_1_right = pg.PlotDataItem(pen=mkPen('w', width=3), name='Straddle')
    line_11_2_right = pg.PlotDataItem(pen=mkPen('r', width=3), name='VWAP')
    vb_right_11.addItem(line_11_1_right)
    vb_right_11.addItem(line_11_2_right)

    plot11.getAxis('right').linkToView(vb_right_11)

    vb_right_11.setXLink(plot11)
    vb_right_11.setYLink(None)

    plot11.showAxis('right')
    plot11.getAxis('right').setStyle(showValues=False)
    
    plot11.addLegend()
    plot11.legend.addItem(line_11_1_right, 'Abs Straddle')
    plot11.legend.addItem(line_11_2_right, 'VWAP')

    plot11.scene().addItem(vb_right_11)

    def updateViews_11():
        """Ensure the right ViewBox is geometrically aligned with the left ViewBox."""
        vb_right_11.setGeometry(plot11.vb.sceneBoundingRect())
        vb_right_11.setYLink(None)  # Extra safety

    def adjust_right_view_11():
        """Dynamically adjust the range of the right Y-axis for plot11 based on y16_2 and y17_2."""
        combined = []
        if len(y16_2) > 0:
            combined.extend(y16_2)
        if len(y17_2) > 0:
            combined.extend(y17_2)

        if combined:
            min_val = min(combined) * 0.99  # 1% padding below
            max_val = max(combined) * 1.01  # 1% padding above
            vb_right_11.setYRange(min_val, max_val)

    plot11.vb.sigResized.connect(updateViews_11)

    # ViewBox : plot21
    ###################################################################################

    vb_right_21 = pg.ViewBox()

    line_13_2_right = pg.PlotDataItem(pen=mkPen(color='g', width=3, style=QtCore.Qt.DotLine), name='CE OBV')
    line_13_222_right = pg.PlotDataItem(pen=mkPen(color='y', width=3, style=QtCore.Qt.DotLine), name='PE OBV')
    vb_right_21.addItem(line_13_2_right)
    vb_right_21.addItem(line_13_222_right)

    plot21.getAxis('right').linkToView(vb_right_21)

    vb_right_21.setXLink(plot21)
    vb_right_21.setYLink(None)

    plot21.showAxis('right')
    plot21.getAxis('right').setStyle(showValues=False)
    
    plot21.addLegend()
    plot21.legend.addItem(line_13_2_right, 'CE OBV')
    plot21.legend.addItem(line_13_222_right, 'PE OBV')

    plot21.scene().addItem(vb_right_21)

    # Floating Labels
    font = QFont('Arial', 12)
    font.setBold(True)
    label_13_3_right = TextItem(anchor=(1, 0.5))
    label_13_3_right.setFont(font)
    label_13_3_right.setColor(QColor(255, 0, 0))  # Bright red
    vb_right_21.addItem(label_13_3_right)
    label_dict['label_13_3_right'] = label_13_3_right

    font2 = QFont('Arial', 12)
    font2.setBold(True)
    label_13_4_right = TextItem(anchor=(1, 0.5))
    label_13_4_right.setFont(font2)
    label_13_4_right.setColor(QColor(255, 255, 255))  # Bright white
    vb_right_21.addItem(label_13_4_right)
    label_dict['label_13_4_right'] = label_13_4_right

    def updateViews_21():
        """Ensure the right ViewBox is geometrically aligned with the left ViewBox."""
        vb_right_21.setGeometry(plot21.vb.sceneBoundingRect())

    def adjust_right_view_21():
        """Dynamically adjust the range of the right Y-axis for plot21 based on y18_2 and y19_2."""
        combined = []
        if y18_2:
            combined.extend(y18_2)
        if y19_2:
            combined.extend(y19_2)

        if combined:
            min_val = min(combined) * 0.99  # 1% padding below
            max_val = max(combined) * 1.01  # 1% padding above
            vb_right_21.setYRange(min_val, max_val)

    plot21.vb.sigResized.connect(updateViews_21)

    ################################################################################

    # ViewBox : plot02
    ###################################################################################

    vb_right_02 = pg.ViewBox()  # Create custom right-side ViewBox

    line_02_1_right = pg.PlotDataItem(pen=mkPen(color='g', width=3, style=QtCore.Qt.DotLine), name='CE_OI')
    line_02_2_right = pg.PlotDataItem(pen=mkPen(color='y', width=3, style=QtCore.Qt.DotLine), name='PE_OI')
    vb_right_02.addItem(line_02_1_right)
    vb_right_02.addItem(line_02_2_right)

    plot02.getAxis('right').linkToView(vb_right_02)

    vb_right_02.setXLink(plot02)
    vb_right_02.setYLink(None)

    plot02.showAxis('right')
    plot02.getAxis('right').setStyle(showValues=False)

    plot02.addLegend()
    plot02.legend.addItem(line_02_1_right, 'CE_OI')
    plot02.legend.addItem(line_02_2_right, 'PE_OI')

    plot02.scene().addItem(vb_right_02)

    def updateViews_02():
        """Ensure the right ViewBox is geometrically aligned with the left ViewBox."""
        vb_right_02.setGeometry(plot02.vb.sceneBoundingRect())

    def adjust_right_view_02():
        """Dynamically adjust the range of the right Y-axis for plot02 based on y8_3 and y9_3."""
        combined = []
        if y8_3:
            combined.extend(y8_3)
        if y9_3:
            combined.extend(y9_3)

        if combined:
            min_val = min(combined) * 0.99  # 1% padding below
            max_val = max(combined) * 1.01  # 1% padding above
            vb_right_02.setYRange(min_val, max_val)

    plot02.vb.sigResized.connect(updateViews_02)

    #################################################################################


    # ViewBox : plot12
    ###################################################################################

    vb_right_12 = pg.ViewBox()  # Create custom right-side ViewBox

    line_12_1_right = pg.PlotDataItem(pen=mkPen('w', width=3), name='Straddle')
    line_12_2_right = pg.PlotDataItem(pen=mkPen('r', width=3), name='VWAP')
    vb_right_12.addItem(line_12_1_right)
    vb_right_12.addItem(line_12_2_right)

    plot12.getAxis('right').linkToView(vb_right_12)

    vb_right_12.setXLink(plot12)
    vb_right_12.setYLink(None)

    plot12.showAxis('right')
    plot12.getAxis('right').setStyle(showValues=False)

    plot12.addLegend()
    plot12.legend.addItem(line_12_1_right, 'Abs Straddle')
    plot12.legend.addItem(line_12_2_right, 'VWAP')

    plot12.scene().addItem(vb_right_12)

    def updateViews_12():
        """Ensure the right ViewBox is geometrically aligned with the left ViewBox."""
        vb_right_12.setGeometry(plot12.vb.sceneBoundingRect())
        vb_right_12.setYLink(None)  # Prevent accidental Y-axis linking

    def adjust_right_view_12():
        """Dynamically adjust the range of the right Y-axis for plot12 based on y16_3 and y17_3."""
        combined = []
        if len(y16_3) > 0:
            combined.extend(y16_3)
        if len(y17_3) > 0:
            combined.extend(y17_3)

        if combined:
            min_val = min(combined) * 0.99  # 1% padding below
            max_val = max(combined) * 1.01  # 1% padding above
            vb_right_12.setYRange(min_val, max_val)

    plot12.vb.sigResized.connect(updateViews_12)

    # ViewBox : plot22
    ###################################################################################

    vb_right_22 = pg.ViewBox()

    line_13_3_right = pg.PlotDataItem(pen=mkPen(color='g', width=3, style=QtCore.Qt.DotLine), name='CE OBV')
    line_13_333_right = pg.PlotDataItem(pen=mkPen(color='y', width=3, style=QtCore.Qt.DotLine), name='PE OBV')
    vb_right_22.addItem(line_13_3_right)
    vb_right_22.addItem(line_13_333_right)

    plot22.getAxis('right').linkToView(vb_right_22)

    vb_right_22.setXLink(plot22)
    vb_right_22.setYLink(None)

    plot22.showAxis('right')
    plot22.getAxis('right').setStyle(showValues=False)

    plot22.addLegend()
    plot22.legend.addItem(line_13_3_right, 'CE OBV')
    plot22.legend.addItem(line_13_333_right, 'PE OBV')

    plot22.scene().addItem(vb_right_22)

    # Floating Labels for plot22
    font = QFont('Arial', 12)
    font.setBold(True)
    label_13_5_right = TextItem(anchor=(1, 0.5))
    label_13_5_right.setFont(font)
    label_13_5_right.setColor(QColor(255, 0, 0))  # Bright red
    vb_right_22.addItem(label_13_5_right)
    label_dict['label_13_5_right'] = label_13_5_right

    font2 = QFont('Arial', 12)
    font2.setBold(True)
    label_13_6_right = TextItem(anchor=(1, 0.5))
    label_13_6_right.setFont(font2)
    label_13_6_right.setColor(QColor(255, 255, 255))  # Bright white
    vb_right_22.addItem(label_13_6_right)
    label_dict['label_13_6_right'] = label_13_6_right

    # For plot22
    def updateViews_22():
        """Ensure the right ViewBox is geometrically aligned with the left ViewBox."""
        vb_right_22.setGeometry(plot22.vb.sceneBoundingRect())

    def adjust_right_view_22():
        """Dynamically adjust the range of the right Y-axis for plot22 based on y18_3 and y19_3."""
        combined = []
        if y18_3:
            combined.extend(y18_3)
        if y19_3:
            combined.extend(y19_3)

        if combined:
            min_val = min(combined) * 0.99
            max_val = max(combined) * 1.01
            vb_right_22.setYRange(min_val, max_val)

    plot22.vb.sigResized.connect(updateViews_22)

############################################################ 

    xx = locals()

    return xx

##########################################################
ttime = None
x_val00, x_val10, x_val11, x_val12, x_val20, x_val21, x_val22 = [None]*7
y_val00, y_val10, y_val11, y_val12, y_val20, y_val21, y_val22 = [None]*7
aa = 1
def update(dfp,xx,str_curr,str_init,dfpp,obv_df):
    global x_val00, x_val10, x_val11, x_val12, x_val20, x_val21, x_val22, y_val00, y_val10, y_val11, y_val12, y_val20, y_val21, y_val22, aa
    global lin_data_val, quad_data_val
    ns = SimpleNamespace(**xx)
    global y13_1, y13_2, y13_3, y14_1, y16_1, y16_2, y16_3

    ###############################################################

    strad0 = dfpp['straddle0']
    strad0 = strad0.ewm(span=20, adjust=False).mean()
    vwap0 = dfpp['VWAP0']

    strad1 = dfpp['straddle1']
    strad1 = strad1.ewm(span=20, adjust=False).mean()
    vwap1 = dfpp['VWAP1']

    # strad2 = dfpp['straddle2']
    # strad2 = strad2.ewm(span=20, adjust=False).mean()
    # vwap2 = dfpp['VWAP2']

    strad3 = dfpp['straddle3']
    strad3 = strad3.ewm(span=20, adjust=False).mean()
    vwap3 = dfpp['VWAP3']

    strad4 = dfpp['straddle4']
    strad4 = strad4.ewm(span=20, adjust=False).mean()
    vwap4 = dfpp['VWAP4']

    ####################################################################################

    ce_obv_0 = obv_df['ce_obv_0']
    pe_obv_0 = obv_df['pe_obv_0']
    ce_obv_0 = ce_obv_0.ewm(span=20, adjust=False).mean()
    pe_obv_0 = pe_obv_0.ewm(span=20, adjust=False).mean()

    ce_obv_1 = obv_df['ce_obv_1']
    pe_obv_1 = obv_df['pe_obv_1']
    ce_obv_1 = ce_obv_1.ewm(span=20, adjust=False).mean()
    pe_obv_1 = pe_obv_1.ewm(span=20, adjust=False).mean()

    # ce_obv_2 = obv_df['ce_obv_2']
    # pe_obv_2 = obv_df['pe_obv_2']
    # ce_obv_2 = ce_obv_2.ewm(span=20, adjust=False).mean()
    # pe_obv_2 = pe_obv_2.ewm(span=20, adjust=False).mean()

    ce_obv_3 = obv_df['ce_obv_3']
    pe_obv_3 = obv_df['pe_obv_3']
    ce_obv_3 = ce_obv_3.ewm(span=20, adjust=False).mean()
    pe_obv_3 = pe_obv_3.ewm(span=20, adjust=False).mean()

    ce_obv_4 = obv_df['ce_obv_4']
    pe_obv_4 = obv_df['pe_obv_4']
    ce_obv_4 = ce_obv_4.ewm(span=20, adjust=False).mean()
    pe_obv_4 = pe_obv_4.ewm(span=20, adjust=False).mean()

    ####################################################################################

    ###############################################################
    ttime = time_fun()
    ttime1 = ttime[:5]
    vix = str_curr.iloc[14,0]
    ns.time_label.setText(f'{ttime} | Vix : {vix}')

    x_range, y_range = ns.plot00.viewRange()
    x_val00 = x_range[0] + (x_range[1] - x_range[0]) * 0.5
    y_val00 = y_range[0] + (y_range[1] - y_range[0]) * 0.95
    ns.time_label.setPos(x_val00, y_val00)
    ################################################################
    ce_0 = str_curr.iloc[10,0]
    pe_0 = str_curr.iloc[11,0]
    straddle_curr_0 = round((ce_0 + pe_0),2)
    straddle_init_0 = round(str_init.iloc[12,0],2)
    decay_0 = round((straddle_curr_0 - straddle_init_0),2)
    strad_lab_0 = f'{ce_0} + {pe_0} = {straddle_curr_0} ({straddle_init_0}) | Change = {decay_0} ({round(decay_0/straddle_init_0*100,2)} %)'
    ns.straddle_label_0.setText(strad_lab_0)

    x_range, y_range = ns.plot10.viewRange()
    x_val10 = x_range[0] + (x_range[1] - x_range[0]) * 0.5
    y_val10 = y_range[0] + (y_range[1] - y_range[0]) * 0.95
    ns.straddle_label_0.setPos(x_val10, y_val10)
    ################################################################
    ce_1 = str_curr.iloc[10,3]
    pe_1 = str_curr.iloc[11,3]
    straddle_curr_1 = round((ce_1 + pe_1),2)
    straddle_init_1 = round(str_init.iloc[12,3],2)
    decay_1 = round((straddle_curr_1 - straddle_init_1),2)
    strad_lab_1 = f'{ce_1} + {pe_1} = {straddle_curr_1} ({straddle_init_1}) | Change = {decay_1} ({round(decay_1/straddle_init_1*100,2)} %)'
    ns.straddle_label_1.setText(strad_lab_1)

    x_range, y_range = ns.plot11.viewRange()
    x_val11 = x_range[0] + (x_range[1] - x_range[0]) * 0.5
    y_val11 = y_range[0] + (y_range[1] - y_range[0]) * 0.95
    ns.straddle_label_1.setPos(x_val11, y_val11)
    ################################################################
    ce_2 = str_curr.iloc[10,4]
    pe_2 = str_curr.iloc[11,4]
    straddle_curr_2 = round((ce_2 + pe_2),2)
    straddle_init_2 = round(str_init.iloc[12,4],2)
    decay_2 = round((straddle_curr_2 - straddle_init_2),2)
    strad_lab_2 = f'{ce_2} + {pe_2} = {straddle_curr_2} ({straddle_init_2}) | Change = {decay_2} ({round(decay_2/straddle_init_2*100,2)} %)'
    ns.straddle_label_2.setText(strad_lab_2)

    x_range, y_range = ns.plot12.viewRange()
    x_val12 = x_range[0] + (x_range[1] - x_range[0]) * 0.5
    y_val12 = y_range[0] + (y_range[1] - y_range[0]) * 0.95
    ns.straddle_label_2.setPos(x_val12, y_val12)
    ################################################################
    spot0 = str(str_curr.iloc[13,0])
    spot_diff_0 = str(round((str_curr.iloc[13,0]) - (str_init.iloc[13,0]),2))
    spot_init_0 = round(str_init.iloc[13,0],2)
    per_change = round((float(spot_diff_0)/float(spot_init_0))*100,2)
    ns.spot_label_0.setText(f'{spot0} : {spot_diff_0} ({per_change} %) | {dfp.iloc[16,0][-1]}')

    x_range, y_range = ns.plot20.viewRange()
    x_val20 = x_range[0] + (x_range[1] - x_range[0]) * 0.5
    y_val20 = y_range[0] + (y_range[1] - y_range[0]) * 0.95
    ns.spot_label_0.setPos(x_val20, y_val20)
    #################################################################
    spot1 = str(str_curr.iloc[13,3])
    spot_diff_1 = str(round((str_curr.iloc[13,3]) - (str_init.iloc[13,3]),2))
    spot_init_1 = round(str_init.iloc[13,3],2)
    per_change = round((float(spot_diff_1)/float(spot_init_1))*100,2)
    ns.spot_label_1.setText(f'{spot1} : {spot_diff_1} ({per_change} %) | {dfp.iloc[16,3][-1]}')

    x_range, y_range = ns.plot21.viewRange()
    x_val21 = x_range[0] + (x_range[1] - x_range[0]) * 0.5
    y_val21 = y_range[0] + (y_range[1] - y_range[0]) * 0.95
    ns.spot_label_1.setPos(x_val21, y_val21)
    #################################################################
    spot2 = str(str_curr.iloc[13,4])
    spot_diff_2 = str(round((str_curr.iloc[13,4]) - (str_init.iloc[13,4]),2))
    spot_init_2 = round(str_init.iloc[13,4],2)
    per_change = round((float(spot_diff_2)/float(spot_init_2))*100,2)
    ns.spot_label_2.setText(f'{spot2} : {spot_diff_2} ({per_change} %) | {dfp.iloc[16,4][-1]}')

    x_range, y_range = ns.plot22.viewRange()
    x_val22 = x_range[0] + (x_range[1] - x_range[0]) * 0.5
    y_val22 = y_range[0] + (y_range[1] - y_range[0]) * 0.95
    ns.spot_label_2.setPos(x_val22, y_val22)
    #################################################################

    # Get time data for all Four expiries
    time_str_1 = dfp.iloc[15,0]
    time_str_2 = dfp.iloc[15,3]
    time_str_3 = dfp.iloc[15,4]
    time_str_4 = dfp.iloc[15,1]

    # Convert time strings to timestamps
    timestamp_1 = time_string_to_timestamp(time_str_1)
    timestamp_2 = time_string_to_timestamp(time_str_2)
    timestamp_3 = time_string_to_timestamp(time_str_3)
    timestamp_4 = time_string_to_timestamp(time_str_4)

    # Store timestamps in x-axis arrays
    x1 = timestamp_1
    x2 = timestamp_2
    x3 = timestamp_3
    x4 = timestamp_4

    # Update Expiry 1 data
    y0_1 = dfp.iloc[0,0] # ce_ltp
    y1_1 = dfp.iloc[1,0] # pe_ltp
    y6_1 = dfp.iloc[6,0] # ce_iv
    y7_1 = dfp.iloc[7,0] # pe_iv
    y8_1 = dfp.iloc[8,0] # ce_oi
    y9_1 = dfp.iloc[9,0] # pe_oi
    y10_1 = dfp.iloc[10,0] # ce_atm - changes
    y11_1 = dfp.iloc[11,0] # pe_atm - changes
    y12_1 = dfp.iloc[12,0] # atm straddle - changes
    y16_1 = strad0 # ATM STraddle
    y17_1 = vwap0 # Straddle VWAP
    y18_1 = ce_obv_0 # CE OBV
    y19_1 = pe_obv_0 # PE OBV

    # Update Expiry 1 plots with timestamps
    ns.line_0_1.setData(x1, y0_1) # ce_ltp
    ns.line_1_1.setData(x1, y1_1) # pe_ltp
    ns.line_6_1.setData(x1, y6_1) # ce_iv
    ns.line_7_1.setData(x1, y7_1) # pe_iv
    ns.line_00_1_right.setData(x1, y8_1)  # ce_oi
    ns.line_00_2_right.setData(x1, y9_1)  # pe_oi
    ns.line_10_1.setData(x1, y10_1) # ce_atm - changes
    ns.line_11_1.setData(x1, y11_1) # pe_atm - changes
    ns.line_12_1.setData(x1, y12_1) # atm straddle - changes
    ns.line_10_1_right.setData(x1, y16_1)  # ATM Straddle
    ns.line_10_2_right.setData(x1, y17_1)  # Straddle VWAP
    ns.line_13_1_right.setData(x1, y18_1) # CE OBV
    ns.line_13_22_right.setData(x1, y19_1) # PE OBV
    ns.adjust_right_view_00()
    ns.adjust_right_view_10()
    ns.adjust_right_view_20()

    # Update Expiry 2 data
    y0_2 = dfp.iloc[0,3]
    y1_2 = dfp.iloc[1,3]
    y6_2 = dfp.iloc[6,3]
    y7_2 = dfp.iloc[7,3]
    y8_2 = dfp.iloc[8,3]
    y9_2 = dfp.iloc[9,3]
    y10_2 = dfp.iloc[10,3]
    y11_2 = dfp.iloc[11,3]
    y12_2 = dfp.iloc[12,3]
    y16_2 = strad3
    y17_2 = vwap3
    y18_2 = ce_obv_3
    y19_2 = pe_obv_3

    # Update Expiry 2 plots with timestamps
    ns.line_0_2.setData(x2, y0_2)
    ns.line_1_2.setData(x2, y1_2)
    ns.line_6_2.setData(x2, y6_2)
    ns.line_7_2.setData(x2, y7_2)
    ns.line_01_1_right.setData(x2, y8_2)
    ns.line_01_2_right.setData(x2, y9_2)
    ns.line_10_2.setData(x2, y10_2)
    ns.line_11_2.setData(x2, y11_2)
    ns.line_12_2.setData(x2, y12_2)
    ns.line_11_1_right.setData(x2, y16_2)
    ns.line_11_2_right.setData(x2, y17_2)
    ns.line_13_2_right.setData(x2, y18_2)
    ns.line_13_222_right.setData(x2, y19_2)
    ns.adjust_right_view_01()
    ns.adjust_right_view_11()
    ns.adjust_right_view_21()

    # Update Expiry 3 data
    y0_3 = dfp.iloc[0,4]
    y1_3 = dfp.iloc[1,4]
    y6_3 = dfp.iloc[6,4]
    y7_3 = dfp.iloc[7,4]
    y8_3 = dfp.iloc[8,4]
    y9_3 = dfp.iloc[9,4]
    y16_3 = strad4
    y10_3 = dfp.iloc[10,4]
    y11_3 = dfp.iloc[11,4]
    y12_3 = dfp.iloc[12,4]
    y17_3 = vwap4
    y18_3 = ce_obv_4
    y19_3 = pe_obv_4

    # Update Expiry 3 plots with timestamps
    ns.line_0_3.setData(x3, y0_3)
    ns.line_1_3.setData(x3, y1_3)
    ns.line_6_3.setData(x3, y6_3)
    ns.line_7_3.setData(x3, y7_3)
    ns.line_02_1_right.setData(x3, y8_3)
    ns.line_02_2_right.setData(x3, y9_3)
    ns.line_10_3.setData(x3, y10_3)
    ns.line_11_3.setData(x3, y11_3)
    ns.line_12_3.setData(x3, y12_3)
    ns.line_12_1_right.setData(x3, y16_3)
    ns.line_12_2_right.setData(x3, y17_3)
    ns.line_13_3_right.setData(x3, y18_3)
    ns.line_13_333_right.setData(x3, y19_3)
    ns.adjust_right_view_02()
    ns.adjust_right_view_12()
    ns.adjust_right_view_22()


    # Update Expiry 4 data
    y0_4 = dfp.iloc[0,1]
    y1_4 = dfp.iloc[1,1]
    y6_4 = dfp.iloc[6,1]
    y7_4 = dfp.iloc[7,1]
    y8_4 = dfp.iloc[8,1]
    y9_4 = dfp.iloc[9,1]
    y10_4 = dfp.iloc[10,1]
    y11_4 = dfp.iloc[11,1]
    y12_4 = dfp.iloc[12,1]
    y16_4 = strad1
    y17_4 = vwap1
    y18_4 = ce_obv_1
    y19_4 = pe_obv_1

    # Update Expiry 4 plots with timestamps
    ns.line_ce.setData(x4, y0_4)
    ns.line_pe.setData(x4, y1_4)
    ns.line_ce_iv.setData(x4, y6_4)
    ns.line_pe_iv.setData(x4, y7_4)
    ns.plot00_viewbox_ce_oi.setData(x4, y8_4)
    ns.plot00_viewbox_pe_oi.setData(x4, y9_4)
    ns.line_ce_atm.setData(x4, y10_4)
    ns.line_pe_atm.setData(x4, y11_4)
    ns.line_straddle.setData(x4, y12_4)
    ns.plot01_viewbox_abs_straddle.setData(x4, y16_4)
    ns.plot01_viewbox_vwap.setData(x4, y17_4)
    ns.plot02_viewbox_1.setData(x4, y18_4)
    ns.plot02_viewbox_2.setData(x4, y19_4)
    ns.adjust_right_view_win2_00()
    ns.adjust_right_view_win2_01()
    ns.adjust_right_view_win2_02()

    ###############################################################################
    lin_data_val = int(summary.range('A43').value)
    quad_data_val = int(summary.range('B43').value)

    update_regression(strad0, ns.straddle_1, ns.linear_1, ns.quad_1, ns.linear_eqn_1, ns.quad_eqn_1)
    update_regression(strad3, ns.straddle_2, ns.linear_2, ns.quad_2, ns.linear_eqn_2, ns.quad_eqn_2)
    update_regression(strad4, ns.straddle_3, ns.linear_3, ns.quad_3, ns.linear_eqn_3, ns.quad_eqn_3)
    update_regression(strad1, ns.straddle_4, ns.linear_4, ns.quad_4, ns.linear_eqn_4, ns.quad_eqn_4)

    ###############################################################################

    # Update Labels with timestamps for x-position

    if x1 and y0_1:
        # Update Expiry 1 labels
        ns.label_dict['label_0'].setText(f"{y0_1[-1]:.2f}")
        ns.label_dict['label_0'].setPos(x1[-1], y0_1[-1])
        ns.label_dict['label_1'].setText(f"{y1_1[-1]:.2f}")
        ns.label_dict['label_1'].setPos(x1[-1], y1_1[-1])
        ns.label_dict['label_6'].setText(f"{y6_1[-1]:.2f}")
        ns.label_dict['label_6'].setPos(x1[-1], y6_1[-1])
        ns.label_dict['label_7'].setText(f"{y7_1[-1]:.2f}")
        ns.label_dict['label_7'].setPos(x1[-1], y7_1[-1])
        ns.label_dict['label_10'].setText(f"{y10_1[-1]:.2f}")
        ns.label_dict['label_10'].setPos(x1[-1], y10_1[-1])
        ns.label_dict['label_11'].setText(f"{y11_1[-1]:.2f}")
        ns.label_dict['label_11'].setPos(x1[-1], y11_1[-1])
        ns.label_dict['label_12'].setText(f"{y12_1[-1]:.2f}")
        ns.label_dict['label_12'].setPos(x1[-1], y12_1[-1])
        ns.label_dict['label_13_1_right'].setText(f"{y18_1.iloc[-1]/1000000:.2f}")
        ns.label_dict['label_13_1_right'].setPos(x1[-1], y18_1.iloc[-1])
        ns.label_dict['label_13_2_right'].setText(f"{y19_1.iloc[-1]/1000000:.2f}")
        ns.label_dict['label_13_2_right'].setPos(x1[-1], y19_1.iloc[-1])

    if x2 and y0_2:
        # Update Expiry 2 labels
        ns.label_dict['label_0_2'].setText(f"{y0_2[-1]:.2f}")
        ns.label_dict['label_0_2'].setPos(x2[-1], y0_2[-1])
        ns.label_dict['label_1_2'].setText(f"{y1_2[-1]:.2f}")
        ns.label_dict['label_1_2'].setPos(x2[-1], y1_2[-1])
        ns.label_dict['label_6_2'].setText(f"{y6_2[-1]:.2f}")
        ns.label_dict['label_6_2'].setPos(x2[-1], y6_2[-1])
        ns.label_dict['label_7_2'].setText(f"{y7_2[-1]:.2f}")
        ns.label_dict['label_7_2'].setPos(x2[-1], y7_2[-1])
        ns.label_dict['label_10_2'].setText(f"{y10_2[-1]:.2f}")
        ns.label_dict['label_10_2'].setPos(x2[-1], y10_2[-1])
        ns.label_dict['label_11_2'].setText(f"{y11_2[-1]:.2f}")
        ns.label_dict['label_11_2'].setPos(x2[-1], y11_2[-1])
        ns.label_dict['label_12_2'].setText(f"{y12_2[-1]:.2f}")
        ns.label_dict['label_12_2'].setPos(x2[-1], y12_2[-1])
        ns.label_dict['label_13_3_right'].setText(f"{y18_2.iloc[-1]/1000000:.2f}")
        ns.label_dict['label_13_3_right'].setPos(x1[-1], y18_2.iloc[-1])
        ns.label_dict['label_13_4_right'].setText(f"{y19_2.iloc[-1]/1000000:.2f}")
        ns.label_dict['label_13_4_right'].setPos(x1[-1], y19_2.iloc[-1])

    if x3 and y0_3:
        # Update Expiry 3 labels
        ns.label_dict['label_0_3'].setText(f"{y0_3[-1]:.2f}")
        ns.label_dict['label_0_3'].setPos(x3[-1], y0_3[-1])
        ns.label_dict['label_1_3'].setText(f"{y1_3[-1]:.2f}")
        ns.label_dict['label_1_3'].setPos(x3[-1], y1_3[-1])
        ns.label_dict['label_6_3'].setText(f"{y6_3[-1]:.2f}")
        ns.label_dict['label_6_3'].setPos(x3[-1], y6_3[-1])
        ns.label_dict['label_7_3'].setText(f"{y7_3[-1]:.2f}")
        ns.label_dict['label_7_3'].setPos(x3[-1], y7_3[-1])
        ns.label_dict['label_10_3'].setText(f"{y10_3[-1]:.2f}")
        ns.label_dict['label_10_3'].setPos(x3[-1], y10_3[-1])
        ns.label_dict['label_11_3'].setText(f"{y11_3[-1]:.2f}")
        ns.label_dict['label_11_3'].setPos(x3[-1], y11_3[-1])
        ns.label_dict['label_12_3'].setText(f"{y12_3[-1]:.2f}")
        ns.label_dict['label_12_3'].setPos(x3[-1], y12_3[-1])
        ns.label_dict['label_13_5_right'].setText(f"{y18_3.iloc[-1]/1000000:.2f}")
        ns.label_dict['label_13_5_right'].setPos(x1[-1], y18_3.iloc[-1])
        ns.label_dict['label_13_6_right'].setText(f"{y19_3.iloc[-1]/1000000:.2f}")
        ns.label_dict['label_13_6_right'].setPos(x1[-1], y19_3.iloc[-1])

    if x4 and y0_4:
        # Update Expiry 4 labels
        ns.label_dict['label_plot02_ce'].setText(f"{y18_4.iloc[-1]/1000000:.2f}")
        ns.label_dict['label_plot02_ce'].setPos(x4[-1], y18_4.iloc[-1])
        ns.label_dict['label_plot02_pe'].setText(f"{y19_4.iloc[-1]/1000000:.2f}")
        ns.label_dict['label_plot02_pe'].setPos(x4[-1], y19_4.iloc[-1])

    # Set X range for each column of plots separately based on timestamps
    if x1 and len(x1) > 1:
        # Set X range for Expiry 1 plots (first column)
        ns.plot00.setXRange(x1[0], x1[-1], padding=0.02)
        ns.plot10.setXRange(x1[0], x1[-1], padding=0.02)
        ns.plot20.setXRange(x1[0], x1[-1], padding=0.02)

    if x2 and len(x2) > 1:
        # Set X range for Expiry 2 plots (second column)
        ns.plot01.setXRange(x2[0], x2[-1], padding=0.02)
        ns.plot11.setXRange(x2[0], x2[-1], padding=0.02)
        ns.plot21.setXRange(x2[0], x2[-1], padding=0.02)

    if x3 and len(x3) > 1:
        # Set X range for Expiry 3 plots (third column)
        ns.plot02.setXRange(x3[0], x3[-1], padding=0.02)
        ns.plot12.setXRange(x3[0], x3[-1], padding=0.02)
        ns.plot22.setXRange(x3[0], x3[-1], padding=0.02)

    if x4 and len(x4) > 1:
        # Set X range for Expiry 3 plots (third column)
        ns.win2_plot00.setXRange(x4[0], x4[-1], padding=0.02)
        ns.win2_plot01.setXRange(x4[0], x4[-1], padding=0.02)
        ns.win2_plot02.setXRange(x4[0], x4[-1], padding=0.02)

##########################################################

def check_data(initial_data, current_data, exp_list):
    
    initial_df = pd.DataFrame(initial_data).reset_index(drop=True)
    current_df = pd.DataFrame(current_data).reset_index(drop=True)

    df_concat = pd.concat([initial_df, current_df], axis=1)
    re_order = df_concat.columns.to_list()
    column_index = [0,5,1,6,2,7,3,8,4,9]
    column_index_order = [re_order[i] for i in column_index]
    df_concat = df_concat[column_index_order]
    df_concat.index = ['CE Side LTP', 'PE Side LTP', 'CE Side Theta', 'PE Side Theta', 'CE Side Vega', 'PE Side Vega', 'CE Side IV', 'PE Side IV', 'CE Side OI', 'PE Side OI', 'CE ATM LTP', 'PE ATM LTP', 'ATM Straddle', 'Spot Price', 'India VIX']

    df_concat.columns = ['1_Initial', '1_Current', '2_Initial', '2_Current', '3_Initial', '3_Current', '4_Initial', '4_Current', '5_Initial', '5_Current']

    den_zero = [(df_concat.iloc[8,1] - df_concat.iloc[8,0]), (df_concat.iloc[9,1] - df_concat.iloc[9,0]), (df_concat.iloc[8,3] - df_concat.iloc[8,2]), (df_concat.iloc[9,3] - df_concat.iloc[9,2]), (df_concat.iloc[8,5] - df_concat.iloc[8,4]), (df_concat.iloc[9,5] - df_concat.iloc[9,4]), (df_concat.iloc[8,7] - df_concat.iloc[8,6]), (df_concat.iloc[9,7] - df_concat.iloc[9,6]), (df_concat.iloc[8,9] - df_concat.iloc[8,8]), (df_concat.iloc[9,9] - df_concat.iloc[9,8])]

    if all(val != 0 and pd.notna(val) for val in den_zero): # pd.notna(val) is True if val is not NaN &&&&& False if val is NaN
        ab = round((df_concat.iloc[8,1] - df_concat.iloc[8,0]) / (df_concat.iloc[9,1] - df_concat.iloc[9,0]),2)
        ba = round((df_concat.iloc[9,1] - df_concat.iloc[9,0]) / (df_concat.iloc[8,1] - df_concat.iloc[8,0]),2)

        bc = round((df_concat.iloc[8,3] - df_concat.iloc[8,2]) / (df_concat.iloc[9,3] - df_concat.iloc[9,2]),2)
        cb = round((df_concat.iloc[9,3] - df_concat.iloc[9,2]) / (df_concat.iloc[8,3] - df_concat.iloc[8,2]),2)

        cd = round((df_concat.iloc[8,5] - df_concat.iloc[8,4]) / (df_concat.iloc[9,5] - df_concat.iloc[9,4]),2)
        dc = round((df_concat.iloc[9,5] - df_concat.iloc[9,4]) / (df_concat.iloc[8,5] - df_concat.iloc[8,4]),2)

        de = round((df_concat.iloc[8,7] - df_concat.iloc[8,6]) / (df_concat.iloc[9,7] - df_concat.iloc[9,6]),2)
        ed = round((df_concat.iloc[9,7] - df_concat.iloc[9,6]) / (df_concat.iloc[8,7] - df_concat.iloc[8,6]),2)

        ef = round((df_concat.iloc[8,9] - df_concat.iloc[8,8]) / (df_concat.iloc[9,9] - df_concat.iloc[9,8]),2)
        fe = round((df_concat.iloc[9,9] - df_concat.iloc[9,8]) / (df_concat.iloc[8,9] - df_concat.iloc[8,8]),2)

    else:
        ab=ba=bc=cb=cd=dc=de=ed=ef=fe=None

    df_concat['1_Diff'] = [df_concat.iloc[0,1] - df_concat.iloc[0,0], 
                           df_concat.iloc[1,1] - df_concat.iloc[1,0], 
                           df_concat.iloc[2,0] - df_concat.iloc[2,1], 
                           df_concat.iloc[3,0] - df_concat.iloc[3,1], 
                           df_concat.iloc[4,1] - df_concat.iloc[4,0], 
                           df_concat.iloc[5,1] - df_concat.iloc[5,0], 
                           df_concat.iloc[6,1] - df_concat.iloc[6,0], 
                           df_concat.iloc[7,1] - df_concat.iloc[7,0], 
                           f'{df_concat.iloc[8,1] - df_concat.iloc[8,0]}  ({ab})', 
                           f'{df_concat.iloc[9,1] - df_concat.iloc[9,0]}  ({ba})',
                           df_concat.iloc[10,1] - df_concat.iloc[10,0], 
                           df_concat.iloc[11,1] - df_concat.iloc[11,0],
                           df_concat.iloc[12,1] - df_concat.iloc[12,0], 
                           df_concat.iloc[13,1] - df_concat.iloc[13,0],
                           df_concat.iloc[14,1] - df_concat.iloc[14,0]]


    df_concat['2_Diff'] = [df_concat.iloc[0,3] - df_concat.iloc[0,2], 
                           df_concat.iloc[1,3] - df_concat.iloc[1,2], 
                           df_concat.iloc[2,2] - df_concat.iloc[2,3], 
                           df_concat.iloc[3,2] - df_concat.iloc[3,3], 
                           df_concat.iloc[4,3] - df_concat.iloc[4,2], 
                           df_concat.iloc[5,3] - df_concat.iloc[5,2], 
                           df_concat.iloc[6,3] - df_concat.iloc[6,2], 
                           df_concat.iloc[7,3] - df_concat.iloc[7,2], 
                           f'{df_concat.iloc[8,3] - df_concat.iloc[8,2]}  ({bc})',
                           f'{df_concat.iloc[9,3] - df_concat.iloc[9,2]}  ({cb})',
                           df_concat.iloc[10,3] - df_concat.iloc[10,2], 
                           df_concat.iloc[11,3] - df_concat.iloc[11,2],
                           df_concat.iloc[12,3] - df_concat.iloc[12,2], 
                           df_concat.iloc[13,3] - df_concat.iloc[13,2],
                           df_concat.iloc[14,3] - df_concat.iloc[14,2]]

    df_concat['3_Diff'] = [df_concat.iloc[0,5] - df_concat.iloc[0,4], 
                           df_concat.iloc[1,5] - df_concat.iloc[1,4], 
                           df_concat.iloc[2,4] - df_concat.iloc[2,5], 
                           df_concat.iloc[3,4] - df_concat.iloc[3,5], 
                           df_concat.iloc[4,5] - df_concat.iloc[4,4], 
                           df_concat.iloc[5,5] - df_concat.iloc[5,4], 
                           df_concat.iloc[6,5] - df_concat.iloc[6,4], 
                           df_concat.iloc[7,5] - df_concat.iloc[7,4], 
                           f'{df_concat.iloc[8,5] - df_concat.iloc[8,4]}  ({cd})',
                           f'{df_concat.iloc[9,5] - df_concat.iloc[9,4]}  ({dc})',
                           df_concat.iloc[10,5] - df_concat.iloc[10,4], 
                           df_concat.iloc[11,5] - df_concat.iloc[11,4],
                           df_concat.iloc[12,5] - df_concat.iloc[12,4], 
                           df_concat.iloc[13,5] - df_concat.iloc[13,4],
                           df_concat.iloc[14,5] - df_concat.iloc[14,4]]

    df_concat['4_Diff'] = [df_concat.iloc[0,7] - df_concat.iloc[0,6], 
                           df_concat.iloc[1,7] - df_concat.iloc[1,6], 
                           df_concat.iloc[2,6] - df_concat.iloc[2,7], 
                           df_concat.iloc[3,6] - df_concat.iloc[3,7], 
                           df_concat.iloc[4,7] - df_concat.iloc[4,6], 
                           df_concat.iloc[5,7] - df_concat.iloc[5,6], 
                           df_concat.iloc[6,7] - df_concat.iloc[6,6], 
                           df_concat.iloc[7,7] - df_concat.iloc[7,6], 
                           f'{df_concat.iloc[8,7] - df_concat.iloc[8,6]}  ({de})',
                           f'{df_concat.iloc[9,7] - df_concat.iloc[9,6]}  ({ed})',
                           df_concat.iloc[10,7] - df_concat.iloc[10,6], 
                           df_concat.iloc[11,7] - df_concat.iloc[11,6],
                           df_concat.iloc[12,7] - df_concat.iloc[12,6], 
                           df_concat.iloc[13,7] - df_concat.iloc[13,6],
                           df_concat.iloc[14,7] - df_concat.iloc[14,6]]

    df_concat['5_Diff'] = [df_concat.iloc[0,9] - df_concat.iloc[0,8], 
                           df_concat.iloc[1,9] - df_concat.iloc[1,8], 
                           df_concat.iloc[2,8] - df_concat.iloc[2,9], 
                           df_concat.iloc[3,8] - df_concat.iloc[3,9], 
                           df_concat.iloc[4,9] - df_concat.iloc[4,8], 
                           df_concat.iloc[5,9] - df_concat.iloc[5,8], 
                           df_concat.iloc[6,9] - df_concat.iloc[6,8], 
                           df_concat.iloc[7,9] - df_concat.iloc[7,8], 
                           f'{df_concat.iloc[8,9] - df_concat.iloc[8,8]}  ({ef})',
                           f'{df_concat.iloc[9,9] - df_concat.iloc[9,8]}  ({fe})',
                           df_concat.iloc[10,9] - df_concat.iloc[10,8], 
                           df_concat.iloc[11,9] - df_concat.iloc[11,8],
                           df_concat.iloc[12,9] - df_concat.iloc[12,8], 
                           df_concat.iloc[13,9] - df_concat.iloc[13,8],
                           df_concat.iloc[14,9] - df_concat.iloc[14,8]]


    df_concat = df_concat[['1_Initial', '1_Current', '1_Diff', '2_Initial', '2_Current', '2_Diff', '3_Initial', '3_Current', '3_Diff', '4_Initial', '4_Current', '4_Diff', '5_Initial', '5_Current', '5_Diff']]
    df_concat = df_concat.rename(columns={'1_Diff':exp_list[0], '2_Diff':exp_list[1], '3_Diff':exp_list[2], '4_Diff':exp_list[3], '5_Diff':exp_list[4]})
    return df_concat

counter = 1
last_triggered_minute = None
xyz=0
exe_speed = None
def chain(instrument_key,expiry_date,counter):

        global structure_initial, structure_current, past_data, past_data_obv, initialize, xyz, exe_speed
        
        url1 = 'https://api.upstox.com/v2/option/chain'
        url2 = 'https://api.upstox.com/v2/market-quote/ltp?instrument_key=NSE_INDEX|India VIX'

        params = {
                'instrument_key': instrument_key,
                'expiry_date': expiry_date
        }
        headers = {
            'Accept': 'application/json',
            'Authorization': f'Bearer {access}'
        }

        time.sleep(exe_speed)

        xy = time.time()
        req_per_sec = 1/(xy - xyz)
        prt = f'{round((xy - xyz),2)} : {round((req_per_sec)*60*30)}'
        print(f'{prt}')
        xyz = xy

        
        while True:
            response_options = requests.get(url1, params=params, headers=headers)  # your actual API call
            response_vix = requests.get(url2, headers=headers)
            if response_options.status_code == 200 and response_vix.status_code == 200:
                options = response_options.json()
                vix = response_vix.json()
                if 'data' in options and 'data' in vix:
                    option_df = pd.json_normalize(options['data'])
                    india_vix = vix['data']['NSE_INDEX:India VIX']['last_price']
                    break
                else:
                    print("Response OK but 'data' key missing, retrying...")
            else:
                print(f"HTTP Error {response_options.status_code} : (No Response from Server), retrying...")

            time.sleep(5)  # avoid spamming the server

        # option_df.to_excel('option.xlsx')
        time_stamp = datetime.now().strftime("%H:%M:%S")
        option_df = option_df[['expiry', 'pcr', 'strike_price', 'underlying_spot_price', 'call_options.instrument_key', 'call_options.market_data.ltp', 'call_options.market_data.oi', 'call_options.option_greeks.vega', 'call_options.option_greeks.theta', 'call_options.option_greeks.gamma', 'call_options.option_greeks.delta', 'call_options.option_greeks.iv', 'put_options.instrument_key', 'put_options.market_data.ltp', 'put_options.market_data.oi', 'put_options.option_greeks.vega', 'put_options.option_greeks.theta', 'put_options.option_greeks.gamma', 'put_options.option_greeks.delta', 'put_options.option_greeks.iv', 'call_options.market_data.volume', 'put_options.market_data.volume']]
        option_df = option_df.rename(columns={'call_options.instrument_key' : 'CE_instrument_key', 'call_options.market_data.ltp' : 'CE_ltp', 'call_options.market_data.oi' : 'CE_oi', 'call_options.option_greeks.vega' : 'CE_vega', 'call_options.option_greeks.theta' : 'CE_theta', 'call_options.option_greeks.gamma' : 'CE_gamma', 'call_options.option_greeks.delta' : 'CE_delta', 'call_options.option_greeks.iv' : 'CE_iv', 'put_options.instrument_key' : 'PE_instrument_key', 'put_options.market_data.ltp' : 'PE_ltp', 'put_options.market_data.oi' : 'PE_oi', 'put_options.option_greeks.vega' : 'PE_vega', 'put_options.option_greeks.theta' : 'PE_theta', 'put_options.option_greeks.gamma' : 'PE_gamma', 'put_options.option_greeks.delta' : 'PE_delta', 'put_options.option_greeks.iv' : 'PE_iv', 'underlying_spot_price' : 'spot_price', 'call_options.market_data.volume':'CE_volume', 'put_options.market_data.volume':'PE_volume'})
        option_df = option_df[['expiry','pcr','CE_instrument_key','CE_delta','CE_oi','CE_iv','CE_vega','CE_theta','CE_volume','CE_ltp','strike_price','PE_ltp','PE_volume','PE_theta','PE_vega','PE_iv','PE_oi','PE_delta','PE_instrument_key','spot_price']]

        option_df['diff'] = abs(option_df['spot_price'] - option_df['strike_price'])
        ce = option_df.loc[option_df['diff'].idxmin(),'CE_ltp']
        strike = option_df.loc[option_df['diff'].idxmin(),'strike_price']
        pe = option_df.loc[option_df['diff'].idxmin(),'PE_ltp']

        fut_spot_price = ce-pe+strike

        option_df['spot_price'] = fut_spot_price
        option_df['diff'] = abs(option_df['spot_price'] - option_df['strike_price'])
        option_df['prem_diff'] = option_df['CE_ltp'] - option_df['PE_ltp']
        option_df['CE/PE'] = round((option_df['CE_ltp'] / option_df['PE_ltp']),2)
        atm_strike = option_df.loc[option_df['diff'].idxmin(), 'strike_price']

        ce_atm_ltp = int(option_df[option_df['strike_price'] == atm_strike].iloc[0]['CE_ltp'])
        pe_atm_ltp = int(option_df[option_df['strike_price'] == atm_strike].iloc[0]['PE_ltp'])
        straddle = int(ce_atm_ltp + pe_atm_ltp)

        ce_atm_vol = int(option_df[option_df['strike_price'] == atm_strike].iloc[0]['CE_volume'])
        pe_atm_vol = int(option_df[option_df['strike_price'] == atm_strike].iloc[0]['PE_volume'])
        straddle_volume = int(ce_atm_vol + pe_atm_vol)

        x = option_df['strike_price'].diff().mode()[0]
        upper_limit = atm_strike + 8*x
        lower_limit = atm_strike - 8*x
        option_df = option_df[(option_df['strike_price'] >= lower_limit) & (option_df['strike_price'] <= upper_limit)]

        ce_df = option_df[option_df['strike_price'] >= atm_strike]
        pe_df = option_df[option_df['strike_price'] <= atm_strike]

        ce_ltp_sum = round(ce_df['CE_ltp'].sum(),2)
        pe_ltp_sum = round(pe_df['PE_ltp'].sum(),2)
        ce_theta_sum = round(ce_df['CE_theta'].sum(),2)
        pe_theta_sum = round(pe_df['PE_theta'].sum(),2)
        ce_vega_sum = round(ce_df['CE_vega'].sum(),2)
        pe_vega_sum = round(pe_df['PE_vega'].sum(),2)
        ce_iv_sum = round(ce_df['CE_iv'].sum(),2)
        pe_iv_sum = round(pe_df['PE_iv'].sum(),2)
        ce_oi_sum = round(ce_df['CE_oi'][0:5].sum(),2)
        pe_oi_sum = round(pe_df['PE_oi'][-5:].sum(),2)

        try:
            with open(f'Credentials/Data/{tdate}_initial_values.json', 'r') as file_read:
                structure_initial = json.load(file_read)
        except:
            if counter<=2:

                structure_initial[f'{instrument_key}_{expiry_date}_initial'] = {'ce_ltp_init' : ce_ltp_sum,
                                                                                'pe_ltp_init' : pe_ltp_sum,
                                                                                'ce_theta_init' : ce_theta_sum,
                                                                                'pe_theta_init' : pe_theta_sum,
                                                                                'ce_vega_init' : ce_vega_sum,
                                                                                'pe_vega_init' : pe_vega_sum,
                                                                                'ce_iv_init' : ce_iv_sum,
                                                                                'pe_iv_init' : pe_iv_sum,
                                                                                'ce_oi_init' : ce_oi_sum,
                                                                                'pe_oi_init' : pe_oi_sum,
                                                                                'ce_atm_ltp' : ce_atm_ltp,
                                                                                'pe_atm_ltp' : pe_atm_ltp,
                                                                                'atm_straddle' : (ce_atm_ltp + pe_atm_ltp),
                                                                                'spot price' : fut_spot_price,
                                                                                'india vix' : india_vix
                                                                                }

        structure_current[f'{instrument_key}_{expiry_date}_Current'] = {'ce_ltp_current' : ce_ltp_sum,
                                                                        'pe_ltp_current' : pe_ltp_sum,
                                                                        'ce_theta_current' : ce_theta_sum,
                                                                        'pe_theta_current' : pe_theta_sum,
                                                                        'ce_vega_current' : ce_vega_sum,
                                                                        'pe_vega_current' : pe_vega_sum,
                                                                        'ce_iv_current' : ce_iv_sum,
                                                                        'pe_iv_current' : pe_iv_sum,
                                                                        'ce_oi_current' : ce_oi_sum,
                                                                        'pe_oi_current' : pe_oi_sum,
                                                                        'ce_atm_ltp' : ce_atm_ltp,
                                                                        'pe_atm_ltp' : pe_atm_ltp,
                                                                        'atm_straddle' : (ce_atm_ltp + pe_atm_ltp),
                                                                        'spot price' : fut_spot_price,
                                                                        'india vix' : india_vix
                                                                        }

        ce_ltp_diff = round((ce_ltp_sum - structure_initial[f'{instrument_key}_{expiry_date}_initial']['ce_ltp_init']),2)
        pe_ltp_diff = round((pe_ltp_sum - structure_initial[f'{instrument_key}_{expiry_date}_initial']['pe_ltp_init']),2)
        ce_theta_diff = round((structure_initial[f'{instrument_key}_{expiry_date}_initial']['ce_theta_init'] - ce_theta_sum),2)
        pe_theta_diff = round((structure_initial[f'{instrument_key}_{expiry_date}_initial']['pe_theta_init'] - pe_theta_sum),2)
        ce_vega_diff = round((ce_vega_sum - structure_initial[f'{instrument_key}_{expiry_date}_initial']['ce_vega_init']),2)
        pe_vega_diff = round((pe_vega_sum - structure_initial[f'{instrument_key}_{expiry_date}_initial']['pe_vega_init']),2)
        ce_iv_diff = round((ce_iv_sum - structure_initial[f'{instrument_key}_{expiry_date}_initial']['ce_iv_init']),2)
        pe_iv_diff = round((pe_iv_sum - structure_initial[f'{instrument_key}_{expiry_date}_initial']['pe_iv_init']),2)
        ce_oi_diff = round((ce_oi_sum - structure_initial[f'{instrument_key}_{expiry_date}_initial']['ce_oi_init']),2)
        pe_oi_diff = round((pe_oi_sum - structure_initial[f'{instrument_key}_{expiry_date}_initial']['pe_oi_init']),2)
        ce_atm_diff = round((ce_atm_ltp - structure_initial[f'{instrument_key}_{expiry_date}_initial']['ce_atm_ltp']),2)
        pe_atm_diff = round((pe_atm_ltp - structure_initial[f'{instrument_key}_{expiry_date}_initial']['pe_atm_ltp']),2)
        atm_straddle_diff = round(((ce_atm_ltp + pe_atm_ltp) - structure_initial[f'{instrument_key}_{expiry_date}_initial']['atm_straddle']),2)
        spot_price_diff = round((fut_spot_price - structure_initial[f'{instrument_key}_{expiry_date}_initial']['spot price']),2)
        india_vix_diff = round((india_vix - structure_initial[f'{instrument_key}_{expiry_date}_initial']['india vix']),2)

        main = {'CE Side LTP':ce_ltp_diff, 'PE Side LTP':pe_ltp_diff, 'CE Side Theta':ce_theta_diff, 'PE Side Theta':pe_theta_diff, 'CE Side Vega':ce_vega_diff, 'PE Side Vega':pe_vega_diff, 'CE Side IV':ce_iv_diff, 'PE Side IV':pe_iv_diff, 'CE Side OI':ce_oi_diff, 'PE Side OI':pe_oi_diff, 'CE ATM LTP':ce_atm_diff, 'PE ATM LTP':pe_atm_diff, 'Atm Straddle':atm_straddle_diff, 'Spot Price': spot_price_diff, 'India Vix': india_vix_diff, 'Time': time_stamp}

        expiry_name = option_df.iloc[0,0]

        main_df = pd.DataFrame([main], index=[expiry_name]).T

        try:
            if (instrument_key == instrument_key_nifty) and (expiry_date == expiry_list_nifty[0]):
                with open(f'Credentials/Data/{tdate}_past_data.json', 'r') as file_read:
                    past_data = json.load(file_read)
                with open(f'Credentials/Data/{tdate}_past_data_obv.json', 'r') as file_read:
                    past_data_obv = json.load(file_read)
                initialize=2
        except:
            pass

        if initialize==1:
            past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}'] = {'ce_ltp': [], 'pe_ltp': [], 'ce_theta' : [], 'pe_theta' : [], 'ce_vega' : [], 'pe_vega' : [], 'ce_iv' : [], 'pe_iv' : [], 'ce_oi' : [], 'pe_oi' : [], 'ce_atm' : [], 'pe_atm' : [], 'atm_straddle' : [], 'spot_price':[], 'india_vix':[], 'time' : [], 'strike':[], 'straddle':[], 'straddle_volume':[], 'ce_atm_ltp':[], 'pe_atm_ltp':[], 'ce_atm_vol':[], 'pe_atm_vol':[]}
                  
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['ce_ltp'].append(main_df.iloc[0,0])
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['pe_ltp'].append(main_df.iloc[1,0])
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['ce_theta'].append(main_df.iloc[2,0])
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['pe_theta'].append(main_df.iloc[3,0])
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['ce_vega'].append(main_df.iloc[4,0])
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['pe_vega'].append(main_df.iloc[5,0])
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['ce_iv'].append(main_df.iloc[6,0])
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['pe_iv'].append(main_df.iloc[7,0])
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['ce_oi'].append(main_df.iloc[8,0])
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['pe_oi'].append(main_df.iloc[9,0])
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['ce_atm'].append(main_df.iloc[10,0])
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['pe_atm'].append(main_df.iloc[11,0])
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['atm_straddle'].append(main_df.iloc[12,0])
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['spot_price'].append(main_df.iloc[13,0])
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['india_vix'].append(india_vix)
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['time'].append(main_df.iloc[15,0])
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['strike'].append(atm_strike)
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['straddle'].append(straddle)
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['straddle_volume'].append(straddle_volume)
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['ce_atm_ltp'].append(ce_atm_ltp)
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['pe_atm_ltp'].append(pe_atm_ltp)
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['ce_atm_vol'].append(ce_atm_vol)
        past_data[f'Today : {t_date} | {instrument_key} | Expiry : {expiry_date}']['pe_atm_vol'].append(pe_atm_vol)
        
        return option_df, main_df, expiry_name

def obv(dfp):

    df1 = dfp.iloc[[16]]
    df2 = dfp.iloc[19:,:]
    df = pd.concat([df1, df2])

    exp0 = pd.DataFrame({'strike0':pd.Series(df.iloc[0,0]), 'ce_atm_ltp0':pd.Series(df.iloc[1,0]), 'pe_atm_ltp0':pd.Series(df.iloc[2,0]), 'ce_atm_vol0':pd.Series(df.iloc[3,0]), 'pe_atm_vol0':pd.Series(df.iloc[4,0])})
    exp1 = pd.DataFrame({'strike1':pd.Series(df.iloc[0,1]), 'ce_atm_ltp1':pd.Series(df.iloc[1,1]), 'pe_atm_ltp1':pd.Series(df.iloc[2,1]), 'ce_atm_vol1':pd.Series(df.iloc[3,1]), 'pe_atm_vol1':pd.Series(df.iloc[4,1])})
    exp2 = pd.DataFrame({'strike2':pd.Series(df.iloc[0,2]), 'ce_atm_ltp2':pd.Series(df.iloc[1,2]), 'pe_atm_ltp2':pd.Series(df.iloc[2,2]), 'ce_atm_vol2':pd.Series(df.iloc[3,2]), 'pe_atm_vol2':pd.Series(df.iloc[4,2])})
    exp3 = pd.DataFrame({'strike3':pd.Series(df.iloc[0,3]), 'ce_atm_ltp3':pd.Series(df.iloc[1,3]), 'pe_atm_ltp3':pd.Series(df.iloc[2,3]), 'ce_atm_vol3':pd.Series(df.iloc[3,3]), 'pe_atm_vol3':pd.Series(df.iloc[4,3])})
    exp4 = pd.DataFrame({'strike4':pd.Series(df.iloc[0,4]), 'ce_atm_ltp4':pd.Series(df.iloc[1,4]), 'pe_atm_ltp4':pd.Series(df.iloc[2,4]), 'ce_atm_vol4':pd.Series(df.iloc[3,4]), 'pe_atm_vol4':pd.Series(df.iloc[4,4])})
    df = pd.concat([exp0, exp1, exp2, exp3, exp4], axis=1)

    for i in range(0,5):
        df[f'sstrike_{i}'] = df[f'strike{i}'] != df[f'strike{i}'].shift(1)
        df[f'group_{i}'] = df[f'sstrike_{i}'].cumsum()

        df[f'ce_volume_{i}'] = np.where(df[f'sstrike_{i}']==True, df[f'ce_atm_vol{i}'], df[f'ce_atm_ltp{i}'] - df[f'ce_atm_ltp{i}'].shift(1))
        df[f'ce_volume_{i}'] = np.where(df[f'ce_volume_{i}']>0, df[f'ce_atm_vol{i}'], np.where(df[f'ce_volume_{i}']<0, -df[f'ce_atm_vol{i}'],0))
        df[f'ce_obv_{i}'] = df.groupby(f'group_{i}')[f'ce_volume_{i}'].cumsum()

        df[f'pe_volume_{i}'] = np.where(df[f'sstrike_{i}']==True, df[f'pe_atm_vol{i}'], df[f'pe_atm_ltp{i}'] - df[f'pe_atm_ltp{i}'].shift(1))
        df[f'pe_volume_{i}'] = np.where(df[f'pe_volume_{i}']>0, df[f'pe_atm_vol{i}'], np.where(df[f'pe_volume_{i}']<0, -df[f'pe_atm_vol{i}'],0))
        df[f'pe_obv_{i}'] = df.groupby(f'group_{i}')[f'pe_volume_{i}'].cumsum()


    df = df[['ce_obv_0', 'pe_obv_0', 'ce_obv_1', 'pe_obv_1', 'ce_obv_2', 'pe_obv_2', 'ce_obv_3', 'pe_obv_3', 'ce_obv_4', 'pe_obv_4']]
    
    return df

one=True
xx=None

def vwap(dfp):
    df = dfp.iloc[16:19,:]
    exp0 = pd.DataFrame({'strike0':pd.Series(df.iloc[0,0]), 'straddle0':pd.Series(df.iloc[1,0]), 'volume0':pd.Series(df.iloc[2,0])})
    exp1 = pd.DataFrame({'strike1':pd.Series(df.iloc[0,1]), 'straddle1':pd.Series(df.iloc[1,1]), 'volume1':pd.Series(df.iloc[2,1])})
    exp2 = pd.DataFrame({'strike2':pd.Series(df.iloc[0,2]), 'straddle2':pd.Series(df.iloc[1,2]), 'volume2':pd.Series(df.iloc[2,2])})
    exp3 = pd.DataFrame({'strike3':pd.Series(df.iloc[0,3]), 'straddle3':pd.Series(df.iloc[1,3]), 'volume3':pd.Series(df.iloc[2,3])})
    exp4 = pd.DataFrame({'strike4':pd.Series(df.iloc[0,4]), 'straddle4':pd.Series(df.iloc[1,4]), 'volume4':pd.Series(df.iloc[2,4])})
    df = pd.concat([exp0, exp1, exp2, exp3, exp4], axis=1)
    for i in range(0,5):
        df[f'vol_diff{i}'] = df[f'volume{i}'].diff()
        mask = df[f'strike{i}'] != df[f'strike{i}'].shift(1)
        df.loc[mask, f'vol_diff{i}'] = np.nan

        df[f'pv{i}'] = df[f'straddle{i}'] * df[f'vol_diff{i}']
        df[f'cum_vol{i}'] = df[f'vol_diff{i}'].cumsum()
        df[f'cum_pv{i}'] = df[f'pv{i}'].cumsum()

        df[f'VWAP{i}'] = df[f'cum_pv{i}'] / df[f'cum_vol{i}']

        df.drop([f'vol_diff{i}', f'pv{i}', f'cum_vol{i}', f'cum_pv{i}', f'strike{i}', f'volume{i}'], axis=1, inplace=True)

    dfpp = df.ffill().bfill()

    return dfpp

def call():

    global a,b,c,d,e,one,xx, initialize, exe_speed

    exe_speed = float(summary.range('A40').value)
    screen()

    nifty_0_chain, nifty_0_main_df, expiry_name_0 = chain(instrument_key_nifty,expiry_list_nifty[0],a)
    nifty_1_chain, nifty_1_main_df, expiry_name_1 = chain(instrument_key_nifty,expiry_list_nifty[1],b)
    nifty_3_chain, nifty_3_main_df, expiry_name_2 = chain(instrument_key_nifty,expiry_list_nifty[2],c)
    bnf_0_chain, bnf_0_main_df, expiry_name_3 = chain(instrument_key_bnf,expiry_list_bnf[0],d)
    sensex_0_chain, sensex_0_main_df, expiry_name_4 = chain(instrument_key_sensex,expiry_list_sensex[0],e)


    initialize=2

    exp_list = [expiry_name_0, expiry_name_1, expiry_name_2, expiry_name_3, expiry_name_4]

    # df_concat = check_data(structure_initial,structure_current, exp_list)

    if a==b==c==d==e==3:
        with open(f'Credentials/Data/{tdate}_initial_values.json', 'w') as file_write:
            json.dump(structure_initial, file_write)

    with open(f'Credentials/Data/{tdate}_past_data.json', 'w') as file_write:
        json.dump(past_data, file_write)

    df = pd.DataFrame(past_data)

    for i in range(0,5):
        df.iloc[0,i] = round(pd.Series(df.iloc[0,i]).ewm(span=300, adjust=False).mean(),5).tolist()
        df.iloc[1,i] = round(pd.Series(df.iloc[1,i]).ewm(span=300, adjust=False).mean(),5).tolist()
        df.iloc[2,i] = round(pd.Series(df.iloc[2,i]).ewm(span=300, adjust=False).mean(),5).tolist()
        df.iloc[3,i] = round(pd.Series(df.iloc[3,i]).ewm(span=300, adjust=False).mean(),5).tolist()
        df.iloc[4,i] = round(pd.Series(df.iloc[4,i]).ewm(span=300, adjust=False).mean(),5).tolist()
        df.iloc[5,i] = round(pd.Series(df.iloc[5,i]).ewm(span=300, adjust=False).mean(),5).tolist()
        df.iloc[6,i] = round(pd.Series(df.iloc[6,i]).ewm(span=100, adjust=False).mean(),5).tolist()
        df.iloc[7,i] = round(pd.Series(df.iloc[7,i]).ewm(span=100, adjust=False).mean(),5).tolist()
        df.iloc[8,i] = round(pd.Series(df.iloc[8,i]).ewm(span=100, adjust=False).mean(),5).tolist()
        df.iloc[9,i] = round(pd.Series(df.iloc[9,i]).ewm(span=100, adjust=False).mean(),5).tolist()
        df.iloc[10,i] = round(pd.Series(df.iloc[10,i]).ewm(span=100, adjust=False).mean(),5).tolist()
        df.iloc[11,i] = round(pd.Series(df.iloc[11,i]).ewm(span=100, adjust=False).mean(),5).tolist()
        df.iloc[12,i] = round(pd.Series(df.iloc[12,i]).ewm(span=100, adjust=False).mean(),5).tolist()
        df.iloc[13,i] = round(pd.Series(df.iloc[13,i]).ewm(span=100, adjust=False).mean(),5).tolist()
        df.iloc[14,i] = round(pd.Series(df.iloc[14,i]).ewm(span=50, adjust=False).mean(),5).tolist()

    dfp = df

    expiry_names = dfp.columns.tolist()
    str_curr = pd.DataFrame(structure_current)
    str_init = pd.DataFrame(structure_initial)

    if one:
        xx = one_time(expiry_names)
        one=False

    dfpp = vwap(dfp)    
    obv_df = obv(dfp)

    update(dfp,xx,str_curr,str_init,dfpp,obv_df)

    try:
        summary.range('C2').value = nifty_0_main_df
        summary.range('F2').value = nifty_1_main_df
        summary.range('I2').value = nifty_3_main_df
        summary.range('L2').value = bnf_0_main_df
        summary.range('O2').value = sensex_0_main_df

        # summary.range('A20').value = df_concat

        nifty_0.range('A1').value = nifty_0_chain
        nifty_1.range('A1').value = nifty_1_chain
        nifty_3.range('A1').value = nifty_3_chain
        bnf_0.range('A1').value = bnf_0_chain
        sensex_0.range('A1').value = sensex_0_chain
    except Exception as e:
        print(f'Error Occured while accessing Excel Sheet : {e}')
    
    if a<=3:
        a=a+1
        b=b+1
        c=c+1
        d=d+1
        e=e+1

    exit_graph = summary.range('C43').value

    t_time = datetime.now().time().replace(microsecond=0)

    # print(f'\rCurrent Time : {t_time} | Market Will Close at {end_time}', end='', flush=True)
    print(f'\rCurrent Time : {t_time} | Market Will Close at {end_time}', end='', flush=True)
    
    if (exit_graph=='E') or (t_time > end_time):
        exporter = ImageExporter(win.scene())
        exporter.parameters()['width'] = 1600  # Optional: Set resolution
        exporter.export(f"Credentials/Data/{tdate}_plot_snapshot.jpg")
        print(f"\n\nPlot saved as {tdate}_plot_snapshot.jpg")
        if t_time > end_time:
            print(f'\rMarket Closed at : {end_time}, Current Time : {t_time} | Program Autoclosed', end='', flush=True)
        if exit_graph=='E':
            print(f'\nProgram Closed Manually at : {t_time} from Excel')
        summary.range('C43').value = None
        wb.macro("StopMonitoring")()
        wb.save()
        # if (t_time > end_time):
        wb.close()
        app_analysis.quit()
        time.sleep(1)
        app.quit()

    elif exit_graph==None:
        check_excel_for_full_screen()
        QTimer.singleShot(0, call)


QTimer.singleShot(0, call)


# Start timer
# timer = QTimer()
# timer.timeout.connect(time_fun2)
# timer.start(1000)

# Connect the keyPressEvent function to the window
win.keyPressEvent = keyPressEvent

# Run app
# app.exec_()
sys.exit(app.exec_())