import tkinter as tk
from tkinter import ttk
import pandas as pd
import threading, queue, os, requests
from datetime import date, timedelta
from time import sleep
from pathlib import Path
from subprocess import Popen

import smtplib, ssl
import mimetypes
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

dirname=os.path.dirname(os.path.abspath(__file__))
tday=date.today().strftime('%m-%d-%y') #get today's date
yday=(date.today()-timedelta(days=1)).strftime('%m-%d-%y') # get yesterday's date for use in links

progress = 0

class GUI(tk.Tk):
    
    def __init__(self):
        tk.Tk.__init__(self)
        self.queue = queue.Queue()
        self.title("Daily Rewards")
        # Window geometry
        ws = self.winfo_screenwidth() # width of the screen
        hs = self.winfo_screenheight() # height of the screen
        w = 344 # window width
        h = 400 # window height
        x = int(ws/2 - w/2)
        y = int(hs/2 - h/2)
        self.geometry(f'{w}x{h}+{x}+{y}')
        # Window elements
        self.labelOperator=tk.Label(self, text="Operator: ", font=('Helvetica', 11))
        self.labelOperator.place(x=45, y=10)
        self.operator_text=tk.StringVar()
        self.entryOperator=tk.Entry(self, text=self.operator_text, width=20, font=('Helvetica', 11))
        self.entryOperator.place(x=125, y=12)
        # Radio Buttons
        self.casino=tk.IntVar()
        self.casino.set(1)
        self.rbtn1=tk.Radiobutton(self, text="EXTREME", font=('Helvetica', 11, "bold"), bg='orange red', width=8, variable=self.casino, value=1)
        self.rbtn2=tk.Radiobutton(self, text="BRANGO", font=('Helvetica', 11, "bold"), bg='orange', width=8, variable=self.casino, value=2)
        self.rbtn3=tk.Radiobutton(self, text="YABBY", font=('Helvetica', 11, "bold"), bg='deep sky blue', width=8, variable=self.casino, value=3)
        self.rbtn1.place(x=10, y=45)
        self.rbtn2.place(x=10, y=80)
        self.rbtn3.place(x=10, y=115)
        # Check Buttons
        self.s20=tk.BooleanVar()
        self.s20.set(False)
        self.cbtn1=tk.Checkbutton(self, text="20 Spins", font=('Helvetica', 11, "bold"), variable=self.s20, onvalue=True, offvalue=False)
        self.cbtn1.place(x=120, y=45)
        self.s40=tk.BooleanVar()
        self.s40.set(False)
        self.cbtn2=tk.Checkbutton(self, text="40 Spins", font=('Helvetica', 11, "bold"), variable=self.s40, onvalue=True, offvalue=False)
        self.cbtn2.place(x=230, y=45)
        self.c10=tk.BooleanVar()
        self.c10.set(False)
        self.cbtn3=tk.Checkbutton(self, text="$10 Chips", font=('Helvetica', 11, "bold"), variable=self.c10, onvalue=True, offvalue=False)
        self.cbtn3.place(x=120, y=80)
        self.c20=tk.BooleanVar()
        self.c20.set(False)
        self.cbtn4=tk.Checkbutton(self, text="$20 Chips", font=('Helvetica', 11, "bold"), variable=self.c20, onvalue=True, offvalue=False)
        self.cbtn4.place(x=230, y=80)
        self.c40=tk.BooleanVar()
        self.c40.set(False)
        self.cbtn5=tk.Checkbutton(self, text="$40 Chips", font=('Helvetica', 11, "bold"), variable=self.c40, onvalue=True, offvalue=False)
        self.cbtn5.place(x=120, y=115)
        self.c50=tk.BooleanVar()
        self.c50.set(False)
        self.cbtn6=tk.Checkbutton(self, text="$50 Chips", font=('Helvetica', 11, "bold"), variable=self.c50, onvalue=True, offvalue=False)
        self.cbtn6.place(x=230, y=115)
        # Buttons
        self.btn1=tk.Button(self, text="Get Info", width = 8, font=('Helvetica', 9), command=self.getdeps)
        self.btn1.place(x=115, y=365)
        self.btn2=tk.Button(self, text="Rewards", width = 8, font=('Helvetica', 9), state='disabled', command=self.processrewards)
        self.btn2.place(x=185, y=365)
        self.btn3=tk.Button(self, text="Reset", width = 8, font=('Helvetica', 9), command=self.reset)
        self.btn3.place(x=265, y=365)
        # Listbox
        self.printer=tk.Listbox(self, width=53, height=10)
        self.printer.place(x=10, y=150)
        # ProgressBar
        self.pbar=ttk.Progressbar(self, orient='horizontal', length = 323, mode='determinate')
        self.pbar.place(x=10, y=325)

    def getdeps(self):
        # Disable Form
        self.btn1.config(state="disabled")
        self.btn3.config(state="disabled")
        self.rbtn1.config(state="disabled")
        self.rbtn2.config(state="disabled")
        self.rbtn3.config(state="disabled")
        # Start Thread
        self.thread=GetDepositors(self.queue, self.casino.get())
        self.thread.start()
        self.periodiccall1()

    def periodiccall1(self):
        self.checkqueue()
        if self.thread.is_alive():
            self.after(100, self.periodiccall1)
        else:
            self.btn2.config(state="active")
            self.btn3.config(state="active")

    def processrewards(self):
        # Disable Form
        self.btn2.config(state="disabled")
        self.btn3.config(state="disabled")
        self.cbtn1.config(state="disabled")
        self.cbtn2.config(state="disabled")
        self.cbtn3.config(state="disabled")
        self.cbtn4.config(state="disabled")
        self.cbtn5.config(state="disabled")
        self.cbtn6.config(state="disabled")
        # Start Thread
        self.dep_range = []
        if self.s20.get(): self.dep_range.extend(range(10,20))
        if self.s40.get(): self.dep_range.extend(range(20,50))
        if self.c10.get(): self.dep_range.extend(range(50,70))
        if self.c20.get(): self.dep_range.extend(range(70,100))
        if self.c40.get(): self.dep_range.extend(range(100,150))
        if self.c50.get(): self.dep_range.extend(range(150,500))
        self.thread=ProcessRewards(self.queue, self.casino.get(), self.dep_range)
        self.thread.start()
        self.periodiccall2()

    def periodiccall2(self):
        self.checkqueue()
        if self.thread.is_alive():
            self.after(100, self.periodiccall2)
        else:
            self.btn1.config(state="active")
            self.btn2.config(state="disabled")
            self.btn3.config(state="active")
            self.cbtn1.config(state="active")
            self.cbtn2.config(state="active")
            self.cbtn3.config(state="active")
            self.cbtn4.config(state="active")
            self.cbtn5.config(state="active")
            self.cbtn6.config(state="active")

    def checkqueue(self):
        while self.queue.qsize():
            # Get messages
            msg = self.queue.get(0)
            self.printer.insert(0, msg)
            self.printer.index(0)
            # Move Progressbar
            global progress
            self.pbar.step(progress)
            if progress > 0: progress = 0
    
    def reset(self):
        self.btn1.config(state="active")
        self.btn2.config(state="disabled")
        self.btn3.config(state="active")
        self.cbtn1.config(state="active")
        self.cbtn2.config(state="active")
        self.cbtn3.config(state="active")
        self.cbtn4.config(state="active")
        self.cbtn5.config(state="active")
        self.cbtn6.config(state="active")
        self.rbtn1.config(state="active", bg='orange red')
        self.rbtn2.config(state="active", bg='orange')
        self.rbtn3.config(state="active", bg='deep sky blue')
        self.printer.delete(0, tk.END)


class GetDepositors(threading.Thread):

    def __init__(self, queue, casino):
        threading.Thread.__init__(self)
        self.queue = queue
        self.queue.put('Get Depositors Module - Activated')
        self.queue.put('==========================')
        self.report_df = pd.DataFrame(columns=['status','pid','login','email','firstname','deposits','withdrawals','eligible','reward','skip', 'processed'])
        Path(f"Reports/{date.today().strftime('%m.%d.%y')}").mkdir(parents=True,exist_ok=True) # Create Report Folder
        # Casino Variables
        if casino == 1:
            self.backend="https://admin.casinoextreme.eu/ALEKSMENWNQMDKTOIDJG/RTGWebAPI/api/"
            self.certificate=os.path.join(dirname, 'certificates', 'casextreme.pem')
            self.casinoname="Extreme"
            self.csv_path=os.path.join(dirname, 'Reports', date.today().strftime('%m.%d.%y'), 'extreme.csv')
        elif casino == 2:
            self.backend="https://admin.casinobrango.com/BRNG2USDZDPYTSZWZTQP/RTGWebAPI/api/"
            self.certificate=os.path.join(dirname, 'certificates', 'usdbrng2.pem')
            self.casinoname="Brango"
            self.csv_path=os.path.join(dirname, 'Reports', date.today().strftime('%m.%d.%y'), 'brango.csv')
        elif casino == 3:
            self.backend="https://admin.yabbycasino.com/YABBYECVSUGMOQMOIPQO/RTGWebAPI/api/"
            self.certificate=os.path.join(dirname, 'certificates', 'mccyabby.pem')
            self.casinoname="Yabby"
            self.csv_path=os.path.join(dirname, 'Reports', date.today().strftime('%m.%d.%y'), 'yabby.csv')

    def getrequest(self, call="", params={}, attempts=3):
        k = 0
        r = requests.get(url=self.backend+call, params=params, cert=self.certificate)
        while True:
            if r.status_code == 200: return r.json()
            if k == attempts: 
                self.status = '0. FAILED'
                self.queue.put(f'[{self.count}/{self.total}] Failed: {self.login}.')
                self.failed += 1
                self.skip = 'x'
                return False
            k+=1
            sleep(3)
            r = requests.get(url=self.backend+call, params=params, cert=self.certificate)

    def run(self):
        while os.path.exists(self.csv_path): # delete csv file if exists
            try:
                os.remove(self.csv_path)
            except:
                self.queue.put('Please close Excel Report...')
                sleep(10)
                continue
        
        self.count = 0
        self.total = 0
        self.failed = 0
        self.login = "Getting depositors!"
        q = self.getrequest(call="reports/depositors", params={"startDate": f"{yday} 00:00", "endDate":f"{yday} 23:59"})
        if q != False:
            self.total = len(q)
            self.count = 0
            for json in q:
                self.count+=1
                self.details(json)
        else:
            self.queue.put('Failed getting depositors. Check your internet connection.')
        
        self.report_df.sort_values(by=['status', 'deposits'], ascending=[True, False], ignore_index=True, inplace=True)
        self.report_df.to_csv(path_or_buf=self.csv_path, index=False)

        self.queue.put(f'==========================')
        if self.failed > 0: 
            self.queue.put('Same procedure as with Manuals.')
            self.queue.put(f'{self.failed} CORRECTIONS NEEDED!')
        self.queue.put('To reward and email them manually, leave "x" signs.')
        self.queue.put('3) save then close the file and click "Rewards" button.')
        self.queue.put('2) remove "x" signs from their "skip" column,')
        self.queue.put('1) open the report file, add the codes you want give them,')
        self.queue.put('To automatically process rewards for Manual customers:')
        self.queue.put(f'==========================')
        self.queue.put(f'Get Depositors Module - Finished')
        self.queue.put(f'==========================')
        Popen(self.csv_path, shell=True)

    def details(self, json):
        self.login=json['login'].strip()
        email=json['email'].strip()
        firstname=json['first_name'].strip().title()
        self.status='3. Eligible'
        self.skip=''
        eligible = True

        q = self.getrequest('accounts/playerid', {'login':self.login}) # get Player ID
        if q != False:
            pid = q.strip()
            
        q = self.getrequest(f'players/{pid}/balance-summary', {'startDate':f"{yday} 00:00",'endDate':f"{tday} 23:59"}) # get withdrawals sum
        if q != False:
            withdrawals = q[0]['real_withdrawals_amount']
                
        q = self.getrequest('cashier/cashback',{'login':self.login,'startDate':f"{yday} 00:00",'endDate':f'{yday} 23:59'}) # get deposits sum and pending withdrawal
        if q != False:
            if q['has_pending_withdrawal'] == True: withdrawals+=1
            deposits = q['sum_all_deposits'] + q['sum_bonus_deposits']
            # Determine Reward
            if deposits < 10:
                reward = ""
                self.status = "5. Ineligible"
                eligible = False
                self.skip = "x"
            elif deposits < 20:
                reward = "DAILY-S!20"
            elif deposits < 50:
                reward = "DAILY-S!40"
            elif deposits < 70:
                reward = "DAILY-C!10"
            elif deposits < 100:
                reward = "DAILY-C!20"
            elif deposits < 150:
                reward = "DAILY-C!40"
            elif deposits < 500:
                reward = "DAILY-C!50"
            else: 
                reward = ""
                self.status = "1. Manual"
                self.skip = "x"

        if eligible: 
            if json['account_status']=='ACTIVE' and json['ban_status']=='NOT BANNED' and withdrawals == 0: # is deactivated or banned, has withdrawals 
                
                excluded = open(os.path.join(dirname,"excluded.txt")).read().splitlines() # if still eligible, check excluded customers
                for x in excluded:
                    if self.login.casefold() == x.replace(' ', '').casefold():
                        self.status = '5. Ineligible'
                        eligible = False
                        self.skip = 'x'
                        break
            else: 
                self.status = '5. Ineligible'
                eligible = False
                self.skip = 'x'

        if self.status != '0. FAILED': # if no calls failed, print success
            self.queue.put(f'[{self.count}/{self.total}] Success: {self.login}.')
        
        # move progress bar and append info
        global progress
        progress+=100/self.total
        self.report_df=self.report_df.append(ignore_index=True,other={'status':self.status, 'pid':pid, 'login':self.login, 'email':email, 'firstname':firstname, 'deposits':deposits, 'withdrawals':withdrawals, 'eligible':eligible, 'reward':reward, 'skip':self.skip, 'processed':False})


class ProcessRewards(threading.Thread):

    def __init__(self, queue, casino, deposit_range):
        threading.Thread.__init__(self)
        self.queue = queue
        self.deposit_range = deposit_range
        self.queue.put('Process Rewards Module - Activated')
        self.queue.put('==========================')
        self.casino = casino
        # Casino Variables
        if self.casino == 1:
            self.backend="https://admin.casinoextreme.eu/ALEKSMENWNQMDKTOIDJG/RTGWebAPI/api/"
            self.certificate=os.path.join(dirname, 'certificates', 'casextreme.pem')
            self.casinoname="Extreme"
            self.csv_path=os.path.join(dirname, 'Reports', date.today().strftime('%m.%d.%y'), 'extreme.csv')
        elif self.casino == 2:
            self.backend="https://admin.casinobrango.com/BRNG2USDZDPYTSZWZTQP/RTGWebAPI/api/"
            self.certificate=os.path.join(dirname, 'certificates', 'usdbrng2.pem')
            self.casinoname="Brango"
            self.csv_path=os.path.join(dirname, 'Reports', date.today().strftime('%m.%d.%y'), 'brango.csv')
        elif self.casino == 3:
            self.backend="https://admin.yabbycasino.com/YABBYECVSUGMOQMOIPQO/RTGWebAPI/api/"
            self.certificate=os.path.join(dirname, 'certificates', 'mccyabby.pem')
            self.casinoname="Yabby"
            self.csv_path=os.path.join(dirname, 'Reports', date.today().strftime('%m.%d.%y'), 'yabby.csv')
        self.df = pd.read_csv(self.csv_path)
        self.total = len(self.df.index)
        self.failed = 0
    
    def run(self):
        for self.index, row in self.df.iterrows():

            self.count = self.index + 1 # get count
            # get details from row
            self.login = row['login']
            self.pid = row['pid']
            self.eligible = row['eligible']
            self.status = row['status']
            self.coupon = row['reward']
            self.email = row['email']
            self.firstname = row['firstname']
            deposits = row['deposits']

            global progress
            progress += 100/self.total

            if row['skip']=='x': 
                self.queue.put(f'[{self.count}/{self.total}] Skipped: {self.login}.')
                self.df.at[self.index, 'processed'] = False # note that no reward was processed
                continue
            if row['status']=='5. Ineligible' or self.eligible == False:
                self.queue.put(f'[{self.count}/{self.total}] Ineligible: {self.login}.')
                self.df.at[self.index, 'status'] = '5. Ineligible'
                self.df.at[self.index, 'processed'] = False # note that no reward was processed
                continue
            if deposits <= 500 and not(int(deposits) in self.deposit_range): # if reward is not being processed today
                self.queue.put(f'[{self.count}/{self.total}] Ineligible: {self.login}.')
                self.df.at[self.index, 'status'] = '4. Skipped'
                self.df.at[self.index, 'processed'] = False # note that no reward was processed
                self.df.at[self.index, 'skip'] = 'x'
                continue
            
            q = self.getrequest(f'players/{self.pid}/balance',{'forMoney':'True'})
            if q != False: # Get Balance
                self.balance = q[0]['balance']

            if self.balance < 1: # Process Coupon
                
                q = requests.post(url=self.backend+f'v2/players/{self.pid}/coupons',params={'couponCode':self.coupon},cert=self.certificate)
                if q.status_code == 200:
                    self.df.at[self.index, 'status'] = '2. Processed'
                    self.df.at[self.index, 'processed'] = True
                    send_email(self.casino, 1, self.email, self.firstname) # send processed email
                    self.queue.put(f'[{self.count}/{self.total}] Processed Reward: {self.login}.')
                    continue
                elif q.status_code == 409:
                    if q.json()['Status'] == 'previous_coupon_pending':
                        self.put_comment()
                        continue
                    elif q.json()['Status'] == 'player_excluded_from_redeeming_all_coupons':
                        self.df.at[self.index, 'status'] = '5. Ineligible'
                        self.df.at[self.index, 'eligible'] = False
                        self.df.at[self.index, 'processed'] = False
                        self.df.at[self.index, 'skip'] = 'x'
                        self.queue.put(f'[{self.count}/{self.total}] Ineligible: {self.login}.')
                        continue
                    else: self.put_comment() # for other 409 statuses
                else: self.put_comment() # for other response codes
            else: self.put_comment() # active balance

        self.df.sort_values(by=['deposits'], ignore_index=True, ascending=False, inplace=True)
        
        while os.path.exists(self.csv_path): # delete csv file if exists
            try:
                os.remove(self.csv_path)
            except:
                self.queue.put('Please close Excel Report so I can email it!')
                sleep(10)
                continue

        self.df.to_csv(path_or_buf=self.csv_path, index=False)
        send_email(self.casino, 3, "risk@casinoextreme.com", "Risk", attach=self.csv_path)
        self.queue.put('==========================')
        self.queue.put('Process Rewards Module - Finished')
        self.queue.put('==========================')

    def getrequest(self, call="", params={}, attempts=3):
        k = 0
        r = requests.get(url=self.backend+call, params=params, cert=self.certificate)
        while True:
            if r.status_code == 200: return r.json()
            if k == attempts: 
                self.df.at[self.index, 'status'] = '0. FAILED'
                self.df.at[self.index, 'processed'] = False
                self.queue.put(f'[{self.count}/{self.total}] Failed Reward: {self.login}.')
                self.failed += 1
                self.df.at[self.index, 'skip'] = 'x'
                return False
            k+=1
            sleep(3)
            r = requests.get(url=self.backend+call, params=params, cert=self.certificate)

    def put_comment(self):
        k=0
        r = requests.put(url=self.backend+f'players/{self.pid}/comp-points',data={'comment':f'<a href="couponPlayerRedeem.asp?PID={self.pid}&amp;couponCode={self.coupon}&amp;show=1">REDEEM {self.coupon}</a> (<b>{tday}</b> Daily Reward)'},cert=self.certificate)
        while True:
            if r.status_code == 200:
                self.df.at[self.index, 'status'] = '3. Pending'
                self.df.at[self.index, 'processed'] = True
                self.queue.put(f'[{self.count}/{self.total}] Pending Reward: {self.login}.')
                break
            if k == 3:
                self.failed += 1
                self.df.at[self.index, 'status'] = '3. Pending'
                self.df.at[self.index, 'processed'] = False
                self.queue.put(f'[{self.count}/{self.total}] Failed comment for {self.login}. Please add manually.')
                break
            k+=1
            sleep(3)
            r = requests.put(url=self.backend+f'players/{self.pid}/comp-points',data={'comment':f'<a href="couponPlayerRedeem.asp?PID={self.pid}&amp;couponCode={self.coupon}&amp;show=1">REDEEM {self.coupon}</a> (<b>{tday}</b> Daily Reward)'})
        
        send_email(self.casino, 2, self.email, self.firstname) # send pending email


def send_email(casino, mail, recepient_email, recepient_name, attach=""):
    if casino == 1:
        smtp_server = "mail.casinoextreme.eu"
        sender_email = "promotions@casinoextreme.com"
        support_email = "support@casinoextreme.com"
        password = "WtHttBcf1ySs"
        subject = f"{recepient_name.title()}, your Loyalty Reward is ready!"
        signed = "Anna"
        cas_text = "Casino Extreme"
        cas_link = "www.casinoextreme.eu"
        phone = "1-800-532-4561"
        filename = "extreme.csv" # for attachement part
    elif casino == 2:
        smtp_server = "mail.casinobrango.com"
        sender_email = "promotions@casinobrango.com"
        support_email = "support@casinobrango.com"
        password = "h0vgKNfeOFvX"
        subject = f"{recepient_name.title()}, your Loyalty Reward is ready!"
        signed = "Anna"
        cas_text = "Casino Brango"
        cas_link = "www.casinobrango.com"
        phone = "1-800-532-4561"
        filename = "brango.csv" # for attachement part
    elif casino == 3:
        smtp_server = "mail.yabbycasino.com"
        sender_email = "promotions@yabbycasino.com"
        support_email = "support@yabbycasino.com"
        password = "TZPjAzhRyTdZ"
        subject = f"{recepient_name.title()}, you got a Loyalty Reward! Congrats!"
        signed = "Ethan"
        cas_text = "Yabby Casino"
        cas_link = "www.yabbycasino.com"
        phone = "1-800-876-3456"
        filename = "yabby.csv" # for attachement part

    if mail == 1: ### REWARD MESSAGE CONTENT
        text = """\
        Dear %s,
        Hope you are well.
        Thank you for your patronage. A loyalty reward has been added to your account, but you can only use it today! We hope it brings you luck!
        Please let us know if you need any assistance.
        Sincerely,
        %s
        %s
        24/7 Live Support
        Chat: %s
        E-mail: %s
        Phone: %s""" % (recepient_name, signed, cas_text, cas_link, support_email, phone,)
        html="""\
        <html>
            <body>
                <p>
                    Dear %s,
                    <br><br>Hope you are well.
                    <br><br>Thank you for your patronage. <b>A loyalty reward has been added to your account</b>, but you can only use it today! We hope it brings you luck!
                    <br><br>Please let us know if you need any assistance.
                    <br><br>Sincerely,
                    <br>%s
                    <br><b>%s</b>
                    <br><br><i>24/7 Live Support
                    <br>Chat: %s
                    <br>E-mail: %s 
                    <br>Phone: %s</i>
                </p>
            </body>
        </html>
        """ % (recepient_name, signed, cas_text, cas_link, support_email, phone,)
    elif mail == 2: ### PENDING MESSAGE CONTENT
        text = """\
        Dear %s,
        Hope you are well.
        Thank you for your patronage. A loyalty reward has been prepared for your account, but you can only use it today! We didn't add it just yet so as not to disturb your current play.
        Once you are done, be sure contact our friendly Customer Service team, and they will have it added to your account. We hope it will prove lucky!
        Please let us know if you need any assistance.
        Sincerely,
        %s
        %s
        24/7 Live Support
        Chat: %s
        E-mail: %s
        Phone: %s""" % (recepient_name, signed, cas_text, cas_link, support_email, phone,)
        html="""\
        <html>
            <body>
                <p>
                    Dear %s,
                    <br><br>Hope you are well.
                    <br><br>Thank you for your patronage. <b>A loyalty reward has been prepared for your account</b>, but you can only use it today! We didn't add it just yet so as not to disturb your current play.
                    <br><br>Once you are done, be sure contact our friendly Customer Service team, and they will have it added to your account. We hope it will prove lucky!
                    <br><br>Please let us know if you need any assistance.
                    <br><br>Sincerely,
                    <br>%s
                    <br><b>%s</b>
                    <br><br><i>24/7 Live Support
                    <br>Chat: %s
                    <br>E-mail: %s 
                    <br>Phone: %s</i>
                </p>
            </body>
        </html>
        """ % (recepient_name, signed, cas_text, cas_link, support_email, phone,)
    elif mail == 3: ### RERPORT MESSAGE CONTENT
        subject = f"Notice: {cas_text} Daily Rewards Report"
        recepient_email = "risk@casinoextreme.com"
        text = """\
        Hello!

        Please find the report of today's Daily Rewards in the attached CSV file.

        Kind Regards,
        Stefan 
        """
    
    if attach == "": # customer emails
        message = MIMEMultipart("alternative")
        message["Subject"] = subject
        message["From"] = f"{cas_text} <{sender_email}>"
        message["To"] = recepient_email
        part1 = MIMEText(text, "plain")
        part2 = MIMEText(html, "html")
        message.attach(part1)
        message.attach(part2)
    else: # attachment mail
        message = MIMEMultipart()
        message["From"] = f"{cas_text} <{sender_email}>"
        message["To"] = recepient_email
        message["Subject"] = subject
        message.premable = subject

        part1 = MIMEText(text, "plain")
        message.attach(part1)

        ctype, encoding = mimetypes.guess_type(attach)
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"

        maintype, subtype = ctype.split("/", 1)

        if maintype == "text":
            fp = open(attach)
            attachment = MIMEText(fp.read(), _subtype=subtype)
            fp.close()
        else:
            fp = open(attach, "rb")
            attachment = MIMEBase(maintype, subtype)
            attachment.set_payload(fp.read())
            fp.close()
            encoders.encode_base64(attachment)
        attachment.add_header("Content-Disposition", "attachment", filename=filename)
        message.attach(attachment)

    port = 465
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(host=smtp_server,port=port,context=context) as s:
        s.login(sender_email, password)
        s.sendmail(from_addr=sender_email, to_addrs=recepient_email, msg=message.as_string())

app = GUI()
app.mainloop()