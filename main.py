
'''
Arrival Report Scraper
Aug 2020

'''

import requests
from bs4 import BeautifulSoup
from datetime import datetime
import xlsxwriter
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkcalendar import Calendar, DateEntry
import os
#import subprocess
from shutil import rmtree
from sys import exit

class Res:
  def __init__(self, name='', boatname='', boattype='', loa='', arrival='', \
    departure='', nights='', assign='', power='', notes='', specreq=''):
      self.name = name
      self.boatname = boatname
      self.boattype = boattype
      self.loa = loa
      self.arrival = arrival
      self.departure = departure
      self.nights = nights
      self.assign = assign
      self.power = power
      self.notes = notes
      self.specreq=specreq

  def __repr__(self):
    return ((f'{self.__class__.__name__}('f'{self.name!r}, {self.boatname!r}, '
     f'{self.boattype!r}, {self.loa!r}, {self.arrival!r}, {self.departure!r}, '
     f'{self.nights!r}, {self.assign!r}, {self.power!r}, {self.notes!r},'
     f'{self.speqreq!r})'))

#GET DATE AND PASSWORD                                                          |
# Dialogue Box: get date, password, sorting method
password = ''
date = datetime.date(datetime.now())
sortby = 'LOA'

class Pop_up(tk.Frame):
  def __init__(self, master):
    super().__init__(master)
    self.pack()
    
    self.label = Label(self ,text = "temp")
    self.label.grid(row=0,column=0)
    self.labeltext = StringVar()
    self.label["textvariable"] = self.labeltext

    # Exit conditions
    self.btn = ttk.Button(self ,text="Quit", command=self.master.destroy)
    self.btn.grid(row=1,column=0)

class App(tk.Frame):
  def __init__(self, master):
    super().__init__(master)
    self.pack()

    self.clicked = 0
    self.label1 = Label(self ,text = "Password:").grid(row = 2,column = 0)
    self.label2 = Label(self ,text = "Date:").grid(row = 0,column = 0)
    self.lable3 = Label(self ,text = "Sort By:").grid(row = 1,column = 0)
    self.entry1 = tk.Entry(self)
    self.entry1.grid(row = 2,column = 1)
    self.entry1.config(show="*")
    self.entry1.focus()
    self.entry2 = Calendar(self, firstweekday='sunday', 
                           date_pattern='yyyy-mm-dd')
    self.entry2.grid(row = 0,column = 1)
    self.entry3 = ttk.Combobox(self, values=['LOA', 'Boat Name'])
    self.entry3.grid(row = 1,column = 1)
    # Create the application variable.
    self.cont1 = tk.StringVar()
    self.cont2 = tk.StringVar()
    self.cont3 = tk.StringVar()
    # Tell the entry widget to watch this variable.
    self.entry1["textvariable"] = self.cont1
    self.entry2["textvariable"] = self.cont2
    self.entry3["textvariable"] = self.cont3
    
    # Exit conditions
    self.btn = ttk.Button(self ,text="Create Report", 
     command=self.create_report_btn).grid(row=3,column=1)
    self.entry1.bind('<Key-Return>',
                    self.create_report)
    self.entry2.bind('<Key-Return>',
                    self.create_report)
    self.entry3.bind('<Key-Return>',
                    self.create_report)
  def create_report(self, event):
    self.clicked = 1
    self.master.destroy()
  def create_report_btn(self):
    self.clicked = 1
    self.master.destroy()

def create_error():
    errorroot = tk.Tk()
    errorform = Pop_up(errorroot)
    errorroot.geometry(f'200x50+{int(errorroot.winfo_screenwidth()/2)-100}'
                       f'+{int(errorroot.winfo_screenheight()/2)-50}')
    return errorform

while(1):
  root = tk.Tk()
  reportform = App(root)
  root.geometry(f'350x300+{int(root.winfo_screenwidth()/2)-175}'
                f'+{int(root.winfo_screenheight()/2)-150}')
  reportform.cont1.set(password)
  reportform.cont2.set(date)
  reportform.cont3.set(sortby)
  reportform.mainloop()
  
  if not reportform.clicked: sys.exit()
  password = reportform.cont1.get()
  date = reportform.cont2.get()
  sortby = reportform.cont3.get()
  
  #GET RESERVATION INFO
  # Login to session
  URL = 'https://dockwa.com'
  LOGIN_ROUTE = '/users/sign_in'
  with requests.Session() as sess:
    try:
      token = (BeautifulSoup(sess.get(URL + LOGIN_ROUTE).text, 'html.parser').find
       ('input', attrs={'name':'authenticity_token', 'type':'hidden'})['value'])
    except:
      temp = create_error()
      temp.labeltext.set("Connection Failed")
      temp.mainloop()
      continue
    
    login_payload = {
      'user[email]': 'mbroman@deerharbormarina.com',
      'user[password]': password,
      'authenticity_token': token
    }
    sess.post(URL + LOGIN_ROUTE,data=login_payload)
      
    #Create list of reservations
    reservations = []
    resnum = 0
    numpage = 0
    req = (sess.get(f'{URL}/manage/ywcgqw-deer-harbor-marina/'
      f'reservations/night_ajax/{date}?arrival=true&page={str(numpage)}'))
    if not req.text:
      temp = create_error()
      temp.labeltext.set("No arrivals. Slow day?")
      temp.mainloop()
      continue
      
    #Multiple pages potentially, 10 rows per page
    while numpage<6:
      soup = BeautifulSoup(req._content.decode('utf-8'), 'html.parser')
      if soup.title:
        temp = create_error()
        temp.labeltext.set("Wrong Password")
        temp.mainloop()
        numpage = -1
        break
      rows = soup.find_all('tr')
  
      #Create reservation instances
      for eachrow in rows:
        newres = Res()
        if eachrow.has_attr('data-reservation-note'):
          newres.notes = eachrow['data-reservation-note']
        if eachrow.has_attr('data-special-request'):
          newres.specreq = eachrow['data-special-request']
        stuff = eachrow.find_all('td')
        if stuff[1].find('div', attrs={'class':'visible-print'}):
          newres.assign = stuff[1].find('div', attrs={'class':'visible-print'}).text
        newres.name = stuff[2].text
        newres.boatname = stuff[3].text
        newres.boattype = stuff[4].text
        newres.loa = stuff[5].text
        newres.arrival = stuff[8].text
        newres.departure = stuff[9].text
        newres.nights = stuff[10].text
        newres.power = stuff[11].text
        reservations.append(newres)
  
      #Get next page and check if empty
      numpage+=1
      req = (sess.get(f'{URL}/manage/ywcgqw-deer-harbor-marina/reservations/'
        f'night_ajax/{date}?arrival=true&page={str(numpage)}'))
      if not req.text:
        #Infinite blocker, max 60 reservation assumed
        numpage = 6
    if numpage == -1:
      continue
    
    if sortby == 'LOA':
      reservations.sort(key=lambda x: x.loa, reverse=True)
    elif sortby == 'Boat Name':
      reservations.sort(key=lambda x: x.boatname)
    else:
      temp = create_error()
      temp.labeltext.set(f'"{sortby}" sorting method doesn\'t exist.')
      temp.mainloop()
      continue
    #print(reservations)
    #print(len(reservations))
  
  #CREATE EXCEL SPREADSHEET
  # Create a workbook and add a worksheet.
  if not os.path.exists(f'{datetime.now().date()}_Reports'):
    os.makedirs(f'{datetime.now().date()}_Reports')
  workpath = f'{datetime.now().date()}_Reports\ArrivalReport{date}.xlsx'
  if os.path.exists(workpath):
    dupnum = 1
    while os.path.exists(f'{datetime.now().date()}_Reports'
                         f'\ArrivalReport{date}-D{dupnum}.xlsx'):
      dupnum += 1
    workpath = (f'{datetime.now().date()}_Reports'
                f'\ArrivalReport{date}-D{dupnum}.xlsx')
  workbook = xlsxwriter.Workbook(workpath)
  worksheet = workbook.add_worksheet()
  
  # Add Formats
  lite_basic = workbook.add_format({'valign':'vcenter', 'font_name':'Arial',
    'text_wrap':1})
  dark_basic = workbook.add_format({'valign':'vcenter', 'font_name':'Arial',
    'text_wrap':1, 'bg_color':'EEEEFF'})
  title = workbook.add_format({'align':'center', 'valign':'vcenter', 
    'font_name':'Arial', 'bold':1})
  lite_center = workbook.add_format({'align':'center', 'valign':'vcenter', 
    'font_name':'Arial', 'text_wrap':1})
  dark_center = workbook.add_format({'align':'center', 'valign':'vcenter', 
    'font_name':'Arial', 'text_wrap':1, 'bg_color':'EEEEFF'})
  lite_slip = workbook.add_format({'align':'center', 'valign':'vcenter', 
    'font_name':'Arial', 'text_wrap':1, 'left':1, 'right':1})
  dark_slip = workbook.add_format({'align':'center', 'valign':'vcenter', 
    'font_name':'Arial', 'text_wrap':1, 'bg_color':'EEEEFF', 'left':1, 'right':1})
  lite_spec = workbook.add_format({'valign':'vcenter', 'font_name':'Arial'})
  dark_spec = workbook.add_format({'valign':'vcenter', 'font_name':'Arial',
    'bg_color':'EEEEFF'})
  filler = workbook.add_format({'font_name':'Arial', 'font_size':17,
    'font_color':'FFFFFF'})
  
  # Adjust the column width.
  worksheet.set_column(1, 1, 26)
  worksheet.set_column(2, 2, 19)
  worksheet.set_column(3, 3, 4.9)
  worksheet.set_column(4, 4, 4.3)
  worksheet.set_column(5, 6, 6.8)
  worksheet.set_column(7, 7, 4.6)
  worksheet.set_column(8, 8, 44)
  worksheet.set_column(9, 9, 44)
  
  # Starting row
  row = 2
  
  # Create Header
  worksheet.merge_range(0, 1, 0, 2, f'ARRIVAL REPORT FOR {date}',title)
  worksheet.write('B2', 'Name', title)
  worksheet.write('C2', 'Boat Name', title)
  worksheet.write('D2', 'Type', title)
  worksheet.write('E2', 'LOA', title)
  #worksheet.write('E2', 'Arrive', title)
  worksheet.write('F2', 'Depart', title)
  worksheet.write('G2', 'Nights', title)
  worksheet.write('H2', 'Slip', title)
  worksheet.write('I2', 'Notes', title)
  worksheet.write('J2', 'Special Requests (not printed)', title)
  
  
  # Iterate over the data and write it out row by row.
  for res in reservations:
    arrival = res.arrival.split('/')
    departure = res.departure.split('/')
  
    if row % 2:
      basetemp = lite_basic
      centertemp = lite_center
      sliptemp = lite_slip
      spectemp = lite_spec
    else:
      basetemp = dark_basic
      centertemp = dark_center
      sliptemp = dark_slip
      spectemp = dark_spec
  
    worksheet.write(row, 0, 'fill', filler)
    worksheet.write(row, 1, res.name, basetemp)
    worksheet.write(row, 2, res.boatname, basetemp)
    worksheet.write(row, 3, res.boattype[0]+'B', centertemp)
    worksheet.write(row, 4, int(res.loa.split("'")[0]), centertemp)
    #worksheet.write(row, 4, arrival[0]+"/"+arrival[1], centertemp)
    worksheet.write(row, 5, departure[0]+"/"+departure[1], centertemp)
    worksheet.write(row, 6, int(res.nights), centertemp)
    worksheet.write(row, 7, res.assign, sliptemp)
    worksheet.write(row, 8, res.notes, basetemp)
    worksheet.write(row, 9, res.specreq, spectemp)
    row += 1
  
  # Print Setup
  worksheet.print_area(0, 1, len(reservations)+1, 8)
  worksheet.set_landscape()
  workbook.close()
  
  # File Management and open spreadsheet
  #subprocess.Popen(workpath, shell=True)
  os.startfile(workpath)
  for x in os.listdir():
    dir = x.split('_')
    if len(dir) > 1:
      if datetime.strptime(dir[0], '%Y-%m-%d').date() < datetime.now().date():
        rmtree(x)
sys.exit()



























