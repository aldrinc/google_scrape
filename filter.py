# coding=utf-8

import requests
from bs4 import BeautifulSoup
import tkinter as tk
import re
import pandas as pd
import os.path
from append_df_to_excel import append_df_to_excel





global l 
l = []

dir_path = os.path.dirname(os.path.realpath(__file__))


def getLeftNav(url): 
    req = requests.get(url)
    soup = BeautifulSoup(req.text, "lxml")
    
    #Scrap  
    i1 = soup.find_all('td', {"id": "leftnav"})
    i2 = i1[0].find_all('ul')
    i3 = [i.find_all('li') for i in i2]
    
    
    
    try: 
        if 'Effacer' in str(i3[len(i3)-1][1]): 
            i3 = i3[0:len(i3)-1]
    except: 
        pass
    
    
    # dict Prix:jusqu'à 300€
    d = {}
    
    # dict jusqu'à 300€:pdtr0
    d2 = {}
    
    # Filling of the 2 dictionaries 
    for j in range (1,len(i3)): 
        i4 = [i.find_all('a') for i in i3[j]]
        p = re.compile('(?<=\")(.*?)(?=\")')
        k = [p.findall(str(i)) for i in i4]
        k = k[1:len(k)]
        i5 = [i.text for i in i3[j]]
        key = i5[0]
        values = i5[1:] 
        values = [i.replace('\xa0', ' ') for i in values]
        values = [i for i in values if i not in 'Effacer']
        
        for e in range(len(values)):
            try: 
                url = 'https://www.google.fr' + k[e][0].replace('&amp;', '&')
                d2[values[e]] = url
                
            except: 
                pass
    
        d[key] = values
        
    return soup, d, d2



def visu(soup,d,d2):
    # Scrollbar management
    class ScrolledFrame(tk.Frame):
    
        def __init__(self, parent, vertical=True, horizontal=False):
            super().__init__(parent)
    
            # canvas for inner frame
            self._canvas = tk.Canvas(self)
            self._canvas.grid(row=0, column=0, sticky='news') # changed
    
            # create right scrollbar and connect to canvas Y
            self._vertical_bar = tk.Scrollbar(self, orient='vertical', command=self._canvas.yview)
            if vertical:
                self._vertical_bar.grid(row=0, column=1, sticky='ns')
            self._canvas.configure(yscrollcommand=self._vertical_bar.set)
    
            # create bottom scrollbar and connect to canvas X
            self._horizontal_bar = tk.Scrollbar(self, orient='horizontal', command=self._canvas.xview)
            if horizontal:
                self._horizontal_bar.grid(row=1, column=0, sticky='we')
            self._canvas.configure(xscrollcommand=self._horizontal_bar.set)
    
            # inner frame for widgets
            self.inner = tk.Frame(self._canvas, bg='white')
            self._window = self._canvas.create_window((0, 0), window=self.inner, anchor='nw')
    
            # autoresize inner frame
            self.columnconfigure(0, weight=1) # changed
            self.rowconfigure(0, weight=1) # changed
    
            # resize when configure changed
            self.inner.bind('<Configure>', self.resize)
            self._canvas.bind('<Configure>', self.frame_width)
    
        def frame_width(self, event):
            # resize inner frame to canvas size
            canvas_width = event.width
            self._canvas.itemconfig(self._window, width = canvas_width)
    
        def resize(self, event=None): 
            self._canvas.configure(scrollregion=self._canvas.bbox('all'))
    
    
    
    # Radio Button
    def create(rb, lb):
        
        manufacturers = rb

        lbl1 = tk.Label(window.inner, text=lb,font='Calibri 14 bold')
        lbl1.pack()
        
        # create an empty dictionary to fill with Radiobutton widgets
        man_select = dict()
        # create a variable class to be manipulated by radiobuttons
        man_var = tk.StringVar()
        '''
        try: 
            for e in l: 
                if e in manufacturers: 
                    i = manufacturers.index(e)
                    man_var.set(manufacturers[i])
        except : 
            pass
            
        '''
        def sel():
            #print(d2[str(man_var.get())])
            global X 
            X = d2[str(man_var.get())]
            if (str(man_var.get())) not in l:
                l.append(str(man_var.get()))
            root.destroy()

       
        # fill radiobutton dictionary with keys from manufacturers list with Radiobutton
        # values assigned to corresponding manufacturer name
        for man in manufacturers:
            man_select[man] = tk.Radiobutton(window.inner, text=man, variable=man_var, value=man, command=sel) 
            #display
            man_select[man].pack(anchor = 'w')
        
        try: 
            return sel()
        except: 
            pass
        
        
    root = tk.Tk()
    root.geometry("300x1000")
    window = ScrolledFrame(root)
    window.pack(expand=True, fill='both')
        
    lb = d.keys()
    
    # Display 
    for k in d.keys(): 
        lb = k 
        rb = d[k]
        create(rb, lb)
        lbl = tk.Label(window.inner, text='------------------------------')
        lbl.pack()
        
    
    def caract_excel():
        # Scrap characterisctic on each product 
        j1 = soup.find_all('div', {'class':'g'})
        j2 = [j.find_all('div', {'class':'pslires'}) for j in j1]
        
        data = {'Name':[], 'Price':[], 'Retailer':[], 'Stars':[], 'Description':[]}
        
        for i in range(len(j2)):
            try: 
                name = j2[i][0].find_all('a')[1].text
                data['Name'].append(name)
            except: 
                data['Name'].append('none')
            
            try: 
                price = j2[i][0].find_all('div')[2].text.replace('\xa0', ' ')
                reta = j2[i][0].find_all('div')[3].text.replace('\xa0', ' ')
                data['Price'].append(price)
                data['Retailer'].append(reta)
            except: 
                data['Price'].append('none')
                data['Retailer'].append('none')
            
            try: 
                desc = j2[i][0].find_all('div')[5].text
                data['Description'].append(desc)
            except: 
                data['Description'].append('none')
                
            try: 
                stars = str(j2[i][0].find_all('div')[6]).split('"')[1].replace('\xa0', ' ')
                if 'étoiles' not in stars: 
                    data['Stars'].append('none')
                else: 
                    stars = stars.split('sur')
                    stars = stars[0]
                    data['Stars'].append(stars)
                
            except: 
                data['Stars'].append('none')
        
        df = pd.DataFrame(data=data)
        writer = pd.ExcelWriter(dir_path + '/data/'+req+'.xlsx')
        df.to_excel(writer,'Products', index=False)
        
        writer.save()
        
        
        s = ''
        for i in l:
            s = s + ' + ' + i 
        
        datafilter = {'Selected filter ': [s]}
        dfilter = pd.DataFrame(data=datafilter)
        append_df_to_excel(dir_path + '/data/'+req+'.xlsx', df=dfilter, 
                   sheet_name='Characteristics', engine='openpyxl',index=False, 
                   startcol=0, startrow=0)
        
        
        col_number = 2 
        for k in d.keys():
            d1 = {k:d[k]}
            df = pd.DataFrame(data=d1)
            append_df_to_excel(dir_path + '/data/'+req+'.xlsx', df=df, 
                   sheet_name='Characteristics', engine='openpyxl',index=False, 
                   startcol=col_number, startrow=0)
            col_number = col_number + 1

        
        if os.path.exists(dir_path + '/data/'+req+'.xlsx'): 
            root.destroy()
         
       
    b = tk.Button(root, text="SEARCH", width= 20, highlightbackground='green')
    b['command'] = caract_excel
    b.pack()
    
    def doSomething():

        root.destroy()
    
    root.protocol('WM_DELETE_WINDOW', doSomething)  # root is your root window
    
    root.mainloop()   

    return X
  

# ------------------------------------------------------------------------------ #

# EXECUTION 
    
req = input('Products :\n')

url = 'https://www.google.fr/search?q='+req+'&source=lnms&tbm=shop&start=0'

while not os.path.exists(dir_path + '/data/'+req+'.xlsx'):
    soup, d, d2 = getLeftNav(url)
 
    url = visu(soup,d,d2)
