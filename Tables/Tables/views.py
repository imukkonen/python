"""
Routes and views for the flask application.
"""

# -*- coding: iso-8859-1 -*

from datetime import datetime
from flask import render_template, request, redirect, url_for
from Tables import app
import win32com.client
from Tables.count import count

Excel = win32com.client.Dispatch("Excel.Application")
month=0
months={1:'tammikuu', 2: 'helmikuu', 3: 'maaliskuu', 4:'huhtikuu', 5:'toukokuu', 6:'kesäkuu', 7:'heinäkuu', 8:'elokuu',9:'syyskuu', 10:'lokakuu',11:'marraskuu', 12:'joulukuu'}

@app.route('/home')
def home():
    """Renders the home page."""
    return render_template(
        'index.html',
        title='Home Page',
        year=datetime.now().year,
    )

@app.route('/', methods=['post', 'get'])
def main():
    """Pääsivusto."""
    global wb
    global month
  # kuukauden valinta ja tarkastus, että palkka on laskettu
    message='Valitse kuukausi'
    month=datetime.now().month
    if request.method == 'POST':
        month_s = int(request.form.get('month_select'))
        if month_s < month:
            month=month_s
            return redirect(url_for('worker', month=month_s))
        else:
            message='Valittu kuukausi ei vielä tullut loppuun'
            
    return render_template(
              'main.html',
               title='Palkinnon laske',
               text='Laskelma',
               months=months,
               message=message
                )

@app.route('/wagecount', methods=['post', 'get'])
def wagecount():
    """Pääsivusto."""
    global wb
    global month
  # kuukauden valinta ja tarkastus, että palkka on laskettu
    message='Valitse kuukausi'
    month=datetime.now().month
    if request.method == 'POST':
        month_s = int(request.form.get('month_select'))
        if month_s < month:
            month=month_s
            count(month)
            message='Palkka laskittu ja on tiedostossa Laskelma'+str(month)+'2019'
        else:
            message='Valittu kuukausi ei vielä tullut loppuun'
            
    return render_template(
              'wagecount.html',
               title='Palkinnon laske',
               text='Laskelma',
               months=months,
               message=message
                )
            
  
@app.route('/worker/<month>', methods=['GET', 'POST'])
def worker(month):
    """Renders työntekijän sivusto."""
    global wb
    
    global wsheet
    tv_name='TV'+str(month)+'2019.xlsm'  #työvuorojen laskun excel kirjan nimi
    
    wb = Excel.Workbooks.Open(tv_name)  #avaamme excel kirjan
    
    wsheet=wb.Sheets("Työntekijät")     #valitse sivu, jossa on työntekijän nimet
    lb_name='Laskelma_'+str(month)+'2019.xlsx'  #palkan excel kirjan nimi kuukaden mukaan
    lb = Excel.Workbooks.Open(lb_name)          # avaa excel kirja
    count = int(float(wsheet.Cells(1,6).value))
    names = [r[0].value for r in wsheet.Range("B2:B86")]
    id1= [r[0].value for r in wsheet.Range("A2:A86")]
    ids = [int(item) for item in id1]
    hwages =[r[0].value for r in wsheet.Range("C2:C86")]
    namesdict=dict(zip(ids,names))      #luodaan sanakirja (dictionary) {id:sukunimi}
    wagesdict=dict(zip(ids,hwages))         #dictionary {id: tuntipalkka}
    ans=''
    id= ''
    name=''
    wage_h=''
    hours=''
    product=''
    night_h=''
    wage_n=''
    wage=''
    wage_p=''
    wage_t=''
    if request.method == 'POST':
        select = request.form.get('wr_select')
        id= str(select)
        name=namesdict.get(int(select))         #valitun työntekijän sukunimi
        wage_h=wagesdict.get(int(select))       #valitun työntekijän tuntipalkka
        lsheet=lb.Sheets(id)                   #valitun työntekijän sivu excel kirjassa
        #palkan ja kuukauden tietoja työntekijästä
        hours=lsheet.Range("B41")               
        product=lsheet.Range("B43").value
        product=round(product,3)
        night_h=lsheet.Range("K42")
        wage_n=lsheet.Range("K43")
        wage=float(lsheet.Range("J41").value)
        wage_p=round(lsheet.Range("K41").value,2)
        wage_t=round(lsheet.Range("K45").value,2)
   
    return render_template(
        'worker.html',
        title='Valitsit '+ months.get(int(month)),
        message='Valitse työntekijä',
        ids=ids,
        names=namesdict,
        id= id,
        name=name,
        wage_h=wage_h,
        hours=hours,
        product=product,
        night_h=night_h,
        wage_n=wage_n,
        wage=wage,
        wage_p=wage_p,
        wage_t=wage_t
        )



@app.route('/contact')
def contact():
    """Renders the contact page."""
    return render_template(
        'contact.html',
        title='Contact',
        year=datetime.now().year,
        message='Your contact page.'
    )

@app.route('/about')
def about():
    """Renders the about page."""
    return render_template(
        'about.html',
        title='About',
        year=datetime.now().year,
        message='Your application description page.'
    )
