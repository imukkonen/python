"""
Routes and views for the flask application.
"""

# -*- coding: iso-8859-1 -*

from datetime import datetime
from flask import render_template, request, redirect, url_for
from Tables import app
import win32com.client
from Tables.count import count, sh_exist

Excel = win32com.client.Dispatch("Excel.Application")
month=0
months={1:'tammikuu', 2: 'helmikuu', 3: 'maaliskuu', 4:'huhtikuu', 5:'toukokuu', 6:'kesäkuu', 7:'heinäkuu', 8:'elokuu',9:'syyskuu', 10:'lokakuu',11:'marraskuu', 12:'joulukuu'}

@app.route('/home', methods=['post', 'get'])
@app.route('/', methods=['post', 'get'])
def home():
    """Pääsivusto."""
    global wb
    global month

  # kuukauden valinta ja tarkastus, että kuukausi on jo mennyt ja palkka saa laskinta
    message='Valitse kuukausi'
    month=datetime.now().month
    if request.method == 'POST':
        month_s = int(request.form.get('month_select'))
        if month_s < month:
            month=month_s
            lb_name='Laskelma_'+str(month)+'2019.xlsx'  #palkan excel kirjan nimi kuukaden mukaan
            lb = Excel.Workbooks.Open(lb_name)          # avaa excel kirja
            cc = lb.Sheets.Count
            if cc==1:
                count(month)
            message='Palkka laskittu ja on tiedostossa Laskelma'+str(month)+'2019.xlsx'
        else:
            message='Valittu kuukausi ei vielä tullut loppuun'
            
    return render_template(
              'wagecount.html',
               title='Palkinnon laske',
               text='Laskelma',
               months=months,
               message=message
                )

@app.route('/wpages', methods=['post', 'get'])
def wpages():
    """Workers wages."""
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
               title='Työntekijän palkkalaskelma',
               text='Laskelma',
               months=months,
               message=message
                )


            
  
@app.route('/worker/<month>', methods=['GET', 'POST'])
def worker(month):
    """Renders työntekijän sivusto."""
    global wb
    
    global wsheet
    virhe=''
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
        wage_h=round(wagesdict.get(int(select)),2)       #valitun työntekijän tuntipalkka
        if sh_exist(lb, id):                    #tarkistamme, että sivu on olemassa
            lsheet=lb.Sheets(id)                   #valitun työntekijän sivu excel kirjassa
            #palkan ja kuukauden tietoja työntekijästä
            hours=lsheet.Range("B41").value               
            product=lsheet.Range("B43").value
            product=round(product*100,3)
            night_h=lsheet.Range("K42").value
            wage_n=round(lsheet.Range("K43").value,2)
            wage=float(lsheet.Range("J41").value)
            wage_p=round(lsheet.Range("K41").value,2)
            wage_t=round(lsheet.Range("K45").value,2)
        else:
            virhe='Työntekijän sivu ei löydy. Luoda palkkalaskinnon sivut Home välilehdessä'
   
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
        wage_t=wage_t,
        virhe=virhe
        )



@app.route('/contact')
def contact():
    """Renders the contact page."""
    return render_template(
        'contact.html',
        title='Contact',
        year=datetime.now().year,
        message='Ota yhteyttä'
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
