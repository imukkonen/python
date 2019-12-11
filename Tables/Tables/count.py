from datetime import datetime
import win32com.client

def count(month):
    Excel = win32com.client.Dispatch("Excel.Application")
    tv_name='TV'+str(month)+'2019.xlsm'  #työvuorojen laskun excel kirjan nimi
    
    wb = Excel.Workbooks.Open(tv_name)  #avaamme excel kirjan
    
    wsheet=wb.Sheets("Työntekijät")     #valitse sivu, jossa on työntekijän nimet
    lb_name='Laskelma_'+str(month)+'2019.xlsx'  #palkan excel kirjan nimi kuukaden mukaan
    lb = Excel.Workbooks.Open(lb_name)          # avaa excel kirja
    #wt=int(wsheet.Range("F1").value)           #työntekijän määrä
    wt=25
    timestamp=datetime.now()
    timestamp=timestamp.strftime('%d/%m/%Y')
    #silmukka työntekijälle
    for i in range(2, wt):
        tt_id =''
        tt_id = int(wsheet.Cells(i, 1).Value)
        #silmukka päivällä (kirjan sivulla)
        for j in range(1, 31):
            if j < 10: # kuukausipäivät
                sh_name = '0' + str(j) #sivun nimi TV kirjassa
            else:
                sh_name = str(j)
            for k in range(5, 53, 4): # silmukka työkoneilla
                if wb.Sheets(sh_name).Cells(5, k).Value == tt_id: # aamuvuoro
                    if not(sh_exist(lb, str(tt_id))):        #jos työntekijän sivua ei ole, kopioidaan malli-sivu ja nimeään sen tt id-llä
                        ls=lb.Worksheets.Add()
                        ls.Name=str(tt_id)
                        ls=lb.Worksheets(str(tt_id))
                        lb.Worksheets("malli").Range("A1:L50").Copy(ls.Range("A1:L50"))
                        lb.Worksheets(str(tt_id)).Cells(7, 2).Value = tt_id
                        lb.Worksheets(str(tt_id)).Cells(4, 12).Value = wb.Sheets("Työntekijät").Cells(i, 3).Value
                    #kopioidaan vuoron tehtyä tt -sivulle
                    lb.Sheets(str(tt_id)).Cells(7, 3).Value = wb.Sheets("Työntekijät").Cells(i, 2).Value
                    lb.Sheets(str(tt_id)).Cells(j + 9, 2).Value = wb.Sheets(sh_name).Cells(7, k).Value
                    lb.Sheets(str(tt_id)).Cells(j + 9, 4).Value = wb.Sheets(sh_name).Cells(7, k + 2).Value
                    lb.Sheets(str(tt_id)).Cells(j + 9, 5).Value = wb.Sheets(sh_name).Cells(7, k + 3).Value
                    lb.Sheets(str(tt_id)).Cells(j + 9, 8).Value = wb.Sheets(sh_name).Cells(53, k).Value + wb.Sheets(sh_name).Cells(53, k + 2).Value
                    lb.Sheets(str(tt_id)).Cells(47, 4).Value = timestamp
                    break
                else:
                    if wb.Sheets(sh_name).Cells(58, k).Value == tt_id: #päivävuoro
                        if not(sh_exist(lb, str(tt_id))): 
                            ls=lb.Worksheets.Add()
                            ls.Name=str(tt_id)
                            ls=lb.Worksheets(str(tt_id))
                            lb.Worksheets("malli").Range("A1:L50").Copy(ls.Range("A1:L50"))
                            lb.Worksheets(str(tt_id)).Cells(7, 2).Value = tt_id
                            lb.Worksheets(str(tt_id)).Cells(4, 12).Value = wb.Sheets("Työntekijät").Cells(i, 3).Value
                        lb.Sheets(str(tt_id)).Cells(7, 3).Value = wb.Sheets("Työntekijät").Cells(i, 2).Value
                        lb.Sheets(str(tt_id)).Cells(j + 9, 2).Value = wb.Sheets(sh_name).Cells(60, k).Value
                        lb.Sheets(str(tt_id)).Cells(j + 9, 4).Value = wb.Sheets(sh_name).Cells(60, k + 2).Value
                        lb.Sheets(str(tt_id)).Cells(j + 9, 6).Value = wb.Sheets(sh_name).Cells(60, k + 3).Value
                        lb.Sheets(str(tt_id)).Cells(j + 9, 8).Value = wb.Sheets(sh_name).Cells(106, k).Value + wb.Sheets(sh_name).Cells(106, k + 2).Value
                        lb.Sheets(str(tt_id)).Cells(47, 4).Value = timestamp
                        break
                    else:
                        if wb.Sheets(sh_name).Cells(111, k).Value == tt_id: #yövuoro
                            if not(sh_exist(lb, str(tt_id))): 
                                ls=lb.Worksheets.Add()
                                ls.Name=str(tt_id)
                                ls=lb.Worksheets(str(tt_id))
                                lb.Worksheets("malli").Range("A1:L50").Copy(ls.Range("A1:L50"))
                                lb.Worksheets(str(tt_id)).Cells(7, 2).Value = tt_id
                                lb.Worksheets(str(tt_id)).Cells(4, 12).Value = wb.Sheets("Työntekijät").Cells(i, 3).Value
                            lb.Sheets(str(tt_id)).Cells(7, 3).Value = wb.Sheets("Työntekijät").Cells(i, 2).Value
                            lb.Sheets(str(tt_id)).Cells(j + 9, 2).Value = wb.Sheets(sh_name).Cells(113, k).Value
                            lb.Sheets(str(tt_id)).Cells(j + 9, 4).Value = wb.Sheets(sh_name).Cells(113, k + 2).Value
                            lb.Sheets(str(tt_id)).Cells(j + 9, 7).Value = wb.Sheets(sh_name).Cells(113, k + 3).Value
                            lb.Sheets(str(tt_id)).Cells(j + 9, 8).Value = wb.Sheets(sh_name).Cells(159, k).Value + wb.Sheets(sh_name).Cells(159, k + 2).Value
                            lb.Sheets(str(tt_id)).Cells(47, 4).Value = timestamp
                            break
    lb.Save()
    wb.Close()
                        
def sh_exist(wb, sName): 
    found=False
    for sh in wb.Sheets:
        if sh.Name==sName:
            found=True
            break
    return found