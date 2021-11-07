import xlsxwriter
from datetime import datetime

def writetoexcel(linestring, datestring,row):
  cell0 = linestring[0:4]
  cell1 = linestring[4:21]
  cell2 = linestring[21:52]
  cell3 = linestring[52:57]
  cell4 = linestring[57:68]
  cell5 = linestring[68:73]
  cell6 = linestring[73:83]
  cell7 = linestring[83:85]
  cell8 = linestring[85:89]
  cell9 = linestring[89:96]
  cell10 = linestring[96:108]
  cell11 = linestring[108:110]
  cell12 = linestring[110:123]
  cell13 = linestring[123:]
  
  date_time = datetime.strptime(datestring, '%m/%d/%Y')
  
  #print(linestring)
  #print(cell1+cell2+cell3+cell4+cell5+cell6+cell7+cell8+cell9+cell10+cell11+cell12+cell13+cell14)
  worksheet.write_string(row,0,cell0)
  worksheet.write_string(row,1,cell1)
  worksheet.write_string(row,2,cell2)
  worksheet.write_string(row,3,cell3)
  worksheet.write_string(row,4,cell4)
  worksheet.write_string(row,5,cell5)
  worksheet.write_datetime(row,6,date_time,date_format)
  worksheet.write_string(row,7,cell7)
  worksheet.write_string(row,8,cell8)
  worksheet.write(row,9,cell9)
  worksheet.write(row,10,cell10)
  worksheet.write_string(row,11,cell11)
  worksheet.write(row,12,cell12)
  worksheet.write(row,13,cell13)
	
	

def isbadline(line,header1,header2,header3,useless1,useless2,useless3,useless4,useless5,useless6):
  if header1 in line: return True
  if header2 in line: return True
  if header3 in line: return True
  if useless1 in line: return True
  if useless2 in line: return True
  if useless3 in line: return True
  if useless4 in line: return True
  if useless5 in line: return True
  if useless6 in line: return True
  if line.isspace(): return True
  return False 

def getmonth(monthstring):
  if monthstring=="Jan":
    return "01"
  elif monthstring=="Feb":
    return "02"
  elif monthstring=="Mar":
    return "03"
  elif monthstring=="Apr":
    return "04"
  elif monthstring=="May":
    return "05"
  elif monthstring=="Jun":
    return "06"
  elif monthstring=="Jul":
    return "07"
  elif monthstring=="Aug":
    return "08"
  elif monthstring=="Sep":
    return "09"
  elif monthstring=="Oct":
    return "10"
  elif monthstring=="Nov":
    return "11"
  elif monthstring=="Dec":
    return "12"
  else:
    return monthstring

def getyear(yearstring):
  if (int(yearstring)>79):
    return "19"+yearstring
  else:
    return "20"+yearstring

  
infile = open('open2.txt', 'r+')
outfile = open('outputNEW.txt', 'w+')
wh =""
whlist = []
totallist = []
grandtotal= ""
whtotal = ""
row=0
workbook = xlsxwriter.Workbook('OutputNEW2.xlsx')
worksheet = workbook.add_worksheet()
date_format = workbook.add_format({'num_format': 'mm/dd/yyyy'})

worksheet.set_column(0,0,6)
worksheet.set_column(1,1,24)
worksheet.set_column(2,2,38)
worksheet.set_column(3,3,7)
worksheet.set_column(4,4,13)
worksheet.set_column(5,5,6)
worksheet.set_column(6,6,12)
worksheet.set_column(7,7,3)
worksheet.set_column(8,8,9)
worksheet.set_column(9,9,12)
worksheet.set_column(10,10,17)
worksheet.set_column(11,11,6)
worksheet.set_column(12,12,20)
worksheet.set_column(13,13,12)
worksheet.set_column(14,14,9)
worksheet.set_column(15,15,24)

for line in infile:
        header1= "S E R I A L #   L I S T I N G"
        header2= "REQUESTED BY:"
        header3= "WARE---ITEM#"
        useless1 = "ITEM# TOTALS:"
        useless2 = "AVG / S/N:"
        useless3 = "LOT# TOTALS:"
        useless4 = "MFGR# TOTALS:"
        useless5 = "DATE   STS       # OF       QTY"
        useless6 = "* LANDED COSTS *"
        
        grand = "<<< GRAND TOTALS >>>"
        ware = "WARE# TOTALS:"
        
        
        
        tempbool= isbadline(line,header1,header2,header3,useless1,useless2,useless3,useless4,useless5,useless6)
        if (tempbool==False):
          
          temp=line[:3]
          
          if (temp != "   "): 
            wh=temp
          
          if grand in line:
            grandtotal = line
            continue
            
          if ware in line:
            whlist.append(wh)
            totallist.append(line)
            continue
            
          
          if ( len(line) > 130):
          	mystring= line[73:80]
          	if mystring=="   0000":
          		month="01"
          		day="01"
          		year="1979"
          		#year="00"
          	else:
          		month = getmonth(mystring[0:3])
          		day = mystring[3:5]
          		year = getyear(mystring[5:7])
          		#year=mystring[5:7]
          		
          	temp = line[:73]+month+"/"+day+"/"+year+line[80:108] + ' ' + line[108:]
          	writetoexcel(line,month+"/"+day+"/"+year,row)
          	row=row+1
          	line=temp
          	outfile.write(line)
          
          
          

blank= '\n'
outfile.write(blank)
outfile.write(blank)
outfile.write(blank)
row+=4

for x in range(len(whlist)):
  wh = whlist[x]
  whtotal= totallist[x]
  #print(wh)
  templine = whtotal[113:]
  templine = templine.strip()
  newline="    " + wh
  newline = newline.ljust(136)
  newline = newline + templine
  #print(newline)
  outfile.write(newline)
  outfile.write(blank)
  worksheet.write(row,14,wh)
  worksheet.write(row,15,whtotal[113:])
  #print(wh,"            ",whtotal[113:])
  row+=1
  

row+=2
outfile.write(blank)
outfile.write(blank)
templine = grandtotal[113:]
templine = templine.strip()
newline = "    " + "Grand Total"
newline = newline.ljust(136)
newline = newline + templine
#print(newline)  
outfile.write(newline)
worksheet.write(row,14,"TOTAL:")
worksheet.write(row,15,grandtotal[113:])



infile.close()  
outfile.close()
workbook.close()        
        
        
