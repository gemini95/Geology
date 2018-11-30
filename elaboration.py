import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
import numpy as np
from sklearn import datasets, linear_model
from sklearn.metrics import mean_squared_error, r2_score
from sklearn.tree import DecisionTreeClassifier
from sklearn.ensemble import RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn import cross_validation


#from sklearn.feature_selection import VarianceThresholds
import xlrd, xlwt, datetime, statistics
import pandas as pd

prepath = "C:/Users/Alan"
PerModello	=	xlrd.open_workbook(prepath+"/Dropbox/Geology/Dati Grezzi Normalizzati/per_modello.xls","r")
AltriValori	=	xlrd.open_workbook(prepath+"/Dropbox/Geology/Dati Grezzi Normalizzati/dati_mercurio_altri.xls","r")
DatiRete	=	xlrd.open_workbook(prepath+"/Dropbox/Geology/Dati Grezzi Normalizzati/dati_mercurio_rete.xls","r")

sheetPerModello		= 	PerModello. 	sheet_by_index(0)
sheetAltriValori	=	AltriValori. 	sheet_by_index(0)
sheetDatiRete		=	DatiRete. 		sheet_by_index(0)


date_format = xlwt.XFStyle()
date_format.num_format_str = 'dd/mm/yyyy'

def nAcquifero(profondita):
	if profondita>=15 and profondita<60:
		return 1
	elif profondita>=60 and profondita<90:
		return 2
	elif profondita>=100 and profondita<120:
		return 3
	elif profondita>=130 and profondita<140:
		return 4
	elif profondita>=145 and profondita<160:
		return 5
	elif profondita>=210 and profondita<220:
		return 7
	elif profondita>=230 and profondita<260:
		return 8
	elif profondita>=270 and profondita<310:
		return 9
	elif profondita>310:
		return 10
	else:
		return 0

def cercaElemento(lista,elemento):
	trovato=0
	for i in range(len(lista)):
		if lista[i]==elemento:
			trovato=1
	return trovato
def creaModello(fileExcel, LetteraModello):
	PerModello=xlrd.open_workbook(prepath+"/Dropbox/Geology/Dati Grezzi Normalizzati/per_modello.xls","r")
	sheetPerModello = PerModello.sheet_by_index(0)

	#fileExcel = xlrd.open_workbook(excel,'r')
	foglio = fileExcel.sheet_by_index(0)

	newModel = xlwt.Workbook()
	sheetModel = newModel.add_sheet('Data')

	for row in range(foglio.nrows):
		riga=foglio.row_values(row)
		if row!=0:
			data = xlrd.xldate_as_tuple(riga[1],fileExcel.datemode)
		for i in range(len(riga)):
			if row!=0:
				if i==1:
					sheetModel.write(row,i, datetime.date(int(data[0]),int(data[1]),int(data[2])),date_format)
				else:
					sheetModel.write(row,i,riga[i])
			else:
				sheetModel.write(row,i,riga[i])

	sheetModel.write(0,i+1,"modello"+LetteraModello)

	if LetteraModello=="A":
		for rowPM in range(sheetPerModello.nrows):
			rigaPM=sheetPerModello.row_values(rowPM)
			for row in range(foglio.nrows):
				riga=foglio.row_values(row)
				if rigaPM[0]==riga[0]:
					sheetModel.write(row,len(riga), rigaPM[2])
	elif LetteraModello =="B":
		for rowPM in range(sheetPerModello.nrows):
			rigaPM=sheetPerModello.row_values(rowPM)
			for row in range(foglio.nrows):
				riga=foglio.row_values(row)
				if rigaPM[0]==riga[0]:
					sheetModel.write(row,len(riga), rigaPM[3])
	elif LetteraModello == "C":
		for rowPM in range(sheetPerModello.nrows):
			rigaPM=sheetPerModello.row_values(rowPM)
			for row in range(foglio.nrows):
				riga=foglio.row_values(row)
				if rigaPM[0]==riga[0]:
					sheetModel.write(row,len(riga), rigaPM[4])
	else:
		print("Errore nella lettura del modello")

	newModel.save("Altro/Modello/"+LetteraModello+".xls")
def aggiungiModello(fileExcel, LetteraModello, nome):
	PerModello=xlrd.open_workbook(prepath+"/Dropbox/Geology/Dati Grezzi Normalizzati/per_modello.xls","r")
	sheetPerModello = PerModello.sheet_by_index(0)

	#fileExcel = xlrd.open_workbook(excel,'r')
	foglio = fileExcel.sheet_by_index(0)

	newModel = xlwt.Workbook()
	sheetModel = newModel.add_sheet('Data')

	for row in range(foglio.nrows):
		riga=foglio.row_values(row)
		if row!=0:
			data = xlrd.xldate_as_tuple(riga[1],fileExcel.datemode)
		for i in range(len(riga)):
			if row!=0:
				if i==1:
					sheetModel.write(row,i, datetime.date(int(data[0]),int(data[1]),int(data[2])),date_format)
				else:
					sheetModel.write(row,i,riga[i])
			else:
				sheetModel.write(row,i,riga[i])

	sheetModel.write(0,i+1,"modello"+LetteraModello)

	if LetteraModello=="A":
		for rowPM in range(sheetPerModello.nrows):
			rigaPM=sheetPerModello.row_values(rowPM)
			for row in range(foglio.nrows):
				riga=foglio.row_values(row)
				if rigaPM[0]==riga[0]:
					sheetModel.write(row,len(riga), rigaPM[2])
	elif LetteraModello =="B":
		for rowPM in range(sheetPerModello.nrows):
			rigaPM=sheetPerModello.row_values(rowPM)
			for row in range(foglio.nrows):
				riga=foglio.row_values(row)
				if rigaPM[0]==riga[0]:
					sheetModel.write(row,len(riga), rigaPM[3])
	elif LetteraModello == "C":
		for rowPM in range(sheetPerModello.nrows):
			rigaPM=sheetPerModello.row_values(rowPM)
			for row in range(foglio.nrows):
				riga=foglio.row_values(row)
				if rigaPM[0]==riga[0]:
					sheetModel.write(row,len(riga), rigaPM[4])
	else:
		print("Errore nella lettura del modello")
	newModel.save(nome+"Modello.xls")
def aggiungiAcquifero(fileExcel, finestra, nome):
	VenetoDati = xlrd.open_workbook(prepath+"/Dropbox/Geology/Dati Grezzi Normalizzati/veneto dati.xls","r")
	sheetVD = VenetoDati.sheet_by_index(1)

	PerModello=xlrd.open_workbook(prepath+"/Dropbox/Geology/Dati Grezzi Normalizzati/per_modello.xls","r")
	sheetPerModello = PerModello.sheet_by_index(0)

	#fileExcel = xlrd.open_workbook(excel,'r')
	foglio = fileExcel.sheet_by_index(0)

	newModel = xlwt.Workbook()
	sheetModel = newModel.add_sheet('Data')

	primaRiga = foglio.row_values(0)
	for i in range(len(primaRiga)):
		sheetModel.write(0,i,primaRiga[i])
	sheetModel.write(0,i+1,"Acquifero")

	for row in range(1,foglio.nrows):
		riga=foglio.row_values(row)
		if row!=0:
			data = xlrd.xldate_as_tuple(riga[1],fileExcel.datemode)
		for i in range(len(riga)):
			if row!=0:
				if i==1:
					sheetModel.write(row,i, datetime.date(int(data[0]),int(data[1]),int(data[2])),date_format)
				else:
					sheetModel.write(row,i,riga[i])
			else:
				sheetModel.write(row,i,riga[i])

		pozzo = riga[0]
		for row1 in range(1,sheetVD.nrows):
			riga1 = sheetVD.row_values(row1)
			pozzo1 = riga1[2][4]+riga1[2][5]+riga1[2][6]+riga1[2][7]
			profondita = riga1[4]
			if int(pozzo)==int(pozzo1):
				modello = "no"
				for i in finestra:
					if nAcquifero(profondita)==i:
						modello="sì"
		sheetModel.write(row,len(riga),modello)
	newModel.save(nome+"Acquifero.xls")

def isSpring(data):
	if data[1]>=3 and data[1]<6:
		return 1
	else:
		return 0
def isSummer(data):
	if data[1]>=6 and data[1]<9:
		return 1
	else:
		return 0
def isAutumn(data):
	if data[1]>=9 and data[1]<12:
		return 1
	else:
		return 0
def isWinter(data):
	if data[1]==12 or data[1]<3:
		return 1
	else:
		return 0
def isWhatsSeason(data):
	if isSpring(data)==1:
		return 1
	elif isSummer(data)==1:
		return 2
	elif isAutumn(data)==1:
		return 3
	elif isWinter(data)==1:
		return 4
	else:
		print("ERROR! WHAT THE FUCK OF SEASON IS IT?")

def contaStagioni(sheet, fileExcel):
	data 	= []
	# Stampa prima riga
	for row in range(1,sheet.nrows):
		riga = sheet.row_values(row)
		d 	 = xlrd.xldate_as_tuple(riga[1],fileExcel.datemode)
		flag=1
		for i in range(len(data)):
			if data[i]==d:
				flag=0
		if flag==1:
			data = data + [xlrd.xldate_as_tuple(riga[1],fileExcel.datemode)]
	
	#data 	=ordinaLista(data)
	data.sort()
	primaData=data[0]
	ultimaData=data[len(data)-1]

	# Scrivi il range delle stagioni
	#File = open("Seasons.txt","w")
	fileSeasons = xlwt.Workbook()
	File 		= fileSeasons.add_sheet('Data')
	Stagione=0
	Anno=0
	
	if isSpring(primaData):
		File.write(0,0,"03/"+str(primaData[0]))
		File.write(0,1,"05/"+str(primaData[0]))
		Stagione=1
	elif isSummer(primaData):
		File.write(0,0,"06/"+str(primaData[0]))
		File.write(0,1,"08/"+str(primaData[0]))
		Stagione=2
	elif isAutumn(primaData):
		File.write(0,0,"09/"+str(primaData[0]))
		File.write(0,1,"11/"+str(primaData[0]))
		Stagione=3
	elif isWinter(data[1]):
		if(primaData[1]==12):
			File.write(0,0,"12/"+str(primaData[0]))
			File.write(0,1,"02/"+str(primaData[0]+1))
		else:
			File.write(0,0,"12/"+str(primaData[0]-1))
			File.write(0,1,"02/"+str(primaData[0]))
			Anno=Anno-1
		Stagione=4
	else:
		File.write("ERROR!\n")
	Stagione=Stagione+1


	nSeasons=1
	fine=1
	while fine==1:
		if Stagione>4:
			Stagione=Stagione%4
			Anno=Anno+1
		if Stagione==1:
			File.write(nSeasons,0,"03/"+str(primaData[0]+Anno))
			File.write(nSeasons,1,"05/"+str(primaData[0]+Anno))
		elif Stagione==2:
			File.write(nSeasons,0,"06/"+str(primaData[0]+Anno))
			File.write(nSeasons,1,"08/"+str(primaData[0]+Anno))
		elif Stagione==3:
			File.write(nSeasons,0,"09/"+str(primaData[0]+Anno))
			File.write(nSeasons,1,"11/"+str(primaData[0]+Anno))
		elif Stagione==4:
			File.write(nSeasons,0,"12/"+str(primaData[0]+Anno))
			File.write(nSeasons,1,"02/"+str(primaData[0]+Anno+1))

		else:
			File.write("ERROR!")
		Stagione=Stagione+1
		nSeasons=nSeasons+1
		if ultimaData[0]==primaData[0]+Anno and Stagione==isWhatsSeason(ultimaData):
			fine=0


	if Stagione>4:
		Stagione=Stagione%4
		Anno=Anno+1
	if Stagione==1:
		File.write(nSeasons,0,"03/"+str(primaData[0]+Anno))
		File.write(nSeasons,1,"05/"+str(primaData[0]+Anno))
	elif Stagione==2:
		File.write(nSeasons,0,"06/"+str(primaData[0]+Anno))
		File.write(nSeasons,1,"08/"+str(primaData[0]+Anno))
	elif Stagione==3:
		File.write(nSeasons,0,"09/"+str(primaData[0]+Anno))
		File.write(nSeasons,1,"11/"+str(primaData[0]+Anno))
	elif Stagione==4:
		File.write(nSeasons,0,"12/"+str(primaData[0]))
		File.write(nSeasons,1,"02/"+str(primaData[0]+1))
	else:
		File.write("ERROR!")
	nSeasons=nSeasons+1
			
	fileSeasons.save("Seasons.xls")

	return nSeasons
def createTrueSeason(listaFogli, listaExcel):
	Inverno = xlwt.Workbook()
	Primavera = xlwt.Workbook()
	Estate  = xlwt.Workbook()
	Autunno = xlwt.Workbook()

	sheetInverno = Inverno.add_sheet("Data")
	sheetPrimavera = Primavera.add_sheet("Data")
	sheetEstate  = Estate.add_sheet("Data")
	sheetAutunno = Autunno.add_sheet("Data")


	riga=listaFogli[0].row_values(0)
	for i in range(len(riga)):
		sheetPrimavera.write(0,i,riga[i])
		sheetEstate.write(0,i,riga[i])
		sheetAutunno.write(0,i,riga[i])
		sheetInverno.write(0,i,riga[i])


	count= [0]*4
	for i in range(len(listaFogli)):
		if i%4==0:
			for row in range(1,listaFogli[i].nrows):
				riga=listaFogli[i].row_values(row)
				data = xlrd.xldate_as_tuple(riga[1],listaExcel[i].datemode)
				count[0]+=1
				for j in range(len(riga)):
					if j==1:
						sheetInverno.write(count[0],j, datetime.date(int(data[0]),int(data[1]),int(data[2])),date_format )
					else:
						sheetInverno.write(count[0], j, riga[j])
		elif i%4==1:
			for row in range(1,listaFogli[i].nrows):
				riga=listaFogli[i].row_values(row)
				data = xlrd.xldate_as_tuple(riga[1],listaExcel[i].datemode)
				count[1]+=1
				for j in range(len(riga)):
					if j==1:
						sheetPrimavera.write(count[1],j, datetime.date(int(data[0]),int(data[1]),int(data[2])),date_format )
					else:
						sheetPrimavera.write(count[1], j, riga[j])
		elif i%4==2:
			for row in range(1,listaFogli[i].nrows):
				riga=listaFogli[i].row_values(row)
				data = xlrd.xldate_as_tuple(riga[1],listaExcel[i].datemode)
				count[2]+=1
				for j in range(len(riga)):
					if j==1:
						sheetEstate.write(count[2],j, datetime.date(int(data[0]),int(data[1]),int(data[2])),date_format )
					else:
						sheetEstate.write(count[2], j, riga[j])
		elif i%4==3:
			for row in range(1,listaFogli[i].nrows):
				riga=listaFogli[i].row_values(row)
				data = xlrd.xldate_as_tuple(riga[1],listaExcel[i].datemode)
				count[3]+=1
				for j in range(len(riga)):
					if j==1:
						sheetAutunno.write(count[3],j, datetime.date(int(data[0]),int(data[1]),int(data[2])),date_format )
					else:
						sheetAutunno.write(count[3], j, riga[j])

	Primavera.save("Stagioni/Primavera.xls")
	Estate.save("Stagioni/Estate.xls")
	Autunno.save("Stagioni/Autunno.xls")
	Inverno.save("Stagioni/Inverno.xls")
def cercaElementoPerIndice(lista,elemento):
	trovato=0
	for i in range(len(lista)):
		if lista[i]==elemento:
			trovato=1
			indice=i
			break
	if trovato==1:
		return indice
def leggiSigla(sheet):
	row=1
	Sigla =[]
	riga=sheet.row_values(row)
	Sigla = Sigla + [riga[3]]
	while row in range(sheet.nrows):	
		riga=sheet.row_values(row)
		nSigle=0
		for i in range(len(Sigla)):
			if riga[3]!=Sigla[i]:
				nSigle = nSigle +1
		if nSigle==len(Sigla):
			Sigla = Sigla + [riga[3]]
		row=row+1
	Sigla.sort()
	return Sigla
def ToOrizzontale(Sigla,sheet,nome): 
	ToOrizzontale	= xlwt.Workbook()
	foglioToOrizzontale =	ToOrizzontale.add_sheet('Data')

	# Scrivo la prima riga del foglio
	foglioToOrizzontale.write(0,0,'Pozzo')
	foglioToOrizzontale.write(0,1,'Data')
	for i in range(0,len(Sigla)):
		foglioToOrizzontale.write(0,i+2,Sigla[i])

	count = 0
	row   = 1
	while row in range(sheet.nrows):	
		Pozzo1	= str(sheet.cell_value(row,0))	# concatenza numero pozzo
		dateCell	= 	sheet.cell_value(row,1)
		Data1 	=	xlrd.xldate_as_tuple(dateCell,DatiRete.datemode)
		stop	=	0
		Istanza = []
		Istanza = Istanza +[Pozzo1]
		Istanza = Istanza + [Data1]
		# scrivo pozzo e data nel foglio
		foglioToOrizzontale.write(count+1,0, int(sheet.cell_value(row,0)))
		foglioToOrizzontale.write(count+1,1, datetime.date(Istanza[1][0],Istanza[1][1],Istanza[1][2]),date_format)
		while stop==0 and row in range(sheet.nrows):
			riga=sheet.row_values(row)
			if Pozzo1==str(riga[0]) and Data1==xlrd.xldate_as_tuple(riga[1],DatiRete.datemode):
				Istanza= Istanza + [riga[sheet.ncols-1]]
				for i in range(len(Sigla)):
					if riga[3]==Sigla[i]:
						break
				foglioToOrizzontale.write(count+1,i+2,riga[sheet.ncols-1])
				row=row+1
			else:
				stop=1
		count=count+1
	ToOrizzontale.save(nome+'.xls')
def foglioIntersezione(sheet1,sheet2,file1,file2):
	intersezione = xlwt.Workbook()
	sheetIntersezione = intersezione.add_sheet('Data')

	rigaFoglioUno = sheet1.row_values(0)
	rigaFoglioDue = sheet2.row_values(0)
	# trovo le colonne di intersezione
	colonneIntersezione = []
	for i in range(len(rigaFoglioUno)):
		if rigaFoglioUno[i] in rigaFoglioDue:
			if rigaFoglioUno[i] not in colonneIntersezione:
				colonneIntersezione=colonneIntersezione+ [rigaFoglioUno[i]]

	# scrivo la prima riga
	for i in range(len(colonneIntersezione)):
		sheetIntersezione.write(0,i,colonneIntersezione[i])

	# copio gli altri valori del foglio 1
	for row in range(1,sheet1.nrows):
		riga= sheet1.row_values(row)
		data = xlrd.xldate_as_tuple(riga[1],file1.datemode)
		nCol = 0
		for i in range(len(riga)):
			if rigaFoglioUno[i] in colonneIntersezione:
				if i==1:
					sheetIntersezione.write(row, nCol, datetime.date(int(data[0]),int(data[1]),int(data[2])),date_format )
				else:
					sheetIntersezione.write(row, nCol, riga[i])
				nCol=nCol+1

	# copio gli altri valori del foglio 2
	for rows in range(1,sheet2.nrows):
		riga=sheet2.row_values(rows)
		data = xlrd.xldate_as_tuple(riga[1],file2.datemode)
		nCol=0
		for i in range(len(riga)):
			if rigaFoglioDue[i] in colonneIntersezione:
				if i==1:
					sheetIntersezione.write(rows+row, nCol, datetime.date(int(data[0]),int(data[1]),int(data[2])),date_format)
				else:
					sheetIntersezione.write(rows+row, nCol,riga[i])
				nCol=nCol+1
	return intersezione
def preProcess(sheet,fileExcel, nome):
	newExcel = xlwt.Workbook()
	newSheet = newExcel.add_sheet("data")


	colonne = [1] * sheet.ncols
	nNulli	= [0] * sheet.ncols
	# Controllo quali colonne hanno elementi nulli
	for row in range(sheet.nrows):
		riga=sheet.row_values(row)
		for i in range(len(riga)):
			if riga[i]=="":
				colonne[i]=0
				nNulli[i]=nNulli[i]+1


	# Gli elementi con meno del 10 % di valori nulli per colonna vengono riempiti con le medie
	percentualeNulli	= [0]*sheet.ncols
	colonnePochiZeri	= [0]*sheet.ncols
	colonneTantiZeri	= [0]*sheet.ncols
	for j in range(2,sheet.ncols):
		percentualeNulli[j]=nNulli[j]*100/sheet.nrows
		if percentualeNulli[j]<10:
			colonnePochiZeri[j]=1
		else:
			colonneTantiZeri[j]=1


	# calcolo le medie delle colonne con pochi valori nulli
	mediane		= [0]*sheet.ncols
	varianza	= [0]*sheet.ncols
	for j in range(2,sheet.ncols):
		newColonna =[]
		if colonnePochiZeri[j]==1:
			colonna = sheet.col_values(j)
			del colonna[0]
			for i in range(len(colonna)):
				if colonna[i]!="":
					newColonna=newColonna+[colonna[i]]
		if newColonna!=[]:
			mediane[j]	= round(statistics.median(newColonna),	2)
			varianza[j]	= round(statistics.stdev(newColonna),	2)

	#print(varianza)
	# Calcolo le colonne da tenere
	for col in range(sheet.ncols):
		if colonnePochiZeri[col]==1 and colonneTantiZeri[col]==0 and varianza[col]>0.5:
			colonne[col]=1
		else:
			colonne[col]=0

	# TOLGO IL MERCURIO (Hg)
	primaRiga=sheet.row_values(0)
	colonnaMercurio = cercaElementoPerIndice(primaRiga,"Hg")
	colonne[colonnaMercurio]=0


	colonne[0]=1
	colonne[1]=1
	# Stampa prima riga
	count=0
	riga = sheet.row_values(0)
	for j in range(len(riga)):
		if colonne[j]==1:
			newSheet.write(0,count,riga[j])
			count=count+1

	# Riscrivo il foglio con le sole colonne da tenere
	data 	= []
	for row in range(1,sheet.nrows):
		count=0
		for j in range(sheet.ncols):
			if colonne[j]==1:
				riga=sheet.row_values(row)
				data = xlrd.xldate_as_tuple(riga[1],fileExcel.datemode)
				if j==1:	# scrivo la data
					newSheet.write(row,count, datetime.date(int(data[0]),int(data[1]),int(data[2])),date_format)
				elif j==0:	# scrivo il pozzo
					newSheet.write(row,count, int(riga[j]))
				elif riga[j]=="":
					newSheet.write(row,count,mediane[j])
				else:
					newSheet.write(row,count,riga[j])
				count=count+1


	newExcel.save(nome)
def excelToArff(nomeFile, doveSalvarlo):
	#print("Sto convertendo il file "+nomeFile+".xls in .arff")
	excelDaLeggere = nomeFile + ".xls"
	#if sheetIsEmpty(excelDaLeggere)==1:
	#	print("Il file è vuoto")

	File = xlrd.open_workbook(excelDaLeggere,'r')
	sheetFile = File.sheet_by_index(0)
	arffFile = open(doveSalvarlo+".arff","w")
	arffFile.write("@relation "+nomeFile+"\n\n")
	primaRiga=sheetFile.row_values(0)
	for i in range(len(primaRiga)):
		if i==1:
			if primaRiga[i]=="Data":
				arffFile.write("@attribute\t"+str(primaRiga[i])+"\tdate 'yyyy-MM-dd'\n")
			else:
				arffFile.write("@attribute\t"+str(primaRiga[i])+"\tnumeric\n")

		elif i==len(primaRiga)-1:
			arffFile.write("@attribute\t"+str(primaRiga[i])+"\t{1,0}\n")
		else:
			arffFile.write("@attribute\t"+str(primaRiga[i])+"\tnumeric\n")

	arffFile.write("\n@data\n")
	for row in range(1,sheetFile.nrows):
		riga=sheetFile.row_values(row)
		if primaRiga[1]=="Data":
			data = xlrd.xldate_as_tuple(riga[1],File.datemode)
			anno=str(data[0])
			if data[1]<10:
				mese="0"+str(data[1])
			else:
				mese=str(data[1])

			if data[2]<10:
				giorno="0"+str(data[2])
			else:
				giorno=str(data[2])
		for i in range(len(riga)):
			if i==1:
				if primaRiga[1]=="Data":
					arffFile.write(anno+"-"+mese+"-"+giorno+",")
				else:
					arffFile.write(str(riga[i])+",")
			elif i==0:
				arffFile.write(str(int(riga[0])))
				arffFile.write(",")
			elif riga[i]=="":
				arffFile.write("?,")
			elif i==len(primaRiga)-1:
				if riga[i]=="sì":
					arffFile.write("1,")
				elif riga[i]=="no":
					arffFile.write("0,")
				else:
					arffFile.write("?,")
			else:
				arffFile.write(str(riga[i])+",")
		arffFile.write("\n")
	arffFile.close()
def testAndTraining(sheet,foglio, modello, nome):
	# Creo un file di train e uno di test
	Training = xlwt.Workbook()
	validation = xlwt.Workbook()
	trainingSheet = Training.add_sheet("Data")
	validationSheet = validation.add_sheet("Data")

	countTrain=0
	countValidation=0
	for row in range(sheet.nrows):
		riga = sheet.row_values(row)
		if row!=0:
			data = xlrd.xldate_as_tuple(riga[1], foglio.datemode)
		# scrivi nel validation
		if riga[len(riga) - 1] == '' or riga[len(riga) - 1] == "incerto":
			for j in range(len(riga)):
				if row!=0:
					if j == 1:
						validationSheet.write(countValidation, j,datetime.date(int(data[0]), int(data[1]), int(data[2])), date_format)
					elif j == 0:
						validationSheet.write(countValidation, j, int(riga[j]))
					else:
						validationSheet.write(countValidation, j, riga[j])
				else:
					validationSheet.write(countValidation, j, riga[j])
			countValidation+=1
		else:
			for j in range(len(riga)):
				if row!=0:
					if j == 1:
						trainingSheet.write(countTrain, j, datetime.date(int(data[0]), int(data[1]), int(data[2])), date_format)
					elif j == 0:
						trainingSheet.write(countTrain, j, int(riga[j]))
					else:
						trainingSheet.write(countTrain, j, riga[j])
				else:
					trainingSheet.write(countTrain, j, riga[j])
			countTrain+=1

	Training.save("Altro/Training/"+nome+"Training" + modello + ".xls")
	validation.save("Altro/Validation/"+nome+"Validation" + modello + ".xls")
	return Training, trainingSheet
def notDominated(sheet, nome):
	print("Elaboro i non dominati...")
	NotDominated = xlwt.Workbook()
	sheetNotDominated = NotDominated.add_sheet("Data")

	count=0
	sheetNotDominated.write(0,0,sheet.cell_value(0,0))
	for row in range(1,sheet.nrows):
		val = sheet.cell_value(row,0)
		parametri = sheet.cell_value(row,1)
		dominated=0
		for wor in range(1, sheet.nrows):
			lav = sheet.cell_value(wor,0)
			parameters = sheet.cell_value(wor,1)
			#print(len(parametri.split(",")), val,"\t",len(parameters.split(",")), lav)
			if len(parametri.split(","))==len(parameters.split(",")) and val<lav:
				dominated = 1
				break
			elif len(parametri.split(","))>len(parameters.split(",")) and val<lav:
				dominated = 1
				break
		if dominated == 0:
			count+=1
			sheetNotDominated.write(count,0, val)
			sheetNotDominated.write(count,1, parametri)
	NotDominated.save("NotDominated/"+nome+"NotDominated.xls")
def evaluate(sheet, foglio, bestModello):

	names = ["Decision Tree", "Random Forest", "Logistic Regression"]
	classifiers = [
		DecisionTreeClassifier(max_depth=5),
		RandomForestClassifier(max_depth=5, n_estimators=10, max_features=1),
		LogisticRegression()
	]

	for obj in ["Modello", "Acquifero"]:
		print("\n"+obj)
		df = pd.read_excel("Unico/"+obj+"/Processed" + obj + ".xls")
		if obj == "Modello":
			yes = df.loc[df['modello' + bestModello] == 'sì']
			no = df.loc[df['modello' + bestModello] == 'no']
		elif obj == "Acquifero":
			yes = df.loc[df['Acquifero'] == 'sì']
			no = df.loc[df['Acquifero'] == 'no']

		for stagione in ["Primavera", "Estate", "Autunno", "Inverno", "Unico"]:
			print(stagione)
			for name, clf in zip(names, classifiers):
				count = 0
				print(name)
				valutatore = xlwt.Workbook()
				sheetValutatore = valutatore.add_sheet("Data")
				sheetValutatore.write(0,0, name)

				df = pd.read_excel(open("Altro/Training/"+stagione+"Training"+obj+ ".xls", "rb"), sheet_name='Data')
				target = df[df.columns[-1]]

				for i in range(2, len(df.columns) - 1):
					for j in range(i + 1, len(df.columns) - 1):
						set = df[[df.columns[i], df.columns[j]]]
						set = set.values
						target = pd.Series(target).values
						train_set = set[:-int(len(set) * .2)]
						test_set = set[int(len(set) * .8):]
						train_target = target[:-int(len(target) * .2)]
						test_target = target[int(len(target) * .8):]
						# target non deve avere valori nulli
						clf.fit(set, target)
						predicted = clf.predict(test_set)
						# Cross Validation
						n_samples, n_feature = set.shape
						cv = cross_validation.ShuffleSplit(n_samples, n_iter=10, test_size=0.4, random_state=0)
						scores = cross_validation.cross_val_score(clf, set, target, cv=cv, scoring='adjusted_rand_score')
						val = sum(scores) / cv.n_iter
						print("Evaluate: "+df.columns[i], df.columns[j], "\t",val)

						count+=1
						sheetValutatore.write(count,0,val)
						sheetValutatore.write(count,1,df.columns[i]+","+df.columns[j])
						# Valuto le triple
						for k in range(j + 1, len(df.columns) - 1):
							#print("Evaluate: "+df.columns[i], df.columns[j], df.columns[k])
							set = df[[df.columns[i], df.columns[j], df.columns[k]]]
							set = set.values
							target = pd.Series(target).values
							train_set = set[:-int(len(set) * .2)]
							test_set = set[int(len(set) * .8):]
							train_target = target[:-int(len(target) * .2)]
							test_target = target[int(len(target) * .8):]
							# target non deve avere valori nulli
							clf.fit(set, target)
							predicted = clf.predict(test_set)
							# Cross Validation
							n_samples, n_feature = set.shape
							cv = cross_validation.ShuffleSplit(n_samples, n_iter=10, test_size=0.4, random_state=0)
							scores = cross_validation.cross_val_score(clf, set, target, cv=cv, scoring='adjusted_rand_score')
							# print(scores)
							val = sum(scores) / cv.n_iter
							count += 1
							sheetValutatore.write(count, 0, val)
							sheetValutatore.write(count, 1, df.columns[i] + "," + df.columns[j]+"," + df.columns[k])
				valutatore.save("Altro/"+stagione+"/"+obj+"/"+name+".xls")
				valutatore = xlrd.open_workbook("Altro/"+stagione+"/"+obj+"/"+name+".xls")
				sheetValutatore = valutatore.sheet_by_index(0)

				notDominated(sheetValutatore,   stagione+"/"+obj+"/"+name)

				ND = xlrd.open_workbook("NotDominated/"+stagione+"/"+obj+"/"+name+"NotDominated.xls")
				sheetND = ND.sheet_by_index(0)

				print("Grafico i non dominati...")
				for row in range(1, sheetND.nrows):
					parametri = sheetND.cell_value(row, 1).split(",")
					print(parametri)
					if len(parametri) == 2:
						X_yes = yes[parametri[0]]
						Y_yes = yes[parametri[1]]
						X_no = no[parametri[0]]
						Y_no = no[parametri[1]]
						plt.title(sheetND.cell_value(row, 0))
						plt.xlabel(parametri[0])
						plt.ylabel(parametri[1])
						plt.scatter(X_yes, Y_yes, color='red')
						plt.scatter(X_no, Y_no, color='blue')
						plt.savefig("Graph/"+stagione+"/"+obj+"/" + parametri[0] + "_" + parametri[1] + ".png")
						plt.close()
					elif len(parametri) == 3:
						X_yes = yes[parametri[0]]
						Y_yes = yes[parametri[0]]
						X_no = no[parametri[1]]
						Y_no = no[parametri[1]]
						Z_yes = yes[parametri[2]]
						Z_no = no[parametri[2]]
						fig = plt.figure()
						plt.title(sheetND.cell_value(row, 0))
						ax = fig.add_subplot(111, projection='3d')
						ax.scatter(X_yes, Y_yes, Z_yes, c='r', marker='o')
						ax.scatter(X_no, Y_no, Z_no, c='b', marker='o')
						ax.set_xlabel(parametri[0])
						ax.set_ylabel(parametri[1])
						ax.set_zlabel(parametri[2])
						plt.savefig("Graph/"+stagione+"/"+obj+"/" + parametri[0] + "_" + parametri[1] + "_" + parametri[2] + ".png")
						plt.close()
	
	for obj in ["Modello", "Acquifero"]:
		Tipo = xlwt.Workbook()
		sheetTipo = Tipo.add_sheet("Data")
		count = 0
		lista = []
		for stagione in ["Primavera", "Estate", "Autunno", "Inverno", "Unico"]:
			for name in names:
				Valutazione = xlrd.open_workbook("NotDominated/"+stagione+"/"+obj+"/"+name+"NotDominated.xls")
				sheetValutazione = Valutazione.sheet_by_index(0)
				for i in range(1,sheetValutazione.nrows):
					val = sheetValutazione.cell_value(i,1)
					if cercaElemento(lista, val)==0:
						lista = lista + [val]
		for i in range(len(lista)):
			sheetTipo.write(count,0, lista[i])
			count+=1
		Tipo.save("NotDominated/"+obj+"NotDominated.xls")

	Acquifero = xlrd.open_workbook("NotDominated/AcquiferoNotDominated.xls")
	sheetAcquifero = Acquifero.sheet_by_index(0)

	Modello = xlrd.open_workbook("NotDominated/ModelloNotDominated.xls")
	sheetModello = Modello.sheet_by_index(0)

	colonnaAcquifero = sheetAcquifero.col_values(0)
	colonnaModello = sheetModello.col_values(0)

	elemComuni=0
	for i in range(len(colonnaAcquifero)):
		if cercaElemento(colonnaModello, colonnaAcquifero[i])==1:
			elemComuni+=1
	Confronto = elemComuni/(max(len(colonnaAcquifero), len(colonnaModello)))
	print("Il modello e l'acquifero hanno un valore di somiglianza: ",round(Confronto,2),"(",elemComuni,"/",max(len(colonnaAcquifero),len(colonnaModello)),"elementi comuni )")



def crea4Stagioni(sheet, fileExcel, Classe):
    Primavera = xlwt.Workbook()
    Estate    = xlwt.Workbook()
    Autunno   = xlwt.Workbook()
    Inverno   = xlwt.Workbook()

    sheetPrimavera = Primavera.add_sheet("Data")
    sheetEstate    = Estate.add_sheet("Data")
    sheetAutunno   = Autunno.add_sheet("Data")
    sheetInverno   = Inverno.add_sheet("Data")

    primaRiga = sheet.row_values(0)
    for i in range(len(primaRiga)):
        sheetPrimavera.write(0,i,primaRiga[i])
        sheetEstate.write(0,i,primaRiga[i])
        sheetAutunno.write(0,i,primaRiga[i])
        sheetInverno.write(0,i,primaRiga[i])

    count = [0]*4
    for row in range(1,sheet.nrows):
        riga = sheet.row_values(row)
        data = xlrd.xldate_as_tuple(riga[1], fileExcel.datemode)
        if isSpring(data)==1:
            count[0]+=1
            for j in range(len(riga)):
                if j == 1:
                    sheetPrimavera.write(count[0], j,datetime.date(int(data[0]), int(data[1]), int(data[2])),date_format)
                else:  # scrivo i valori delle sigle
                    sheetPrimavera.write(count[0], j, riga[j])
        elif isSummer(data)==1:
            count[1]+=1
            for j in range(len(riga)):
                if j == 1:
                    sheetEstate.write(count[1], j,datetime.date(int(data[0]), int(data[1]), int(data[2])),date_format)
                else:  # scrivo i valori delle sigle
                    sheetEstate.write(count[1], j, riga[j])
        elif isAutumn(data)==1:
            count[2]+=1
            for j in range(len(riga)):
                if j == 1:
                    sheetAutunno.write(count[2], j,datetime.date(int(data[0]), int(data[1]), int(data[2])),date_format)
                else:  # scrivo i valori delle sigle
                    sheetAutunno.write(count[2], j, riga[j])
        elif isWinter(data)==1:
            count[3]+=1
            for j in range(len(riga)):
                if j == 1:
                    sheetInverno.write(count[3], j,datetime.date(int(data[0]), int(data[1]), int(data[2])),date_format)
                else:  # scrivo i valori delle sigle
                    sheetInverno.write(count[3], j, riga[j])
    Primavera.save("Stagioni/"+Classe+"/Primavera.xls")
    Estate.save("Stagioni/"+Classe+"/Estate.xls")
    Autunno.save("Stagioni/"+Classe+"/Autunno.xls")
    Inverno.save("Stagioni/"+Classe+"/Inverno.xls")
SiglaDR = leggiSigla(sheetDatiRete)
SiglaAV = leggiSigla(sheetAltriValori)
ToOrizzontale(SiglaDR, sheetDatiRete,	"Altro/DatiReteOrizzontale" 				)
ToOrizzontale(SiglaAV, sheetAltriValori,"Altro/AltriValoriOrizzontale"				)
DatiReteOrizzontale 	= xlrd.open_workbook(r'Altro/DatiReteOrizzontale.xls'	  	)
AltriValoriOrizzontale 	= xlrd.open_workbook(r'Altro/AltriValoriOrizzontale.xls'	)
sheetDatiReteOrizzontale 		= 	DatiReteOrizzontale. 	sheet_by_index(0)
sheetAltriValoriOrizzontale 	= 	AltriValoriOrizzontale. sheet_by_index(0)
merge = foglioIntersezione(sheetDatiReteOrizzontale, sheetAltriValoriOrizzontale, DatiReteOrizzontale, AltriValoriOrizzontale)
merge.save("Dataset.xls")

Dataset 		= xlrd.open_workbook(r'Dataset.xls')
sheetDataset	= Dataset.sheet_by_index(0)

preProcess(sheetDataset,Dataset, 'Altro/Processed.xls')

Dataset 		= xlrd.open_workbook(r'Altro/Processed.xls')
sheetDataset	= Dataset.sheet_by_index(0)

############################################ Creo i 3 modelli da confrontare ###########################
creaModello(Dataset,"A")
creaModello(Dataset,"B")
creaModello(Dataset,"C")


A = xlrd.open_workbook("Altro/Modello/A.xls","r")
B = xlrd.open_workbook("Altro/Modello/B.xls","r")
C = xlrd.open_workbook("Altro/Modello/C.xls","r")

sheetA = A.sheet_by_index(0)
sheetB = B.sheet_by_index(0)
sheetC = C.sheet_by_index(0)

testAndTraining(sheetA,A,"A","")
testAndTraining(sheetB,B,"B","")
testAndTraining(sheetC,C,"C","")

#excelToArff("Altro/Training/TrainingA", "Altro/Training/TraningA")
#excelToArff("Altro/Training/TrainingB", "Altro/Training/TraningB")
#excelToArff("Altro/Training/TrainingC", "Altro/Training/TraningC")

######################################Classifico i modelli##########################################
#Splitto il train e il test secondo la logica 80-20
names = ["Decision Tree", "Random Forest", "Logistic Regression"]
modelli = ["A", "B", "C"]
medie = [0,0,0]
classifiers = [
	DecisionTreeClassifier(max_depth=5),
	RandomForestClassifier(max_depth=5, n_estimators=10, max_features=1),
	LogisticRegression()
]

for modello, media in zip(modelli, range(len(modelli))):
	print(modello)
	medie[media]=0
	for name, clf in zip(names, classifiers):
		df = pd.read_excel(open("Altro/Training/Training"+modello + ".xls", "rb"), sheet_name='Data')
		target = df['modello' + modello]
		set = df[df.columns[2:-1]].values
		target = pd.Series(target).values
		train_set = set[:-int(len(set) * .2)]
		test_set = set[int(len(set) * .8):]
		train_target = target[:-int(len(target) * .2)]
		test_target = target[int(len(target) * .8):]
		#print(name)
		clf.fit(set, target)
		predicted = clf.predict(test_set)
		# Cross Validation
		n_samples, n_feature = set.shape
		cv = cross_validation.ShuffleSplit(n_samples, n_iter = 10, test_size = 0.4, random_state = 0)
		scores = cross_validation.cross_val_score(clf, set, target, cv =cv, scoring = 'adjusted_rand_score')
        #print(scores)
		val = sum(scores)/cv.n_iter
		#print(val)
		medie[media]+=val
	medie[media]/=len(names)
	print("Media:\t",medie[media],"\n")
#print(medie)

####################### bestModello #################
if max(medie)==medie[0]:
	bestModello="A"
elif max(medie)==medie[1]:
	bestModello="B"
elif max(medie)==medie[2]:
	bestModello="C"
else:
	print("Error to evaluate best Model")
print("\nBest Modello:",bestModello)





Dataset = xlrd.open_workbook("Altro/Processed.xls")
sheetDataset = Dataset.sheet_by_index(0)

Finestra = [6,7,8]

aggiungiModello(Dataset, bestModello, "Unico/Modello/Processed")
aggiungiAcquifero(Dataset, Finestra, "Unico/Acquifero/Processed")

Acquifero = xlrd.open_workbook("Unico/Acquifero/ProcessedAcquifero.xls")
sheetAcquifero = Acquifero.sheet_by_index(0)
crea4Stagioni(sheetAcquifero, Acquifero, "Acquifero")

ProcessedBestModello = xlrd.open_workbook("Unico/Modello/ProcessedModello.xls")
sheetProcessed = ProcessedBestModello.sheet_by_index(0)
crea4Stagioni(sheetProcessed, ProcessedBestModello, "Modello")

testAndTraining(sheetAcquifero, Acquifero, "Acquifero", "")
testAndTraining(sheetProcessed, ProcessedBestModello, "Modello", "")

for esempi in ["Modello", "Acquifero"]:
	for stagione in ["Primavera", "Estate", "Autunno", "Inverno"]:
		Stagione = xlrd.open_workbook("Stagioni/"+esempi+"/"+stagione+".xls")
		sheetStagione = Stagione.sheet_by_index(0)

		testAndTraining(sheetStagione, Stagione, esempi, stagione)
testAndTraining(sheetProcessed, ProcessedBestModello, "Modello", "Unico")
testAndTraining(sheetAcquifero, Acquifero, "Acquifero", "Unico")

evaluate(sheetDataset, Dataset, bestModello)

#evaluate(sheetDataset, Dataset, bestModello, "", "Stagioni", "Modello")

#evaluate(sheetAcquifero, Acquifero, bestModello, "", "Unico", "Acquifero")
#evaluate(sheetAcquifero, Acquifero, bestModello, "", "Stagioni", "Acquifero")