import openpyxl
import os

celda = []
fileName = ""


def ReadFile(): # 1 -- leer el archivo excel
    theFile = openpyxl.load_workbook( 'CodesNames.xlsx' )
    allSheetNames = theFile.sheetnames

    print( "===========================================" )
    print( "Extracting data from {}".format( theFile.sheetnames ) )
    print( "===========================================" )

    for sheet in allSheetNames:
        currentSheet = theFile[sheet]
        for row in range( 1, currentSheet.max_row + 1 ):
            for column in "C":  # al agregar mas letras, agrego más columnas
                C = "{}{}".format( column, row ) # toda la columna A
                ColumnC = currentSheet[C].value

                Code_NAME( ColumnC )

def Code_NAME( codes_names ):
    if codes_names != "":
        celda.append( codes_names )



def loadName():   
    for name in celda:
        dato = name.index('_') # busca el signo _
        fileName = name[:dato]+".jpg" # guarda lo que hay antes del _   ejemplo: 0036
        
        n = name.rstrip() #quita los espacios adelante y atras del nombre
        newName = n+".jpg" #nuevo nombre, ejemplo 0268_ALFREDO E. ROJAS  CORDERO.jpg
        
        try:
            os.rename(r'C:\\PhotoScan\\Pictures\\'+fileName+'',r'C:\\PhotoScan\\Pictures\\'+newName+'')
        except:
            print("error")
            




ReadFile()
loadName()