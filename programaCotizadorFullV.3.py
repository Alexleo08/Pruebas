import pandas as pd
import requests
import json
from datetime import datetime
import sys

def formatoFecha(date):
    months = ("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    day = date.day
    month = months[date.month - 1]
    year = date.year
    hour = date.hour
    minute = date.minute
    messsage = "{} de {} del {} a las {}:{}".format(day, month, year, hour, minute)
    return messsage

def servicio():
    parameters = {'passengers':passengers,
    'serviceType':serviceType,
    'endingDate': f'{endingDate}',
    'mainSecurityAgent':mainSecurityAgent,
    'longitude':longitude,
    'lastName': f'{lastName}',
    'idEmergencyContact':idEmergencyContact,
    'mainChildSeat':mainChildSeat,
    'idCustomer':idCustomer,
    'fullDayActivity':{'isFlight':isFlight,
        'longitude':longitudeF,
        'time':f'{time}',
        'fullAddress':f'{fullAddress}',
        'nextDay':nextDay,
        'deviceId':f'{deviceId}',
        'idCustomer1':idCustomer1,
        'latitude':latitudeF,
        'observation':f'{observation}'},
    'idLocalApp':f'{idLocalApp}',
    'idServiceConfiguration':idServiceConfiguration,
    'startingDateTime': f'{startingDateTime}',
    'advancedvehicle':advancedvehicle,
    'deviceid':f'{deviceid}',
    'trackingvehicle':trackingvehicle,
    'trackingarmored':trackingarmored,
    'surname':f'{surname}',
    'email':f'{email}',
    'idCountryCode':idCountryCode,
    'mobileNumber':mobileNumber,
    'mainArmored':mainArmored,
    'trackingSegurityAgent':trackingSegurityAgent,
    'latitude':latitude,
    'name':f'{name}'}
    #URL = 'https://dev103b.gcsecurity.mx/api/ServiceSierra/quotation2' 
    URL = 'https://api-ep.gcsecurity.mx/api/ServiceSierra/quotation2' 
    data = requests.post(URL, auth=('angel', '123456'), json=parameters)
    #print("Código de Respuesta" , data.status_code)
    dataJson = data.json()
    #print(parameters)
    #print(dataJson)
    return dataJson

#archivoExcel = pd.read_excel('C:/Users/aleja/Desktop/PP/ejemploFullPruebas.xlsx') #para leer archivo excel
archivoExcel = pd.read_excel('C:/Users/aleja/Desktop/PP/ejemploFullPruebasProduccionDEF.xlsx') #para leer archivo excel
resultadoJson= pd.DataFrame() #se crea variable tipo DataFrame para el registro de las respuestas del servidor


    #Respuesta de la peticion que se hace al servidor
try:
    for index, row in archivoExcel.iterrows(): 
        passengers= row['passengers']
        serviceType= row['serviceType']
        endingDate= row['endingDate']
        mainSecurityAgent= row['mainSecurityAgent']
        longitude= row['longitude']
        lastName= row['lastName']
        idEmergencyContact= row['idEmergencyContact']
        mainChildSeat= row['mainChildSeat']
        idCustomer= row['idCustomer']
        isFlight= row['isFlight']
        longitudeF= row['longitudeF']
        time= row['time']
        fullAddress= row['fullAddress']
        nextDay= row['nextDay']
        deviceId= row['deviceId']
        idCustomer1= row['idCustomer1']
        latitudeF= row['latitudeF']
        observation= row['observation']
        idLocalApp= row['idLocalApp']
        idServiceConfiguration= row['idServiceConfiguration']
        startingDateTime= row['startingDateTime']
        advancedvehicle= row['advancedvehicle']
        deviceid= row['deviceid']
        trackingvehicle= row['trackingvehicle']
        trackingarmored= row['trackingarmored']
        surname= row['surname']
        email= row['email']
        idCountryCode= row['idCountryCode']
        mobileNumber= row['mobileNumber']
        mainArmored= row['mainArmored']
        trackingSegurityAgent= row['trackingSegurityAgent']
        latitude= row['latitude']
        name= row['name']
        resultadoJson = resultadoJson.append(servicio(),ignore_index=True)
        print('Consulta número: ', index)
except requests.exceptions.RequestException as ex: 
    f = open ('C:/Users/aleja/Desktop/PP/LogsCotizador.txt','a')
    now = datetime.now()
    f.write('\n' + formatoFecha(now) + '** CotizadorFull **' +'\n' + 'Nombre del error :'  + str(ex))
    f.close()
    print('Error al ejecutar la peticion al servidor:', ex)

    #Se utiliza para colocar nombre a las columnas y quitar los caracteres de los renglones de la consulta al json
resultadoJson['data'] =resultadoJson['data'].astype(str)
columnData =resultadoJson['data'].str.split(',', expand=True) #para dividir por columnas en base a ","
columData= columnData.rename(columns ={0:'retentionPercentage',1:'costOfServiceMXN',2:'retentionOfServiceMXN',3:'totalMXN',
                  4:'totalWithDiscountMXN',5:'costOfServiceUSD', 6:'retentionOfServiceUSD',7:'totalUSD',8:'totalWithDiscountUSD',
                  9:'tripZeroTerminal',10:'tripOneTerminal', 11:'meetingDateTimeDepartureFlight',12:'description',
                  13:'idQuotation',14:'relocationTime'}, inplace=True)
columnData['retentionPercentage'] = columnData['retentionPercentage'].str.replace('[','', regex=True)
columnData['retentionPercentage'] = columnData['retentionPercentage'].str.replace('{','', regex=True)
columnData['retentionPercentage'] = columnData['retentionPercentage'].str.replace('''retentionPercentage''','')
columnData['retentionPercentage'] = columnData['retentionPercentage'].str.replace("'",'')
columnData['retentionPercentage'] = columnData['retentionPercentage'].str.replace(": ",'')
columnData['costOfServiceMXN'] = columnData['costOfServiceMXN'].str.replace('''costOfServiceMXN''','')
columnData['costOfServiceMXN'] = columnData['costOfServiceMXN'].str.replace("'",'')
columnData['costOfServiceMXN'] = columnData['costOfServiceMXN'].str.replace(": ",'')
columnData['retentionOfServiceMXN'] = columnData['retentionOfServiceMXN'].str.replace('''retentionOfServiceMXN''','')
columnData['retentionOfServiceMXN'] = columnData['retentionOfServiceMXN'].str.replace("'",'')
columnData['retentionOfServiceMXN'] = columnData['retentionOfServiceMXN'].str.replace(": ",'')
columnData['totalMXN'] = columnData['totalMXN'].str.replace('''totalMXN''','')
columnData['totalMXN'] = columnData['totalMXN'].str.replace("'",'')
columnData['totalMXN'] = columnData['totalMXN'].str.replace(": ",'')
columnData['totalWithDiscountMXN'] = columnData['totalWithDiscountMXN'].str.replace('''totalWithDiscountMXN''','')
columnData['totalWithDiscountMXN'] = columnData['totalWithDiscountMXN'].str.replace("'",'')
columnData['totalWithDiscountMXN'] = columnData['totalWithDiscountMXN'].str.replace(": ",'')
columnData['costOfServiceUSD'] = columnData['costOfServiceUSD'].str.replace('''costOfServiceUSD''','')
columnData['costOfServiceUSD'] = columnData['costOfServiceUSD'].str.replace("'",'')
columnData['costOfServiceUSD'] = columnData['costOfServiceUSD'].str.replace(": ",'')
columnData['retentionOfServiceUSD'] = columnData['retentionOfServiceUSD'].str.replace('''retentionOfServiceUSD''','')
columnData['retentionOfServiceUSD'] = columnData['retentionOfServiceUSD'].str.replace(": ",'')
columnData['retentionOfServiceUSD'] = columnData['retentionOfServiceUSD'].str.replace("'",'')
columnData['totalUSD'] = columnData['totalUSD'].str.replace('''totalUSD''','')
columnData['totalUSD'] = columnData['totalUSD'].str.replace(": ",'')
columnData['totalUSD'] = columnData['totalUSD'].str.replace("'",'')
columnData['totalWithDiscountUSD'] = columnData['totalWithDiscountUSD'].str.replace('''totalWithDiscountUSD''','')
columnData['totalWithDiscountUSD'] = columnData['totalWithDiscountUSD'].str.replace(": ",'')
columnData['totalWithDiscountUSD'] = columnData['totalWithDiscountUSD'].str.replace("'",'')
columnData['tripZeroTerminal'] = columnData['tripZeroTerminal'].str.replace('''tripZeroTerminal''','')
columnData['tripZeroTerminal'] = columnData['tripZeroTerminal'].str.replace(": ",'')
columnData['tripZeroTerminal'] = columnData['tripZeroTerminal'].str.replace("'",'')

    #para concatenar en un solo DataFrame la respuesta del Servidor
union = pd.concat([resultadoJson, columnData], axis=1)
union=union.drop('data',axis=1)

    #para concatenar en un solo DataFrame la respuesta del Servidor con las peticiones que se hizo del archivo excel
archivoFinal = pd.concat([archivoExcel, union], axis=1)
print("******La Unión en un solo DataFrame se ha realizado*******"+ '\n')


#resultadoJson.to_excel('C:/Users/aleja/Desktop/PP/RespuestaCotizador/respuestaCotizadorFull.xlsx')
print("****** Convirtiendo el DataFrame a Excel*******" + '\n')
archivoFinal.to_excel('C:/Users/aleja/Desktop/PP/RespuestaCotizador/respuestaCotizadorFullProduccion1.xlsx')
print("****** Archivo excel terminado*******")

