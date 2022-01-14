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
    messsage = "{} de {} del {} a las {}:{}".format(day, month, year, hour,minute)
    return messsage

def servicio():
    parameters = {'mainsecurityagent':mainsecurityagent,
    'idlocalapp' : f'{idlocalapp}',
    'trackingsegurityagent': trackingsegurityagent,
    'servicetype': servicetype,
    'mainchildseat' : mainchildseat, 
    'mainarmored' : mainarmored,
    'idcustomer' : idcustomer,
    'idserviceconfiguration' : idserviceconfiguration, 
    'modality' : modality,
    'idemergencycontact' : idemergencycontact, 
    'idcountrycode' : idcountrycode,
    'surname' :f'{surname}',
    'latitude' : latitude, 
    'trackingarmored' : trackingarmored, 
    'trackingvehicle' : trackingvehicle, 
    'name' : f'{name}',
    'lastname' : f'{lastname}', 
    'email' : f'{email}',
    'passengers' : passengers, 
    'longitude' : longitude, 
    'advancedVehicle' : advancedVehicle, 
    'mobileNumber' : mobileNumber, 
    'deviceID' : f'{deviceID}',
    'tripOne':{'originFullAddress':f'{originFullAddress}',
        'meetingDateTime':f'{zmeetingDateTime}',
        'destinationLongitude':destinationLongitude,
        'destinationFullAddress':f'{destinationFullAddress}',
        'originLatitude':originLatitude,
        'destinationLatitude':destinationLatitude,
        'originLongitude':originLongitude},
    'tripZero':{'originFullAddress':f'{zoriginFullAddress}',
        'originLongitude':zoriginLongitude,
        'meetingDateTime': f'{meetingDateTime}',
        'originLatitude':zoriginLatitude,
        'destinationLongitude':zdestinationLongitude,
        'destinationLatitude':zdestinationLatitude,
        'destinationFullAddress':f'{zdestinationFullAddress}'}}
    #URL = 'https://dev103b.gcsecurity.mx/api/ServiceSierra/quotation2'
    URL = 'https://api-ep.gcsecurity.mx/api/ServiceSierra/quotation2' 
    data = requests.post(URL, auth=('angel', '123456'), json=parameters)
    #print("Código de Respuesta" , data.status_code)
    dataJson = data.json()
    #print(dataJson)
    return dataJson

#archivoExcel = pd.read_excel('C:/Users/aleja/Desktop/PP/ejemploRedondoPruebas.xlsx') #para leer archivo excel
archivoExcel = pd.read_excel('C:/Users/aleja/Desktop/PP/ejemploRedondo1.xlsx') #para leer archivo excel
resultadoJson = pd.DataFrame() #se crea variable tipo DataFrame para el registro de las respuestas del servidor

try:
        #Respuesta de la peticion que se hace al servidor
    for index, row in archivoExcel.iterrows(): 
        mainsecurityagent = row['mainsecurityagent']
        idlocalapp = ['idlocalapp']
        trackingsegurityagent = row['trackingsegurityagent']
        servicetype = row['servicetype']
        mainchildseat = row['mainchildseat']
        mainarmored = row['mainarmored']
        idcustomer = row['idcustomer']
        idserviceconfiguration = row['idserviceconfiguration']
        modality = row['modality']
        idemergencycontact = row['idemergencycontact']
        idcountrycode = row['idcountrycode']
        surname = row['surname']
        latitude = row['latitude']
        trackingarmored = row['trackingarmored']
        trackingvehicle = row['trackingvehicle']
        name = row['name']
        lastname = row['lastname']
        email = row['email']
        passengers = row['passengers']
        longitude = row['longitude']
        advancedVehicle = row['advancedVehicle']
        mobileNumber = row['mobileNumber']
        deviceID = row['deviceID']
        originFullAddress = row['originFullAddress']
        meetingDateTime = row['meetingDateTime']
        destinationLongitude = row['destinationLongitude']
        destinationFullAddress = row['destinationFullAddress']
        originLatitude = row['originLatitude']
        destinationLatitude = row['destinationLatitude']
        originLongitude = row['originLongitude']
        zoriginFullAddress = ['zoriginFullAddress']
        zoriginLongitude = row['zoriginLongitude']
        zmeetingDateTime = row['zmeetingDateTime']
        zoriginLatitude = row['zoriginLatitude']
        zdestinationLongitude = row['zdestinationLongitude']
        zdestinationLatitude = row['zdestinationLatitude']
        zdestinationFullAddress = row['zdestinationFullAddress']
        resultadoJson = resultadoJson.append(servicio(),ignore_index=True)
        print('Consulta número: ', index)
except requests.exceptions.RequestException as ex: 
    f = open ('C:/Users/aleja/Desktop/PP/LogsCotizador.txt','a')
    now = datetime.now()
    f.write('\n' + formatoFecha(now) + '** CotizadorRedondo **' +'\n' + 'Nombre del error :'  + str(ex))
    f.close()
    print('Error al ejecutar la peticion al servidor:', ex)

#Se utiliza para colocar nombre a las columnas y quitar los caracteres de los renglones de la consulta al json
resultadoJson['data'] =resultadoJson['data'].astype(str)
columnData =resultadoJson['data'].str.split(',', expand=True)#para dividir por columnas en base a ","
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

    #para concatenar en un solo DataFrame por columnas la respuesta del Servidor (Si son filas axis=0)
union = pd.concat([resultadoJson, columnData], axis=1)
union=union.drop('data',axis=1)

    #para concatenar en un solo DataFrame la respuesta del Servidor con las peticiones que se hizo del archivo excel
archivoFinal = pd.concat([archivoExcel, union], axis=1)
print("******La Unión en un solo DataFrame se ha realizado*******"+ '\n')

#resultadoJson.to_excel('C:/Users/aleja/Desktop/PP/RespuestaCotizador/respuestaCotizadorRedondo.xlsx')
print("****** Convirtiendo el DataFrame a Excel*******" + '\n')
archivoFinal.to_excel('C:/Users/aleja/Desktop/PP/RespuestaCotizador/redondoV1.xlsx')
print("****** Archivo excel terminado*******")
