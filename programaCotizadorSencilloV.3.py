import pandas as pd
import requests
import json as js
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
    parameters = {'email':f'{email}',
              'mainSecurityAgent': mainSecurityAgent,
              'mobileNumber': mobileNumber,
              'idEmergencyContact':idEmergencyContact,
              'longitude':longitude,
              'lastName':f'{lastName}',
              'serviceType':serviceType,
              'idLocalApp':f'{idLocalApp}',
              'idServiceConfiguration':idServiceConfiguration,
              'trackingVehicle':trackingVehicle,
              'deviceID':f'{deviceID}',
              'surname':f'{surname}',
              'mainChildSeat':mainChildSeat,
              'passengers':passengers,
              'mainArmored':mainArmored,
              'trackingSegurityAgent':trackingSegurityAgent,
              'advancedVehicle':advancedVehicle,
              'latitude':latitude,
              'tripZero':{'meetingDateTime': f'{meetingDateTime}',
                          'originLatitude':originLatitude,
                          'destinationLatitude':destinationLatitude,
                          'destinationLongitude':destinationLongitude,
                          'originFullAddress':originFullAddress,
                          'originLongitude':originLongitude,
                          'destinationFullAddress':destinationFullAddress},
              'name':f'{name}',
              'idCustomer':idCustomer,
              'idCountryCode':idCountryCode,
              'modality':modality,
              'trackingArmored':trackingArmored}
    #URL = 'https://dev103b.gcsecurity.mx/api/ServiceSierra/quotation2' 
    URL = 'https://api-ep.gcsecurity.mx/api/ServiceSierra/quotation2' 
    data = requests.post(URL, auth=('angel', '123456'), json=parameters) 
    #print("Código de Respuesta" , data.status_code)
    #print(parameters)
    dataJson = data.json()
    #print(dataJson)
    return dataJson

#archivoExcel = pd.read_excel('C:/Users/aleja/Desktop/PP/Respaldos Ptyhon/ejemploSencilloPruebas.xlsx') #para leer archivo excel
archivoExcel = pd.read_excel('C:/Users/aleja/Desktop/PP/SENCILLORevisionPasajeros.xlsx') #para leer archivo excel
resultadoJson = pd.DataFrame() #se crea variable tipo DataFrame para el registro de las respuestas del servidor

    #Respuesta de la peticion que se hace al servidor
try:
    for index, row in archivoExcel.iterrows(): 
        email = row['email']
        mainSecurityAgent = row['mainSecurityAgent']
        mobileNumber = row['mobileNumber']
        idEmergencyContact = row['idEmergencyContact']
        longitude = row['longitude']
        lastName = row['lastName']
        serviceType = row['serviceType']
        idLocalApp = row['idLocalApp']
        idServiceConfiguration = row['idServiceConfiguration']
        trackingVehicle = row['trackingVehicle']
        deviceID = row['deviceID']
        surname = row['surname']
        mainChildSeat = row['mainChildSeat']
        passengers = row['passengers']
        mainArmored = row['mainArmored']
        trackingSegurityAgent = row['trackingSegurityAgent']
        advancedVehicle = row['advancedVehicle']
        latitude = row['latitude']
        meetingDateTime = row['meetingDateTime']
        originLatitude = row['originLatitude']
        destinationLatitude = row['destinationLatitude']
        destinationLongitude = row['destinationLongitude']
        originFullAddress = row['originFullAddress']
        originLongitude = row['originLongitude']
        destinationFullAddress = row['destinationFullAddress']
        name = row['name']
        idCustomer = row['idCustomer']
        idCountryCode = row['idCountryCode']
        modality = row['modality']
        trackingArmored = row['trackingArmored']
        resultadoJson = resultadoJson.append(servicio(),ignore_index=True)
        print('Consulta número: ', index)
except requests.exceptions.RequestException as ex: 
    f = open ('C:/Users/aleja/Desktop/PP/LogsCotizador.txt','a')
    now = datetime.now()
    f.write('\n' + formatoFecha(now) + '** CotizadorSencillo **' +'\n' + 'Nombre del error :'  + str(ex))
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

#resultadoJson.to_excel('C:/Users/aleja/Desktop/PP/RespuestaCotizador/respuestaCotizadorSencillo.xlsx')
print("****** Convirtiendo el DataFrame a Excel*******" + '\n')
archivoFinal.to_excel('C:/Users/aleja/Desktop/PP/RespuestaCotizador/SencilloPasajeros.xlsx')
print("****** Archivo excel terminado*******")
