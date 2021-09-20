#!/usr/bin/python3.7
'''
This script collecting data from sensors
'''

import troykahat
import time
import statistics
import logging
from time import sleep

#Logger configuration
logger=logging.getLogger()
logger.setLevel(logging.INFO)
formatter=logging.Formatter('%(asctime)s %(name)s %(levelname)s: %(message)s')
#Logger to file
fh=logging.FileHandler('/home/pi/projects/results/logs/Chili_sensors.log')
fh.setLevel(logging.INFO)
fh.setFormatter(formatter)
logger.addHandler(fh)
#Logger to console
ch=logging.StreamHandler()
ch.setLevel(logging.INFO)
ch.setFormatter(formatter)
logger.addHandler(ch)

logger.info('Script started')

def SoilMoisture_sensor (PIN_AP_SOIL,PIN_WP_MOSF):
    '''
    This module connecting resistive soil moisture sensor via mosfet N-type (normally disabled) and reading indications.
    '''
    try:
        #Digital pin for controlling mosfet
        wp = troykahat.wiringpi_io()
        wp.pinMode(PIN_WP_MOSF, wp.OUTPUT)
        #Analog pin for soil moisture sensor 
        ap = troykahat.analog_io()
        ap.pinMode(PIN_AP_SOIL, ap.INPUT)
        
        try:
            #Changing mosfet statement
            wp.digitalWrite(PIN_WP_MOSF, True) #set True for enable
            logger.info(f'Digital pin {PIN_WP_MOSF} enabled, mosfet opened')
            sleep(1.0)
            
            #List for collecting sensor's indications
            result_list=[]
            #Reading 10 indications
            i=0
            while (i<10):
                i=i+1
                #Reading sensor
                SoilMoisture=ap.analogRead(PIN_AP_SOIL)
                logger.debug(f'Analog pin {PIN_AP_SOIL} with SoilMoisture sensor read information: {SoilMoisture}')
                #Changing view of result and append final list
                SoilMoisture_human_readable = SoilMoisture*100
                logger.info(f'Analog pin {PIN_AP_SOIL} with SoilMoisture sensor got value: {SoilMoisture_human_readable}')
                result_list.append(SoilMoisture_human_readable)
                sleep(0.2)
        
        finally:
            #Changing mosfet statement
            wp.digitalWrite(PIN_WP_MOSF, False)
            logger.info(f'Digital pin {PIN_WP_MOSF} disabled, mosfet closed')
        
        #Find the average value
        average_SoilMoisture=(statistics.mean (result_list))
        logger.info(f'Average value for this measuring session is: {average_SoilMoisture}')
        #Current time
        tym = time.localtime()
        current_time = time.strftime("%d/%m/%Y %H:%M:%S",tym)
        current_time_short = time.strftime("%d/%m/%Y %H:%M",tym)
        logger.debug(f'Current time is: {current_time}')
        #Write data to file
        path=f'/home/pi/projects/results/SoilMoisture-{PIN_AP_SOIL}.csv'
        with open(path,'a') as file:
            file.write (f'{PIN_AP_SOIL}|{current_time_short}|{average_SoilMoisture}\n')
            logger.debug(f'Saved data to {path}')

    except Exception as e:
        logging.exception('Exception occured')

def Themperature_sensor(PIN_AP_TEMP):
    '''
    This module reading indications of themperature sensor.
    '''
    try:
        #Analog pin for themperature sensor 
        ap = troykahat.analog_io()
        ap.pinMode(PIN_AP_TEMP, ap.INPUT)
        
        #Reading sensor
        TempValue=ap.analogRead(PIN_AP_TEMP)
        logger.debug(f'Analog pin {PIN_AP_TEMP} with themperature sensor read information: {TempValue}')
        #Changing view of result
        TempValue_V=3.3*TempValue
        TempValue_mV=1000*TempValue_V
        TempValue_C = TempValue_V*100-50
        logger.info(f'Themperature is: {TempValue_C}')
        #Current time
        tym = time.localtime()
        current_time = time.strftime("%d/%m/%Y %H:%M:%S",tym)
        current_time_short = time.strftime("%d/%m/%Y %H:%M",tym)
        logger.debug(f'Current time is: {current_time}')
        #Write data to file
        path='/home/pi/projects/results/Themperature.csv'
        with open(path,'a') as file:
            file.write (f'{PIN_AP_TEMP}|{current_time_short}|{TempValue_C}\n')
            logger.debug(f'Saved data to {path}')
        
    except Exception as e:
        logging.exception('Exception occured')

plant_list=(1,2,3,4,5,6)

def main():
    for each_plant in plant_list:
        SoilMoisture_sensor(each_plant,0)
    Themperature_sensor(0)
    
if __name__=='__main__':
    main()

logger.info('Script finished')