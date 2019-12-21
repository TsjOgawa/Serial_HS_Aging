import serial
import time
'''
---------------------------------------------
Python serial(Take systems)

'''

class Take_serial(object):
    #Communication 
    STX=chr(0x02)#STX command
    ETX=chr(0x03)#ETX command
    RTX=chr(0x13)#RTX command
    ONcom='ON03FFFF'
    CHset='CH000001'
    OFcom='OF03FFFF'
    data=None
    command=(STX+ONcom+ETX).encode('utf-8') # ON command setting
    OFFCOM =(STX+OFcom+ETX).encode('utf-8')  #OFF command setting
    Channel=(STX+CHset+ETX).encode('utf-8') #channel setting
    ReadCOM=(STX+'CR14B1'+ETX).encode('utf-8')#Read command setting
#bit map image output command  PT0500##  '##' is imageNo
    def __init__(self,name):
        self.ser =serial.Serial(name,115200,timeout=1)
    def Serial_Write(self,Tx_data):
        '''
        ------------------------------------
        レジスタのWrite/Readを実行する
        -----------------------------------
        '''
        self.data=''
        self.ser.write(Tx_data)
        time.sleep(2)
        while  self.data == '':
            self.data=''
            self.data = self.ser.read(self.ser.in_waiting or 1)
           # print(self.data.decode('utf-8'))
        return self.data.decode('utf-8')

    def Serial_read(self):
        return 0
    def Serial_close(self):
        #Portを閉じます
        self.ser.close()
        return 1
    def Serial_Open(self):
        self.ser =serial.Serial('COM48',9600,timeout=None)
        self.ser.write(self.command)#ON command
        return 0

if __name__ == "__main__":
    result = Take_serial('COM48')
    err=result.Serial_Write((chr(0x02)+'ON03FFFF'+chr(0x03)).encode('utf-8'))
    print(err)
    time.sleep(1)
    com=result.Serial_Write((chr(0x02)+'CR14B1'+chr(0x03)).encode('utf-8'))
   # com=result.Serial_Write((chr(0x02)+'CR14B1'+chr(0x03)).encode('utf-8'))
   # com=result.Serial_Write(Take_serial.ReadCOM)
    print('read:'+com)
    result.Serial_close()


