import serial
import time
'''
---------------------------------------------
Python serial(Take systems)
[Overview]
JPADPT29-03のシリアル通信を行う

'''

class Take_serial(object):
    #Communication 
    self.STX=chr(0x02)#STX command
    self.ETX=chr(0x03)#ETX command
    self.RTX=chr(0x13)#RTX command
    self.com1='CL'
    self.com2='CR'
    ONcom='ON03FFFF'
    CHset='CH000001'
    OFcom='OF03FFFF'
    data=''
    command=(STX+ONcom+ETX).encode('utf-8') # ON command setting
    OFFCOM =(STX+OFcom+ETX).encode('utf-8')  #OFF command setting
    Channel=(STX+CHset+ETX).encode('utf-8') #channel setting
    ReadCOM=(STX+'CR14B1'+ETX).encode('utf-8')#Read command setting
#bit map image output command  PT0500##  '##' is imageNo
    def __init__(self,name):
        #Serial Poprtread
        use_port=search_com_port()

    def Serial_Write(self,Tx_data):
        '''
        ------------------------------------
        レジスタのWrite/Readを実行する
        -----------------------------------
        '''
        self.ser.write(Tx_data)
        while  self.data =='':
            self.data = self.ser.read(self.ser.in_waiting or 1)
            print(self.data.decode('utf-8'))
        return self.data.decode('utf-8')

    def Serial_read(self):
        return 0
    def Serial_close(self):
        #Portを閉じます
        self.ser.close()
        return 1
    def Serial_Open(self):
        self.ser.write(self.command)#ON command
        return 0

    def search_com_port(self):
        coms = serial.tools.list_ports.comports()
        comlist = []
        for com in coms:
            comlist.append(com.device)
        print('Connected COM ports: ' + str(comlist))
        use_port = comlist[0]
        print('Use COM port: ' + use_port)

        return use_port

if __name__ == "__main__":
    result = Take_serial()
    result.Serial_Open()
    print(result)


