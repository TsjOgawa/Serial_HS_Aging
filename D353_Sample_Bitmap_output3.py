 #-*- coding: utf-8 -*-
import serial
import time
#Communication 
STX=chr(0x02)#STX command
ETX=chr(0x03)#ETX command
RTX=chr(0x13)#RTX command
ONcom='ON03FFFF'
CHset='CH000001'
OFcom='OF03FFFF'
command=(STX+ONcom+ETX).encode('utf-8') # ON command setting
OFFCOM=(STX+OFcom+ETX).encode('utf-8')  #OFF command setting
Channel=(STX+CHset+ETX).encode('utf-8') #channel setting
ReadCOM=(STX+'CR14B1'+ETX).encode('utf-8')#Read command setting
#bit map image output command  PT0500##  '##' is imageNo
BMP_1 	=	(STX+'PT050000'+ETX).encode('utf-8') #BMP image1 select
BMP_2	=	(STX+'PT050001'+ETX).encode('utf-8') #BMP image2 select
BMP_3	=	(STX+'PT050002'+ETX).encode('utf-8') #BMP image3 select
BMP_4	=	(STX+'PT050003'+ETX).encode('utf-8') #BMP image4 select
BMP_5	=	(STX+'PT050004'+ETX).encode('utf-8') #BMP image5 select
BMP_6	=	(STX+'PT050005'+ETX).encode('utf-8') #BMP image6 select
BMP_7	=	(STX+'PT050006'+ETX).encode('utf-8')#BMP image7 select
BMP_8	=	(STX+'PT050007'+ETX).encode('utf-8')#BMP image8 select
BMP_9	=	(STX+'PT050008'+ETX).encode('utf-8')#BMP image9 select
BMP_10	=	(STX+'PT050009'+ETX).encode('utf-8')#BMP image10 select
BMP_11	=	(STX+'PT05000A'+ETX).encode('utf-8')#BMP image11 select
BMP_12	=	(STX+'PT05000B'+ETX).encode('utf-8')#BMP image12 select
BMP_13	=	(STX+'PT05000C'+ETX).encode('utf-8')#BMP image13 select
BMP_14	=	(STX+'PT05000D'+ETX).encode('utf-8')#BMP image14 select
BMP_15	=	(STX+'PT05000E'+ETX).encode('utf-8')#BMP image15 select
BMP_16	=	(STX+'PT05000F'+ETX).encode('utf-8')#BMP image16 select
BMP_17	=	(STX+'PT050010'+ETX).encode('utf-8')#BMP image17 select
BMP_box=['PT050001',
'PT050002','PT050003','PT050004','PT050005','PT050006','PT050007','PT050008','PT050009','PT05000A','PT05000B','PT05000C','PT05000D','PT05000E','PT05000F',
'PT050010','PT050011','PT050012','PT050013','PT050014','PT050015','PT050016','PT050017','PT050018','PT050019','PT05001A','PT05001B','PT05001C','PT05001D','PT05001E','PT05001F',
'PT050020','PT050021','PT050022','PT050023','PT050024','PT050025','PT050026','PT050027','PT050028','PT050029','PT05002A','PT05002B','PT05002C','PT05002D','PT05002E','PT05002F',
'PT050030','PT050031','PT050032','PT050033','PT050034','PT050035','PT050036','PT050037','PT050038','PT050039','PT05003A','PT05003B','PT05003C','PT05003D','PT05003E','PT05003F',
'PT050040','PT050041','PT050042','PT050043','PT050044','PT050045','PT050046','PT050047','PT050048','PT050049','PT05004A']

ser =serial.Serial('COM49',9600,timeout=1)#Win  serial port setting
ser.write(command)#ON command
#ser.write(RED)#Fill image (red)
#ser.write(Channel)# channel select
#ser.write(ReadCOM)# read register

for bitdata in BMP_box:
    ser.write((STX+bitdata+ETX).encode('utf-8'))
    data = ser.read(ser.in_waiting or 1)
#ser.write(OFcom)    
ser.close()
