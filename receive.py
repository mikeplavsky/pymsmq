import win32com.client	

qi = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
qi.PathName = r".\Private$\Tasks"

from constants import *
queue = qi.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)

while True:

	msg = queue.Receive(ReceiveTimeout='20000')

	if msg: 

		print( msg.Label )
		print( msg.Body )
		
	else: 

		print ( 'waiting for messages' )
