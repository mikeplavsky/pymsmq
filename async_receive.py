import win32com.client, pythoncom

qi = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
qi.PathName = r".\Private$\Tasks"

from constants import *
queue = qi.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)

class Evs:
	
	def OnArrived(self,queue,cursor):
		
		msg = queue.Receive()
		print( msg.Label )

ev = win32com.client.DispatchWithEvents("MSMQ.MSMQEvent",Evs)
queue.EnableNotification( Event = ev, ReceiveTimeout = 10000 )

while True:

	pythoncom.PumpWaitingMessages()
	
	import time
	time.sleep(2)
	
	