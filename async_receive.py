import win32com.client, pythoncom

qi = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
qi.PathName = r".\Private$\Tasks"

from constants import *
queue = qi.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)

class Evs:

	Called = False

	def OnArrived(self,q,c):
		
		msg = queue.Receive()
		print( msg.Label )		
		
		print ('got message')
		Evs.Called = True

ev = win32com.client.DispatchWithEvents("MSMQ.MSMQEvent",Evs)		
queue.EnableNotification( Event = ev, ReceiveTimeout = 10000 )

while True:

	if Evs.Called: 
	
		print ('setting event')
		
		ev = win32com.client.DispatchWithEvents("MSMQ.MSMQEvent",Evs)		
		queue.EnableNotification( Event = ev, ReceiveTimeout = 10000 )
		
		Evs.Called = False
	
	pythoncom.PumpWaitingMessages()	
	
	import time
	time.sleep(0.2)
	
	
	
	