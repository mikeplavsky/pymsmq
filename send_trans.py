import win32com.client	

qi = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
qi.PathName = r".\Private$\Tasks"

from constants import *
queue = qi.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)

msg = win32com.client.Dispatch("MSMQ.MSMQMessage")

d = win32com.client.Dispatch("MSMQ.MSMQTransactionDispenser")
tr = d.BeginTransaction()

for i in range(0,10):

	msg.Label = "Task " + str(i)
	msg.Body = "{report:12}"

	msg.Send( queue, tr )
	
tr.Commit()
	
	