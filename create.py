import win32com.client	

qi = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
qi.PathName = r".\Private$\Tasks"

qi.Create()
