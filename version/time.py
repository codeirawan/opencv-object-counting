from datetime import datetime
import time

start_time = datetime.now()
print(start_time)

while True:
	time.sleep(10)
	current_time = datetime.now()
	#get timedelta
	diff = current_time-start_time
	minutes = round(diff.total_seconds() / 60,2)
	print('Total difference in minutes: ', minutes)