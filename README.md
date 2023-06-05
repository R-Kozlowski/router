# router
A robot capable of configuring network connections in a TMS. 

The idea behind building this kind of robot is to reduce the user's manual work. The program base on a START excel master file, which is optimized by VBA code to reduce user errors. All the user has to do is enter the data to be inserted in the TMS and choose link to the TMS environment. 
The robot mainly base on browser support with the Selenium module. It takes 12 seconds to enter a single network row, but by using the threading processes module and running two browsers at the same moment thanks to it, this time was reduced by 50%. The biggest advantage of the Selenium module is that user can still usign his computer at the same time and do something completely different.
The program uses waiting modules to reduce the number of wrong clicks done by robot and try/except conditions to handle errors and not interrupt his job.
The program handles critical errors and saves the data to a Crash table excel file, which can be recovered in the START file automatically by VBA code and continue the robot job in main.py file.
