import sys, os, pendulum, glob, traceback
from win32com.shell import shell, shellcon

def log(logtype, errortype, arg1, arg2, arg3):
	x = shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0)
	f = open(x + '\\Lab_Calcs_Log.txt', 'a')
	if logtype == 'start':
		f.write('\n{0} Started program'.format(pendulum.now()))
	elif logtype == 'userend':
		f.write('\n{0} Program terminated properly by user'.format(pendulum.now()))
	elif logtype == 'Cellplatecalc':
		f.write('\n{0} Cell plating calculation complete. Want = {1}, Have = {2}, VolNeeded = {3}'.format(pendulum.now(), str(arg1), str(arg2), str(arg3)))
	elif logtype == 'UV':
		f.write('\n{0} UV calculation. Flux = {1}, Fluence = {2}, Time(sec) = {3}'.format(pendulum.now(), str(arg1), str(arg2), str(arg3)))
	else:
		if errortype == 'invalidinput':
			f.write('\n{0} ***Invalid input for plate cell calculations'.format(pendulum.now()))
		elif errortype == 'typeerror':
			f.write('\n{0} TypeError for choosing place cell calculation'.format(pendulum.now()))
		elif errortype == 'unexpectedtypeend':
			f.write('\n{0} ***TypeError for input to end program'.format(pendulum.now()))
		elif errortype == 'ZeroDivisionError':
			f.write('\n{0} ***Amazingly, user was able to fenagle a divide by zero. Somehow.'.format(pendulum.now()))
		else:
			f.write('\n{0} ***Undocumented error'.format(pendulum.now()))
		
def UV_exposuretimecalc():
	flux = float(input('What is the current reading on the detector? Detector units should be in middle position.\n')) / 100
	fluence = float(input('How much UVA do you want? (in J/m2)?\n'))
	time = fluence / flux
	print('Time (sec):', time)
	print('Time (min):', int(time/60), 'minutes', int(time - int(time/60)*60), 'seconds')
	log('UV', '', flux, fluence, time)

def platecellcount_nonadjusted():
	whatwehavenotadjusted = input('What was the cell count (not adjusted for trypan blue)?\n')
	whatwewant = input('How many total cells do we want in the end?\n')
	finalvol = float(whatwewant) / (float(whatwehavenotadjusted)/2)
	print(finalvol)
	log('cellplatecalc', '', whatwewant, round(float(whatwehavenotadjusted)/2, 4), round(float(finalvol), 4))
	
def platecellcount_adjusted():
	whatwehaveadjusted = input('What was the cell count after trypan blue adjustment per mL?\n')
	whatwewant = input('How many total cells do we want in the end?\n')
	finalvol = float(whatwewant) / float(whatwehaveadjusted)
	print(finalvol)
	log('cellplatecalc', '', whatwewant, round(float(whatwehavenotadjusted)/2, 4), round(float(finalvol), 4))

	
if __name__ == '__main__':
	
	x = shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0)
	filenum = glob.glob(x + '\\Lab_Calcs_Log.txt')
	if len(filenum) == 0:
		f = open(x + '\\Lab_Calcs_Log.txt', 'w+')
	#This returns a FileNotFound error, I need to find a way to make the Lab_Calcs_Log.txt file here somehow.
	
	log('start', '', '', '', '')
	while True:
		
		#Cell plating calculations here
		mainchoice = input('\n"p"\t--->\tCell count calculations for plating.\n"u"\t--->\tUVA time calculation\n\nInput a command: ')
		if mainchoice == 'p' or mainchoice == 'P':
			try:
				plateinput1 = input('Have you adjusted for trypan blue yet? (Type y if Countess is not being used for cell count) y/n: ')
				if plateinput1.lower() == 'y':
					platecellcount_adjusted()
				elif plateinput1.lower() == 'n':
					platecellcount_nonadjusted()
				else:
					print('\nInvalid input, returning to main menu.')
					log('error', 'invalidinput', '', '', '')
			except TypeError as localerr:
				print('***Unexpected character type, try again. \nError details: ', localerr, 'n')
				log('error', 'typeerror', '', '', '')
			except ZeroDivisionError as localerr:
				print('***I\'m not sure how you managed to get this error message. \nError details: ', localerr, 'n')
				log('error', 'ZeroDivisionError', '', '', '')
			except:
				print('***Unexpected error')
				log('error', '', '', '', '')
		
		elif mainchoice == 'u' or mainchoice == 'U':
			try:
				UV_exposuretimecalc()
			except:
				print('***Unexpected error')
				log('error', '', '', '', '')
		else:
			print('Invalid input, try again')
		
		#Asks the user if they want to do another calculation or close the program
		exitinput = input('Do another calculation? y/n: ')
		if exitinput == 'n' or exitinput == 'N':
			try:
				break
			except TypeError as localerr:
				print('***Unexpected character type, exiting program. \nError details: ', localerr, '\n')
				log('error', 'unexpectedtypeend', '', '', '')
				break
		else:
			pass
	
	log('userend', '', '', '', '')
	sys.exit()