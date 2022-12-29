import numpy as np
import math
from pandas import DataFrame
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator as tick
import os
import openpyxl


date = '221127'
liquid_list = ['Explosion','DI water', 'SSW', 'Tap water']


for liquid_type in liquid_list:
	plt.rcParams.update({'font.family':'th sarabun new', 'font.size':'14'})
	fig, ax = plt.subplots()
	second_ax = ax.twinx()

	MainDirectory = "C:/Users/pongkornmee/Downloads/P-nui master/Master suppression using Arduino/" + date
	PartName = date + '-' + liquid_type
	print(PartName)

	path = MainDirectory + '/' + PartName + '.xlsx'
	print(path)
	sheet = pd.read_excel(path, sheet_name = 'Sheet1')

	Cell_temp_list_1 = np.array(sheet['Cell-positive'])
	Cell_temp_list_2 = np.array(sheet['Cell-middle'])
	Cell_temp_list_3 = np.array(sheet['Cell-negative'])
	Time_list = np.array(sheet['Time'])
	Voltage_list = np.array(sheet['Voltage'])
	# deltalist_von = np.array(dfvon['delta'])
	# vonlist = np.array(dfvon['VM'])
	    
	line_1 = ax.plot(Time_list, Cell_temp_list_1, color='tab:red', label='Cell-positive')
	line_2 = ax.plot(Time_list, Cell_temp_list_2, color='tab:orange', label='Cell-middle')
	line_3 = ax.plot(Time_list, Cell_temp_list_3, color='tab:blue', label='Cell-negative')
	line_4 = second_ax.plot(Time_list, Voltage_list, color='tab:green',label='Voltage')
	lns = line_1 + line_2 + line_3 + line_4    
	labs = [l.get_label() for l in lns]
	# ax.legend(lns, labs, frameon=False, ncol=2, fontsize=12, loc='lower left')

   
	ax.tick_params(which='both', direction='in')
	ax.tick_params(which='major', length=2.5)
	second_ax.tick_params(which='both', direction='in')
	second_ax.tick_params(which='major', length=2.5)
	if liquid_type == 'Explosion':
		ax.set_ylim(0,max(Cell_temp_list_2)*1.2)
		ax.legend(lns, labs, frameon=False, ncol=1, fontsize=12, loc='center left')
	else:
		ax.set_ylim(0,max(Cell_temp_list_2)*1.2)
		ax.legend(lns, labs, frameon=False, ncol=2, fontsize=12, loc='lower left')
	ax.set_xlim(0,max(Time_list))
	second_ax.set_ylim(0,5)

	ax.set_xlabel('Time [s]', fontsize=12)
	ax.set_ylabel('Temperature [degC]', fontsize=12)
	second_ax.set_ylabel('Voltage [V]', fontsize=12)
	plt.title(liquid_type + ' ' + 'liquid submersion')

	plt.savefig(liquid_type + ' ' + 'liquid submersion.jpeg')
	# plt.show()
