#!/usr/bin/env python

'''
Title:		Delivery Note Creation
Author:		Rob Phoenix (rob.phoenix@bt.com)
Usage: 		Creates multiple excel workbooks, each one a delivery note 
			based on each worksheet in a main inventory database excel 
			workbook.  Copies over store name, address & code; lists the 
			serial numbers of every device shipping to each store, grouped 
			by model name; and lists the quantities of each model and the 
			total number of boxes shipping.
Platform:	Windows
Warning:	Only works with .xls files, NOT with .xlsx files
Date: 		23/12/2014
Version: 	2.2
''' 

import re
import csv
import os
import xlrd
import xlwt
from xlwt import Workbook, easyxf
from xlrd import open_workbook,XL_CELL_TEXT
import win32gui
from win32com.shell import shell, shellcon

#Define a 'visual demarcation' line
vis_demarc_line = '\n\n' + ('=' * 100)

#Define output workbook styles
common_style = (
	'font: name Calibri, bold off, height 220;'
	'borders: left medium, right medium'
	)
all_borders_right_align = xlwt.easyxf(common_style + 
	', top medium, bottom medium;'
	'alignment: horizontal right, vertical centre;'
	)
all_borders_left_align = xlwt.easyxf(common_style + 
	', top medium, bottom medium;'
	'alignment: horizontal left, vertical centre;'
	)
all_borders_centre_align = xlwt.easyxf(common_style + 
	', top medium, bottom medium;'
	'alignment: horizontal centre, vertical centre;'
	)
left_and_right_border_left_align = xlwt.easyxf(common_style + 
	';''alignment: horizontal left, vertical centre;'
	)
delivery_note_title = xlwt.easyxf(
	'font: name Calibri, italic on, height 360;'
	'alignment: horizontal centre, vertical centre;'
	)

def format_delivery_note(deliverynote):	
	#Insert IT Services image
	deliverynote.insert_bitmap('ITSERVICES.bmp', 0, 1)
	#Create text & fields (row1,row2,col1,col2)
	deliverynote.write_merge(7,7,2,6,"DELIVERY NOTE",
		delivery_note_title)
	deliverynote.write_merge(9,9,0,3,"Delivery Address",
		all_borders_left_align)
	deliverynote.write_merge(9,9,5,6,"Shipment Date:",
		all_borders_right_align)
	deliverynote.write_merge(9,9,7,9,"",
		all_borders_centre_align)
	deliverynote.write_merge(11,11,5,6,"Consignment Type:",
		all_borders_right_align)
	deliverynote.write_merge(11,11,7,9,"",
		all_borders_centre_align)
	deliverynote.write_merge(13,13,5,6,"Your Ref:",
		all_borders_right_align)
	deliverynote.write_merge(13,13,7,9,"",
		all_borders_centre_align)
	deliverynote.write_merge(15,15,5,6,"Our Ref:",
		all_borders_right_align)
	deliverynote.write_merge(15,15,7,9,"",
		all_borders_centre_align)
	deliverynote.write_merge(17,17,5,6,"FAO:",
		all_borders_right_align)
	deliverynote.write_merge(17,17,7,9,"",
		all_borders_centre_align)
	deliverynote.write_merge(19,19,0,1,"Quantity",
		all_borders_centre_align)
	deliverynote.write_merge(19,19,2,5,"Description",
		all_borders_centre_align)
	deliverynote.write_merge(19,19,6,9,"Serial Number",
		all_borders_centre_align)

def main():
	#open browser window to enable user to choose file to create delivery 
	#notes from
	desktop_pidl = shell.SHGetFolderLocation (0, shellcon.CSIDL_DESKTOP, 
		0, 0)
	pidl, display_name, image_list = shell.SHBrowseForFolder (
	  win32gui.GetDesktopWindow (),
	  desktop_pidl,
	  "Please select the file you would \
like to create your delivery notes from.",
	  shellcon.BIF_BROWSEINCLUDEFILES,
	  None,
	  None
	)
	if (pidl, display_name, image_list) == (None, None, None):
	  print vis_demarc_line + "\n\n\tNothing selected"
	else:
	  fin = shell.SHGetPathFromIDList (pidl)
	#open Excel workbook
	workbook_in = xlrd.open_workbook(fin)
	#Iterate through worksheets
	print vis_demarc_line + "\n\n\t"
	for sheet in workbook_in.sheets():
		store = sheet.name
		worksheet = workbook_in.sheet_by_name(store)
		#Grab data from input workbook
		store_name = worksheet.cell(0,1)
		store_address_cell = sheet.cell(1,1)
		store_code = worksheet.cell(3,1)
		store_address = re.split("[,.]",store_address_cell.value)
		#Create dictionary to store address
		address_dict = {}
		#Load address into dictionary
		i = 1
		for address_line in store_address:
			address_dict[i] = address_line
			i += 1
		#Define devices in each worksheet, where all_devices is a 
		#list of devices
		switch_col = worksheet.col_values(4,6,15)
		ap_col = worksheet.col_values(1,17,36)
		list_of_switches = list(set(switch_col))
		list_of_switches.remove('')
		list_of_aps = list(set(ap_col))
		all_devices = list_of_switches + list_of_aps
		#Save sheet as temporary csv file for easier data collection
		with open(store + '-temp.csv', 'w') as csv_in:
			csv_temp = csv.writer(csv_in)
			for row in range(worksheet.nrows):
				csv_temp.writerow(worksheet.row_values(row))
		#Parse csv file for device and serial data
		device1 = all_devices[0]
		device2 = all_devices[1]
		device3 = None
		if len(all_devices) > 3:
			device3 = all_devices[2]
		device4 = all_devices[-1]
		#Create empty device lists, where the name of each list is a 
		#different device model
		device1_list = []
		device2_list = []
		device3_list = []
		device4_list = []
		#Build device lists with device serial numbers
		with open(store + '-temp.csv', 'r') as csv_out:
			lines = csv_out.readlines()
			for line in lines:
				if device1 in line and \
				(line.split(',')[-2] != ''):
					device1_list.append(line.split(',')[-2])
				elif device2 in line and \
				(line.split(',')[-2] != ''):
					device2_list.append(line.split(',')[-2])
				elif (len(all_devices) > 3) and \
				device3 in line and \
				(line.split(',')[-2] != ''):
					device3_list.append(line.split(',')[-2])
				elif device4 in line and \
				(line.split(',')[2] != ''):
					device4_list.append(line.split(',')[2])
		#Show progress to user
		print '\n\tCreating %s (%s) delivery note...\n' % (
			store_name.value, store_code.value
			)
		#Create delivery note workbook
		workbook_out = Workbook()
		del_note = workbook_out.add_sheet('Delivery Note',
			cell_overwrite_ok=True
			)
		#Format delivery note workbook
		format_delivery_note(del_note)
		#Write store name and address to delivery note
		store_code_statement = 'Store code: %s' % store_code.value
		del_note.write_merge(10,10,0,3,
			store_name.value,left_and_right_border_left_align)
		del_note.write_merge(11,11,0,3,
			address_dict.get(1),left_and_right_border_left_align)
		del_note.write_merge(12,12,0,3,
			address_dict.get(2),left_and_right_border_left_align)
		del_note.write_merge(13,13,0,3,
			address_dict.get(3),left_and_right_border_left_align)
		del_note.write_merge(14,14,0,3,
			address_dict.get(4),left_and_right_border_left_align)
		del_note.write_merge(15,15,0,3,
			address_dict.get(5),left_and_right_border_left_align)
		del_note.write_merge(16,16,0,3,
			address_dict.get(6),left_and_right_border_left_align)
		del_note.write_merge(17,17,0,3,
			store_code_statement,all_borders_centre_align)
		#Specify device quantities
		i = 20
		j = i + len(device1_list)
		k = j + len(device2_list)
		l = k + len(device3_list)
		total_no_of_devices = (len(device1_list) + len(device2_list) + 
			len(device3_list) + len(device4_list)
			)
		#write device model names to delivery note
		del_note.write_merge(i,i,2,5,device1,all_borders_centre_align)
		del_note.write_merge(j,j,2,5,device2,all_borders_centre_align)
		del_note.write_merge(k,k,2,5,device3,all_borders_centre_align)
		del_note.write_merge(l,l,2,5,device4,all_borders_centre_align)
		#Write device quantities to delivery note
		del_note.write_merge(i,i,0,1,len(device1_list),
			all_borders_centre_align)
		del_note.write_merge(j,j,0,1,len(device2_list),
			all_borders_centre_align)
		del_note.write_merge(k,k,0,1,len(device3_list),
			all_borders_centre_align)
		del_note.write_merge(l,l,0,1,len(device4_list),
			all_borders_centre_align)
		#Write device lists to excel file
		for item in device1_list:
			del_note.write_merge(i,i,6,9,item,all_borders_centre_align)
			i += 1
		for item in device2_list:
			del_note.write_merge(j,j,6,9,item,all_borders_centre_align)
			j += 1
		for item in device3_list:
			del_note.write_merge(k,k,6,9,item,all_borders_centre_align)
			k += 1
		for item in device4_list:
			del_note.write_merge(l,l,6,9,item,all_borders_centre_align)
			l += 1
		#Write number of boxes to delivery note
		total_boxes = '%d BOXES IN TOTAL' % total_no_of_devices
		row1 = total_no_of_devices + 20
		row2 = row1 + 1
		del_note.write_merge(row1,row2,0,9,
			total_boxes,all_borders_centre_align)
		#Delete temp csv files
		os.remove(store + '-temp.csv')
		#create 'Delivery Notes' folder within the same directory as the 
		#Excel spreadsheet if it doesn't already exist
		del_notes_dir = os.path.dirname(fin) + '\Delivery Notes\\'
		if not os.path.exists(del_notes_dir):
			os.makedirs(del_notes_dir)
		#Save excel file
		workbook_out.save(del_notes_dir + store_code.value + '-' 
			+ total_boxes + '.xls'
			)
	print (
		vis_demarc_line + 
		"\n\n\tYour delivery notes have been created here:\n\n\t%s" 
		% del_notes_dir + vis_demarc_line + '\n\n'
		)

if __name__ == '__main__':
	main()