#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import sys
import ttk
import json
import time
import win32print
import win32ui
import win32con
import time
import threading
import qrcode 
import tkFileDialog
import Tkinter
import win32gui
from PIL import Image, ImageWin,ImageDraw,ImageFont

reload(sys) 
sys.setdefaultencoding('gb18030') 

mcuType = ''
font1 = ImageFont.truetype('msyh.ttf', 16)
font2 = ImageFont.truetype('msyh.ttf', 20)
font3 = ImageFont.truetype('msyh.ttf', 30)
currentDate = time.strftime("%Y-%m-%d", time.localtime())

def printerList():
	printerList = []
	for printerItem in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1):
	 	flags, desc, name, comment = printerItem
	 	printerList.append(name.encode('utf-8'))

	return printerList


def print2Printer(devNo):
	qr1=qrcode.QRCode(version=1,  
                 error_correction=qrcode.constants.ERROR_CORRECT_L,  
                 box_size=8,  
                 border=4,  
                 )  
	qr1.add_data(devNo)  
	qr1.make(fit=True)  
	img1=qr1.make_image() 

	img1 = img1.resize((80,80),Image.ANTIALIAS)
	newImg  = Image.new("RGB",(128,128),(255,255,255))
	img1Size = img1.size
	box = (0, 0, img1Size[0], img1Size[1])
	box1 = (24,24,104,104)
	region = img1.crop(box)
	newImg.paste(region,box1)
	newImgDraw = ImageDraw.Draw(newImg)
	newImgDraw.ink = 0 + 0 * 256 + 0 * 256 * 256
	newImgDraw.text((0,100),devNo[3:],font = font1)
	newImg.resize((80,80),Image.ANTIALIAS)


	#the second tag
	tempStringModel = modelChosen.get()
	tempStringNo = devNo

	#tempStringModel = tempStringModel.replace("LY", "L  Y")
	#tempStringModel = tempStringModel.replace("PA", "P   A")
	#tempStringNo = tempStringNo.replace("LY", "L  Y")
	#tempStringNo = tempStringNo.replace("PA", "P   A")

	deviceString = '品名: '.decode('utf-8') + deviceChosen.get()
	modelString = '型号: '.decode('utf-8') + tempStringModel
	serialString= '序列号: '.decode('utf-8') + tempStringNo[3:]
	companyString = '深圳市丰利源节能科技有限公司'.decode('utf-8')

	qr2 = qrcode.QRCode(version=1,  
                 error_correction=qrcode.constants.ERROR_CORRECT_L,  
                 box_size=8,  
                 border=4,  
                 )  
	qr2.add_data(devNo+';'+currentDate)  
	qr2.make(fit=True)
	img2 = qr2.make_image() 
	img2 = img2.resize((80,80),Image.ANTIALIAS)

	newImg1  = Image.new("RGB",(350,200),(255,255,255))
	imgSize1 = img2.size
	box = (0, 0, imgSize1[0], imgSize1[1])
	box1 = (220,50,300,130)
	region = img2.crop(box)
	newImg1.paste(region,box1)
	a2 = ImageDraw.Draw(newImg1)
	a2.ink = 0 + 0 * 256 + 0 * 256 * 256
	a2.text((100,20),"合格证".decode('utf-8'),font=font3)
	a2.text((10,60),deviceString,font=font2)
	a2.text((10,90),modelString,font=font2)
	a2.text((10,120),serialString,font=font2)
	a2.text((10,150),companyString,font=font2)
	newImg1.resize((175,100),Image.ANTIALIAS)

	hDC = win32ui.CreateDC ()
	hDC.CreatePrinterDC (printerChosen.get())
	hDC.StartDoc ('qrcode')
	hDC.StartPage ()
	hDC.SetMapMode (win32con.MM_TWIPS)

	'''
	hDC.DrawText (serialString.encode('gbk'), (300, -800, 1440 * 8, -1300), win32con.DT_LEFT)  #win32con.DT_CENTER
	hDC.DrawText ("品名:温度传感器".decode('utf-8').encode('gbk'), (300, -200, 4200, -700), win32con.DT_CENTER)  #win32con.DT_CENTER
	hDC.DrawText (deviceString.encode('gbk'), (300, -200, 4200, -700), win32con.DT_LEFT)  #win32con.DT_CENTER
	hDC.DrawText (modelString.encode('gbk'), (300, -500, 1440 * 8, -900), win32con.DT_LEFT)  #win32con.DT_CENTER
	hDC.DrawText (serialString.encode('gbk'), (300, -800, 1440 * 8, -1300), win32con.DT_LEFT)  #win32con.DT_CENTER
	hDC.DrawText (companyString.encode('gbk'), (300, -1500, 1440 * 8, -1800), win32con.DT_LEFT)  #win32con.DT_CENTER
	'''

	dib = ImageWin.Dib (newImg)
	dib.draw (hDC.GetHandleOutput (), (400, 0, 1600, -1200))

	dib1 = ImageWin.Dib (newImg1)
	dib1.draw (hDC.GetHandleOutput (), (2200, 0, 4300, -1200))

	hDC.EndPage ()
	hDC.EndDoc ()
	hDC.DeleteDC ()

	return

def deviceChosenEvent(event):
	configFile = open("./devices.json",'r')
	configList = json.load(configFile)

	currentDevice = deviceChosen.get()
	modelList = []
	for device in configList:
		if currentDevice==device['deviceName']:
			modelList =  device['deviceModels']
			break

	deviceNewModels = []	
	for model in modelList:
		deviceNewModels.append( model['deviceModel'] ) 
	configFile.close()
	
	modelChosen['values'] = deviceNewModels
	modelChosen.current(0)
	root.update()
	return

def modelChosenEvent(event):
	global mcuType
	configFile = open("./devices.json",'r')
	configList = json.load(configFile)

	currentDevice = deviceChosen.get()
	currentModel = modelChosen.get()

	#匹配设备
	modelList = []
	for device in configList:
		if currentDevice==device['deviceName']:
			modelList =  device['deviceModels']
			break

	#匹配设备型号
	for model in modelList:
		if currentModel == model['deviceModel']:
			mcuType = model['mcuType']
			break
	configFile.close()

	return


def btnChooseFirmWare(event):
	filePath = tkFileDialog.askopenfilename(title=u"选择文件",initialdir=(os.path.expanduser('./')))
	fileInfo.set(filePath)
	root.mainloop()

	return	


def btnPrintDevInfo(event):
	btnState['text'] = 'WAIT'
	btnState['background'] = 'white'
	root.update()

	toolPath  = ".\commander"

	deviceName = mcuType

	cmdGetDevInfo = toolPath + '\commander.exe device info --device ' + deviceName
	devinfo=os.popen(cmdGetDevInfo)  	#popen与system可以执行指令,popen可以接受返回对象  
	devinfo=devinfo.read() 				#读取输出
	stateText.delete(0.0, Tkinter.END) 
	stateText.insert(Tkinter.END,devinfo) 
	
	if 'ERROR' in devinfo:
		btnState['text'] = 'ERROR'
		btnState['background'] = 'red'
		root.update()
	else :
		btnState['text'] = 'PASS'
		btnState['background'] = 'green'
		root.update()

		devId = devinfo[-22:-6]
		#printData='S/N:' + devId
		print2Printer(devId.upper())


	root.mainloop()

	return

def btnUnlockDev(event):
	btnState['text'] = 'WAIT'
	btnState['background'] = 'white'
	root.update()

	toolPath  = ".\commander"
	#deviceName = mcuChosen.get()
	deviceName = 'EFR32'      #mcuType

	outputInfo = ''
	cmdUnlockDev = toolPath + '\commander.exe device lock --debug disable --device ' + deviceName
	unlockInfo=os.popen(cmdUnlockDev)  	#popen与system可以执行指令,popen可以接受返回对象  
	unlockInfo=unlockInfo.read() 			#读取输出 
	outputInfo = unlockInfo
	stateText.delete(0.0, Tkinter.END)
	stateText.insert(Tkinter.END,unlockInfo)

	if 'ERROR' in outputInfo:
		btnState['text'] = 'ERROR'
		btnState['background'] = 'red'
		root.update()
	else :
		btnState['text'] = 'PASS'
		btnState['background'] = 'green'
		root.update()

	root.mainloop()

	return	


def btnEraseDev(event):
	btnState['text'] = 'WAIT'
	btnState['background'] = 'white'
	root.update()

	toolPath  = ".\commander"
	
	#deviceName = mcuChosen.get()
	deviceName = 'EFR32'      #mcuType

	outputInfo = ''

	cmdEraseDev = toolPath + '\commander.exe device masserase ' + '--device ' + deviceName
	eraseInfo=os.popen(cmdEraseDev)  	#popen与system可以执行指令,popen可以接受返回对象  
	eraseInfo=eraseInfo.read() 			#读取输出 
	outputInfo = outputInfo+eraseInfo
	stateText.delete(0.0, Tkinter.END)
	stateText.insert(Tkinter.END,eraseInfo)

	if 'ERROR' in outputInfo:
		btnState['text'] = 'ERROR'
		btnState['background'] = 'red'
		root.update()
	else :
		btnState['text'] = 'PASS'
		btnState['background'] = 'green'
		root.update()

	root.mainloop()

	return	

def btnflashDev(event):
	btnState['text'] = 'WAIT'
	btnState['background'] = 'white'
	root.update()

	toolPath  = ".\commander"
	flashFilePath = firewareEntry.get()
	print(flashFilePath)
	if flashFilePath[-3:] != 'hex' and flashFilePath[-3:] != 'bin' :
		btnState['text'] = 'ERROR'
		btnState['background'] = 'red'
		root.mainloop()
		return 

	#deviceName = mcuChosen.get()
	deviceName = mcuType

	outputInfo = ''
	cmdFlashDev = toolPath + '\commander.exe flash ' + flashFilePath + ' --address 0x0 --device ' + deviceName
	flashInfo=os.popen(cmdFlashDev)  	#popen与system可以执行指令,popen可以接受返回对象  
	flashInfo=flashInfo.read() 			#读取输出 
	outputInfo = outputInfo+flashInfo
	stateText.delete(0.0, Tkinter.END)
	stateText.insert(Tkinter.END,flashInfo)

	if 'ERROR' in outputInfo:
		btnState['text'] = 'ERROR'
		btnState['background'] = 'red'
		root.update()
	else :
		btnState['text'] = 'PASS'
		btnState['background'] = 'green'
		root.update()

	root.mainloop()

	return	

def btnLockDev(event):
	btnState['text'] = 'WAIT'
	btnState['background'] = 'white'
	root.update()

	toolPath  = ".\commander"
	
	#deviceName = mcuChosen.get()
	deviceName = 'EFR32'      #mcuType

	outputInfo = ''
	cmdLockDev = toolPath + '\commander.exe device lock --debug enable --device ' + deviceName
	lockInfo=os.popen(cmdLockDev)  	#popen与system可以执行指令,popen可以接受返回对象  
	lockInfo=lockInfo.read() 			#读取输出 
	outputInfo = outputInfo+lockInfo
	stateText.delete(0.0, Tkinter.END)
	stateText.insert(Tkinter.END,lockInfo)

	if 'ERROR' in outputInfo:
		btnState['text'] = 'ERROR'
		btnState['background'] = 'red'
		root.update()
	else :
		btnState['text'] = 'PASS'
		btnState['background'] = 'green'
		root.update()

	root.mainloop()

	return		

def closeWindow():
	os._exit(0)


root = Tkinter.Tk()                     # 创建窗口对象的背景色
root.title('条码工具')
root.geometry('800x600+500+300')
root.resizable(False, False)
root.attributes('-topmost',1)

configFile = open("./devices.json",'r')
configList = json.load(configFile)

devicesList = []
for device in configList:
	devicesList.append( device['deviceName'])

deviceModels = []
for model in configList[0]['deviceModels']:
	deviceModels.append( model['deviceModel'] )

temp =  configList[0]['deviceModels']
mcuType = temp[0]['mcuType']
configFile.close()

printLabel=Tkinter.Label(root,text='请选择连接的打印机型号：', font=(10))
printLabel.place(x = 50, y = 30)   

printerList = printerList()
printerType = Tkinter.StringVar()
printerChosen = ttk.Combobox(root, width=12, textvariable=printerType)
printerChosen['values'] = printerList
printerChosen.place(x = 300, y = 30)   
printerChosen.current(0)    					 	# 设置下拉列表默认显示的值


deviceLabel=Tkinter.Label(root,text='请选择连接的设备种类：',font=(10))
deviceLabel.place(x = 50, y = 80)

deviceName = Tkinter.StringVar()
deviceChosen = ttk.Combobox(root, width=12, textvariable=deviceName)
deviceChosen['values'] = devicesList
deviceChosen.place(x = 300, y = 80)   
deviceChosen.current(0)
deviceChosen.bind("<<ComboboxSelected>>",deviceChosenEvent) 

modelLabel=Tkinter.Label(root,text='请选择连接的设备型号：',font=(10))
modelLabel.place(x = 50, y = 130)

modelType = Tkinter.StringVar()
modelChosen = ttk.Combobox(root, width=12, textvariable=modelType)
modelChosen['values'] = deviceModels
modelChosen.place(x = 300, y = 130)
modelChosen.current(0)
modelChosen.bind("<<ComboboxSelected>>",modelChosenEvent) 

firewareLabel=Tkinter.Label(root,text='请选择需要烧录的固件：',font=(10))
firewareLabel.place(x = 50, y = 180)
fileInfo = Tkinter.StringVar() 
#firewareEntry = Entry(root,text = '')
firewareEntry = Tkinter.Entry(root,textvariable=fileInfo)
firewareEntry['state'] = 'readonly'
firewareEntry.place(x = 300, y = 180)
firewareBtn = Tkinter.Button(root,text = '选择固件',height = 1)
firewareBtn.place(x = 450, y = 175)
firewareBtn.bind("<Button-1>", btnChooseFirmWare)

stateText = Tkinter.Text(root, height=15, width=100)
stateText.place(x = 50, y = 230)

btnStateText = Tkinter.StringVar() 
btnState = Tkinter.Button(root,text='WAIT',width = 40 ,height = 6)
btnStateText.set('WAIT')
btnState.place(x = 50,y = 450)

btnPrint = Tkinter.Button(root,text = '打印',width = 15,height = 3)
btnPrint.place(x = 380,y = 475)
btnPrint.bind("<Button-1>", btnPrintDevInfo)

btnUnlock= Tkinter.Button(root,text = '解密',width = 8,height = 1)
btnUnlock.place(x = 530,y = 475)
btnUnlock.bind("<Button-1>", btnUnlockDev)

btnErase = Tkinter.Button(root,text = '擦除',width = 8,height = 1)
btnErase.place(x = 620,y = 475)
btnErase.bind("<Button-1>", btnEraseDev)

btnFlash = Tkinter.Button(root,text = '编程',width = 8,height = 1)
btnFlash.place(x = 530,y = 515)
btnFlash.bind("<Button-1>", btnflashDev)

btnLock = Tkinter.Button(root,text = '加密',width = 8,height = 1)
btnLock.place(x = 620,y = 515)
btnLock.bind("<Button-1>", btnLockDev)

'''
btnPrint = Tkinter.Button(root,text = '编程',width = 15,height = 3)
btnPrint.place(x = 530,y = 475)
btnPrint.bind("<Button-1>", btnflashDev)
'''

try:
	root.protocol('WM_DELETE_WINDOW', closeWindow)
	root.mainloop()								# 进入消息循环
except Exception as e:
	os._exit(0)