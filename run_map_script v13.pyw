# -*- coding: cp1252 -*-
import arcpy, codecs, zipfile
from arcpy import env
from Tkinter import *
from ttk import *
from tkFileDialog import *
import tkMessageBox
import tkFont
import datetime
import datepick
import os
import mapsupportv2


# aed 26/9/18 - Added binding to buttons to activate them on Return/Enter
# aed 26/9/18 - Added 2 to all grid rows >=84
# AED 10/02/21 - Removed Durham Bird Club checks & added new consultants to drop down
# AED 21/04/21 - Added Hartlepool LGS
# AED 20/07/21 - Check whether search in area of Northumberland Bat Group
# AED 27/06/22 - Check whether search in area of Durham Bat Group (script version 13)

class Application(Frame):

# aed 25/09/18
# Functions to facilitate tabbing out of text boxes
    def _focusNext(self, widget):
        '''Return the next widget in tab order'''
        widget = self.tk.call('tk_focusNext', widget._w)
        if not widget: return None
        return self.nametowidget(widget.string)

    def _focusPrev(self, widget):
        '''Return the previous widget in tab order'''
        widget = self.tk.call('tk_focusPrev', widget._w)
        if not widget: return None
        return self.nametowidget(widget.string)

    def OnTextTab(self, event):
        '''Move focus to next widget'''
        widget = event.widget
        next = self._focusNext(widget)
        next.focus()
        return "break"

    def OnTextShiftTab(self, event):
        '''Move focus to previous widget'''
        widget = event.widget
        next = self._focusPrev(widget)
        next.focus()
        return "break"

    def createWidgets(self):

        # client dictionary - keys are names that will appear in the dropdown, values are [folder name, debtor list name]


# AED 10/02/21 - added new clients
        self.clientDict = {
                    'A1 Ecology' : ['A1 Ecology', 'A1 Ecology'],
                    'Access Ecology' : ['Access Ecology', 'Access Ecology'],
                    'Acer Ecology' : ['Acer Ecology', 'Acer Ecology'],
                    'ADAS' : ['ADAS UK', 'ADAS'],
                    'Aecol' : ['Aecol', 'Aecol'],
                    'AECOM' : ['AECOM', 'AECOM'],
                    'AES' : ['AES', 'Applied Ecological Services Ltd'],
                    'AJT Environmental Consultants' : ['AJT Environmental Consultants', 'AJT Environmental Consultants'],
                    'All About Trees' : ['All_About_Trees', 'All About Trees'],
                    'Amec' : ['AMEC', 'Amec Foster Wheeler Environment & Infrastructure'],
                    'Amey Consulting' : ['Amey Consulting', 'Amey Consulting'],
                    'APEM Limited' : ['APEM Limited', 'APEM Limited'],
                    'Appletons UK' : ['Appletons UK', 'Appletons UK'],
                    'Arbtech' : ['Arbtech Consulting', 'Arbtech'],
                    'Arcadis UK' : ['Arcadis UK', 'Arcadis UK'],
                    'Archer Ecology' : ['Archer_Ecology', 'Archer Ecology'],
                    'Arcus Ecology' : ['Arcus_Ecology', 'Arcus Ecology'],
                    'Argus Ecology' : ['Argus_Ecology', 'Argus Ecology'],
                    'Artemis Ecological' : ['Artemis Ecological Consulting', 'Artemis Ecological Consulting Ltd'],
                    'Arup' : ['Arup', 'Arup'],
                    'Atkins' : ['Atkins', 'Atkins'],
                    'Avian Ecology' : ['Avian', 'Avian Ecology'],
                    'Baker Consultants' : ['Baker Consultants', 'Baker Consultants'],
                    'Barrett Environmental' : ['Barrett_Environmental', 'Barrett Environmental'],
                    'BDP' : ['BDP', 'BDP'],
                    'BeatyMadine' : ['BeatyMadine', 'BeatyMadine'],
                    'Biodiverse Consulting' : ['Biodiverse Consulting', 'Biodiverse Consulting'],
                    'BL Ecology' : ['BL Ecology', 'BL Ecology'],
                    'Bridgeway Consulting' : ['Bridgeway Consulting', 'Bridgeway Consulting'],
                    'Brindle & Green' : ['Brindle & Green', 'Brindle & Green Ltd'],
                    'Brooks Ecological' : ['Brooks_Ecological', 'Brooks Ecological'],
                    'BSG Ecology' : ['Baker_Shepherd_Gillespie', 'BSG Ecology'],
                    'BT Ecological Services' : ['BT Ecological Services', 'BT Ecological Services'],
                    'Budhaig Environmental' : ['Budhaig', 'Budhaig Environmental'],
                    'Capita' : ['Capita', 'Capita'],
                    'Cascade Consulting' : ['Cascade Consulting', 'Cascade Consulting'],
                    'CH2M' : ['CH2M HILL', 'CH2M'],
                    'CSCA Ecology' : ['CSCA Ecology', 'CSCA Ecology'],
                    'David Dodds' : ['David Dodds Associates Ltd', 'David Dodds Associates'],
                    'Delta-Simons' : ['Delta-Simons', 'Delta-Simons'],
                    'Dendra' : ['Dendra_Consulting', 'Dendra Consulting'],
                    'Derek Hilton-Brown' : ['Derek Hilton-Brown', 'Derek Hilton-Brown'],
                    'Veronica Howard' : ['Veronica_Howard', 'Dr Veronica Howard'],
                    'DWS' : ['Durham Wildlife Services', 'DWS Ecology'],
                    'E3 Ecology' : ['E3_Ecology', 'E3 Ecology'],
                    'Eco-Fish Consultants' : ['Eco-Fish Consultants', 'Eco-Fish Consultants'],
                    'Ecological Services' : ['Ecological Services', 'Ecological Services'],
                    'Ecology Services UK Ltd' : ['Ecology Services UK Ltd', 'Ecology Services UK Ltd'],
                    'Ecology Solutions' : ['Ecology Solutions', 'Ecology Solutions'],
                    'EcoNorth' : ['Eco_North', 'EcoNorth'],
                    'Ecosurv Ltd' : ['Eco_Surv', 'Ecosurv Ltd'],
                    'Ecus Ltd' : ['ECUS', 'Ecus Ltd'],
                    'Elite Ecology' : ['Elite Ecology', 'Elite Ecology'],
                    'Elliott Environmental' : ['Elliott_Environmental_Surveyors', 'Elliott Environmental'],
                    'Elmer Ecology' : ['Elmer Environmental Consulting', 'Elmer Ecology'],
                    'Envirotech' : ['Envirotech', 'Envirotech'],
                    'Enzygo' : ['Enzygo', 'Enzygo'],
                    'ERAP' : ['ERAP', 'ERAP Ltd'],
                    'Falco Ecology' : ['Falco Ecology', 'Falco Ecology'],
                    'FPCR' : ['FPCR', 'FPCR Environment & Design'],
                    'George Dodds' : ['George Dodds & Co', 'George Dodds & Co'],
                    'GN Megson' : ['GN Megson Ecology', 'GN Megson Ecology Ltd'],
                    'Grays Ecology' : ['Grays Ecology', 'Grays Ecology'],
                    'GSL' : ['GSL', 'GSL'],
                    'Hadrian Ecology' : ['Hadrian Ecology', 'Hadrian Ecology'],
                    'Harwood Biology' : ['Harwood Biology', 'Harwood Biology'],
                    'Haycock & Jay Associates' : ['Haycock & Jay Associates', 'Haycock & Jay Associates'],
                    'INCA' : ['INCA', 'INCA'],
                    'Innovation Environmental' : ['Innovation Group Environmental Services', 'Innovation Group Environmental Services'],
                    'Jacobs' : ['Jacobs UK', 'Jacobs ltd'],
                    'JBA Consulting' : ['JBA Consulting', 'JBA Consulting'],
                    'JCA Limited' : ['JCA Limited', 'JCA Limited'],
                    'John Drewett Ecology' : ['John Drewett Ecology', 'John Drewett Ecology'],
                    'John Durkin' : ['JL Durkin', 'John Durkin'],
                    'JP Environmental Solutions' : ['JP Environmental Solutions', 'JP Environmental Solutions'],
                    'Julia Quinonez' : ['Julia Quinonez', 'Julia Quinonez'],
                    'Longrigg Smith' : ['Longrigg Smith', 'Longrigg Smith'],
                    'MAB Ecology' : ['MAB Ecology', 'MAB Ecology'],
                    'Middlemarch Environmental' : ['Middlemarch', 'Middlemarch Environmental Ltd'],
                    'Milner Ecology' : ['Milner_Ecology', 'Milner Ecology'],
                    'Mott MacDonald' : ['Mott_MacDonald', 'Mott MacDonald'],
                    'MW Surveys' : ['MW Surveys', 'MW Surveys'],
                    'Natural Power' : ['Natural Power', 'Natural Power'],
                    'Naturally Wild' : ['Naturally_Wild', 'Naturally Wild'],
                    'Newcastle City Council' : ['Newcastle City Council', 'Newcastle City Council'],
                    'Northumberland County Council' : ['Northumberland County Council', 'Northumberland County Council'],
                    'Northumbrian Water' : ['Northumbrian_Water', 'Northumbrian Water'],
                    'OS Ecology' : ['OS Ecology', 'OS Ecology'],
                    'Parnassus Ecology' : ['Parnassus Ecology', 'Parnassus Ecology'],
                    'PBA Applied Ecology' : ['PBA Applied Ecology', 'PBA Applied Ecology'],
                    'Peak Ecology' : ['Peak Ecology', 'Peak Ecology'],
                    'Penn Associates' : ['Penn_Associates_Ecology', 'Penn Associates'],
                    'Peter Brett Associates' : ['Peter Brett Associates', 'Peter Brett Associates'],
                    'Peter Tattersfield' : ['Peter Tattersfield', 'Peter Tattersfield'],
                    'Quants Environmental' : ['Quants Environmental', 'Quants Environmental'],
                    'Rachel Flannery' : ['Rachel Flannery', 'Rachel Flannery'],
                    'Rachel Hacking Ecology Ltd' : ['Rachel Hacking Ecology Ltd', 'Rachel Hacking Ecology Ltd'],
                    'Rachel Hepburn' : ['Rachel Hepburn', 'Rachel Hepburn'],
                    'REC Ltd' : ['REC', 'REC Ltd'],
                    'RH Ecological Services' : ['Rachel Hepburn', 'RH Ecological Services'],
                    'Rigby Jerram' : ['Rigby Jerram', 'Rigby Jerram'],
                    'RPS' : ['RPS', 'RPS'],
                    'RSK' : ['RSK', 'RSK'],
                    'Ryal Soil and Ecology' : ['Ryal_Soil_&_Ecology', 'Ryal Soil and Ecology'],
                    'SK Environmental' : ['SK Environmental Solutions', 'SK Environmental Solutions'],
                    'SLR Consulting' : ['SLR', 'SLR Consulting'],
                    'Smeeden Foreman' : ['Smeeden_Foreman', 'Smeeden Foreman'],
                    'Sol Environmental' : ['Sol Environment', 'Sol Environmental'],
                    'Sterna Ecology' : ['Sterna Ecology', 'Sterna Ecology'],
                    'Strutt and Parker' : ['Strutt & Parker', 'Strutt and Parker'],
                    'Stuart Johnson' : ['Stuart Johnson', 'Stuart Johnson'],
                    'Surface Property' : ['Surface Property', 'Surface Property'],
                    'TEP' : ['TEP', 'TEP'],
                    'Tees River Trust' : ['Tees River Trust', 'Tees River Trust'],
                    'The Ecological Consultancy' : ['The Ecological Consultancy', 'The Ecological Consultancy'],
                    'The Ecology Consultancy' : ['The Ecology Consultancy', 'The Ecology Consultancy'],
                    'Tina Wiffen' : ['Tina Wiffen', 'Tina Wiffen'],
                    'Total Ecology' : ['Total_Ecology', 'Total Ecology'],
                    'TSP' : ['TSP Projects', 'TSP Projects'],
                    'Tweed Ecology' : ['Tweed Ecology', 'Tweed Ecology'],
                    'Tweed Forum' : ['Tweed Forum', 'Tweed Forum'],
                    'Tyler Grange' : ['Tyler_Grange', 'Tyler Grange'],
                    'Tyne Ecology' : ['Tyne Ecology', 'Tyne Ecology'],
                    'UES' : ['UES', 'UES'],
                    'Urban Green Space' : ['Urban Green Space', 'Urban Green Space'],
                    'Veronica Howard' : ['Veronica Howard', 'Veronica Howard'],
                    'Vickers & Barrass' : ['Vickers and Barrass', 'Vickers & Barrass'],
                    'Wardell Armstrong' : ['Wardell_Armstrong', 'Wardell Armstrong'],
                    'Waterman' : ['Waterman', 'Waterman Infrastructure & Environment Ltd'],
                    'Whitcher Wildlife' : ['Whitcher Wildlife', 'Whitcher Wildlife Ltd'],
                    'Whittingham Ecology' : ['Whittingham Ecology', 'Whittingham Ecology'],
                    'Wold Ecology' : ['Wold_Ecology', 'Wold Ecology'],
                    'Wood PLC' : ['Wood PLC', 'Wood PLC'],
                    'WRM' : ['WRM', 'WRM'],
                    'WSP' : ['WSP', 'WSP UK CPL'],
                    'WYG' : ['WYG', 'WYG'] }


# aed 26/9/18 By default tab order depends on order controls are created.

        self.sub = Frame(self, width=600, height=400)
        self.sub['borderwidth'] = 2
        self.sub['relief'] = 'sunken'

        self.lbl9 = Label(self.sub, text='Site name')
        self.siteName = StringVar()
        self.entSiteName = Entry(self.sub, textvariable=self.siteName)

        self.lbl10 = Label(self.sub, text='Consultancy')
        self.consultancy = StringVar()
        self.selConsultancy = Combobox(self.sub, textvariable=self.consultancy)
        self.selConsultancy['values'] = sorted(self.clientDict.keys())

# aed 02/10/18 Flag whether consultant has an SLA
        self.lblSLA = Label(self.sub, text="SLA")
        self.SLA = IntVar()
        self.SLA.set(0)
        self.chkSLA = Checkbutton(self.sub, variable=self.SLA)


        self.lbl11 = Label(self.sub, text='Consultant name')
        self.consultantName = StringVar()
        self.entConsultantName = Entry(self.sub, textvariable=self.consultantName)

        self.lbl12 = Label(self.sub, text='Request date')
        self.reqDate = StringVar()
        self.entRequestDate = Entry(self.sub, textvariable=self.reqDate)
        self.lbl12a = Label(self.sub, text='(dd-mm-yyyy)')

# aed 25/9/18
# A button to set the date to today
        self.btnToday = Button(self.sub, text='Today', width=10, command=self.setToday)
        self.btnToday.bind("<Return>", self.setToday)

# aed 26/9/18
        self.btnPickDate = Button(self.sub, text='Pick date...', width=10, command=self.pickDate)
        self.btnPickDate.bind("<Return>", self.pickDate)

        self.lbl13 = Label(self.sub, text='Output folder')
# aed 25/9/18
# Enable tabbing out of the field
# AED - 27/7/20
        self.txtOutputFolder = Text(self.sub, width=100, height=2, wrap='word')
        self.txtOutputFolder.bind("<Tab>", self.OnTextTab)
        self.txtOutputFolder.bind("<Shift-Tab>", self.OnTextShiftTab)
# aed 26/9/18
        self.btnSelectOutput = Button(self.sub, text='Choose...', width=10, command=self.selectOF)
        self.btnSelectOutput.bind("<Return>", self.selectOF)

        self.lbl14 = Label(self.sub, text='Name prefix')
        self.namePrefix = StringVar()
        self.entNamePrefix = Entry(self.sub, textvariable=self.namePrefix)

        self.sep3 = Separator(self.sub, orient=HORIZONTAL)

        self.lbl1 = Label(self.sub, text='Search by')

        self.srchType = StringVar()
        self.srchType.set('gr')
        self.rbGridRef = Radiobutton(self.sub, text='Grid reference', variable=self.srchType, value='gr')
        self.rbShapefile = Radiobutton(self.sub, text='Shapefile', variable=self.srchType, value='sf')

        self.sep = Separator(self.sub, orient=HORIZONTAL)

        self.lbl3a = Label(self.sub, text='Search area')

        self.lbl4 = Label(self.sub, text='Grid reference')
        self.gridRef = StringVar()
        self.entGridRef = Entry(self.sub, textvariable=self.gridRef)

# aed 25/9/18
# Enable tabbing out of the field
        self.lbl5 = Label(self.sub, text='Shapefile')
# AED - 27/7/20
        self.txtShapefile = Text(self.sub, width=100, height=2, wrap='word')
        self.txtShapefile.bind("<Tab>", self.OnTextTab)
        self.txtShapefile.bind("<Shift-Tab>", self.OnTextShiftTab)
# aed 25/9/18
        self.btnPickFile = Button(self.sub, text='Choose...', width=10, command=self.selectSF)
        self.btnPickFile.bind("<Return>", self.selectSF)

        self.lbl6 = Label(self.sub, text='Shapefile is:')

        self.sfType = StringVar()
        self.rbSiteBound = Radiobutton(self.sub, text='Site boundary', variable=self.sfType, value='sb')
        self.rbSearchArea = Radiobutton(self.sub, text='Full search area (i.e. including buffer)', variable=self.sfType, value='sa')

        self.lbl40 = Label(self.sub, text='Shapefile description')
        self.sfTitle = StringVar()
        self.entSfTitle = Entry(self.sub, textvariable=self.sfTitle)


        self.lbl8a = Label(self.sub, text='Search radius (km)')
        self.searchRadius = StringVar()
        self.searchRadius.set('2')
        self.spSearchRad = Spinbox(self.sub, from_=1, to=10, textvariable=self.searchRadius)

        self.lbl8b = Label(self.sub, text='Shapefile provided by client')
        self.clientSF = IntVar()
        self.clientSF.set(0)
        self.chkClientSF = Checkbutton(self.sub, variable=self.clientSF)


        self.sep2 = Separator(self.sub, orient=HORIZONTAL)

        self.lbl8z = Label(self.sub, text='Request details')


        self.lbl18 = Label(self.sub, text='Reports')

        self.repType = StringVar()
        self.repType.set('pn')
        self.rbPandN = Radiobutton(self.sub, text='Protected & Notable', variable=self.repType, value='pn')
        self.rbAllSpec = Radiobutton(self.sub, text='All Species', variable=self.repType, value='as')
# aed 26/9/18 - Single species
        self.rbSingleSpec = Radiobutton(self.sub, text='Single Species', variable=self.repType, value='ss')

        self.lblSS = Label(self.sub, text='Species name')
        self.singleSpec = StringVar()
        self.entSingleSpec = Entry(self.sub, textvariable=self.singleSpec)

        self.lbl21 = Label(self.sub, text='LWS citations')
        self.LWS = IntVar()
        self.LWS.set(0)
        self.chkLWS = Checkbutton(self.sub, variable=self.LWS)

        self.sep4 = Separator(self.sub, orient=HORIZONTAL)

        self.lbl25 = Label(self.sub, text='Maps')
# aed 26/9/18
        self.btnIgnoreLayers = Button(self.sub, text='Ignore layers...', width=15, command=self.selectIgnoreLayers)
        self.btnIgnoreLayers.bind("<Return>", self.selectIgnoreLayers)

        self.lbl26 = Label(self.sub, text='Statutory sites')
        self.statMap = IntVar()
        self.statMap.set(0)
        self.chkStat = Checkbutton(self.sub, variable=self.statMap)

        self.lbl27 = Label(self.sub, text='Non-statutory sites')
        self.nonStatMap = IntVar()
        self.nonStatMap.set(0)
        self.chkNonStat = Checkbutton(self.sub, variable=self.nonStatMap)

        self.lbl28 = Label(self.sub, text='Priority habitats')
        self.phMap = IntVar()
        self.phMap.set(0)
        self.chkPH = Checkbutton(self.sub, variable=self.phMap)

        self.sep5 = Separator(self.sub, orient=HORIZONTAL)
# AED 27/7/20
        self.txtLog = Text(self.sub, width=100, height=8, wrap='word')
        self.txtLog['state']='disabled'
        self.scrBar = Scrollbar(self.sub, command=self.txtLog.yview)
        self.txtLog['yscrollcommand']=self.scrBar.set

# aed 26/9/18
        self.btnClear = Button(self.sub, text='Clear', width=10, command=self.clearFields)
        self.btnClear.bind("<Return>", self.clearFields)
        self.btnOK = Button(self.sub, text='OK', width=10, command=self.submit)
        self.btnOK.bind("<Return>", self.submit)

        self.btnExit = Button(self.sub, text='Quit', width=10, command=self.checkExit)
        self.btnExit.bind("<Return>", self.checkExit)
# AED - 27/7/20
#        self.btnQuit = Button(self, text = 'Quit', width=10, command=self.checkExit)
#        self.btnQuit.bind("<Return>", self.checkExit)

#### GRID
# AED - 27/7/20
        #self.btnQuit.grid(row=10, column=0, sticky='SW')

        self.lbl1.grid(row=50, column=0, sticky=W)

# Search by frame
        self.rbGridRef.grid(row=52, column=0, sticky=W, padx=5, pady=3)
        self.rbShapefile.grid(row=52, column=1, sticky=W, padx=5, pady=3)
        self.rbGridRef.grid(row=50, column=1, sticky=W, padx=5, pady=3)
        self.rbShapefile.grid(row=50, column=2, sticky=W, padx=5, pady=3)


        self.sep.grid(row=56, columnspan=5, sticky='EW', padx=5, pady=3)

# Search area frame
        self.lbl3a.grid(row=58, column=0, sticky=W)
#AED - 27/7/20
        self.lbl4.grid(row=58, column=0, sticky=E)
        self.entGridRef.grid(row=58, column=1, sticky=W, padx=5, pady=3)

        self.lbl5.grid(row=58, column=0, sticky='NE', pady=3)
        self.txtShapefile.grid(row=58, column=1, columnspan=3, sticky='EW', padx=5, pady=3)
        self.btnPickFile.grid(row=58, column=4, sticky=E, padx=5)

        self.lbl6.grid(row=62, column=0, sticky=E)

        self.rbSiteBound.grid(row=62, column=1, sticky=W, padx=5, pady=3)
        self.rbSearchArea.grid(row=68, column=1, sticky=W, padx=5, pady=3)

        self.lbl40.grid(row=62, column=2, sticky=E, padx=5, pady=3)
        self.entSfTitle.grid(row=62, column=3, sticky=W, padx=5, pady=3)


        self.lbl8a.grid(row=70, column=0, sticky=E)
        self.spSearchRad.grid(row=70, column=1, sticky=W, padx=5, pady=3)
        self.lbl8b.grid(row=72, column=0, sticky=E)
        self.chkClientSF.grid(row=72, column=1, sticky=W, padx=5, pady=3)

        self.sep2.grid(row=74, columnspan=5, sticky='EW', padx=5, pady=3)

# Request details frame

        self.lbl8z.grid(row=17, column=0, sticky=W)
        self.lbl9.grid(row=18, column=0, sticky=E)
        self.entSiteName.grid(row=18, column=1, columnspan=3, padx=5, pady=3, sticky='EW')
        self.lbl10.grid(row=20, column=0, sticky=E)
        self.selConsultancy.grid(row=20, column=1, columnspan=3, padx=5, pady=3, sticky='EW')
# aed - 2/10/18
        self.lblSLA.grid(row=20, column=4, sticky=W)
        self.chkSLA.grid(row=20, column=4, sticky=E)

        self.lbl11.grid(row=22, column=0, sticky=E)
        self.entConsultantName.grid(row=22, column=1, columnspan=3, padx=5, pady=3, sticky='EW')
        self.lbl12.grid(row=24, column=0, sticky=E)
        self.entRequestDate.grid(row=24, column=1, padx=5, pady=3, sticky='EW')
        self.lbl12a.grid(row=24, column=2,  sticky=W)
# aed 25/9/18
# Display the today button
        self.btnToday.grid(row=24, column=3, padx=5, pady=3)
        self.btnPickDate.grid(row=24, column=4, padx=5, pady=3)
        self.lbl13.grid(row=26, column=0, sticky='NE', pady=3)
        self.txtOutputFolder.grid(row=26, column=1, columnspan=3, padx=5, pady=3, sticky='EW')
        self.btnSelectOutput.grid(row=26, column=4, sticky=E, padx=5)
        self.lbl14.grid(row=28, column=0, sticky=E)
        self.entNamePrefix.grid(row=28, column=1, columnspan=3, padx=5, pady=3, sticky='EW')

        self.sep3.grid(row=32, columnspan=5, sticky='EW', padx=5, pady=3)

# Reports frame

        self.lbl18.grid(row=80, column=0, sticky=W)
#AED - 27/7/20
        self.rbPandN.grid(row=80, column=1, sticky=W, padx=5, pady=3)
        self.rbAllSpec.grid(row=80, column=2, sticky=W, padx=5, pady=3)
# aed 26/9/18
# Single species
#AED - 27/7/20
        self.rbSingleSpec.grid(row=80, column=2, sticky=E, padx=5, pady=3)

        self.lblSS.grid(row=84, column=2, sticky=E, padx=5, pady=3)
        self.entSingleSpec.grid(row=84, column=3, sticky=W, padx=5, pady=3)

#AED - 27/7/20
        self.lbl21.grid(row=84, column=0, sticky=E)
        self.chkLWS.grid(row=84, column=1, sticky=W, padx=5, pady=3)

        self.sep4.grid(row=88, columnspan=5, sticky='EW', padx=5, pady=3)

# Maps frame
        self.lbl25.grid(row=92, column=0, sticky=W)
        self.btnIgnoreLayers.grid(row=92, column=2, sticky=E, padx=5, pady=3)
#AED - 27/7/20
        self.lbl26.grid(row=92, column=0, sticky=E)
        self.chkStat.grid(row=92, column=1, sticky=W, padx=5, pady=3)

        self.lbl27.grid(row=94, column=0, sticky=E)
        self.chkNonStat.grid(row=94, column=1, sticky=W, padx=5, pady=3)

        self.lbl28.grid(row=96, column=0, sticky=E)
        self.chkPH.grid(row=96, column=1, sticky=W, padx=5, pady=3)


        self.sep5.grid(row=102, columnspan=5, sticky='EW', padx=5, pady=3)

# Output log frame
        self.txtLog.grid(row=104, column=0, columnspan=3, rowspan=4, sticky='NSEW', padx=5, pady=3)
        self.scrBar.grid(row=104, column=3, rowspan=4, sticky='NSW')
#AED - 27/7/20
        self.btnClear.grid(row=104, column=4, sticky=W, padx=5, pady=3)
        self.btnOK.grid(row=105, column=4, sticky=W, padx=5, pady=3)

#AED - 27/7/20
        self.btnExit.grid(row=106, column=4, sticky=W, padx=5, pady=3)

        self.sub.grid(row=0, column=1, rowspan=10, sticky='nse')

        self.srchType.trace("w", self.srchTypeUpdated)
        self.sfType.trace("w", self.sfTypeUpdated)
# aed 26/9/18
        self.repType.trace("w", self.repTypeUpdated)

        self.siteName.trace("w", self.updateNamePrefix)
        self.consultancy.trace("w", self.updateNamePrefix)
        self.reqDate.trace("w", self.updateNamePrefix)

        self.entNamePrefix.bind("<FocusOut>", self.checkNamePrefix)
        self.entRequestDate.bind("<FocusOut>", self.checkDateSeps)
        self.entGridRef.bind("<FocusOut>", self.checkGridRef)
        self.entSfTitle.bind("<FocusOut>", self.checkSfTitle)

        self.lbl5.grid_remove()
        self.txtShapefile.grid_remove()
        self.btnPickFile.grid_remove()
        self.lbl6.grid_remove()
        self.rbSiteBound.grid_remove()
        self.rbSearchArea.grid_remove()
        self.lbl40.grid_remove()
        self.entSfTitle.grid_remove()
        self.lbl8b.grid_remove()
        self.chkClientSF.grid_remove()
#aed 26/9/18
        self.lblSS.grid_remove()
        self.entSingleSpec.grid_remove()

        self.columnconfigure(0, weight=3)
        self.columnconfigure(1, weight=3)
# AED - 27/7/20
        self.rowconfigure(0, weight=3, minsize=500)
        self.sub.columnconfigure(0, weight=5, minsize=160)

# aed 26/9/18
# set focus to the first field
        self.entSiteName.focus()

    def srchTypeUpdated(self, *args):
        if self.srchType.get() == 'sf':
            self.lbl4.grid_remove()
            self.entGridRef.grid_remove()
            self.lbl5.grid()
            self.txtShapefile.grid()
            self.btnPickFile.grid()
            self.lbl6.grid()
            self.rbSiteBound.grid()
            self.rbSearchArea.grid()
            self.lbl8b.grid()
            self.chkClientSF.grid()
            self.sfTypeUpdated()

        if self.srchType.get() == 'gr':
            self.lbl4.grid()
            self.entGridRef.grid()
            self.lbl8a.grid()
            self.spSearchRad.grid()
            self.spSearchRad.configure(state='normal')
            self.lbl5.grid_remove()
            self.txtShapefile.grid_remove()
            self.btnPickFile.grid_remove()
            self.lbl6.grid_remove()
            self.sfType.set('')
            self.rbSiteBound.grid_remove()
            self.rbSearchArea.grid_remove()
            self.lbl8a.grid()
            self.spSearchRad.grid()
            self.spSearchRad.configure(state='normal')
            self.lbl8b.grid_remove()
            self.chkClientSF.grid_remove()
            self.clientSF.set(0)
            self.lbl40.grid_remove()
            self.entSfTitle.grid_remove()
            self.entSfTitle.state(['disabled'])

# aed 26/9/18

    def repTypeUpdated(self, *args):
        if self.repType.get() == 'ss':
            self.lblSS.grid()
            self.entSingleSpec.grid()
        else:
            self.lblSS.grid_remove()
            self.entSingleSpec.grid_remove()


    def selectSF(self, *args):
        currFolder = self.txtOutputFolder.get('1.0','end').strip()
        initdir = ""
        if currFolder == "":
# aed 25/9/18
# Check which drive the server is mapped to
            for drive in "GHIJKLMNOPQRSTUVWXYZ":
                if os.path.exists(drive + ":\ERIC North East\ERIC Data Services\Commercial Data Requests\\"):
                    initdir = drive + ":\ERIC North East\ERIC Data Services\Commercial Data Requests\\"
                else:
                    initdir = os.curdir
            cons = self.consultancy.get()
            if cons <> "" and cons in self.clientDict.keys():
                consFolder = self.clientDict.get(cons)[0]
                initdir += consFolder + "\\"
        else:
            initdir = currFolder
        self.file_opt = options = {}
        options['filetypes'] = [('Shapefiles', '.shp'),('All files', '.*')]
        options['initialdir'] = initdir
        fileSF = askopenfilename(**self.file_opt)
        if fileSF <> "":
            fileSFWithPath = os.path.normpath(fileSF)
            self.txtShapefile.delete('1.0', 'end')
            self.txtShapefile.insert('1.0', fileSFWithPath)

            # get shapefile type
            desc = arcpy.Describe(fileSFWithPath)
            if desc.shapeType == "Polygon":
                self.defaultSfTitle = "Site Boundary"
            elif desc.shapeType == "Polyline":
                self.defaultSfTitle = "Route"
            elif desc.shapeType == "Point":
                self.defaultSfTitle = "Point"
            self.sfTitle.set(self.defaultSfTitle)

            if currFolder == "":
                fileFolder = os.path.dirname(os.path.abspath(fileSF))
                self.txtOutputFolder.delete('1.0', 'end')
                self.txtOutputFolder.insert('1.0', fileFolder)


    def sfTypeUpdated(self, *args):
        if self.sfType.get() == 'sb':
            self.lbl8a.grid()
            self.spSearchRad.grid()
            self.spSearchRad.configure(state='normal')
            self.lbl40.grid()
            self.entSfTitle.grid()
            self.entSfTitle.state(['!disabled'])
        else:
            self.lbl8a.grid_remove()
            self.spSearchRad.grid_remove()
            self.spSearchRad.configure(state='disabled')
            self.lbl40.grid_remove()
            self.entSfTitle.grid_remove()
            self.entSfTitle.state(['disabled'])


    def pickDate(self, *args):
        self.wdw = Toplevel()
        self.wdw.geometry('+%d+%d' % (root.winfo_x()+20, root.winfo_y()+30))
        self.dateScreen = datepick.DatePicker(self.wdw, self)
        self.dateScreen.grid()
        self.wdw.transient(self)
        self.wdw.grab_set()
        self.wait_window(self.wdw)

# aed 25/9/18
# Set the date to today
    def setToday(self, *args):
        self.reqDate.set(datetime.date.today().strftime("%d-%m-%Y"))

    def newDate(self, data):
        self.wdw.destroy()
        if data <> None:
            self.reqDate.set(data.strftime("%d-%m-%Y"))


    def updateNamePrefix(self, *args):
        if self.siteName.get() <> "" and self.consultancy.get() <> "" and self.reqDate.get() <> "":
            self.namePrefix.set(self.consultancy.get() + " " + self.reqDate.get() + " " + self.siteName.get())

    def checkNamePrefix(self, *args):
        if self.namePrefix.get() == "":
            self.updateNamePrefix()

    def checkDateSeps(self, *args):
        self.reqDate.set(self.reqDate.get().replace('/','-').replace('.','-'))


    def checkGridRef(self, *args):
        self.gridRef.set(self.gridRef.get().strip())

    def checkSfTitle(self, *args):
        if self.sfTitle.get() == "":
            self.sfTitle.set(self.defaultSfTitle)


    def selectOF(self, *args):
        currFolder = self.txtOutputFolder.get('1.0','end').strip()
        initdir = ""
# aed 25/9/18
        if currFolder == "":
# aed 03/08/20 - for working at home use init dir from config file
            exec(self.dir_config)
            initdir = "I:\ERIC\Data requests\Script"
            initdir = dir_root
        else:
            initdir = currFolder
        fileLoc = askdirectory(initialdir=initdir)
        if fileLoc <> "":
            self.txtOutputFolder.delete('1.0', 'end')
            self.txtOutputFolder.insert('1.0', os.path.normpath(fileLoc))


    def selectIgnoreLayers(self, *args):
        self.wdw = Toplevel()
        self.wdw.geometry('+%d+%d' % (root.winfo_x()+20, root.winfo_y()+30))
        self.layersScreen = LayerPicker(self.wdw, self)
        self.layersScreen.grid()
        self.wdw.transient(self)
        self.wdw.grab_set()
        self.wait_window(self.wdw)


    def ignoreDone(self, data):
        self.wdw.destroy()
        if data <> None:
            self.ignoreLayerNames = data


    def clearFields(self, *args):
        proceed = tkMessageBox.askyesno('Confirm Clear', 'Clear all fields?')
        if proceed:
            self.gridRef.set('')
            self.txtShapefile.delete('1.0','end')
            self.sfType.set('')
            self.srchType.set('gr')
            self.searchRadius.set('2')
            self.clientSF.set(0)
            self.siteName.set('')
            self.consultancy.set('')
            self.consultantName.set('')
            self.reqDate.set('')
            self.txtOutputFolder.delete('1.0','end')
            self.namePrefix.set('')
            self.repType.set('pn')
            self.LWS.set(0)
            self.statMap.set(0)
            self.nonStatMap.set(0)
            self.phMap.set(0)
            self.txtLog['state']='normal'
            self.txtLog.delete('1.0','end')
            self.txtLog['state']='disabled'
            self.ignoreLayerNames=[]
            self.lbl40.grid_remove()
            self.entSfTitle.grid_remove()
            self.entSfTitle.state(['disabled'])
# aed 26/9/18
            self.singleSpec.set('')
            self.SLA.set(0)


    def submit(self, *args):
        self.txtLog['state']='normal'
        self.txtLog.delete('1.0','end')
        self.txtLog['state']='disabled'

        if self.siteName.get() == "":
            tkMessageBox.showerror('Error','Please enter site name')
            return

        if self.consultancy.get() == "":
            tkMessageBox.showerror('Error','Please enter consultancy')
            return

        if self.consultantName.get() == "":
            tkMessageBox.showerror('Error','Please enter consultant name')
            return

        if self.reqDate.get() == "":
            tkMessageBox.showerror('Error','Please choose request date')
            return

        if self.txtOutputFolder.get('1.0','end').strip() == "":
            tkMessageBox.showerror('Error','Please choose output folder')
            return

        if self.srchType.get() == 'gr':
            if self.gridRef.get() == "":
                tkMessageBox.showerror('Error','Please select grid reference\n(Or choose to search by Shapefile and enter details)')
                return
# aed 26/9/18 Strip spaces from grid ref before validating
            gridref = self.gridRef.get().replace(" ","")
            if not mapsupportv2.validateGridRef(gridref):
            # if not mapsupportv2.validateGridRef(self.gridRef.get()):
                tkMessageBox.showerror('Error','Please enter a valid grid reference')
                return
            try:
                radius = float(self.searchRadius.get())
            except ValueError:
                tkMessageBox.showerror('Error','Please enter a valid number for the search radius')
                return
        else:
            if self.txtShapefile.get('1.0','end').strip() == "":
                tkMessageBox.showerror('Error','Please choose shapefile\n(Or choose to search by grid reference and enter details)')
                return
            if self.sfType.get() == "":
                tkMessageBox.showerror('Error','Please choose whether shapefile is site boundary or full search area')
                return
            if self.sfType.get() == "sb":
                try:
                    radius = float(self.searchRadius.get())
                except ValueError:
                    tkMessageBox.showerror('Error','Please enter a valid number for the search radius')
                    return

        if self.statMap.get() == 0 and self.nonStatMap.get() == 0 and self.phMap.get() == 0:
            proceed = tkMessageBox.askyesno('Confirm Map Details', 'No map types selected. Proceed with script?')
            if not proceed:
                return

        radius = float(self.searchRadius.get())
        if radius > 5:
            chkMessage = ('Are you sure you want to search with a %skm radius?' % (radius))
            proceed = tkMessageBox.askyesno('Confirm Search Radius', chkMessage)
            if not proceed:
                return
# aed 26/9/18

        if self.repType.get() == 'ss' and self.singleSpec.get() == "":
           tkMessageBox.showerror('Error','Please enter the species name')
           return

        if len(self.ignoreLayerNames) > 0:
            proceed = tkMessageBox.askyesno('Confirm Ignore', ('Proceed with script ignoring layers %s?' % (self.ignoreLayerNames)))
            if not proceed:
                return


        scriptSucceeded = False

        try:
            scriptSucceeded = self.runMapScript()
        except Exception as e:
            tkMessageBox.showerror('Error',('Script failed: %s' % (e)))

        if scriptSucceeded:
            tkMessageBox.showinfo("Complete","Map script completed")

        self.ignoreLayerNames = []


#    def checkExit(self, event=None):
    def checkExit(self, *args):
        if (self.gridRef.get() != '' or self.txtShapefile.get('1.0','end').strip() != '') and self.txtLog.get('1.0','end').strip() == '':
            proceed = tkMessageBox.askyesno('Confirm Quit','Quit without running script?')
            if not proceed:
                return
        self.quit()


    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.setUpConstants()
        if self.getConfigInfo(): # AED
            self.defaultSfTitle='Site Boundary'
            self.createWidgets()
            self.grid()
        else:
            self.quit()


    def writeToLog(self, logString):
        self.txtLog['state']='normal'
        self.txtLog.insert('end', logString+"\n")
        self.txtLog.see('end')
        self.txtLog['state']='disabled'
        self.txtFile.write(logString + "\n")
        self.update_idletasks()


    def setUpConstants(self):
        #
        # Layer names - need to change these if we ever add to / change layers on the map template
        #

        # Add something to GUI for re-run and allow selection of layer(s) to ignore - e.g. to remove Northumberland LWS when producing map with River Tyne bank
        self.ignoreLayerNames = []


        self.statutorySiteLayerNames = ["AONB North East","Local Nature Reserve",
                                   "Marine Conservation Zone","National Nature Reserve",
                                   "National Park","Ramsar","SSSI",
                                   "Special Area of Conservation","Special Protection Area"]

        self.priorityHabitatsLayerName = "Priority Habitats"

        self.nonStatSiteLayerNames = ["Darlington Local Wildlife Sites",
        "Durham Local Wildlife Sites",
        "Gateshead Local Wildlife Sites",
        "Hartlepool Local Wildlife Sites",
        "Hartlepool LGS",
        "Middlesbrough Local Wildlife Sites",
        "Newcastle Local Wildlife Sites",
        "Newcastle SLCI",
        "North Tyneside Local Wildlife Sites",
        "North Tyneside SLCI",
        "Northumberland Local Wildlife Sites",
        "Redcar and Cleveland LGS",
        "Redcar and Cleveland LWS",
        "South Tyneside Local Wildlife Sites",
        "South Tyneside LGS",
        "Stockton Local Wildlife Sites",
        "Sunderland Local Wildlife Sites"]

        self.siteLayers = self.statutorySiteLayerNames + self.nonStatSiteLayerNames

# AED - 10/02/21
##        self.DurhamBirdDataLayerNames = ["10km Squares", "1km Squares", "2km Squares"]

        self.potentialWaxcapLayerNames = ["potential waxcap sites"]

        self.TBCAreaLayerNames = ["hartlepool", "middlesbrough", "redcar_cleveland", "stockton"]

# AED - 10/02/21
##        self.DBCAreaLayerNames = ["darlington", "durham", "gateshead", "southtyneside", "sunderland"]

        self.NTBCAreaLayerNames = ["Newcastle_City_Boundary", "North_Tyneside_Boundary", "Northumberland"]

        self.LWSCitationAreas = ["Durham Local Wildlife Sites",
                            "Gateshead Local Wildlife Sites",
                            "Hartlepool Local Wildlife Sites",
                            "South Tyneside Local Wildlife Sites",
                            "Sunderland Local Wildlife Sites"]

# AED - 20/07/21
        self.NlandBatGroupArea = ["Northumberland Bat Group area"]
# AED - 27/06/22
        self.DurhamBatGroupArea = ["Durham Bat Group area"]

# AED - 10/02/21
##        self.countyLayers = self.TBCAreaLayerNames + self.DBCAreaLayerNames + self.NTBCAreaLayerNames
        self.countyLayers = self.TBCAreaLayerNames + self.NTBCAreaLayerNames
# To label the polygons
        self.labelFields = {"AONB North East" : "NAME",
                       "Local Nature Reserve" : "LNR_NAME",
                       "Marine Conservation Zone" : "MCZ_NAME",
                       "National Nature Reserve" : "NNR_NAME",
                       "National Park" : "NAME",
                       "Ramsar" : "NAME",
                       "SSSI" : "SSSI_NAME",
                       "Special Area of Conservation" : "SAC_NAME",
                       "Special Protection Area" : "SPA_NAME",
                       "Darlington Local Wildlife Sites" : "SITENAME",
                       "Durham Local Wildlife Sites" : "SITENAME",
                       "Gateshead Local Wildlife Sites" : "SITENAME",
                       "Hartlepool Local Wildlife Sites" : "SITENAME",
                       "Hartlepool LGS" : "SITENAME",
                       "Middlesbrough Local Wildlife Sites" : "SITENAME",
                       "Newcastle Local Wildlife Sites" : "SITE",
                       "Newcastle SLCI" : "SITE",
                       "North Tyneside Local Wildlife Sites" : "SITE_NAME",
                       "North Tyneside SLCI" : "NAME",
                       "Northumberland Local Wildlife Sites" : "SITE_NAME",
                       "Redcar and Cleveland LGS" : "SITENAME",
                       "Redcar and Cleveland LWS" : "SITENAME",
                       "South Tyneside Local Wildlife Sites" : "SITENAME",
                       "South Tyneside LGS" : "SITE_NAME",
                       "Stockton Local Wildlife Sites" : "SITENAME",
                       "Sunderland Local Wildlife Sites" : "SITENAME"}

# For areas with citation docs
        self.idFields = {"Durham Local Wildlife Sites" : "SITEID",
                    "Gateshead Local Wildlife Sites" : "SITEID",
                    "Hartlepool Local Wildlife Sites" : "Code",
                    "South Tyneside Local Wildlife Sites" : "SITEID",
                    "Sunderland Local Wildlife Sites" : "SITEID"}


        self.SACName = "Special Area of Conservation"
        self.SPAName = "Special Protection Area"

        self.ptName = "Point Template"
        self.btName = "Buffer Template"
        self.sbName = "Site Boundary Template"
        self.ltName = "Line Template"

        #  Bird, badger & bat group text for e-mails
# AED 10/02/21
##        self.DurhamBirdText = "\nYour search area contains records from Durham Bird Club. I advise you to contact them, \
##as they may hold additional information of interest to you. "
##        self.DurhamBird_OA_Text = "\nYour search area contains records from Durham Bird Club. I advise you to contact them, \
##as they may hold additional information of interest to you. ***** CHECK - Outside DBC Area *****"
        self.TeesmouthBirdBatBadgerText = "I would also advise you to contact Teesmouth Bird Club and Durham Bat and Badger Groups, \
as they may also hold additional information of interest to you.\n"
        self.TeesmouthBirdText = "I would also advise you to contact Teesmouth Bird Club  \
as they may also hold additional information of interest to you.\n"
        self.DurhamBadgerText = "I would also advise you to contact Durham Badger Group, \
as they may also hold additional information of interest to you.\n"
# AED - 27/06/22 - not needed
#        self.DurhamBatText = "I would also advise you to contact Durham Bat Group, \
#as they may also hold additional information of interest to you.\n"
# AED 20/07/21 - Remove reference to Northumberland Bat Group
# AED - And reference to Durham Bat Group
        self.DurhamBatBadgerNorthumberlandBatBadgerBirdText = "I would also advise you to contact Durham Badger Group, Northumberland Badger Groups and Northumberland \
and Tyneside Bird Club, as they may hold additional information of interest to you.\n"
#self.DurhamBatBadgerNorthumberlandBatBadgerBirdText = "I would also advise you to contact Durham Bat and Badger Groups, Northumberland Badger Groups and Northumberland \
#and Tyneside Bird Club, as they may hold additional information of interest to you.\n"
# IZ 05/07/21 - Changed email text to reflect that we supply Northumberland bat data
        self.NorthumberlandBatBadgerBirdText = "I would also advise you to contact Northumberland Badger Group and Northumberland \
and Tyneside Bird Club, as they may hold additional information of interest to you.\n"
        self.NorthumberlandBadgerText = "I would also advise you to contact Northumberland Badger Group, \
as they may hold additional information of interest to you.\n"
        self.NorthumberlandBatText = "\nPlease note that Northumberland Bat Group data are included in this search.\n"
        self.DurhamBatText = "\nPlease note that Durham Bat Group data are included in this search.\n" #AED - 27/06/22
        self.NorthumberlandDurhamBatText = "\nPlease note that Durham and Northumberland Bat Group data are included in this search.\n" #AED - 27/06/22
#AED - 27/06/22        
        #self.DurhamBatBadgerText = "I would also advise you to contact Durham Bat and Badger Groups, as they may also hold \
#additional information of interest to you.\n"
        self.NorthumberlandBirdText = "I would also advise you to contact Northumberland and Tyneside Bird Club, as they may hold \
additional information of interest to you.\n"


    def archiveFiles(self,filePath):
        archivePath = filePath + "\\archive"
        if os.path.exists(archivePath):
            proceed = tkMessageBox.askyesno('Confirm Rerun','Script has been rerun previously. Continue with rerun?\n(Archived files will be deleted)')
            if not proceed:
                return False
            else:
                archiveFiles = os.listdir(archivePath)
                for filename in archiveFiles:
                    fullname = archivePath+'\\'+filename
                    if not os.path.isfile(fullname):
                        tkMessageBox.showerror('Error',"Can't remove archive directory containing subfolder(s)")
                        return False
                for filename in archiveFiles:
                    fullname = archivePath+'\\'+filename
                    os.remove(fullname)
        else:
            os.makedirs(archivePath)

        mapFile = filePath+"\\"+self.namePrefix.get()+" Map.pdf"

        if os.path.isfile(mapFile):
            archiveMapFile = archivePath+"\\"+self.namePrefix.get()+" Map.pdf"
            os.rename(mapFile, archiveMapFile)

        PHMapFile = filePath+"\\"+self.namePrefix.get()+" Priority Habitats Map.pdf"

        if os.path.isfile(PHMapFile):
            archivePHMapFile = archivePath+"\\"+self.namePrefix.get()+" Priority Habitats Map.pdf"
            os.rename(PHMapFile, archivePHMapFile)

        emailFile = filePath+"\\"+self.namePrefix.get()+" EmailText.txt"
##        emailFile = filePath+"\\EmailText.txt"

        if os.path.isfile(emailFile):
            archiveEmailFile = archivePath+"\\EmailText.txt"
            os.rename(emailFile, archiveEmailFile)

        LWSFile = filePath+"\\"+self.namePrefix.get()+" LWS.zip"
##        LWSFile = filePath+"\\LWS.zip"

        if os.path.isfile(LWSFile):
            archiveLWSFile = archivePath+"\\LWS.zip"
            os.rename(LWSFile, archiveLWSFile)

        for filename in os.listdir(filePath+"\\workfiles"):
            # Renaming the lock file causes problems - ????
            # Would be nice to free up the locks

            archiveFilename = archivePath+"\\"+os.path.split(filename)[1]
            currFilename = filePath+"\\workfiles\\"+os.path.split(filename)[1]
            os.rename(currFilename, archiveFilename)

        return True


    def getConfigInfo(self):
        #Read file details from config file
        config_file = r"config.ini"

        try:
            with open(config_file) as f:
                read_f = f.read()

            #Get configuration information out of file
            # Assumes map template on first line & LWS mapping file on second
            self.map_config = read_f.split('\n')[0]
            self.lws_config = read_f.split('\n')[1]
            self.dir_config = read_f.split('\n')[2]
            return True
        except:
            # Error reading config file
            tkMessageBox.showerror('Error','Unable to get config')
            return False

    def runMapScript(self):

        self.startTime = datetime.datetime.now()

        filePath = self.txtOutputFolder.get('1.0','end').strip()
        workfilePath = filePath+"\\workfiles"
        if not os.path.exists(workfilePath):
            os.makedirs(workfilePath)
        else:
            proceed = tkMessageBox.askyesno('Confirm Rerun','Script has already created files in this output folder. Rerun?')
            if proceed:
                archiveSuccess = self.archiveFiles(filePath)
                if not archiveSuccess:
                    print('Could not archive')
                    return False
            else:
                print("Don't want to rerun")
                return False



        textFile = workfilePath+"\\MappingScriptOutput.txt"
        self.txtFile = open(textFile,"w")

        self.writeToLog("Mapping script started at %s" % (self.startTime))

    	# Changes by Jonathan Shiell - Mar 19 - to allow template file to be specified in config file
        # Get mapdocument
        exec(self.map_config)
        mxd = arcpy.mapping.MapDocument(map_script)

        # Get dataframe
        df = arcpy.mapping.ListDataFrames(mxd)[0]

        # Set parameters
        showStat = self.statMap.get() == 1
        showNonStat = self.nonStatMap.get() == 1
        showPriorityHabitats = self.phMap.get() == 1
        getLWSData = self.LWS.get() == 1

        pointName = self.gridRef.get().upper()
        siteBoundarySF = self.txtShapefile.get('1.0','end').strip()
        bufferSizeKm = float(self.searchRadius.get())

        # Get template layers
        ptLayer = None
        btLayer = None
        sbLayer = None
        ltLayer = None

        layers = arcpy.mapping.ListLayers(mxd)
        for layer in layers:
            if layer.name == self.ptName:
                ptLayer = layer
            if layer.name == self.btName:
                btLayer = layer
            if layer.name == self.sbName:
                sbLayer = layer
            if layer.name == self.ltName:
                ltLayer = layer



        # Convert date string to date object
        requestDate = datetime.datetime.strptime(self.reqDate.get(), "%d-%m-%Y").date()

        searchType = "Searching by grid reference"
        if self.srchType.get() == 'sf':
            searchType = "Searching by Shapefile of "
            if self.sfType.get() == 'sb':
                searchType+="Site Boundary"
            else:
                searchType+="Search Area"

        # Calculate parameters derived from input parameters

        mapFile = filePath+"\\"+self.namePrefix.get()+" Map"

        if self.sfType.get() == 'sa':  # Searching with a shapefile of full search area (i.e. don't need to create a buffer)
            bufferName = "Search Area"
        else: # Searching by grid ref or shapefile of site boundary
            bufferSize = int(1000 * bufferSizeKm)
            bufferName = str(bufferSize)+"m Search Area"
            bufferFile = workfilePath+"\\"+bufferName


        if showPriorityHabitats:
            PHMapFile = filePath+"\\"+self.namePrefix.get()+" Priority Habitats Map"

        if getLWSData:
            LWSZip = filePath+"\\"+self.namePrefix.get()+" LWS.zip"
##            LWSZip = filePath+"\\LWS.zip"

        # Parameter values are written to a log file in the output folder for checking later
        self.writeToLog("\nStarting map script with parameters:")
        self.writeToLog(searchType)
        self.writeToLog("Point name: " + pointName)
        self.writeToLog("Site boundary: " + siteBoundarySF)
        self.writeToLog("Buffer size: " + str(bufferSizeKm))
        self.writeToLog("Site name: " + self.siteName.get())
        self.writeToLog("Consultancy: " + self.consultancy.get())
        self.writeToLog("Consultant name: " + self.consultantName.get())
        self.writeToLog("Request date: " + requestDate.strftime("%d-%m-%Y"))
        self.writeToLog("File path: " + filePath)
        self.writeToLog("Stat sites: " + str(showStat))
        self.writeToLog("Non-stat site: " + str(showNonStat))
        self.writeToLog("Priority habitats: " + str(showPriorityHabitats))
        self.writeToLog("LWS data: " + str(getLWSData))
        self.writeToLog("Report type: " + self.repType.get())
        self.writeToLog("Client provided SF: " + str(self.clientSF.get()==1) + "\n")

        if len(self.ignoreLayerNames) > 0:
            self.writeToLog("*** Ignoring layers: %s" % (self.ignoreLayerNames))

        # create a point
        #
        if self.srchType.get() == 'gr':
            # remove spaces and commas from grid reference before calculating coordinates
            pointName = pointName.replace(" ","")
            pointName = pointName.replace(",","")
            XYvalues = mapsupportv2.getEastingsAndNorthings(pointName)
            eastings = XYvalues[0]
            northings = XYvalues[1]
            pointFile = workfilePath+"\\"+pointName
            pt = arcpy.Point()
            pt.X = eastings
            pt.Y = northings
            ptGeoms = []
            ptGeoms.append(arcpy.PointGeometry(pt))
            arcpy.CopyFeatures_management(ptGeoms, pointFile)
            pointLayer = arcpy.mapping.Layer(pointFile+".shp")
        else:
            pointLayer = arcpy.mapping.Layer(siteBoundarySF)
            fieldName1 = "xCentroid"
            fieldName2 = "yCentroid"
            fieldPrecision = 7
            # Expressions are calculated using the Shape Field's geometry property
            expression1 = "float(!SHAPE.CENTROID!.split()[0])"
            expression2 = "float(!SHAPE.CENTROID!.split()[1])"
            # Execute AddField
            arcpy.AddField_management(pointLayer, fieldName1, "LONG",
                              fieldPrecision)
            arcpy.AddField_management(pointLayer, fieldName2, "LONG",
                              fieldPrecision)
            # Execute CalculateField
            arcpy.CalculateField_management(pointLayer, fieldName1, expression1,
                                    "PYTHON")
            arcpy.CalculateField_management(pointLayer, fieldName2, expression2,
                                    "PYTHON")
            cursor = arcpy.SearchCursor(pointLayer)
            row = cursor.next()
            self.writeToLog("Using shapefile with centroid: " + str(row.getValue("xCentroid")) + ", " + str(row.getValue("yCentroid")))

        # add buffer
        #
        if self.sfType.get() == 'sa': # i.e. script has been given search buffer
            bufferLayer = arcpy.mapping.Layer(siteBoundarySF)
        else: # need to create search buffer
            arcpy.Buffer_analysis(pointLayer, bufferFile, bufferSize, dissolve_option="ALL")     # Added , dissolve_option="ALL"
            bufferLayer = arcpy.mapping.Layer(bufferFile+".shp")

        # apply symbology to point
        if self.srchType.get() == 'gr':
            #print "Applying symbology to layer '{}' using '{}'.".format(pointLayer.name, ptLayer)
            arcpy.ApplySymbologyFromLayer_management(pointLayer, ptLayer)
            # AED - 27/7/20
            #pointLayer.name = "<bol>" + pointName + "</bol>"
        elif self.sfType.get() == 'sb':
            desc = arcpy.Describe(siteBoundarySF)
            if desc.shapeType == "Polygon":
                arcpy.ApplySymbologyFromLayer_management(pointLayer, sbLayer)
            elif desc.shapeType == "Polyline":
                arcpy.ApplySymbologyFromLayer_management(pointLayer, ltLayer)
            elif desc.shapeType == "Point":
                arcpy.ApplySymbologyFromLayer_management(pointLayer, ptLayer)
            # AED - 27/7/20
            #pointLayer.name = "<bol>" + self.sfTitle.get() + "</bol>"
            pointLayer.name = self.sfTitle.get()

        # apply symbology to buffer
        #print "Applying symbology to layer '{}' using '{}'.".format(bufferLayer.name, btLayer)
        arcpy.ApplySymbologyFromLayer_management(bufferLayer, btLayer)
        # AED - 27/7/20
        #bufferLayer.name = "<bol>" + bufferName + "</bol>"
        #AED - 11/08/20
        bufferLayer.name = bufferName

        # display point and buffer layers on map
        arcpy.mapping.AddLayer(df,bufferLayer,"TOP")
        if self.sfType.get() <> 'sa':
            arcpy.mapping.AddLayer(df,pointLayer,"TOP")

        # Change dataframe extent to match buffer
        df.extent = bufferLayer.getExtent()

        pointCounty = ""
        statFound = False
        nonStatFound = False
        PHFound = False
##        DurhamBirdDataFound = False   AED 10/02/21
        potentialWaxcapAreaFound = False
        LWSFound = False
        NTBCArea = False
        TBCArea = False
##        DBCArea = False   AED 10/02/21
        NlandBatArea = False    # AED 20/07/21
        DurhamBatArea = False   # AED 27/06/22

        consultancyUpper = self.consultancy.get().upper()
        siteUpper = self.siteName.get().upper()
        exportMap = False
        siteTypes = ""

        layers = arcpy.mapping.ListLayers(mxd)
        textElements = arcpy.mapping.ListLayoutElements(mxd,"TEXT_ELEMENT")
        yPos = textElements[4].elementPositionY

        if textElements[0].text == " ":
            textElements[0].text = "� Natural England 2016, reproduced\nwith the permission of Natural England,\nhttp://www.naturalengland.org.uk/copyright/"

        # Determine which bird club(s) may have data in the search area and whether waxcap check is needed

# AED
        for layer in layers:
            if layer.name in self.countyLayers:
                arcpy.SelectLayerByLocation_management(layer,"INTERSECT",bufferLayer)
                if int(arcpy.GetCount_management(layer).getOutput(0)) > 0:
                    #AED 10/02/21
##                    if layer.name in self.DBCAreaLayerNames:
##                        DBCArea = True
##                    elif layer.name in self.NTBCAreaLayerNames:
                    if layer.name in self.NTBCAreaLayerNames:
                        NTBCArea = True
                    elif layer.name in self.TBCAreaLayerNames:
                        TBCArea = True
                arcpy.SelectLayerByAttribute_management(layer, "CLEAR_SELECTION")
                arcpy.SelectLayerByLocation_management(layer,"INTERSECT",pointLayer)
                if int(arcpy.GetCount_management(layer).getOutput(0)) > 0:
                    if pointCounty == "":
                        pointCounty = layer.name
                    else:
                        pointCounty += ", " + layer.name
                arcpy.SelectLayerByAttribute_management(layer, "CLEAR_SELECTION")
# AED 10/02/21
##            elif layer.name in self.DurhamBirdDataLayerNames:
##                arcpy.SelectLayerByLocation_management(layer,"INTERSECT",bufferLayer)
##                if int(arcpy.GetCount_management(layer).getOutput(0)) > 0:
##                    DurhamBirdDataFound = True
##                arcpy.SelectLayerByAttribute_management(layer, "CLEAR_SELECTION")
            elif layer.name in self.potentialWaxcapLayerNames:
                arcpy.SelectLayerByLocation_management(layer,"INTERSECT",bufferLayer)
                if int(arcpy.GetCount_management(layer).getOutput(0)) > 0:
                    potentialWaxcapAreaFound = True
                arcpy.SelectLayerByAttribute_management(layer, "CLEAR_SELECTION")
# AED 20/07/21
# Check for Northumberland Bat data
            elif layer.name in self.NlandBatGroupArea:
                arcpy.SelectLayerByLocation_management(layer,"INTERSECT",bufferLayer)
                if int(arcpy.GetCount_management(layer).getOutput(0)) > 0:
                    NlandBatArea = True
                arcpy.SelectLayerByAttribute_management(layer, "CLEAR_SELECTION")

# AED 27/06/22
# Check for Durham Bat data
            elif layer.name in self.DurhamBatGroupArea:
                arcpy.SelectLayerByLocation_management(layer,"INTERSECT",bufferLayer)
                if int(arcpy.GetCount_management(layer).getOutput(0)) > 0:
                    DurhamBatArea = True
                arcpy.SelectLayerByAttribute_management(layer, "CLEAR_SELECTION")
                
        self.writeToLog("Search point is in: " + pointCounty)

#AED 10/02/21
##        self.writeToLog("Search area includes: " + ("DBC " if DBCArea else "") + ("NTBC " if NTBCArea else "") + ("TBC" if TBCArea else ""))
        self.writeToLog("Search area includes: " + ("NTBC " if NTBCArea else "") + ("TBC" if TBCArea else ""))


        if showPriorityHabitats:
            titleText = "ECOLOGICAL DATA  SEARCH - \n"
            titleText += "PRIORITY HABITATS\n \n"
            titleText += siteUpper + "\n \n"
            titleText += consultancyUpper + "\n \n"
            titleText += "PLOT PRODUCED: <dyn type=\"date\" format=\"long\"/>"
            textElements[4].text = titleText
            textElements[4].elementPositionY = yPos
            for layer in layers:
                if layer.name == self.priorityHabitatsLayerName:
                    arcpy.SelectLayerByLocation_management(layer,"INTERSECT",bufferLayer)
                    if int(arcpy.GetCount_management(layer).getOutput(0)) > 0:
                        PHFound = True
                        layerName = layer.name
                        layerFile = workfilePath+"\\"+layerName
                        arcpy.CopyFeatures_management(layer, layerFile)
                        copyLayer = arcpy.mapping.Layer(layerFile+".shp")
                        tmpName = copyLayer.name
                        copyLayer.name = "Temp"+tmpName
                        #print "Applying symbology to layer '{}' using '{}'.".format(copyLayer.name, layer)
                        arcpy.ApplySymbologyFromLayer_management(copyLayer, layer)
                        copyLayer.name = tmpName
                        arcpy.mapping.AddLayer(df,copyLayer)
                    arcpy.SelectLayerByAttribute_management(layer, "CLEAR_SELECTION")
            origBuffer = arcpy.mapping.ListLayers(mxd, "*"+bufferName+"*", df)[0]
            arcpy.mapping.RemoveLayer(df,origBuffer)
            if self.sfType.get() <> 'sa':
                origPoint = arcpy.mapping.ListLayers(mxd, "*"+pointLayer.name+"*", df)[0]
                arcpy.mapping.RemoveLayer(df,origPoint)
            arcpy.mapping.AddLayer(df,bufferLayer,"TOP")
            if self.sfType.get() <> 'sa':
                arcpy.mapping.AddLayer(df,pointLayer,"TOP")
            if PHFound:
                legend = arcpy.mapping.ListLayoutElements(mxd,"LEGEND_ELEMENT")[0]
                if legend.elementPositionY + (legend.elementHeight * 1.1) > yPos:
                    legendOffset = legend.elementHeight * 1.2
                    legend.elementPositionY = yPos - legendOffset
                arcpy.mapping.ExportToPDF(mxd,PHMapFile,image_quality = "NORMAL",layers_attributes = "LAYERS_AND_ATTRIBUTES")
                self.writeToLog("Exported Priority Habitats map")
                phLayer = arcpy.mapping.ListLayers(mxd, "*"+layerName+"*", df)[0]
                arcpy.mapping.RemoveLayer(df,phLayer)
            else:
                self.writeToLog("No Priority Habitats found")
        else:
            self.writeToLog("Didn't search for Priority Habitats")



        # Check for overlapping data layers
        layers.reverse()

        if showStat or showNonStat or getLWSData:
            for layer in layers:
                if layer.name in self.ignoreLayerNames:
                    self.writeToLog("Ignore %s" % (layer.name))
                    continue
                if layer.name in self.statutorySiteLayerNames and showStat:
                    arcpy.SelectLayerByLocation_management(layer,"INTERSECT",bufferLayer)
                    if int(arcpy.GetCount_management(layer).getOutput(0)) > 0:
                        labelField = self.labelFields.get(layer.name,"NAME")
                        statFound = True
                        layerName = layer.name
                        layerFile = workfilePath+"\\"+layerName
                        arcpy.CopyFeatures_management(layer, layerFile)
                        copyLayer = arcpy.mapping.Layer(layerFile+".shp")
                        #AED - 11/08/20
                        copyLayer.name = "Temp"+copyLayer.name
                        #print "Applying symbology to layer '{}' using '{}'.".format(copyLayer.name, layer)
                        arcpy.ApplySymbologyFromLayer_management(copyLayer, layer)
                        #AED - 27/7/20
                        #copyLayer.name = "<bol>" + layerName + "</bol>"
                        #AED - 11/08/20
                        copyLayer.name = layerName
                        arcpy.mapping.AddLayer(df,copyLayer,"TOP")
                        newLayer = arcpy.mapping.ListLayers(mxd, "*"+layerName+"*", df)[0]
                        if newLayer.supports("LABELCLASSES"):
                            for lblclass in newLayer.labelClasses:
                                expression = lblclass.expression
                                expression = "[" + labelField + "]"
                                lblclass.expression = '"%s" & %s & "%s"' %("<FNT size='18'><BOL>", expression, "</BOL></FNT>")
                                lblclass.showClassLabels = True
                        newLayer.showLabels = True
                        if layerName == self.SPAName: #Need to set NAME to SPA_NAME so that name appears in PDF model tree
                            field1 = labelField
                            field2 = "NAME"
                            cursor = arcpy.UpdateCursor(newLayer)
                            for row in cursor:
                                row.setValue(field2, row.getValue(field1))
                                cursor.updateRow(row)
                        if layerName == self.SACName: #Need to set NAME to SAC_NAME so that name appears in PDF model tree
                            field1 = labelField
                            field2 = "NAME"
                            cursor = arcpy.UpdateCursor(newLayer)
                            for row in cursor:
                                row.setValue(field2, row.getValue(field1))
                                cursor.updateRow(row)
                    arcpy.SelectLayerByAttribute_management(layer, "CLEAR_SELECTION")
                elif layer.name in self.nonStatSiteLayerNames and (showNonStat or getLWSData):
                    arcpy.SelectLayerByLocation_management(layer,"INTERSECT",bufferLayer)
                    if int(arcpy.GetCount_management(layer).getOutput(0)) > 0:
                        labelField = self.labelFields.get(layer.name,"SITENAME")
                        layerName = layer.name
                        layerFile = workfilePath+"\\"+layerName
                        arcpy.CopyFeatures_management(layer, layerFile)
                        copyLayer = arcpy.mapping.Layer(layerFile+".shp")
                        if showNonStat:
                            nonStatFound = True
                            copyLayer.name = "Temp"+copyLayer.name
                            #print "Applying symbology to layer '{}' using '{}'.".format(copyLayer.name, layer)
                            arcpy.ApplySymbologyFromLayer_management(copyLayer, layer)
                            # AED - 27/7/20
                            #copyLayer.name = "<bol>" + layerName + "</bol>"
                            #AED - 11/08/20
                            copyLayer.name = layerName
                            arcpy.mapping.AddLayer(df,copyLayer,"TOP")
                            newLayer = arcpy.mapping.ListLayers(mxd, "*"+layerName+"*", df)[0]
                            if newLayer.supports("LABELCLASSES"):
                                for lblclass in newLayer.labelClasses:
                                    expression = lblclass.expression
                                    expression = "[" + labelField + "]"
                                    lblclass.expression = '"%s" & %s & "%s"' %("<FNT size='18'><BOL>", expression, "</BOL></FNT>")
                                    lblclass.showClassLabels = True
                            newLayer.showLabels = True
                        if getLWSData:
                            if layer.name in self.LWSCitationAreas:
                                if not LWSFound:
                                    #aed LWSMappings = mapsupportv2.getLWSMappings() #Get LWS file mappings
                                    LWSMappings = self.getLWSMappings() #Get LWS file mappings
                                    LWSDir = zipfile.ZipFile(LWSZip, "w") #Create zip file to store citation documents
                                    LWSFound = True #Only set this flag when we have citation data
                                idField = self.idFields.get(layer.name,"SITEID")
                                self.writeToLog("Selected sites from " + layer.name + "\n")
                                cursor = arcpy.SearchCursor(layerFile+".shp")
                                row = cursor.next()
                                while row:
                                    LWSSiteName = row.getValue(labelField)
                                    siteId = row.getValue(idField)
                                    siteKey = LWSSiteName + siteId
                                    siteFile = LWSMappings.get(siteKey,"No citation document found")
                                    self.writeToLog(LWSSiteName + " (" + siteId + ") - " + siteFile + "\n")
                                    if siteKey in LWSMappings.keys():  # Added .keys() in Sep2016 rewrite
                                        siteFileName = os.path.split(siteFile)[1]
                                        if not siteFileName in LWSDir.namelist():
                                            LWSDir.write(siteFile, siteFileName)
                                    row = cursor.next()
                    arcpy.SelectLayerByAttribute_management(layer, "CLEAR_SELECTION")
                elif layer.name == self.priorityHabitatsLayerName:
                    layer.visible = False

            if LWSFound:
                LWSDir.close()
            origBuffer = arcpy.mapping.ListLayers(mxd, "*"+bufferName+"*", df)[0]
            arcpy.mapping.RemoveLayer(df,origBuffer)
            if self.sfType.get() <> 'sa':
                origPoint = arcpy.mapping.ListLayers(mxd, "*"+pointLayer.name+"*", df)[0]
                arcpy.mapping.RemoveLayer(df,origPoint)
            arcpy.mapping.AddLayer(df,bufferLayer,"TOP")
            if self.sfType.get() <> 'sa':
                arcpy.mapping.AddLayer(df,pointLayer,"TOP")

        if showStat:
            if statFound:
                self.writeToLog("Stat sites found\n")
            else:
                self.writeToLog("No stat sites in search area\n")
        else:
            self.writeToLog("Didn't search for stat sites\n")

        if showNonStat:
            if nonStatFound:
                self.writeToLog("Non stat sites found\n")
            else:
                self.writeToLog("No non-stat sites in search area\n")
        else:
            self.writeToLog("Didn't search for non-stat sites\n")

# AED 10/02/21
##        if DurhamBirdDataFound:
##            self.writeToLog("\nDurham Bird Club data found\n")

        # Format map for export

        if statFound:
            exportMap = True
            if nonStatFound:
                siteTypes = "STATUTORY & NON STATUTORY SITES"
            else:
                siteTypes = "STATUTORY SITES"
        else:  # no stat sites
            if nonStatFound:
                exportMap = True
                textElements[0].text = " "   # Don't need NE copyright if no stat sites on the map
                siteTypes = "NON STATUTORY SITES"


        # Export to PDF

        if exportMap:
            titleText = "ECOLOGICAL DATA  SEARCH - \n"
            titleText += siteTypes + "\n \n"
            titleText += siteUpper + "\n \n"
            titleText += consultancyUpper + "\n \n"
            titleText += "PLOT PRODUCED: <dyn type=\"date\" format=\"long\"/>"
            textElements[4].text = titleText
            textElements[4].elementPositionY = yPos
            legend = arcpy.mapping.ListLayoutElements(mxd,"LEGEND_ELEMENT")[0]
            legendOffset = legend.elementHeight * 1.2
            if legend.elementPositionY + (legend.elementHeight * 1.1) > yPos:
                legend.elementPositionY = yPos - legendOffset
            arcpy.mapping.ExportToPDF(mxd,mapFile,image_quality = "NORMAL",layers_attributes = "LAYERS_AND_ATTRIBUTES")
            self.writeToLog("Exported " + siteTypes.title() + " map\n")
        else:
            self.writeToLog("Stat / Non-stat map not produced\n")

        if potentialWaxcapAreaFound:
            self.writeToLog("\n**************************************\nPotential waxcap sites in search area\n**************************************\n")
            
        emailFile = filePath+"\\"+self.namePrefix.get()+" EmailText.txt"
##        emailFile = filePath+"\\EmailText.txt"
        emFile = codecs.open(emailFile,encoding="utf-8",mode="w")

        emDate = requestDate.strftime("%d %B %Y") # format request date for email

        emFile.write("Hi " + self.consultantName.get() + ",\n")
        emFile.write("\n" + self.siteName.get() + "\n")
        emFile.write("\nWith regard to your data request for the above project on " + emDate + " please find enclosed the following:\n")

# aed 26/9/18 - Single species
        if self.repType.get() == 'pn':
            emFile.write("\nA list of protected and notable species found within the search area.\n")
        elif self.repType.get() == 'ss':
# Would be nice to ensure the species name is in lower case and has no trailing "s"
            emFile.write("\nA list of " + self.singleSpec.get().lower() + " records found within the search area.\n")
        else:
            emFile.write("\nLists of protected and notable species and other species found within the search area.\n")

        if statFound:
            if nonStatFound:
                emFile.write("An interactive PDF map of statutory and non-statutory sites found within the search area.\n")
            else:
                emFile.write("An interactive PDF map of statutory sites found within the search area.\n")
                if showNonStat:
                    emFile.write("Please note that no non-statutory sites were found within the search area.\n")
        else:
            if nonStatFound:
                emFile.write("An interactive PDF map of non-statutory sites found within the search area.\n")
                if showStat:
                    emFile.write("Please note that no statutory sites were found within the search area.\n")
            else:
                if showStat:
                    if showNonStat:
                        emFile.write("Please note that no statutory sites or non-statutory sites were found within the search area.\n")
                    else:
                        emFile.write("Please note that no statutory sites were found within the search area.\n")
                else:
                    if showNonStat:
                        emFile.write("Please note that no non-statutory sites were found within the search area.\n")

        if PHFound:
            emFile.write("An interactive PDF map of priority habitats found within the search area.\n")
        else:
            if showPriorityHabitats:
                emFile.write("Please note that no priority habitats were found within the search area.\n")

        if LWSFound:
            emFile.write("Citation documents for local wildlife sites found within the search area.\n")

        if statFound or nonStatFound or PHFound:
            emFile.write("Instructions for accessing interactive PDF maps.\n")

        costCalculated = False
        cost = 0

        if bufferSizeKm == 1:
            cost = 80
            costCalculated = True
        elif bufferSizeKm == 2 or bufferSizeKm == 3:
            cost = 120
            costCalculated = True
# aed 26/9/18
        if self.repType.get() == 'ss':
            cost = 40

        if LWSFound:
            cost += 40

        if self.srchType.get() == 'sf' and self.clientSF.get() <> 1:  # Only charge for shapefile if not provided by client
            cost += 40

        if self.SLA.get() <> 1:
            if costCalculated:
                emFile.write("\nOur invoice for " + unichr(163) + str(cost) + " + VAT will follow in due course.\n\n")
            else:
                emFile.write("\nOur invoice for " + unichr(163) + "****** CALCULATE COST ****** + VAT will follow in due course.\n\n")
        else:
            emFile.write("\n")

        # TODO - when consultancy is selected from database, include something to say whether client should be charged or not (e.g. NWL) and remove cost line if no charge

# Changes made by Jonathan Shiell - March 2019 - to modify text for single species searches

        #### Please populate as required - alternative could be to load in from text file.
        #### Please also use all lower case.

        badger = ['badger']
        bats = ['bat']
        newts = ["great crested newt",]
        bird = ["barn owl"]
        other_non_bird = []

 
        if NTBCArea:

            if ((self.repType.get() != 'ss') or (self.singleSpec.get().lower() not in (newts+badger+bats+other_non_bird+bird))):
                emFile.write(self.NorthumberlandBatBadgerBirdText)
            elif (self.repType.get() == 'ss') and (self.singleSpec.get().lower() in (badger)):
                emFile.write(self.NorthumberlandBadgerText)
##            elif (self.repType.get() == 'ss') and (self.singleSpec.get().lower() in (bats)):
##                emFile.write(self.NorthumberlandBatText)
            elif (self.repType.get() == 'ss') and (self.singleSpec.get().lower() in (bird)):
                emFile.write(self.NorthumberlandBirdText)
        elif TBCArea:
            if ((self.repType.get() != 'ss') or (self.singleSpec.get().lower() not in (newts+badger+bats+other_non_bird+bird))):
                emFile.write(self.TeesmouthBirdBatBadgerText )
            elif (self.repType.get() == 'ss') and (self.singleSpec.get().lower() in (badger)):
                emFile.write(self.DurhamBadgerText)
# AED - 27/06/22
            #elif (self.repType.get() == 'ss') and (self.singleSpec.get().lower() in (bats)):
            #    emFile.write(self.DurhamBatText )

            elif (self.repType.get() == 'ss') and (self.singleSpec.get().lower() in (bird)):
                emFile.write(self.TeesmouthBirdText)
        else:  # In Durham area but not Durham Bird Club Area - unlikely to get here
#AED - 27/06/22
            #if ((self.repType.get() != 'ss') or (self.singleSpec.get().lower() not in (newts+badger+bats+other_non_bird+bird))):
            #    emFile.write(self.DurhamBatBadgerText )
            #elif (self.repType.get() == 'ss') and (self.singleSpec.get().lower() in (badger)):
            if (self.repType.get() == 'ss') and (self.singleSpec.get().lower() in (badger)):
                emFile.write(self.DurhamBadgerText)
# AED - 27/06/22
            #elif (self.repType.get() == 'ss') and (self.singleSpec.get().lower() in (bats)):
            #    emFile.write(self.DurhamBatText)
            
# AED 10/02/21
##            elif (self.repType.get() == 'ss') and (self.singleSpec.get().lower() in (bird)):
##                emFile.write(self.DurhamBirdText)

#AED end
        if potentialWaxcapAreaFound and (self.repType.get() != 'ss'):
            emFile.write ("\nPlease note that there are potential waxcap grassland sites in your search area. We have therefore left waxcap species within your report.\n")

# AED 27/06/22
        if NlandBatArea and DurhamBatArea:
            if (self.repType.get() != 'ss'):
                emFile.write(self.NorthumberlandDurhamBatText)
            elif (self.repType.get() == 'ss') and (self.singleSpec.get().lower() in (bats)):
                    emFile.write(self.NorthumberlandDurhamBatText)      
# AED 20/07/21
        #if NlandBatArea:
        elif NlandBatArea:

            if (self.repType.get() != 'ss'):
                emFile.write(self.NorthumberlandBatText)
            elif (self.repType.get() == 'ss') and (self.singleSpec.get().lower() in (bats)):
                    emFile.write(self.NorthumberlandBatText)            

# AED 27/06/22
        elif DurhamBatArea:

            if (self.repType.get() != 'ss'):
                emFile.write(self.DurhamBatText)
            elif (self.repType.get() == 'ss') and (self.singleSpec.get().lower() in (bats)):
                    emFile.write(self.DurhamBatText)     

        emFile.write("\nWhile the enclosed information is provided in good faith, the Environmental Records Information \
Centre North East accepts no responsibility for the completeness or accuracy of the data provided, or the use to \
which it is put or for any loss which may arise from its use.\n")

        emFile.write("\nI hope you find this information useful.\n\nKind regards,\n")

        emFile.close()

        endTime = datetime.datetime.now()
        elapsedTime = endTime - self.startTime
        self.writeToLog("\nMapping script ended at %s (Elapsed time %s)\n" % (endTime, elapsedTime))
        self.txtFile.close()

        del mxd
        ##del arcpy #aed

        return True

    def getLWSMappings (self):
        "This reads LWS mappings from file and creates a dictionary keyed on site description and id"
        # Replaces the one in mapsupportv2

        exec(self.lws_config)
        mappingsFile = open(lws_mappings)

        mappings = {}
        for line in mappingsFile:
            parts = line.split("\t")
            sitename = parts[0].replace('"','')    # 23/03/2016 added .replace('"','')
            siteid = parts[1]
            sitefile = parts[3].strip().replace('"','')
            key = sitename + siteid
            mappings[key] = sitefile
        return mappings



class LayerPicker(Frame):

    def createWidgets(self):

        self.layerLabels = []
        self.layerChkButtons = []
        self.layerIgn = []

        for layer in self.layers:
            lbl = Label(self, text=layer)
            ign = IntVar()
            if layer in self.currIgnoreLayers:
                ign.set(1)
            else:
                ign.set(0)
            chkignore = Checkbutton(self, variable=ign)
            self.layerLabels.append(lbl)
            self.layerChkButtons.append(chkignore)
            self.layerIgn.append(ign)

        self.btnOK = Button(self, text='OK', width=10, command=self.submit)
        self.btnCancel = Button(self, text='Cancel', width=10, command=self.cancel)

        rowPos = 2

        for i in range(len(self.layers)):
            chk = self.layerChkButtons[i]
            lbl = self.layerLabels[i]
            chk.grid(row=rowPos, column=0, sticky=E, padx=5, pady=5)
            lbl.grid(row=rowPos, column=1, sticky=W, padx=5, pady=5)
            rowPos += 2

        self.btnOK.grid(row = rowPos, column=0, sticky=E, padx=5, pady=5)
        self.btnCancel.grid(row = rowPos, column=1, sticky=E, padx=5, pady=5)


    def submit(self):
        print('Submit')
        ignoreLayers = []
        for i in range(len(self.layers)):
            if self.layerIgn[i].get() == 1:
                ignoreLayers.append(self.layerLabels[i]['text'])
        self.app.ignoreDone(ignoreLayers)


    def cancel(self):
        print('Cancel')
        self.app.ignoreDone(None)


    def __init__(self, master, app):
        Frame.__init__(self, master)
        self.app = app
        self.layers = self.app.siteLayers
        self.currIgnoreLayers = self.app.ignoreLayerNames
        self.createWidgets()





root = Tk()
#AED 27/6/22 - add version number
root.title("ERIC Mapping Script V13")
#AED 27/7/20
root.geometry('1080x610+10+10')
app = Application(master=root)
app.mainloop()
root.destroy()
