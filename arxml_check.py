# -*- coding: utf-8 -*-
"""
Created on Wed Jun 12 09:27:49 2024

@author: liuchangjun

@email: 2857418430@qq.com
"""

import xml.etree.ElementTree as ET
import xlwings as xw
from pathlib import Path


class ArxmlCheck(object):
    def __init__(self, arxml_parh, excel_path):
        self.arxml_parh = arxml_parh
        self.excel_path = excel_path
        
    def get_info(self):
        self.app = xw.App(visible=False, add_book=False)
        try:
            self.wb = self.app.books.open(self.excel_path)
        except:
            self.wb = self.app.books.add()

        tree = ET.parse(self.arxml_parh)
        self.root = tree.getroot()

        # instances = self.root.findall(".//{http://autosar.org/schema/r4.0}AR-PACKAGE")
        # for instance in instances:
        #     if instance.find("{http://autosar.org/schema/r4.0}SHORT-NAME").text == 'Interfaces':
        #         InterfacesElements = instance.find('{http://autosar.org/schema/r4.0}ELEMENTS')
        #         # self.SenderReceiverInterface = InterfacesElements.findall('{http://autosar.org/schema/r4.0}SENDER-RECEIVER-INTERFACE')
                
        #     elif instance.find("{http://autosar.org/schema/r4.0}SHORT-NAME").text == 'Constants':
        #         self.ConstantsElements = instance.find('{http://autosar.org/schema/r4.0}ELEMENTS')
        #         self.ConstantSpecification = self.ConstantsElements.findall('{http://autosar.org/schema/r4.0}CONSTANT-SPECIFICATION')
            
        #     elif instance.find("{http://autosar.org/schema/r4.0}SHORT-NAME").text == 'ComponentTypes':
        #         self.ComponentTypesElements = instance.find('{http://autosar.org/schema/r4.0}ELEMENTS')
        #         self.RPortProtoType = self.ComponentTypesElements.findall('.//{http://autosar.org/schema/r4.0}R-PORT-PROTOTYPE')
        #         self.PPortProtoType = self.ComponentTypesElements.findall('.//{http://autosar.org/schema/r4.0}P-PORT-PROTOTYPE')
        
    def N003_2(self):
        SR_SHORT_NAMES, SR_DATAELEMENT_SHORT_NAMES, UUID = [], [], []   
        self.SenderReceiverInterface = self.root.findall(".//{http://autosar.org/schema/r4.0}ELEMENTS/{http://autosar.org/schema/r4.0}SENDER-RECEIVER-INTERFACE")
        for SenderReceiverInterface in self.SenderReceiverInterface:
            SR_SHORT_NAME = SenderReceiverInterface.find('{http://autosar.org/schema/r4.0}SHORT-NAME').text
            SR_DATAELEMENT_SHORT_NAME = SenderReceiverInterface.find('{http://autosar.org/schema/r4.0}DATA-ELEMENTS//{http://autosar.org/schema/r4.0}SHORT-NAME').text
            if SR_SHORT_NAME != 'IF_'+SR_DATAELEMENT_SHORT_NAME:
                SR_SHORT_NAMES.append(SR_SHORT_NAME)
                SR_DATAELEMENT_SHORT_NAMES.append(SR_DATAELEMENT_SHORT_NAME)
                UUID.append(SenderReceiverInterface.get('UUID'))
        try:
            sheet_ = self.wb.sheets.add('N003(2)', after=self.wb.sheets[-1]) 
            sheet_.name = 'N003(2)'
            sheet_.range('A1').value = ['INDEX', 'SR/SHORT_NAMES', 'SR/DATAELEMENT/SHORT_NAMES', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SR_SHORT_NAMES, SR_DATAELEMENT_SHORT_NAMES, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
        
    def A006_1(self):
        SR_NUM_OF_DATAELEMENT, UUID = [], []
        self.SenderReceiverInterface = self.root.findall(".//{http://autosar.org/schema/r4.0}ELEMENTS/{http://autosar.org/schema/r4.0}SENDER-RECEIVER-INTERFACE")
        for SenderReceiverInterface in self.SenderReceiverInterface:
            if len(SenderReceiverInterface.findall('{http://autosar.org/schema/r4.0}DATA-ELEMENTS')) != 1:
                SR_NUM_OF_DATAELEMENT.append(len(SenderReceiverInterface.findall('{http://autosar.org/schema/r4.0}DATA-ELEMENTS')))
                UUID.append(SenderReceiverInterface.get('UUID'))
        try:
            sheet_ = self.wb.sheets.add('A006(1)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'SR/NUM_OF_DATAELEMENT', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SR_NUM_OF_DATAELEMENT, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)

    def N019_1(self):
        SR_DATAELEMENT_SHORT_NAMES, SR_DATAELEMENT_TYPE_TREFS, UUID = [], [], []
        self.SenderReceiverInterface = self.root.findall(".//{http://autosar.org/schema/r4.0}ELEMENTS/{http://autosar.org/schema/r4.0}SENDER-RECEIVER-INTERFACE")
        for SenderReceiverInterface in self.SenderReceiverInterface:
            I_DataElement = SenderReceiverInterface.find('{http://autosar.org/schema/r4.0}DATA-ELEMENTS//{http://autosar.org/schema/r4.0}VARIABLE-DATA-PROTOTYPE')
            SR_DATAELEMENT_SHORT_NAME = I_DataElement.find('{http://autosar.org/schema/r4.0}SHORT-NAME').text
            SR_DATAELEMENT_TYPE_TREF = I_DataElement.find('{http://autosar.org/schema/r4.0}TYPE-TREF').text
            SR_DATAELEMENT_TYPE_TREF = SR_DATAELEMENT_TYPE_TREF.split('/')[-1]
            if SR_DATAELEMENT_TYPE_TREF  not in [i+SR_DATAELEMENT_SHORT_NAME for i in ['APDT_', 'AADT_', 'ARDT_']]:
                SR_DATAELEMENT_SHORT_NAMES.append(SR_DATAELEMENT_SHORT_NAME)
                SR_DATAELEMENT_TYPE_TREFS.append(SR_DATAELEMENT_TYPE_TREF)
                UUID.append(SenderReceiverInterface.get('UUID'))
        try:
            sheet_ = self.wb.sheets.add('N019(1)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'SR/DATAELEMENT/SHORT_NAMES', 'SR/DATAELEMENT/TYPE_TREFS', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SR_DATAELEMENT_SHORT_NAMES, SR_DATAELEMENT_TYPE_TREFS, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
        
    def A041_1(self):
        self.SenderReceiverInterface = self.root.findall(".//{http://autosar.org/schema/r4.0}ELEMENTS//{http://autosar.org/schema/r4.0}SENDER-RECEIVER-INTERFACE")
        SR_DATAELEMENT_SWCALIBRATIONACCES, UUID = [], []
        for SenderReceiverInterface in self.SenderReceiverInterface:
            SwCalibrationAcces = SenderReceiverInterface.find(".//{http://autosar.org/schema/r4.0}SW-CALIBRATION-ACCESS")
            if SwCalibrationAcces.text  != 'READ-ONLY':
                SR_DATAELEMENT_SWCALIBRATIONACCES.append(SwCalibrationAcces.text)
                UUID.append(SenderReceiverInterface.get('UUID'))
        try:
            sheet_ = self.wb.sheets.add('A041(1)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'SR/SwCALIBRATIONACCESS', 'UUID'] 
            sheet_.range('A1').expand(mode="right").font.bold = True 
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SR_DATAELEMENT_SWCALIBRATIONACCES, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)

    def A042_1(self):
        self.SenderReceiverInterface = self.root.findall(".//{http://autosar.org/schema/r4.0}ELEMENTS/{http://autosar.org/schema/r4.0}SENDER-RECEIVER-INTERFACE")
        SR_DATAELEMENT_SW_IMPL_POLICY, UUID = [], []
        for SenderReceiverInterface in self.SenderReceiverInterface:
            SwImplPolicy = SenderReceiverInterface.find(".//{http://autosar.org/schema/r4.0}SW-IMPL-POLICY")
            if SwImplPolicy.text  != 'STANDARD':
                SR_DATAELEMENT_SW_IMPL_POLICY.append(SwImplPolicy.text)
                UUID.append(SenderReceiverInterface.get('UUID'))
        try:
            sheet_ = self.wb.sheets.add('A042(1)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'SR/SW-IMPL-POLICY', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SR_DATAELEMENT_SW_IMPL_POLICY, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
    
    def A111_1(self):
        self.SenderReceiverInterface = self.root.findall(".//{http://autosar.org/schema/r4.0}ELEMENTS/{http://autosar.org/schema/r4.0}SENDER-RECEIVER-INTERFACE")
        self.ConstantSpecification = self.root.findall(".//{http://autosar.org/schema/r4.0}ELEMENTS/{http://autosar.org/schema/r4.0}CONSTANT-SPECIFICATION")
        SR_DATAELEMENT_SHORT_NAMES, sr_dataelement_short_names, UUID, ERROR, CS_SHORT_NAMES = [], [], [], [], []
        for SenderReceiverInterface in self.SenderReceiverInterface:
            SR_DATAELEMENT_SHORT_NAME = SenderReceiverInterface.find('{http://autosar.org/schema/r4.0}DATA-ELEMENTS//{http://autosar.org/schema/r4.0}SHORT-NAME').text
            SR_DATAELEMENT_SHORT_NAMES.append('IV_'+SR_DATAELEMENT_SHORT_NAME)
            sr_dataelement_short_names.append('iv_'+SR_DATAELEMENT_SHORT_NAME.lower())
        for ConstantSpecification in self.ConstantSpecification:
            CONSTANT_SPECIFICATION_SHORT_NAME = ConstantSpecification.find('{http://autosar.org/schema/r4.0}SHORT-NAME').text
            if CONSTANT_SPECIFICATION_SHORT_NAME not in SR_DATAELEMENT_SHORT_NAMES:
                UUID.append(ConstantSpecification.get('UUID'))
                CS_SHORT_NAMES.append(CONSTANT_SPECIFICATION_SHORT_NAME)
                if CONSTANT_SPECIFICATION_SHORT_NAME.lower() in sr_dataelement_short_names:
                    ERROR.append(str(CONSTANT_SPECIFICATION_SHORT_NAME+':存在但是大小写不一致'))
                else:
                    ERROR.append('Not in SR/DATAELEMENT/SHORT_NAME')
            else:
                FLAG = ConstantSpecification.findall('.//{http://autosar.org/schema/r4.0}APPLICATION-VALUE-SPECIFICATION')
                ErrorInfo = ''
                ErrorSingle = 0
                for ind,flag in enumerate(FLAG):
                    CS_VALUES_PHYS = flag.find('.//{http://autosar.org/schema/r4.0}SW-VALUES-PHYS')
                    if not CS_VALUES_PHYS:
                        ErrorInfo += 'No.'+str(ind+1)+' Undefined SW-VALUES-PHYS;'
                        ErrorSingle = 1
                if ErrorSingle:
                    UUID.append(ConstantSpecification.get('UUID'))
                    ERROR.append(ErrorInfo)
                    CS_SHORT_NAMES.append(CONSTANT_SPECIFICATION_SHORT_NAME)
        try:           
            sheet_ = self.wb.sheets.add('A111(1)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'ERROR', 'CS/SHORT_NAME', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), ERROR, CS_SHORT_NAMES, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
    
    def A131_1(self):
        self.SenderReceiverInterface = self.root.findall(".//{http://autosar.org/schema/r4.0}ELEMENTS/{http://autosar.org/schema/r4.0}SENDER-RECEIVER-INTERFACE")
        SR_DATAELE_INIT_CONSTANT_REFS, SR_DATAELE_SHORT_NAMES, UUID = [], [], []
        for SenderReceiverInterface in self.SenderReceiverInterface:
            SR_DATAELE_INIT_CONSTANT_REF = SenderReceiverInterface.find(".//{http://autosar.org/schema/r4.0}CONSTANT-REF")
            if SR_DATAELE_INIT_CONSTANT_REF.get('DEST') == 'CONSTANT-SPECIFICATION':
                SR_DATAELE_INIT_CONSTANT_REF = (SR_DATAELE_INIT_CONSTANT_REF.text).split('/')[-1]
                VDP_SHORT_NAME = SenderReceiverInterface.find(".//{http://autosar.org/schema/r4.0}VARIABLE-DATA-PROTOTYPE/{http://autosar.org/schema/r4.0}SHORT-NAME").text
                if SR_DATAELE_INIT_CONSTANT_REF != 'IV_'+VDP_SHORT_NAME:
                    SR_DATAELE_INIT_CONSTANT_REFS.append(SR_DATAELE_INIT_CONSTANT_REF)
                    SR_DATAELE_SHORT_NAMES.append(VDP_SHORT_NAME)
                    UUID.append(SenderReceiverInterface.get('UUID'))
            else:
                SR_DATAELE_INIT_CONSTANT_REFS.append('')
                SR_DATAELE_SHORT_NAMES.append('CONSTANT_REF DEST!=CONSTANT-SPECIFICATION')
                UUID.append(SenderReceiverInterface.get('UUID'))
        try:
            sheet_ = self.wb.sheets.add('A131(1)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'SR/DATAELE/INIT/CONSTANT_REFS', 'SR/DATAELE/SHORT_NAMES', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SR_DATAELE_INIT_CONSTANT_REFS, SR_DATAELE_SHORT_NAMES, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
            
    def A067_1(self):
        self.SenderReceiverInterface = self.root.findall(".//{http://autosar.org/schema/r4.0}ELEMENTS/{http://autosar.org/schema/r4.0}SENDER-RECEIVER-INTERFACE")
        SR_INVALID_POLICY, SR_SHORTNAMES, UUID = [], [], []
        for SenderReceiverInterface in self.SenderReceiverInterface:
            SR_InvalidationPolicy = SenderReceiverInterface.find(".//{http://autosar.org/schema/r4.0}INVALIDATION-POLICY")
            SR_SHORTNAME = SenderReceiverInterface.find("{http://autosar.org/schema/r4.0}SHORT-NAME").text
            if SR_InvalidationPolicy:
                SR_INVALID_POLICY.append(SR_InvalidationPolicy.text)
                SR_SHORTNAMES.append(SR_SHORTNAME)
                UUID.append(SenderReceiverInterface.get('UUID'))
        try:
            sheet_ = self.wb.sheets.add('A067(1)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'SR/INVALID_POLICY', 'SR/SHORT_NAME', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SR_INVALID_POLICY, SR_SHORTNAMES, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
        
    def N007_1(self):
        self.RPortProtoType = self.root.findall(".//{http://autosar.org/schema/r4.0}PORTS/{http://autosar.org/schema/r4.0}R-PORT-PROTOTYPE")
        RPortShortNames, RPortDataElementRefs, UUID = [], [], []
        for RPortProtoType in self.RPortProtoType:
            RPortShortName = RPortProtoType.find("{http://autosar.org/schema/r4.0}SHORT-NAME").text
            RPortDataElementRef = RPortProtoType.find(".//{http://autosar.org/schema/r4.0}DATA-ELEMENT-REF")
            try:
                RPortDataElementRef = RPortDataElementRef.text
                RPortDataElementRef = RPortDataElementRef.split('/')[-1]
                if RPortShortName != 'R_'+RPortDataElementRef:
                    RPortShortNames.append(RPortShortName)
                    RPortDataElementRefs.append(RPortDataElementRef)
                    UUID.append(RPortProtoType.get('UUID'))
            except:
                RPortShortNames.append(RPortShortName)
                RPortDataElementRefs.append('Undefined DATA-ELEMENT-REF')
                UUID.append(RPortProtoType.get('UUID'))
        try:
            sheet_ = self.wb.sheets.add('N007(1)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'RPort/ShortName', 'RPort/DataElementRef', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), RPortShortNames, RPortDataElementRefs, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
            
    def N008_1(self):
        self.PPortProtoType = self.root.findall(".//{http://autosar.org/schema/r4.0}PORTS/{http://autosar.org/schema/r4.0}P-PORT-PROTOTYPE")
        PPortShortNames, PPortDataElementRefs, UUID = [], [], []
        for PPortProtoType in self.PPortProtoType:
            PPortShortName = PPortProtoType.find("{http://autosar.org/schema/r4.0}SHORT-NAME").text
            PPortDataElementRef = PPortProtoType.find(".//{http://autosar.org/schema/r4.0}DATA-ELEMENT-REF")
            try:
                PPortDataElementRef = PPortDataElementRef.text
                PPortDataElementRef = PPortDataElementRef.split('/')[-1]
                if PPortShortName != 'P_'+PPortDataElementRef:
                    PPortShortNames.append(PPortShortName)
                    PPortDataElementRefs.append(PPortDataElementRef)
                    UUID.append(PPortProtoType.get('UUID'))
            except:
                PPortShortNames.append(PPortShortName)
                PPortDataElementRefs.append('Undefined DATA-ELEMENT-REF')
                UUID.append(PPortProtoType.get('UUID'))
        try:
            sheet_ = self.wb.sheets.add('N008(1)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'PPort/ShortName', 'PPort/DataElementRef', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), PPortShortNames, PPortDataElementRefs, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
                
    def A030_2(self):
        self.PPortProtoType = self.root.findall(".//{http://autosar.org/schema/r4.0}PORTS/{http://autosar.org/schema/r4.0}P-PORT-PROTOTYPE")
        PPort_Nonqueued_Sender_ConstantRefs, PPort_Nonqueued_Sender_DataElementRefs, UUID = [], [], []
        for PPortProtoType in self.PPortProtoType:
            PPort = PPortProtoType.find("{http://autosar.org/schema/r4.0}PROVIDED-INTERFACE-TREF")
            if PPort.get('DEST') == 'SENDER-RECEIVER-INTERFACE':
                PPortConstantRef = PPortProtoType.find(".//{http://autosar.org/schema/r4.0}CONSTANT-REF").text
                PPortConstantRef = PPortConstantRef.split('/')[-1]
                PPortDataElementRef = PPortProtoType.find(".//{http://autosar.org/schema/r4.0}DATA-ELEMENT-REF").text
                PPortDataElementRef = PPortDataElementRef.split('/')[-1]
                if PPortConstantRef != 'IV_' + PPortDataElementRef:
                    PPort_Nonqueued_Sender_ConstantRefs.append(PPortConstantRef)
                    PPort_Nonqueued_Sender_DataElementRefs.append(PPortDataElementRef)
                    UUID.append(PPortProtoType.get('UUID'))
        try:
            sheet_ = self.wb.sheets.add('A030(2)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'PPort/Nonqueued_Sender/ConstantRefs', 'PPort/Nonqueued_Sender/DataElementRefs', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), PPort_Nonqueued_Sender_ConstantRefs, PPort_Nonqueued_Sender_DataElementRefs, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
                
    def A031_2(self):
        self.RPortProtoType = self.root.findall(".//{http://autosar.org/schema/r4.0}PORTS/{http://autosar.org/schema/r4.0}R-PORT-PROTOTYPE")
        RPort_Nonqueued_Sender_ConstantRefs, RPort_Nonqueued_Sender_DataElementRefs, UUID = [], [], []
        for RPortProtoType in self.RPortProtoType:
            RPort = RPortProtoType.find("{http://autosar.org/schema/r4.0}REQUIRED-INTERFACE-TREF")
            if RPort.get('DEST') == 'SENDER-RECEIVER-INTERFACE':
                RPortConstantRef = RPortProtoType.find(".//{http://autosar.org/schema/r4.0}CONSTANT-REF").text
                RPortConstantRef = RPortConstantRef.split('/')[-1]
                RPortDataElementRef = RPortProtoType.find(".//{http://autosar.org/schema/r4.0}DATA-ELEMENT-REF").text
                RPortDataElementRef = RPortDataElementRef.split('/')[-1]
                if RPortConstantRef != 'IV_' + RPortDataElementRef:
                    RPort_Nonqueued_Sender_ConstantRefs.append(RPortConstantRef)
                    RPort_Nonqueued_Sender_DataElementRefs.append(RPortDataElementRef)
                    UUID.append(RPortProtoType.get('UUID'))
        try:
            sheet_ = self.wb.sheets.add('A031(2)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'RPort/Nonqueued_Receiver/ConstantRefs', 'RPort/Nonqueued_Receiver/DataElementRefs', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), RPort_Nonqueued_Sender_ConstantRefs, RPort_Nonqueued_Sender_DataElementRefs, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
            
    def A032_3(self):
        AliveTimeOuts, ShortNames = [], []
        self.RPort = self.root.findall(".//{http://autosar.org/schema/r4.0}PORTS/{http://autosar.org/schema/r4.0}R-PORT-PROTOTYPE")
        for RPort in self.RPort:
            ShortName = RPort.find(".//{http://autosar.org/schema/r4.0}SHORT-NAME").text
            try:
                AliveTimeOut = RPort.find(".//{http://autosar.org/schema/r4.0}ALIVE-TIMEOUT").text
                if AliveTimeOut != str(0):
                    AliveTimeOuts.append(AliveTimeOut)
                    ShortNames.append(ShortName)
            except:
                AliveTimeOuts.append('Undefined ALIVE-TIMEOUT')
                ShortNames.append(ShortName)
        try:
            sheet_ = self.wb.sheets.add('A032(3)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'RPort/Nonqueued_Receiver/AliveTimeOut', 'RPort/Nonqueued_Receiver/ShortName']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(ShortNames)+1)), AliveTimeOuts, ShortNames]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
      
    def A116_0(self):
        HandleOutOfRanges, ShortNames = [], []
        self.RPort = self.root.findall(".//{http://autosar.org/schema/r4.0}PORTS/{http://autosar.org/schema/r4.0}R-PORT-PROTOTYPE")
        for RPort in self.RPort:
            ShortName = RPort.find(".//{http://autosar.org/schema/r4.0}SHORT-NAME").text
            try:
                HandleOutOfRange = RPort.find(".//{http://autosar.org/schema/r4.0}HANDLE-OUT-OF-RANGE").text
                if HandleOutOfRange != 'NONE':
                    HandleOutOfRanges.append(HandleOutOfRange)
                    ShortNames.append(ShortName)
            except:
                HandleOutOfRanges.append('Undefined HANDLE-OUT-OF-RANGE')
                ShortNames.append(ShortName)
        try:
            sheet_ = self.wb.sheets.add('A116(0)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'RPort/Nonqueued_Receiver/HandleOutOfRange', 'RPort/Nonqueued_Receiver/ShortName']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(ShortNames)+1)), HandleOutOfRanges, ShortNames]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
            
    def A117_0(self):
        EnableUpdates, ShortNames = [], []
        self.RPort = self.root.findall(".//{http://autosar.org/schema/r4.0}PORTS/{http://autosar.org/schema/r4.0}R-PORT-PROTOTYPE")
        for RPort in self.RPort:
            ShortName = RPort.find(".//{http://autosar.org/schema/r4.0}SHORT-NAME").text
            try:
                EnableUpdate = RPort.find(".//{http://autosar.org/schema/r4.0}ENABLE-UPDATE").text
                if EnableUpdate != 'false':
                    EnableUpdates.append(EnableUpdate)
                    ShortNames.append(ShortName)
            except:
                EnableUpdates.append('Undefined ENABLE-UPDATE')
                ShortNames.append(ShortName)
        try:
            sheet_ = self.wb.sheets.add('A117(0)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'RPort/Nonqueued_Receiver/EnableUpdate', 'RPort/Nonqueued_Receiver/ShortName']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(ShortNames)+1)), EnableUpdates, ShortNames]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
          
    def A118_0(self):
        HandleNeverReceiveds, ShortNames = [], []
        self.RPort = self.root.findall(".//{http://autosar.org/schema/r4.0}PORTS/{http://autosar.org/schema/r4.0}R-PORT-PROTOTYPE")

        for RPort in self.RPort:
            ShortName = RPort.find(".//{http://autosar.org/schema/r4.0}SHORT-NAME").text
            try:
                HandleNeverReceived = RPort.find(".//{http://autosar.org/schema/r4.0}HANDLE-NEVER-RECEIVED").text
                if HandleNeverReceived != 'false':
                    HandleNeverReceiveds.append(HandleNeverReceived)
                    ShortNames.append(ShortName)
            except:
                HandleNeverReceiveds.append('Undefined HANDLE-NEVER-RECEIVED')
                ShortNames.append(ShortName)
        try:
            sheet_ = self.wb.sheets.add('A118(0)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'RPort/Nonqueued_Receiver/HandleNeverReceived', 'RPort/Nonqueued_Receiver/ShortName']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(ShortNames)+1)), HandleNeverReceiveds, ShortNames]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
            
    def A119_0(self):
        HandleTimeOutTypes, ShortNames = [], []
        self.RPort = self.root.findall(".//{http://autosar.org/schema/r4.0}PORTS/{http://autosar.org/schema/r4.0}R-PORT-PROTOTYPE")

        for RPort in self.RPort:
            ShortName = RPort.find(".//{http://autosar.org/schema/r4.0}SHORT-NAME").text
            try:
                HandleTimeOutType = RPort.find(".//{http://autosar.org/schema/r4.0}HANDLE-TIMEOUT-TYPE").text
                if HandleTimeOutType != 'NONE':
                    HandleTimeOutTypes.append(HandleTimeOutType)
                    ShortNames.append(ShortName)
            except:
                HandleTimeOutTypes.append('Undefined HANDLE-TIMEOUT-TYPE')
                ShortNames.append(ShortName)
        try:
            sheet_ = self.wb.sheets.add('A119(0)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'RPort/Nonqueued_Receiver/HandleTimeOutType', 'RPort/Nonqueued_Receiver/ShortName']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(ShortNames)+1)), HandleTimeOutTypes, ShortNames]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
            
    def A120_0(self):
        UsesEndToEndProtections, ShortNames = [], []
        self.RPort = self.root.findall(".//{http://autosar.org/schema/r4.0}PORTS/{http://autosar.org/schema/r4.0}R-PORT-PROTOTYPE")

        for RPort in self.RPort:
            ShortName = RPort.find(".//{http://autosar.org/schema/r4.0}SHORT-NAME").text
            try:
                UsesEndToEndProtection = RPort.find(".//{http://autosar.org/schema/r4.0}USES-END-TO-END-PROTECTION").text
                if UsesEndToEndProtection != 'false':
                    UsesEndToEndProtections.append(UsesEndToEndProtection)
                    ShortNames.append(ShortName)
            except:
                UsesEndToEndProtections.append('Undefined USES-END-TO-END-PROTECTION')
                ShortNames.append(ShortName)
        try:
            sheet_ = self.wb.sheets.add('A120(0)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'RPort/Nonqueued_Receiver/UsesEndToEndProtection', 'RPort/Nonqueued_Receiver/ShortName']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(ShortNames)+1)), UsesEndToEndProtections, ShortNames]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
            
    def A122_0(self):
        UsesEndToEndProtections, ShortNames = [], []
        self.PPort = self.root.findall(".//{http://autosar.org/schema/r4.0}PORTS/{http://autosar.org/schema/r4.0}P-PORT-PROTOTYPE")
        for PPort in self.PPort:
            ShortName = PPort.find(".//{http://autosar.org/schema/r4.0}SHORT-NAME").text
            try:
                UsesEndToEndProtection = PPort.find(".//{http://autosar.org/schema/r4.0}USES-END-TO-END-PROTECTION").text
                if UsesEndToEndProtection != 'false':
                    UsesEndToEndProtections.append(UsesEndToEndProtection)
                    ShortNames.append(ShortName)
            except:
                UsesEndToEndProtections.append('Undefined USES-END-TO-END-PROTECTION')
                ShortNames.append(ShortName)
        try:
            sheet_ = self.wb.sheets.add('A122(0)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'PPort/Nonqueued_sender/UsesEndToEndProtection', 'PPort/Nonqueued_sender/ShortName']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(ShortNames)+1)), UsesEndToEndProtections, ShortNames]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
            
    def A123_0(self):
        HandleOutOfRanges, ShortNames = [], []
        self.PPort = self.root.findall(".//{http://autosar.org/schema/r4.0}PORTS/{http://autosar.org/schema/r4.0}P-PORT-PROTOTYPE")

        for PPort in self.PPort:
            ShortName = PPort.find(".//{http://autosar.org/schema/r4.0}SHORT-NAME").text
            try:
                HandleOutOfRange = PPort.find(".//{http://autosar.org/schema/r4.0}HANDLE-OUT-OF-RANGE").text
                if HandleOutOfRange != 'NONE':
                    HandleOutOfRanges.append(HandleOutOfRange)
                    ShortNames.append(ShortName)
            except:
                HandleOutOfRanges.append('Undefined HANDLE-OUT-OF-RANGE')
                ShortNames.append(ShortName)
        try:
            sheet_ = self.wb.sheets.add('A123(0)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'PPort/Nonqueued_sender/HandleOutOfRange', 'PPort/Nonqueued_sender/ShortName']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(ShortNames)+1)), HandleOutOfRanges, ShortNames]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
        
    def N010_3(self):
        IBShortNames, UUID = [], []
        self.IB = self.root.findall(".//{http://autosar.org/schema/r4.0}INTERNAL-BEHAVIORS//{http://autosar.org/schema/r4.0}SWC-INTERNAL-BEHAVIOR")
        for IB in self.IB:
            IBShortName = IB.find("{http://autosar.org/schema/r4.0}SHORT-NAME").text
            if not IBShortName.startswith('IB_'):
                IBShortNames.append(IBShortName)
                UUID.append(IB.get('UUID'))
        try:        
            sheet_ = self.wb.sheets.add('N010(3)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'IB/ShortName', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), IBShortNames, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
            
    def A125_0(self):
        DATATYPE_MAPPING_SETS, IB_DATATYPE_REFS, UUID = [], [], []
        self.IBDatatypeRefs = self.root.findall(".//{http://autosar.org/schema/r4.0}SWC-INTERNAL-BEHAVIOR//{http://autosar.org/schema/r4.0}DATA-TYPE-MAPPING-REF")
        self.DataTypeMappingSet = self.root.findall(".//{http://autosar.org/schema/r4.0}ELEMENTS//{http://autosar.org/schema/r4.0}DATA-TYPE-MAPPING-SET")
        for IBDatatypeRef in self.IBDatatypeRefs:
            if IBDatatypeRef.get('DEST') == 'DATA-TYPE-MAPPING-SET':
                IBDatatypeRef = (IBDatatypeRef.text).split('/')[-1]
                IB_DATATYPE_REFS.append(IBDatatypeRef)
        for i in self.DataTypeMappingSet:
            DataTypeMappingSet_SHORT_NAME = i.find('{http://autosar.org/schema/r4.0}SHORT-NAME').text
            if DataTypeMappingSet_SHORT_NAME not in IB_DATATYPE_REFS:
                DATATYPE_MAPPING_SETS.append(DataTypeMappingSet_SHORT_NAME)
                UUID.append(i.get('UUID'))
        try:           
            sheet_ = self.wb.sheets.add('A125(0)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'DataTypeMappingSet/SHORT_NAME', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), DATATYPE_MAPPING_SETS, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
            
    def A139_0(self):
        ERROR, ShortNames = [], []
        self.IB = self.root.findall(".//{http://autosar.org/schema/r4.0}INTERNAL-BEHAVIORS//{http://autosar.org/schema/r4.0}SWC-INTERNAL-BEHAVIOR")
        for IB in self.IB:
            HandleTerminationAndRestart = IB.find(".//{http://autosar.org/schema/r4.0}HANDLE-TERMINATION-AND-RESTART").text
            SupportsMultipleInstantiation = IB.find(".//{http://autosar.org/schema/r4.0}SUPPORTS-MULTIPLE-INSTANTIATION").text
            ShortName = IB.find(".//{http://autosar.org/schema/r4.0}SHORT-NAME").text
            if HandleTerminationAndRestart != 'NO-SUPPORT' and SupportsMultipleInstantiation != 'false':
                ERROR.append('2 errors')
                ShortNames.append(ShortName)
            elif HandleTerminationAndRestart != 'NO-SUPPORT' and SupportsMultipleInstantiation == 'false':
                ERROR.append('HandleTerminationAndRestart'+'='+HandleTerminationAndRestart)
                ShortNames.append(ShortName)
            elif HandleTerminationAndRestart == 'NO-SUPPORT' and SupportsMultipleInstantiation != 'false':
                ERROR.append('SupportsMultipleInstantiation'+'='+SupportsMultipleInstantiation)
                ShortNames.append(ShortName)
        try:
            sheet_ = self.wb.sheets.add('A139(0)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'ERROR', 'SWC_IB/ShortName']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(ShortNames)+1)), ERROR, ShortNames]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
            
    def N017_3(self):
        SWC_IMPL_ShortNames, SWC_IMPL_BehaviorRefs, UUID = [], [], []
        self.SWC_IMPL = self.root.findall(".//{http://autosar.org/schema/r4.0}ELEMENTS//{http://autosar.org/schema/r4.0}SWC-IMPLEMENTATION")
        for SWC_IMPL in self.SWC_IMPL:
            SWC_IMPL_ShortName = SWC_IMPL.find("{http://autosar.org/schema/r4.0}SHORT-NAME").text
            SWC_IMPL_BehaviorRef = (SWC_IMPL.find("{http://autosar.org/schema/r4.0}BEHAVIOR-REF").text).split('/')[-1]
            if not SWC_IMPL_ShortName[4:] == SWC_IMPL_BehaviorRef[3:]:
                SWC_IMPL_ShortNames.append(SWC_IMPL_ShortName)
                SWC_IMPL_BehaviorRefs.append(SWC_IMPL_BehaviorRef)
                UUID.append(SWC_IMPL.get('UUID'))
        try:        
            sheet_ = self.wb.sheets.add('N017(3)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'SWC_IMPL/ShortNames', 'SWC_IMPL/BehaviorRefs', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SWC_IMPL_ShortNames, SWC_IMPL_BehaviorRefs, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
            
    def N020_1(self):
        SWC_IMPL_Langueges, ShortNames = [], []
        self.SWC_IMPL = self.root.findall(".//{http://autosar.org/schema/r4.0}ELEMENTS//{http://autosar.org/schema/r4.0}SWC-IMPLEMENTATION")
        for SWC_IMPL in self.SWC_IMPL:
            SWC_IMPL_Languege = SWC_IMPL.find("{http://autosar.org/schema/r4.0}PROGRAMMING-LANGUAGE").text
            ShortName = SWC_IMPL.find("{http://autosar.org/schema/r4.0}SHORT-NAME").text
            if not SWC_IMPL_Languege == 'C':
                SWC_IMPL_Langueges.append(SWC_IMPL_Languege)
                ShortNames.append(ShortName)
        try:        
            sheet_ = self.wb.sheets.add('N020(1)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'SWC_IMPL/Langueges', 'SWC_IMPL/ShortName']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(ShortNames)+1)), SWC_IMPL_Langueges, ShortNames]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
            
    def A014_4(self):
        SERVER_CALL_POINTS, ShortNames, UUID = [], [], []
        self.RUNNABLE = self.root.findall(".//{http://autosar.org/schema/r4.0}RUNNABLES/{http://autosar.org/schema/r4.0}RUNNABLE-ENTITY")
        for RUNNABLE in self.RUNNABLE:
            SERVER_CALL_POINT = RUNNABLE.find(".//{http://autosar.org/schema/r4.0}SERVER-CALL-POINTS")
            ShortName = RUNNABLE.find("{http://autosar.org/schema/r4.0}SHORT-NAME").text
            try:
                for child in SERVER_CALL_POINT.getchildren():
                    chlid_name = child.tag.split('}')[-1]
                    if chlid_name != 'SYNCHRONOUS-SERVER-CALL-POINT':
                        SERVER_CALL_POINTS.append(chlid_name)
                        ShortNames.append(ShortName)
                        UUID.append(child.get('UUID'))
                        break
            except:
                SERVER_CALL_POINTS.append("Undefined SERVER-CALL-POINTS")
                ShortNames.append(ShortName)
                UUID.append(RUNNABLE.get('UUID'))
        try:        
            sheet_ = self.wb.sheets.add('A014(4)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'SERVER_CALL_POINT/CHILD', 'RUNNABLE-ENTITY/ShortNames','UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SERVER_CALL_POINTS, ShortNames, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
            
    def N013_4(self):
        SERVER_CALL_ShortNameS, SERVER_CALL_TargetOperationRefS, UUID = [], [], []
        self.SERVER_CALL_POINT = self.root.findall(".//{http://autosar.org/schema/r4.0}SERVER-CALL-POINTS//{http://autosar.org/schema/r4.0}SYNCHRONOUS-SERVER-CALL-POINT")
        for SERVER_CALL_POINT in self.SERVER_CALL_POINT:
            SERVER_CALL_ShortName = SERVER_CALL_POINT.find("{http://autosar.org/schema/r4.0}SHORT-NAME").text
            SERVER_CALL_TargetOperationRef = SERVER_CALL_POINT.find("{http://autosar.org/schema/r4.0}OPERATION-IREF//{http://autosar.org/schema/r4.0}TARGET-REQUIRED-OPERATION-REF").text
            SERVER_CALL_TargetOperationRef = SERVER_CALL_TargetOperationRef.split('/')[-1]
            if SERVER_CALL_ShortName != 'sc_'+SERVER_CALL_TargetOperationRef:
                SERVER_CALL_ShortNameS.append(SERVER_CALL_ShortName)
                SERVER_CALL_TargetOperationRefS.append(SERVER_CALL_TargetOperationRef)
                UUID.append(SERVER_CALL_POINT.get('UUID'))
        try:        
            sheet_ = self.wb.sheets.add('N013(4)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'SERVER_CALL/SYSC/ShortName', 'SERVER_CALL/SYSC/TargetOperationRef', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SERVER_CALL_ShortNameS, SERVER_CALL_TargetOperationRefS, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
            
    def N018_2(self):
        SERVER_CALL_ShortNameS, SERVER_CALL_TimeOutS, UUID = [], [], []
        self.SERVER_CALL_POINT = self.root.findall(".//{http://autosar.org/schema/r4.0}SERVER-CALL-POINTS//{http://autosar.org/schema/r4.0}SYNCHRONOUS-SERVER-CALL-POINT")
        for SERVER_CALL_POINT in self.SERVER_CALL_POINT:
            SERVER_CALL_ShortName = SERVER_CALL_POINT.find("{http://autosar.org/schema/r4.0}SHORT-NAME").text
            SERVER_CALL_TimeOut = SERVER_CALL_POINT.find("{http://autosar.org/schema/r4.0}TIMEOUT").text
            if SERVER_CALL_TimeOut != '0':
                SERVER_CALL_ShortNameS.append(SERVER_CALL_ShortName)
                SERVER_CALL_TimeOutS.append(SERVER_CALL_TimeOut)
                UUID.append(SERVER_CALL_POINT.get('UUID'))
        try:        
            sheet_ = self.wb.sheets.add('N018(2)', after=self.wb.sheets[-1]) 
            sheet_.range('A1').value = ['INDEX', 'SERVER_CALL/SYSC/ShortName', 'SERVER_CALL/SYSC/TimeOut', 'UUID']  
            sheet_.range('A1').expand(mode="right").font.bold = True
            sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SERVER_CALL_ShortNameS, SERVER_CALL_TimeOutS, UUID]
            sheet_.range('A1').expand(mode='table').columns.autofit()
        except Exception as err:
            print(err)
                
    def save(self):
        self.wb.save()
        weight_file = Path.cwd() / self.wb.name
        new_name =  Path.cwd() / 'HTPC_SR_Interface.xlsx'
        self.wb.close()
        try:
            weight_file.rename(new_name)
        except:
            ...
        finally:
            self.app.quit()
            
#%%主程序跳转
if __name__ == '__main__':
    arxml_parh = 'D:/Hua/My_Own_Utilities/arxml/arxml_py/HTPC_2.arxml'
    excel_path  = 'D:/Hua/My_Own_Utilities/arxml/arxml_py/HTPC_SR_Interface.xlsx'
    arxml_check = ArxmlCheck(arxml_parh, excel_path)
    arxml_check.get_info()  # 初始化读取arxml
    #%%下面是不同规则测试
    arxml_check.N003_2()    # SR/SHORT_NAMES = 'IF_'+SR/DATAELEMENT/SHORT_NAMES
    arxml_check.A006_1()    # SR/DATAELEMENT是否唯一
    arxml_check.N019_1()    # ['APDT_', 'AADT_', 'ARDT_']+SR/DATAELEMENT/SHORT_NAMES = SR/DATAELEMENT/TYPE_TREFS
    arxml_check.A041_1()    # SR/DATAELEMENT/SW_CALIBRATION_ACCESS是否配置为READ-ONLY
    arxml_check.A042_1()    # SR/DATAELEMENT/SW_IMPL_POLICY是否配置为STANDARD
    arxml_check.A111_1()    # CONSTANT_SPECIFICATION/SHORT_NAME不在SR/DATAELEMENT/SHORT_NAMES中；或者没有设置SW-VALUES-PHYS
    arxml_check.A131_1()    # SR_DATAELE_INIT_CONSTANT_REFS = 'IV_'+SR_DATAELE_SHORT_NAMES
    arxml_check.A067_1()    # SR不存在INVALIDATION-POLICY节点
    arxml_check.N007_1()    # 'R_'+RPortShortNames = RPortDataElementRef
    arxml_check.N008_1()    # 'P_'+PPortShortNames = PPortDataElementRef
    arxml_check.A030_2()    # Nonqueued_Sender/ConstantRefs = 'IV_'+Nonqueued_Sender_DataElementRefs
    arxml_check.A031_2()    # Nonqueued_Receiver/ConstantRefs = 'IV_'+Nonqueued_Receiver_DataElementRefs
    arxml_check.A032_3()    # Nonqueued_Receiver/AliveTimeOut设置为0
    arxml_check.A116_0()    # Nonqueued_Receiver/HandleOutOfRange设置为None
    arxml_check.A117_0()    # Nonqueued_Receiver/EnableUpdate设置为false
    arxml_check.A118_0()    # Nonqueued_Receiver/HandleNeverReceived设置为false
    arxml_check.A119_0()    # Nonqueued_Receiver/HandleTimeOutType设置为NONE
    arxml_check.A120_0()    # Nonqueued_Receiver/UsesEndToEndProtection设置为false
    arxml_check.A122_0()    # Nonqueued_Sender/UsesEndToEndProtection设置为false
    arxml_check.A123_0()    # Nonqueued_Receiver/HANDLE-OUT-OF-RANGE设置为NONE
    arxml_check.N010_3()    # INTERNAL-BEHAVIORS/SWC-INTERNAL-BEHAVIOR/SHORT-NAME是否'IB_'开头
    arxml_check.A125_0()    # DataTypeMappingSet/SHORT_NAME是否是引用IB/DATA-TYPE-MAPPING-REF
    arxml_check.A139_0()    # IB/SWC-IB/HANDLE-TERMINATION-AND-RESTART和 IB/SWC-IB/SUPPORTS-MULTIPLE-INSTANTIATION配置是否正确
    arxml_check.N017_3()    # 'IMP_'+SWC_IMPL/ShortNames = SWC_IMPL/BehaviorRefs/ShortNames
    arxml_check.N020_1()    # SWC_IMPL/PROGRAMMING-LANGUAGE是否为'C'
    arxml_check.A014_4()    # SERVER_CALL_POINTS是否只配置了‘SYNCHRONOUS-SERVER-CALL-POINT’
    arxml_check.N013_4()    # 'SERVER_CALL/SYSC/ShortName' = 'sc_'+'SERVER_CALL/SYSC/TargetOperationRef'
    arxml_check.N018_2()    # SERVER_CALL/SYSC/TIMEOUT是否配置为0
    #%% 测试运行后，执行保存文件
    arxml_check.save()      # 保存excel
