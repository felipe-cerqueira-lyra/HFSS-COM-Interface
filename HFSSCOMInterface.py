from win32com import client
import pythoncom
from threading import Lock, current_thread
import os
from shutil import copy, rmtree
import pandas as pd
import logging
from uuid import uuid4



class HFSS():
	FILEFORMAT		= {'tab':2, 'sNp':3, 'cit':4, 'm':7}
	COMPLEXFORMAT	= {'Mag/Pha':0, 'Re/Im':1, 'db/Pha':2}
	GEOMETRY		= {"Elevation":["Theta", "Phi", "0deg"], "Azimuth":["Phi", "Theta", "90deg"]}
	
	def __init__(self, projectAddr=None, designName=None, stream=None, inThread=False):
		if inThread:
			self.root = os.getcwd()
			return
		
		if stream:
			pythoncom.CoInitialize()
			stream.Seek(0,0)
			self.unmarshaledInterface = pythoncom.CoUnmarshalInterface(stream, pythoncom.IID_IDispatch)
			self.oApp = client.Dispatch(self.unmarshaledInterface)
			
		else:
			# self.oApp = client.Dispatch("AnsoftHfss.HfssScriptInterface")
			self.oApp = client.Dispatch("Ansoft.ElectronicsDesktop.2019.2")
			self.oApp_id = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, self.oApp)

		self.oDesktop	= self.oApp.GetAppDesktop()
		self.oDesktop.RestoreWindow()
				
		try:
			assert os.path.exists(projectAddr), "Project not found!"

			self.root   = projectAddr[::-1].split('/', maxsplit=1)[-1][::-1]
			projectName = projectAddr.split('/')[-1].split('.')[0]

			self.oDesktop.OpenProject(projectAddr)
			self.oProject = self.oDesktop.SetActiveProject(projectName)

			assert len(self.oProject.GetTopDesignList()) > 0, f"Project must have at least one design!"

			if not designName:
				designName = self.oProject.GetTopDesignList()[0]
			self.oDesign = self.oProject.SetActiveDesign(designName)
		
		except AssertionError as e:
			logging.error(e)
		
		except Exception as e:
			logging.error("Something went wrong! Closing application")
			logging.error(e)
			self.close()
	
	def set_parallel_mode(self, oApp, oDesktop, oProject, oDesign):
		self.oApp = oApp
		self.oDesktop = oDesktop
		self.oProject = oProject
		self.oDesign = oDesign

	def save(self):
		self.oProject.Save()
	
	def open_project(self, project_addr):
		assert os.path.exists(project_addr), "Project not found!"
		self.oDesktop.OpenProject(project_addr)            

	def close_project(self):
		for pjr in self.oDesktop.GetProjects():
			pjr.Save()
			pjr.Close()

		del self.oDesign
		del self.oProject

	def close(self):
		self.close_project()

		self.oDesktop.QuitApplication()
		del self.oDesktop
		del self.oApp
		
	def set_design_variable(self, varDic):
		change = ["NAME:AllTabs",[
			"NAME:LocalVariableTab",
			["NAME:PropServers","LocalVariables"],
			["NAME:ChangedProps"]
		]]
		for key,value in varDic.items(): change[1][2] += [[f'NAME:{key}', "Value:=", value]]
		self.oDesign.ChangeProperty(change)
	
	def set_project_variable(self, varDic):  
		change = ["NAME:AllTabs",[
			"NAME:ProjectVariableTab",
			["NAME:PropServers","ProjectVariables"],
			["NAME:ChangedProps"]
		]]
		for key,value in varDic.items(): change[1][2] += [[f'NAME:{key}', "Value:=", value]]
		self.oProject.ChangeProperty(change)

	def set_variable(self, varDIc):
		projVar = {}
		desVar  = {}
		for key,value in varDIc.items():
			if '$' in key:projVar[key]=value
			else:desVar[key]=value
		self.set_project_variable(projVar)
		self.set_design_variable(desVar)
		
	def edit_material(self, materialName, materialProps):
		'''
		materialProps := {"permittivity":value, 
							"permeability":value, 
							"conductivity": value, 
							"dielectric_loss_tangent":value}
		'''
		change = [f"NAME:{materialName}", "CoordinateSystemType:=", "Cartesian", "BulkOrSurfaceType:=", 1,
					["NAME:PhysicsTypes", "set:=", ["Electromagnetic"]]]
		
		props = [[f"{key}:=",value] for key,value in materialProps.items()]
		for prop in props:change += prop
		
		oDefinitionManager = self.oProject.GetDefinitionManager()
		oDefinitionManager.EditMaterial(materialName, change)

	def analyze(self, setupName):
		self.oDesign.AnalyzeDistributed(setupName)

	def clean_solutions(self):
		self.oDesign.DeleteFullVariation("All", True)
		
	def create_repport(self, repportName, ReportType, DisplayType, solutionName, contextArray, FamiliesArray, xData, yData):
		'''
		oModule.CreateReport(
			ReportName,
			ReportType,
			DisplayType,
			SolutionName,
			ContextArray,
			FamiliesArray,
			ReportDataArray,
			[])
		More information at page 351 of "ANSYS Electronics Desktop™:Scripting Guide"
		'''
		
		oModule = self.oDesign.GetModule("ReportSetup")        
		oModule.CreateReport(
			repportName,#ReportName
			ReportType,#ReportType
			DisplayType,#DisplayType
			solutionName,#SolutionName
			["Context:=", contextArray],#ContextArray
			FamiliesArray,#FamiliesArray
			["X Component:=", xData, "Y Component:=", [yData]],#ReportDataArray
			[]
		)

	def delete_repport(self, repportName):
		oModule = self.oDesign.GetModule("ReportSetup")
		oModule.DeleteReports([repportName])
	  
	def create_near_field_repport(self, repportName, DisplayType, solutionName, context, xData, yData, freq):
		oModule = self.oDesign.GetModule("ReportSetup")
		if xData == "Theta":
			familiesArray = ["Theta:=", ["All"],"Phi:=", ["All"],"Freq:=", [freq]]
		elif xData == "Phi":
			familiesArray = ["Phi:=", ["All"],"Theta:=", ["All"],"Freq:=", [freq]]
		
		self.create_repport(repportName, "Near Fields", DisplayType, solutionName, context, familiesArray, xData, yData)
	
	def create_far_field_repport(self, repportName, DisplayType, solutionName, context, yData, freq):
		oModule = self.oDesign.GetModule("ReportSetup")
		primary, secondary, secondaryAng = self.GEOMETRY[context]
		familiesArray = [f"{primary}:=", ["All"],
						 ["NAME:VariableValues",
						  "Freq:=", freq,
						  f"{secondary}:=", secondaryAng]]
		self.create_repport(repportName, "Far Fields", DisplayType, solutionName, context, familiesArray, primary, yData)
		
	def create_antenna_parameter_repport(self, repportName, DisplayType, solutionName, context, yData):
		oModule = self.oDesign.GetModule("ReportSetup")
		familiesArray = ["Freq:=", ["All"]]
		
		self.create_repport(repportName, "Antenna Parameters", DisplayType, solutionName, context, familiesArray, "Freq", yData)
	
	def export_network_data(self, fileAddr, solutionName, dataType, complexFormat):
		fileFormat = fileAddr.split('.')[-1]
		oModule = self.oDesign.GetModule("Solutions")
		oModule.ExportNetworkData("", [solutionName], self.FILEFORMAT[fileFormat],
									   fileAddr, ["All"], False, 50, dataType, -1, self.COMPLEXFORMAT[complexFormat])
	
	def export_near_field_data(self, fileAddr, solutionName, dataType, xSweep, context, freq):
		self.create_near_field_repport("tempRepport", "Data Table", solutionName, context, xSweep, dataType, freq)
		self.export_report_data("tempRepport", fileAddr)
		self.delete_repport("tempRepport")
	
	def export_far_field_data(self, fileAddr, solutionName, dataType, context, freq):
		self.create_far_field_repport("tempRepport", "Data Table", solutionName, context, dataType, freq)
		self.export_report_data("tempRepport", fileAddr)
		self.delete_repport("tempRepport")
		
	def export_antenna_parameter_data(self, fileAddr, solutionName, dataType, context):
		self.create_antenna_parameter_repport("tempRepport", "Data Table", solutionName, context, dataType)
		self.export_report_data("tempRepport", fileAddr)
		self.delete_repport("tempRepport")
	
	def export_report_data(self, reportName, fileAddr):
		oModule = self.oDesign.GetModule("ReportSetup")
		oModule.ExportToFile(reportName, fileAddr)
		
	def get_network_data(self, dataType, solutionName, complexFormat):
		addr = f"{self.root}/{str(uuid4())}.tab"
		self.export_network_data(addr, solutionName, dataType, complexFormat)
		
		ND = pd.read_csv(addr,skiprows=1, delim_whitespace=True, skipinitialspace=True, index_col=0)
		ND = ND.rename(columns={col:col.rstrip() for col in ND.columns})

		oModule = self.oDesign.GetModule("BoundarySetup")
		nports = oModule.GetNumExcitations()

		if complexFormat == "Re/Im":
			nd_re = ND.iloc[:,[x for x in range (2*nports**2) if x%2==0]]
			nd_im = ND.iloc[:,[x for x in range (2*nports**2) if x%2==1]]
			nd_re = nd_re.rename(columns={col:col.split('_')[0] for col in nd_re.columns})
			nd_im = nd_im.rename(columns={col:col.split('_')[0] for col in nd_im.columns})
			nd = nd_re + 1j*nd_im
		else:
			nd = ND.iloc[:,[x for x in range (2*nports**2)]]
		
		os.remove(addr)
		
		return nd
   
	def get_near_field_data(self, dataType, xSweep, solutionName, context, freq):
		addr = f"{self.root}/{str(uuid4())}.tab"
		
		self.export_near_field_data(addr, solutionName, dataType, xSweep, context, freq)
		NFD = self.read_report_data_from_file(addr)
		os.remove(addr)
		
		return NFD

	def get_far_field_data(self, dataType, solutionName, context, freq):
		addr = f"{self.root}/{str(uuid4())}.tab"
		
		self.export_far_field_data(addr, solutionName, dataType, context, freq) 
		FFD = self.read_report_data_from_file(addr)
		os.remove(addr)
		
		return FFD
	
	def get_antenna_parameter_data(self, dataType, solutionName, context):
		addr = f"{self.root}/{str(uuid4())}.tab"
		
		self.export_antenna_parameter_data(addr, solutionName, dataType, context)
		APD = self.read_report_data_from_file(addr)
		os.remove(addr)
		
		return APD
	
	def read_report_data_from_file(self, addr):
		df = pd.read_csv(addr,skiprows=0, delim_whitespace=True, skipinitialspace=True, index_col=0)
		df = df.rename(columns={col:col.rstrip() for col in df.columns})
		return df
		

class ParallelInterface():
	def __init__(self):
		pythoncom.CoInitialize()
		oApp = client.DispatchEx("Ansoft.ElectronicsDesktop.2019.2")
		self.stream = pythoncom.CreateStreamOnHGlobal()
		pythoncom.CoMarshalInterface(self.stream, 
									 pythoncom.IID_IDispatch, 
									 oApp._oleobj_, 
									 pythoncom.MSHCTX_LOCAL, 
									 pythoncom.MSHLFLAGS_TABLESTRONG)
		del oApp
		self.lock = Lock()

	def open_project(self, projectAddr, nThreads):
		self.projectsAddr = [projectAddr.replace('.aedt', f'T{i}.aedt') for i in range(nThreads)]
		for prjAddrThreads in self.projectsAddr:
			copy(projectAddr, prjAddrThreads)
		
		self.hfss = HFSS(self.projectsAddr[0], stream=self.stream)
		for prjAddrThreads in self.projectsAddr[1:]:
			self.hfss.open_project(prjAddrThreads)

	def close(self):
		self.hfss.close()
		self.stream.Seek(0,0)
		pythoncom.CoReleaseMarshalData(self.stream)
		del self.hfss
		del self.stream
		pythoncom.CoUninitialize()
		for prjAddrThreads in self.projectsAddr:
			os.remove(prjAddrThreads) 
			rmtree(prjAddrThreads.replace('.aedt', '.aedtresults'))

def set_HFSS_parallel(stream, lock, idx):
	lock.acquire()
	pythoncom.CoInitialize()
	stream.Seek(0,0)
	unmarshaledInterface = pythoncom.CoUnmarshalInterface(stream, pythoncom.IID_IDispatch) 
	oApp = client.Dispatch(unmarshaledInterface)
	lock.release()
	oDesktop = oApp.GetAppDesktop()
	oProject = oDesktop.GetProjects()[idx]
	designName = oProject.GetTopDesignList()[0]
	oDesign = oProject.GetDesign(designName)

	hfss = HFSS(inThread=True)
	hfss.set_parallel_mode(oApp, oDesktop, oProject, oDesign)

	return hfss

def run_in_parallel(pI):
	def decorator(function):
		def wrapper(*args, **kwargs):
			idx = current_thread().name.split('_')[-1]
			hfss = set_HFSS_parallel(pI.stream, pI.lock, idx)
			result = function(*args, **kwargs, hfss=hfss)
			del hfss
			return result
		return wrapper
	return decorator