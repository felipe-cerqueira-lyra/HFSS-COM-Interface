from win32com import client
import os
import pandas as pd
import logging

class HFSS():
	FILEFORMAT		= {'tab':2, 'sNp':3, 'cit':4, 'm':7}
	COMPLEXFORMAT	= {'Mag/Pha':0, 'Re/Im':1, 'db/Pha':2}
	GEOMETRY		= {"Elevation":["Theta", "Phi", "0deg"], "Azimuth":["Phi", "Theta", "90deg"]}
	
	def __init__(self, projectAddr, designName=None):
		self.root   = projectAddr[::-1].split('\\', maxsplit=1)[-1][::-1]
		projectName = projectAddr.split('\\')[-1].split('.')[0]
		
		self.oApp		= client.Dispatch("AnsoftHfss.HfssScriptInterface")
		#self.oApp		= client.Dispatch("Ansoft.ElectronicsDesktop.2019.2")
		self.oDesktop	= self.oApp.GetAppDesktop()
		self.oDesktop.RestoreWindow()
				
		try:
			assert os.path.exists(projectAddr), "Project not found!"

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
			#self.close()
		
	def save(self):
		self.oProject.Save()
	
	def open_project(self, project_addr):
		assert os.path.exists(project_addr), "Project not found!"
		self.oDesktop.OpenProject(project_addr)            
	
	def close(self):
		
		for prj in self.oDesktop.GetProjects():prj.Save()
		for prj in self.oDesktop.GetProjects():prj.Close()
		prj = None
			
		self.oDesign = None
		self.oProject = None
		
		self.oDesktop.QuitApplication()
		del self.oDesktop
		
	def set_design_variable(self, varName, varValue):
		change = ["NAME:AllTabs",[
					"NAME:LocalVariableTab",
					["NAME:PropServers","LocalVariables"],
					["NAME:ChangedProps",[f"NAME:{varName}","Value:=",f"{varValue}"]]
				]]
		self.oDesign.ChangeProperty(change)
	
	def set_project_variable(self, varName, varValue):  
		change = ["NAME:AllTabs",[
					"NAME:ProjectVariableTab",
					["NAME:PropServers","ProjectVariables"],
					["NAME:ChangedProps",[f"NAME:{varName}","Value:=",f"{varValue}"]]
				]]
		self.oProject.ChangeProperty(change)
		
	def edit_material(self, materialName, materialProps):
		'''
		materialProps := {"permittivity":value, 
							"permeability":value, 
							"conductivity": value, 
							"dielectric_loss_tangent":value}
		'''
		change = [f"NAME:{materialName}", "CoordinateSystemType:=", "Cartesian", "BulkOrSurfaceType", 1,
					["NAME:PhysicsTypes", "set:=", ["Electromagnetic"]]]
		
		props = [[f"{key}:=",value] for key,value in materialProps.items()]
		for prop in props:change += prop
		
		oDefinitionManager = self.oProject.GetDefinitionManager()
		oDefinitionManager.EditMaterial(materialName, change)

	def analyze(self, setup_name):
		self.oDesign.Analyze(setup_name)
		
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
		oModule.DeleteReports(["tempRepport"])
	
	def export_far_field_data(self, fileAddr, solutionName, dataType, context, freq):
		self.create_far_field_repport("tempRepport", "Data Table", solutionName, context, dataType, freq)
		self.export_report_data("tempRepport", fileAddr)
		oModule.DeleteReports(["tempRepport"])
		
	def export_antenna_parameter_data(self, fileAddr, solutionName, dataType, context):
		self.create_antenna_parameter_repport("tempRepport", "Data Table", solutionName, context, dataType)
		self.export_report_data("tempRepport", fileAddr)
		oModule.DeleteReports(["tempRepport"])
	
	def export_report_data(self, reportName, fileAddr):
		oModule = self.oDesign.GetModule("ReportSetup")
		oModule.ExportToFile(reportName, fileAddr)
		
	def get_network_data(self, dataType, solutionName, complexFormat):
		addr = f"{self.root}/Solution.tab" 
		self.export_network_data(addr, solutionName, dataType, complexFormat)
		
		ND = pd.read_csv(addr,skiprows=1, sep='\t', skipinitialspace=True, index_col=0)
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
		addr = f"{self.root}/Solution.tab"
		
		self.export_near_field_data(addr, solutionName, dataType, xSweep, context, freq)
		NFD = read_report_data_from_file(addr)
		os.remove(addr)
		
		return NFD

	def get_far_field_data(self, dataType, solutionName, context, freq):
		addr = f"{self.root}/Solution.tab"
		
		self.export_far_field_data(addr, solutionName, dataType, context, freq) 
		FFD = read_report_data_from_file(addr)
		os.remove(addr)
		
		return FFD
	
	def get_antenna_parameter_data(self, dataType, solutionName, context):
		addr = f"{self.root}/Solution.tab"
		
		self.export_antenna_parameter_data(addr, solutionName, dataType, context)
		APD = read_report_data_from_file(addr)
		os.remove(addr)
		
		return APD
	
	def read_report_data_from_file(self, addr):
		df = pd.read_csv(addr,skiprows=0, sep='\t', skipinitialspace=True, index_col=0)
		df = df.rename(columns={col:col.rstrip() for col in df.columns})
		return df
		