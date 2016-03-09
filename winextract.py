import optparse
import os
import subprocess
import shutil
import xml.etree.ElementTree

class Environment:		  
	
	# Initialize class variables
	def __init__(self,msUpdate,outputDir,archPatch):
		self.msUpdate  = msUpdate
		self.outputDir = outputDir
		self.archPatch = archPatch
		self.outFolder = ""
		self.cabPath = ""
		self.patchPath = ""
		self.tmp = "Tmp"
		self.patch = "Patch"
		self.other = "Other"
		self.logs = "Logs"
	
	#----------------------------------------------------------------------------
	# Name: PrepareEnv 															-
	# Description: 																-
	# Function processes the .msu file and extracts the contents to a tmp  		-
	# directory for additional processing										-
	#---------------------------------------------------------------------------- 
	def PrepareEnv(self):
		tmp = "Tmp"
		
		# Set current working directory
		currDir = os.getcwd()

		# Valid Responses
		validResponse = ["Y","N"]

			
		# Set current path of msu file
		msUpdatePath = os.path.join(currDir,self.msUpdate)
		
		# Check if msu file exists
		if os.path.exists(msUpdatePath):
			print "File: " + msUpdatePath	
			
			if not os.path.isabs(self.outputDir):
				# Check supplied output folder
				self.outFolder = os.path.join(currDir,self.outputDir)
			else:
				self.outFolder = outputDir
			
			# If file does not exist give user option to create
			if not os.path.exists(self.outFolder):			
				
				# Loop until valid reponse received
				while True:
					try:
						fResp = raw_input( "The " + self.outFolder + " folder does not exist. Create [Y|N]? \n")
						
					except ValueError:
						print "Exception occurred"
						continue
					print ""	
					if fResp not in validResponse:
						print "Invalid response. Choose [Y|N] \n"
						continue
					else:
						break
				
				if fResp == "Y":
					# Create folder
					os.makedirs(os.path.join(currDir,self.outFolder))
				else:
					print "Folder will be created in current working directory"
					os.makedirs(self.outFolder)
				
				for subDir in ["Reources","Patch","Other","Tmp","Logs"]:
					os.makedirs(os.path.join(self.outFolder,self.archPatch,subDir))
					
				# Extract .msu File with expand.exe. Assumes expand.exe in C:\Windows\System32
				cmd = ['expand',"-F:*",msUpdatePath,os.path.join(self.outFolder,self.archPatch,self.tmp)]
				subprocess.Popen(cmd,shell=True).communicate()
				
				self.cabPath = os.path.join(self.outFolder,self.archPatch,self.tmp)
			else:
				print "Folder Already exists. Exiting.........."
				exit(0)
		
	#----------------------------------------------------------------------------
	# Name: ExtractCab 															-
	# Description: 																-
	# The PrepareEnv method does the initial unpacking of the .msu package		-
	# and writes the contents to a Temporary folder in the directory			-
	# specified in the arguments. This functions extract the cab file			- 
	# to the Patch directory.													- 
	#---------------------------------------------------------------------------- 
	def ExtractCab(self):
		strWusscan = "WSUSSCAN"
		self.patchPath = os.path.join(self.outFolder,self.archPatch,"Patch")
		
		for cabFile in os.listdir(os.path.join(self.outFolder,self.archPatch,self.tmp)):
			if cabFile.endswith(".cab") and (strWusscan not in cabFile):
				print cabFile
				self.cabPath = os.path.join(self.outFolder,self.archPatch,"Tmp",cabFile)
			
		# Extract .msu File with expand.exe. Assumes expand.exe in C:\Windows\System32
		cmd = ['expand',"-F:*",self.cabPath,os.path.join(self.outFolder,self.cabPath,self.patchPath)]
		subprocess.Popen(cmd,shell=True).communicate()
	
	#----------------------------------------------------------------------------
	# Name: ProcessExtract														-
	# Description: 																-
	# The ProcessExtract method copies those files we are not interested in to 	-
	# the Other directory 													 	-
	#----------------------------------------------------------------------------
	def ProcessExtract(self):
		misc = [".mum",".manifest",".cat"]
		for miscFile in os.listdir(os.path.join(self.outFolder,self.archPatch,self.patch)):
			
			if miscFile.endswith((".mum",".manifest",".cat")):
				# We don't care about these files. so move to Other
				fullPath = os.path.join(self.outFolder,self.archPatch,self.patch,miscFile)
				shutil.copy(fullPath,os.path.join(self.outFolder,self.archPatch,self.other));
				
				# Remove the copied files
				os.remove(fullPath);
	#----------------------------------------------------------------------------
	# Name: GenerateReport														-
	# Description: 																-
	# The GenerateReport method takes each directory in the Patch directory 	-
	# and lists the associated dll's etc. These are the files that will be the  -
	# of the diffing exercise													-
	#----------------------------------------------------------------------------
	def GenerateReport(self):
		report = "report.csv"
		delim = "|"
		dll = ".dll"
		
		with open(os.path.join(self.outFolder,self.archPatch,self.logs,report), 'w') as rpt:
			for dir in os.listdir(os.path.join(self.outFolder,self.archPatch,self.patchPath)):
				str = dir
				
				for dll in os.listdir(os.path.join(self.outFolder,self.archPatch,self.patchPath,dir)):
						if dll.endswith((".dll",".exe")):
							str += delim + dll
				rpt.write(str + "\n")
				str = ""		
def main():
	
	# Allowed extensions
	extensions = ["x86","x64","WOW64"]
	
	# Set parser options
	parser = optparse.OptionParser('usage %prog -m <.msu patch file> -f <output folder> -a <architecture>' )
	parser.add_option('-m', dest='msUpdate',type='string', help='.msu Update File')
	parser.add_option('-f', dest='outputDir',type='string', help='Output Folder e.g. MSXX-XXX')
	parser.add_option('-a', dest='archPatch',type='string', default='x86', help='Architecture of patch <x86 Defualt>')
		
	# Parse the arguments
	(options,args) = parser.parse_args()
	msUpdate = options.msUpdate
	outputDir = options.outputDir
	pArch = options.archPatch
	
	# Ensure valid argument supplied and check extension
	if(msUpdate is not None) and (msUpdate.endswith(".msu")):
		
		# Initialize class
		env = Environment(msUpdate,outputDir,pArch)
		
		# Extract .msu
		env.PrepareEnv()
		
		# Extract .cab
		env.ExtractCab()
		
		# Clean up junk files
		env.ProcessExtract()
		
		# Generate log
		env.GenerateReport()
	else:
		print parser.usage
		exit(0)

if __name__ == "__main__":
	main()