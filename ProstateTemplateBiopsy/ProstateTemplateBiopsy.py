##=========================================================================

#  Program:   Prostate Template Biopsy - Slicer Module
#  Language:  Python

#  Copyright (c) Brigham and Women's Hospital. All rights reserved.

#  This software is distributed WITHOUT ANY WARRANTY; without even
#  the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR
#  PURPOSE.  See the above copyright notices for more information.

#=========================================================================


import os
# from matplotlib.pyplot import get
import vtk, qt, ctk, slicer, ast
from slicer.ScriptedLoadableModule import *
import time
import glob
import datetime
import os
from SlicerDevelopmentToolboxUtils.helpers import SmartDICOMReceiver
from SlicerDevelopmentToolboxUtils.events import SlicerDevelopmentToolboxEvents
from SlicerDevelopmentToolboxUtils.widgets import IncomingDataWindow, CustomStatusProgressbar
from SlicerDevelopmentToolboxUtils.constants import DICOMTAGS, STYLE
from SlicerDevelopmentToolboxUtils.exceptions import DICOMValueError, UnknownSeriesError
from SlicerDevelopmentToolboxUtils.module.session import StepBasedSession
from DICOMLib import DICOMUtils

class ProstateTemplateBiopsy(ScriptedLoadableModule):
  """Uses ScriptedLoadableModule base class, available at:
  https://github.com/Slicer/Slicer/blob/master/Bakse/Python/slicer/ScriptedLoadableModule.py
  """

  def __init__(self, parent):
    ScriptedLoadableModule.__init__(self, parent)
    self.parent.title = "Prostate Template Biopsy"
    self.parent.categories = ["IGT"]
    self.parent.dependencies = []
    self.parent.contributors = ["Franklin King"]
    self.parent.helpText = """
"""
    self.parent.helpText += self.getDefaultModuleDocumentationLink()
    self.parent.acknowledgementText = """
"""
    # Set module icon from Resources/Icons/<ModuleName>.png
    moduleDir = os.path.dirname(self.parent.path)
    for iconExtension in ['.svg', '.png']:
      iconPath = os.path.join(moduleDir, 'Resources/Icons', self.__class__.__name__ + iconExtension)
      if os.path.isfile(iconPath):
        parent.icon = qt.QIcon(iconPath)
        break

class ProstateTemplateBiopsyWidget(ScriptedLoadableModuleWidget):
  """Uses ScriptedLoadableModuleWidget base class, available at:
  https://github.com/Slicer/Slicer/blob/master/Base/Python/slicer/ScriptedLoadableModule.py
  """

  def __init__(self, parent=None):
    ScriptedLoadableModuleWidget.__init__(self, parent)

  def setup(self):
    ScriptedLoadableModuleWidget.setup(self)

    # ------------------------------------ Initialization UI ---------------------------------------
    # TODO: 
    # - Populate Image List when images are loaded
    initializeCollapsibleButton = ctk.ctkCollapsibleButton()
    initializeCollapsibleButton.text = "Connection"
    self.layout.addWidget(initializeCollapsibleButton)
    initializeLayout = qt.QFormLayout(initializeCollapsibleButton)

    self.caseDirPath = None
    self.caseDICOMPath = None

    self.casesPathBox = qt.QLineEdit("C:/w/data/ProstateBiopsyModuleTest/Cases")
    self.casesPathBox.setReadOnly(True)
    self.casesPathBrowseButton = qt.QPushButton("...")
    self.casesPathBrowseButton.clicked.connect(self.select_directory)
    pathBoxLayout = qt.QHBoxLayout()
    pathBoxLayout.addWidget(self.casesPathBox)
    pathBoxLayout.addWidget(self.casesPathBrowseButton)
    initializeLayout.addRow(pathBoxLayout)

    self.initializeButton = qt.QPushButton("Initialize Case")
    self.initializeButton.toolTip = "Start DICOM Listener and Create Folders"
    self.initializeButton.enabled = True
    self.initializeButton.connect('clicked()', self.initializeCase)
    initializeLayout.addWidget(self.initializeButton)

    self.caseDirLabel = qt.QLabel()
    self.caseDirLabel.text = "Waiting to Initialize Case"
    initializeLayout.addRow("Case Directory: ", self.caseDirLabel)

    self.closeCaseButton = qt.QPushButton("Close Case")
    self.closeCaseButton.toolTip = "Close Case"
    self.closeCaseButton.enabled = False
    self.closeCaseButton.connect('clicked()', self.closeCase)
    initializeLayout.addWidget(self.closeCaseButton)

    self.loadedFiles = []
    self.filesToBeLoaded = []
    self.observationTimer = qt.QTimer()
    self.observationTimer.setInterval(500)
    self.observationTimer.timeout.connect(self.observeDicomFolder)

    self.seriesList = []
    self.seriesTimeStamps = dict()
    slicer.mrmlScene.AddObserver(slicer.mrmlScene.NodeAddedEvent,self.onNodeAddedEvent)
    # -------------------------------------- ----------  --------------------------------------

    # ------------------------------------ Image List UI --------------------------------------
    # TODO: 
    # - Switch images depending on clicking an image (can have a button for it on each row)
    # - Ignore certain images
    imageListCollapsibleButton = ctk.ctkCollapsibleButton()
    imageListCollapsibleButton.text = "Images"
    self.layout.addWidget(imageListCollapsibleButton)
    imageListLayout = qt.QVBoxLayout(imageListCollapsibleButton)

    self.imageRoles = ['N/A', 'CALIBRATION', 'PLANNING', 'CONFIRMATION']

    # Image List Table with combo boxes to set role for images
    self.imageListTableWidget = qt.QTableWidget(0, 3)
    self.imageListTableWidget.setHorizontalHeaderLabels(["Image Description", "Acquisition Time", "Role"])
    imageListLayout.addWidget(self.imageListTableWidget)
    # -------------------------------------- ----------  --------------------------------------

    # ---------------------------------- Registration UI --------------------------------------
    registrationCollapsibleButton = ctk.ctkCollapsibleButton()
    registrationCollapsibleButton.text = "Registration"
    self.layout.addWidget(registrationCollapsibleButton)
    registrationLayout = qt.QFormLayout(registrationCollapsibleButton)


    # ------------------------------------- ----------  --------------------------------------
    
    # ------------------------------------- Planning UI --------------------------------------
    planningCollapsibleButton = ctk.ctkCollapsibleButton()
    planningCollapsibleButton.text = "Images"
    self.layout.addWidget(planningCollapsibleButton)
    planningLayout = qt.QFormLayout(planningCollapsibleButton)


    # -------------------------------------- ----------  --------------------------------------
    # Add vertical spacer
    self.layout.addStretch(1)
    # -------------------------------------- ----------  --------------------------------------
  
  def select_directory(self):
    directory = qt.QFileDialog.getExistingDirectory(self.parent, "Select Cases Directory")
    if directory:
      self.casesPathBox.setText(directory)

  # Phase change enables later steps, but generally does not disable older steps unless those steps would cause problems
  def onPhaseChange(self, phase):
    if phase == "START":
      self.casesPathBrowseButton.setEnabled(True)
      self.initializeButton.setEnabled(True)
      self.closeCaseButton.setEnabled(False)
    elif phase == "REGISTRATION": 
      self.casesPathBrowseButton.setEnabled(False)
      self.initializeButton.setEnabled(False)
      self.closeCaseButton.setEnabled(True)

  def initializeCase(self):
    if not self.casesPathBox.text:
      return
    
    # Create case folder
    path = self.casesPathBox.text
    currentDate = datetime.date.today().strftime("%Y-%m-%d")
    dirNames = [d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d))]
    index = 1
    for dirName in dirNames:
      before, sep, after = dirName.rpartition("_")
      if before == currentDate:
        index += 1
    self.caseDirPath = f'{path}/{currentDate}_{index}'
    os.mkdir(self.caseDirPath)
    self.caseDirLabel.text = self.caseDirPath

    slicer.util.selectModule('DICOM')
    slicer.util.selectModule('ProstateTemplateBiopsy')

    # Set DICOM Database
    slicer.modules.DICOMWidget.updateDatabaseDirectoryFromWidget(self.caseDirPath)

    # Start Listener and Observation Timer
    slicer.modules.DICOMWidget.onToggleListener(True)
    self.observationTimer.start()

    self.onPhaseChange("REGISTRATION")

  def observeDicomFolder(self):
    currentFileList = self.getFileList(f'{self.caseDirPath}/dicom')
    # Files still being added
    if (len(self.loadedFiles) + len(self.filesToBeLoaded)) < len(currentFileList):
      for file in currentFileList:
        if (file not in self.loadedFiles) and (file not in self.filesToBeLoaded):
          self.filesToBeLoaded.append(file)
      # TODO: Mark new files loading as Happening
    # Files no longer being added
    # (len(self.loadedFiles) + len(self.filesToBeLoaded)) >= len(currentFileList)
    else:
      if len(self.filesToBeLoaded) > 0:
        #print(self.filesToBeLoaded)
        self.loadSeries(self.filesToBeLoaded)
        self.loadedFiles += self.filesToBeLoaded
        self.filesToBeLoaded = []
      # TODO: Mark new files loading as Not Happening
    
  def getFileList(self, directory):
    filenames = []
    if os.path.exists(directory):
      for filename in glob.iglob(f'{directory}/**/*.dcm', recursive=True):
        filenames.append(filename.replace('\\', '/'))
    return filenames

  def closeCase(self):
    self.seriesList = []
    self.loadedFiles = []
    self.filesToBeLoaded = []
    self.observationTimer.stop()
    
  def loadSeries(self, newFilesAdded):
    seriesUIDs = []
    for file in newFilesAdded:
      seriesUIDs.append(StepBasedSession.getDICOMValue(file, '0020,000E'))
    loadedNodeIDs = DICOMUtils.loadSeriesByUID(seriesUIDs)
  
  @vtk.calldata_type(vtk.VTK_OBJECT)
  def onNodeAddedEvent(self, caller, event, calldata):
    newNode = calldata
    if not isinstance(newNode, slicer.vtkMRMLScalarVolumeNode):
      return
    if (newNode.GetName() == "CalibrationOutVolume"):
      return
    self.addToImageList(newNode)

  # TODO: Automatically update roles in another function
  # TODO: Set fixed sizes for table/rows
  def addToImageList(self, newNode):
    rowCount = self.imageListTableWidget.rowCount

    imageName = qt.QTableWidgetItem(newNode.GetName())
    imageTime = qt.QTableWidgetItem(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    imageRoleChoice = qt.QComboBox()
    imageRoleChoice.addItems(self.imageRoles) # add some options to the combo box
    imageRoleChoice.currentTextChanged.connect(lambda: self.updateImageListRoles(imageRoleChoice.currentText, rowCount))

    self.imageListTableWidget.insertRow(rowCount)
    self.imageListTableWidget.setItem(rowCount, 0, imageName)
    self.imageListTableWidget.setItem(rowCount, 1, imageTime)
    self.imageListTableWidget.setCellWidget(rowCount, 2, imageRoleChoice) # set the combo box as a cell widget

    self.autoImageRoleAssignment(newNode, imageRoleChoice, rowCount)

  def updateImageListRoles(self, newRoleAssigned, rowCount):
    if newRoleAssigned in ["CALIBRATION", "PLANNING"]:
      for index in range(0, self.imageListTableWidget.rowCount):
        if (index != rowCount) and (newRoleAssigned == (self.imageListTableWidget.cellWidget(index, 2).currentText)):
          self.imageListTableWidget.cellWidget(index, 2).setCurrentIndex(self.imageRoles.index("N/A"))

  def autoImageRoleAssignment(self, newNode, imageRoleChoice, rowCount):
    name = newNode.GetName()
    if "Template" in name:
      imageRoleChoice.setCurrentIndex(self.imageRoles.index("CALIBRATION"))
      self.updateImageListRoles(imageRoleChoice.currentText, rowCount)
    elif "cover" in name:
      imageRoleChoice.setCurrentIndex(self.imageRoles.index("PLANNING"))
      self.updateImageListRoles(imageRoleChoice.currentText, rowCount)