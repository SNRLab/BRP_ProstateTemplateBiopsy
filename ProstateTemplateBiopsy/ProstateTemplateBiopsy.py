##=========================================================================

#  Program:   Prostate Template Biopsy - Slicer Module
#  Language:  Python

#  Copyright (c) Brigham and Women's Hospital. All rights reserved.

#  This software is distributed WITHOUT ANY WARRANTY; without even
#  the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR
#  PURPOSE.  See the above copyright notices for more information.

#=========================================================================
# pip_install('scikit-image')
# pip_install('PyPDF2')
# pip_install('reportlab')
# pip_install('io')

import os
import vtk, qt, ctk, slicer, ast
from slicer.ScriptedLoadableModule import *
import time
import glob
import datetime
import os
import re
import math
import ast
import numpy as np
from skimage.draw import line_nd

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
    self.parent.contributors = ["Franklin King (SNR)"]
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
    self.ignoredVolumeNames = ['MaskedCalibrationVolume', 'MaskedCalibrationLabelMapVolume', 'TempLabelMapVolume']
    self.currentPhase = 'START'
    self.imageRoles = ['N/A', 'CALIBRATION', 'PLANNING', 'CONFIRMATION']
    self.caseDirPath = None
    self.caseDICOMPath = None
    self.zFrameModelNode = None
    self.templateModelNode = None
    self.calibratorModelNode = None
    self.guideHolesModelNode = None
    self.guideHoleLabelsModelNode = None
    self.templateOrigin = None
    self.templateHorizontalOffset = 5
    self.templateVerticalOffset = 5
    self.templateHorizontalLabels = []
    self.templateVerticalLabels = []
    self.templateWorksheetPath = ''
    self.templateWorksheetOverlayPath = ''
    self.worksheetOrigin = [238,404]
    self.worksheetHorizontalOffset = 26
    self.worksheetVerticalOffset = 26
    self.removeNodeByName('ZFrameTransform')
    # TODO: Cleanup function to allow for a Reload to truly restart a case?
    self.ZFrameCalibrationTransformNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLLinearTransformNode", "ZFrameTransform")
    self.increaseThresholdForRepair = False
    self.increaseThresholdForRetry = False
    self.validRegistration = False
    self.biopsyFiducialListNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLMarkupsFiducialNode", "Target")
    self.biopsyFiducialListNode.AddObserver(slicer.vtkMRMLMarkupsNode.PointPositionDefinedEvent, self.onTargetAdded)
    self.biopsyFiducialListNode.AddObserver(slicer.vtkMRMLMarkupsNode.PointModifiedEvent, self.onTargetMoved)

    self.loadedFiles = []
    self.filesToBeLoaded = []
    self.observationTimer = qt.QTimer()
    self.observationTimer.setInterval(500)
    self.observationTimer.timeout.connect(self.observeDicomFolder)

    self.seriesList = []
    self.seriesTimeStamps = dict()
    slicer.mrmlScene.AddObserver(slicer.mrmlScene.NodeAddedEvent,self.onNodeAddedEvent)

    slicer.util.setDataProbeVisible(False)

  def enter(self):
    slicer.util.setDataProbeVisible(False)

  def exit(self):
    slicer.util.setDataProbeVisible(True)

  def setup(self):
    ScriptedLoadableModuleWidget.setup(self)
    moduleDir = os.path.dirname(slicer.util.modulePath(self.__module__))

    # ------------------------------------ Initialization UI ---------------------------------------
    # TODO: 
    # - Change port to 104 [Port can only be changed in DICOM module UI under Query and Retrieve]
    # - Set Default values using a config file
    self.initializeCollapsibleButton = ctk.ctkCollapsibleButton()
    self.initializeCollapsibleButton.text = "Connection"
    #initializeCollapsibleButton.setStyleSheet("ctkCollapsibleButton {background-color: #16417C}")
    self.initializeCollapsibleButton.collapsed = False
    self.layout.addWidget(self.initializeCollapsibleButton)
    initializeLayout = qt.QFormLayout(self.initializeCollapsibleButton)

    initializeFont = qt.QFont()
    initializeFont.setPointSize(18)
    initializeFont.setBold(False)
    self.initializeButton = qt.QPushButton("Initialize Case")
    self.initializeButton.setStyleSheet("QPushButton {background-color: #16417C}")
    self.initializeButton.setFont(initializeFont)
    self.initializeButton.toolTip = "Start DICOM Listener and Create Folders"
    self.initializeButton.enabled = True
    self.initializeButton.connect('clicked()', self.initializeCase)
    initializeLayout.addRow(self.initializeButton)

    self.casesPathBox = qt.QLineEdit("C:/w/data/ProstateBiopsyModuleTest/Cases")
    self.casesPathBox.setReadOnly(True)
    self.casesPathBrowseButton = qt.QPushButton("...")
    self.casesPathBrowseButton.clicked.connect(self.select_directory)
    pathBoxLayout = qt.QHBoxLayout()
    pathBoxLayout.addWidget(self.casesPathBox)
    pathBoxLayout.addWidget(self.casesPathBrowseButton)
    initializeLayout.addRow(pathBoxLayout)

    self.caseDirLabel = qt.QLabel()
    self.caseDirLabel.text = "Waiting to Initialize Case"
    initializeLayout.addRow("Case Directory: ", self.caseDirLabel)
    # -------------------------------------- ----------  --------------------------------------

    # ------------------------------------ Image List UI --------------------------------------
    # TODO: 
    # - Switch images depending on clicking an image (can have a button for it on each row)
    self.imageListCollapsibleButton = ctk.ctkCollapsibleButton()
    self.imageListCollapsibleButton.text = "Images"
    #imageListCollapsibleButton.setStyleSheet("ctkCollapsibleButton {background-color: #16417C}")
    self.imageListCollapsibleButton.collapsed = True
    self.layout.addWidget(self.imageListCollapsibleButton)
    imageListLayout = qt.QVBoxLayout(self.imageListCollapsibleButton)

    # Image List Table with combo boxes to set role for images
    self.imageListTableWidget = qt.QTableWidget(0, 3)
    self.imageListTableWidget.setHorizontalHeaderLabels(["Image Description", "Acquisition Time", "Role"])
    self.imageListTableWidget.horizontalHeader().setSectionResizeMode(qt.QHeaderView.Stretch)
    self.imageListTableWidget.setMaximumHeight(100)
    imageListLayout.addWidget(self.imageListTableWidget)
    self.imageListTableWidget.setSizePolicy(qt.QSizePolicy.MinimumExpanding, qt.QSizePolicy.Minimum)
    # -------------------------------------- ----------  --------------------------------------

    # ---------------------------------- Registration UI --------------------------------------
    # TODO: 
    # - Sometimes Transform sliders have a bug that cause one slider to change another when clicking directly towards it; Might happen with switching transforms
    # - Load in defaults via config file (and make sure to use that in the threshold adjust code)
    self.registrationCollapsibleButton = ctk.ctkCollapsibleButton()
    self.registrationCollapsibleButton.text = "Registration"
    #registrationCollapsibleButton.setStyleSheet("ctkCollapsibleButton {background-color: #16417C}")\
    self.registrationCollapsibleButton.collapsed = True
    self.layout.addWidget(self.registrationCollapsibleButton)
    registrationLayout = qt.QFormLayout(self.registrationCollapsibleButton)

    registerFont = qt.QFont()
    registerFont.setPointSize(18)
    registerFont.setBold(False)
    self.registrationButton = qt.QPushButton("Register")
    self.registrationButton.setStyleSheet("QPushButton {background-color: #16417C}")
    self.registrationButton.setFont(registerFont)
    self.registrationButton.toolTip = "Start registration process for Z-Frame"
    self.registrationButton.enabled = True
    self.registrationButton.connect('clicked()', self.onRegister)
    registrationLayout.addRow(self.registrationButton)

    self.configFileSelectionBox = qt.QComboBox()
    self.configFileSelectionBox.addItems(['Template 001 - Seven Fiducials', 'Template 002 - Nine Fiducials', 'Template 003 - BRP Robot - Nine Fiducials', 'Template 004 - Wide Z-frame - Seven Fiducials'])
    self.configFileSelectionBox.setCurrentIndex(3)
    registrationLayout.addRow('ZFrame Configuration:', self.configFileSelectionBox)

    registrationParametersGroupBox = ctk.ctkCollapsibleGroupBox()
    registrationParametersGroupBox.title = "Automatic Registration Parameters"
    registrationParametersGroupBox.collapsed = True
    registrationParametersLayout = qt.QFormLayout(registrationParametersGroupBox)
    registrationLayout.addRow(registrationParametersGroupBox)

    self.defaultThresholdPercentage = 0.08
    self.thresholdSliderWidget = ctk.ctkSliderWidget()
    self.thresholdSliderWidget.setToolTip("Set range for threshold percentage for isolating registration fiducial markers")
    self.thresholdSliderWidget.setDecimals(2)
    self.thresholdSliderWidget.minimum = 0.00
    self.thresholdSliderWidget.maximum = 1.00
    self.thresholdSliderWidget.singleStep = 0.01
    self.thresholdSliderWidget.value = 0.08
    registrationParametersLayout.addRow("Threshold Percentage:", self.thresholdSliderWidget)

    self.fiducialSizeSliderWidget = ctk.ctkRangeWidget()
    self.fiducialSizeSliderWidget.setToolTip("Set range for fiducial size for isolating registration fiducial markers")
    self.fiducialSizeSliderWidget.setDecimals(0)
    self.fiducialSizeSliderWidget.maximum = 5000
    self.fiducialSizeSliderWidget.minimum = 0
    self.fiducialSizeSliderWidget.singleStep = 1
    self.fiducialSizeSliderWidget.maximumValue = 2000
    self.fiducialSizeSliderWidget.minimumValue = 300
    registrationParametersLayout.addRow("Fiducial Size Range:", self.fiducialSizeSliderWidget)

    self.borderMarginSliderWidget = ctk.ctkSliderWidget()
    self.borderMarginSliderWidget.setToolTip("Set range for threshold percentage for isolating registration fiducial markers")
    self.borderMarginSliderWidget.setDecimals(0)
    self.borderMarginSliderWidget.minimum = 0
    self.borderMarginSliderWidget.maximum = 50
    self.borderMarginSliderWidget.singleStep = 1
    self.borderMarginSliderWidget.value = 15
    registrationParametersLayout.addRow("Border Removal Margin:", self.borderMarginSliderWidget)

    self.removeOrientationCheckBox = qt.QCheckBox("Remove orientation from registration transform")
    self.removeOrientationCheckBox.setChecked(True)
    registrationParametersLayout.addRow(self.removeOrientationCheckBox)

    self.removeBorderIslandsCheckBox = qt.QCheckBox("Remove segment islands on border of volume")
    self.removeBorderIslandsCheckBox.setChecked(True)
    registrationParametersLayout.addRow(self.removeBorderIslandsCheckBox)

    self.repairFiducialImageCheckBox = qt.QCheckBox("Attempt repair of fiducial image")
    self.repairFiducialImageCheckBox.setChecked(True)
    registrationParametersLayout.addRow(self.repairFiducialImageCheckBox)

    self.retryFailedRegistrationCheckBox = qt.QCheckBox("Re-attempt registration with different settings upon failure")
    self.retryFailedRegistrationCheckBox.setChecked(True)
    registrationParametersLayout.addRow(self.retryFailedRegistrationCheckBox)

    self.manualRegistrationGroupBox = ctk.ctkCollapsibleGroupBox()
    self.manualRegistrationGroupBox.title = "Manual Registration"
    self.manualRegistrationGroupBox.collapsed = True
    manualRregistrationLayout = qt.QVBoxLayout(self.manualRegistrationGroupBox)
    registrationLayout.addRow(self.manualRegistrationGroupBox)

    self.manualRegistrationTransformSliders = slicer.qMRMLTransformSliders()
    self.manualRegistrationTransformSliders.setWindowTitle("Translation")
    self.manualRegistrationTransformSliders.TypeOfTransform = slicer.qMRMLTransformSliders.TRANSLATION
    self.manualRegistrationTransformSliders.setDecimals(3)
    manualRregistrationLayout.addWidget(self.manualRegistrationTransformSliders)

    self.manualRegistrationRotationSliders = slicer.qMRMLTransformSliders()
    self.manualRegistrationRotationSliders.setWindowTitle("Rotation")
    self.manualRegistrationRotationSliders.TypeOfTransform = slicer.qMRMLTransformSliders.ROTATION
    self.manualRegistrationRotationSliders.setDecimals(3)
    manualRregistrationLayout.addWidget(self.manualRegistrationRotationSliders)

    self.identityButton = qt.QPushButton("Identity")
    self.identityButton.toolTip = "Set Z-Frame calibration matrix to be Identity"
    self.identityButton.setMaximumWidth(100)
    self.identityButton.connect('clicked()', self.onIdentity)
    manualRregistrationLayout.addWidget(self.identityButton)

    manualRegisterFont = qt.QFont()
    manualRegisterFont.setPointSize(14)
    manualRegisterFont.setBold(False)
    useManualRegistrationButton = qt.QPushButton("Accept Manual Registration")
    useManualRegistrationButton.setStyleSheet("QPushButton {background-color: #16417C}")
    useManualRegistrationButton.toolTip = "Flag module to accept manual registration"
    useManualRegistrationButton.connect('clicked()', self.onUseManualRegistration)
    manualRregistrationLayout.addWidget(useManualRegistrationButton)

    self.manualRegistrationTransformSliders.setMRMLTransformNode(self.ZFrameCalibrationTransformNode)
    self.manualRegistrationRotationSliders.setMRMLTransformNode(self.ZFrameCalibrationTransformNode)
    # ------------------------------------- ----------  --------------------------------------
    
    # ------------------------------------- Planning UI --------------------------------------
    self.planningCollapsibleButton = ctk.ctkCollapsibleButton()
    self.planningCollapsibleButton.text = "Planning"
    #planningCollapsibleButton.setStyleSheet("ctkCollapsibleButton {background-color: #16417C}")
    self.planningCollapsibleButton.collapsed = True
    self.layout.addWidget(self.planningCollapsibleButton)
    planningLayout = qt.QVBoxLayout(self.planningCollapsibleButton)

    addTargetFont = qt.QFont()
    addTargetFont.setPointSize(18)
    addTargetFont.setBold(False)
    self.addTargetButton = qt.QPushButton("Add Target")
    self.addTargetButton.setStyleSheet("QPushButton {background-color: #16417C}")
    self.addTargetButton.setFont(registerFont)
    self.addTargetButton.toolTip = "Add a target for biopsy"
    self.addTargetButton.enabled = True
    self.addTargetButton.connect('clicked()', self.onAddTarget)
    planningLayout.addWidget(self.addTargetButton)

    # Image List Table with combo boxes to set role for images
    self.targetListTableWidget = qt.QTableWidget(0, 5)
    self.targetListTableWidget.setHorizontalHeaderLabels(["Target", "Grid", "Depth\n(cm)", "Position\n(RAS)", "   "])
    self.targetListTableWidget.horizontalHeader().setSectionResizeMode(0, qt.QHeaderView.Stretch)
    self.targetListTableWidget.horizontalHeader().setSectionResizeMode(1, qt.QHeaderView.Stretch)
    self.targetListTableWidget.horizontalHeader().setSectionResizeMode(2, qt.QHeaderView.Stretch)
    self.targetListTableWidget.horizontalHeader().setSectionResizeMode(3, qt.QHeaderView.Stretch)
    self.targetListTableWidget.horizontalHeader().setSectionResizeMode(4, qt.QHeaderView.Fixed)
    self.targetListTableWidget.setMaximumHeight(200)
    planningLayout.addWidget(self.targetListTableWidget)
    self.targetListTableWidget.setSizePolicy(qt.QSizePolicy.MinimumExpanding, qt.QSizePolicy.Minimum)
    self.targetListTableWidget.setColumnWidth(4, 10)
    self.targetListTableWidget.itemChanged.connect(self.onTargetListItemChanged)

    generateWorksheetFont = qt.QFont()
    generateWorksheetFont.setPointSize(14)
    generateWorksheetFont.setBold(False)
    self.generateWorksheetButton = qt.QPushButton("Generate Worksheet")
    #self.generateWorksheetButton.setStyleSheet("QPushButton {background-color: #16417C}")
    self.generateWorksheetButton.setFont(generateWorksheetFont)
    self.generateWorksheetButton.toolTip = "Generate PDF for worksheet"
    self.generateWorksheetButton.enabled = True
    self.generateWorksheetButton.connect('clicked()', self.onGenerateWorksheet)
    planningLayout.addWidget(self.generateWorksheetButton)
    # worksheetLayout.addWidget(self.generateWorksheetButton)
    # -------------------------------------- ----------  --------------------------------------
    line = qt.QFrame()
    line.setFrameShape(qt.QFrame.HLine)
    self.layout.addWidget(line)

    footerWidget = qt.QWidget()
    self.layout.addWidget(footerWidget)
    footerLayout = qt.QGridLayout(footerWidget)

    initializeFont = qt.QFont()
    initializeFont.setPointSize(12)
    initializeFont.setBold(False)
    self.closeCaseButton = qt.QPushButton("Save and Close Case")
    self.closeCaseButton.setStyleSheet("QPushButton {background-color: #16417C}")
    self.closeCaseButton.setMaximumWidth(150)
    self.closeCaseButton.setFont(initializeFont)
    self.closeCaseButton.toolTip = "Close Case"
    self.closeCaseButton.enabled = False
    self.closeCaseButton.connect('clicked()', self.saveAndCloseCase)
    footerLayout.addWidget(self.closeCaseButton, 0, 0, 1, 2)

    rulerButton = qt.QPushButton("")
    rulerButton.setMaximumWidth(50)
    rulerIconPath = os.path.join(moduleDir, 'Resources/Icons', 'MarkupsLine.png')
    rulerIcon = qt.QIcon(rulerIconPath)
    rulerButton.setIcon(rulerIcon)
    rulerButton.toolTip = "Add ruler"
    rulerButton.connect('clicked()', self.addRuler)
    footerLayout.addWidget(rulerButton, 0, 7, 1, 1)

    self.toggleGuideButton = qt.QPushButton("")
    self.toggleGuideButton.setMaximumWidth(50)
    guideIconPath = os.path.join(moduleDir, 'Resources/Icons', 'GuideHoles.png')
    guideIcon = qt.QIcon(guideIconPath)
    self.toggleGuideButton.setIcon(guideIcon)
    self.toggleGuideButton.toolTip = "Toggle guide hole visibility"
    self.toggleGuideButton.setCheckable(True)
    self.toggleGuideButton.setChecked(True)
    self.toggleGuideButton.connect('clicked()', self.toggleGuideHoles)
    footerLayout.addWidget(self.toggleGuideButton, 0, 8, 1, 1)

    self.autoCheckBox = qt.QCheckBox("Auto")
    self.autoCheckBox.setChecked(True)
    footerLayout.addWidget(self.autoCheckBox, 1, 8, 1, 1)

    # Add vertical spacer
    self.layout.addStretch(1)
    # -------------------------------------- ----------  --------------------------------------


  def select_directory(self):
    directory = qt.QFileDialog.getExistingDirectory(self.parent, "Select Cases Directory")
    if directory:
      self.casesPathBox.setText(directory)

  def removeNodeByName(self, nodeName):
    nodes = slicer.util.getNodes(nodeName)
    for node in nodes.values():
      slicer.mrmlScene.RemoveNode(node)

  def numpy_to_vtk_image_data(self, numpy_array):
    image_data = vtk.vtkImageData()
    flat_data_array = numpy_array.transpose(2,1,0).flatten()
    vtk_data =  vtk.util.numpy_support.numpy_to_vtk(num_array=flat_data_array, deep=True, array_type=vtk.VTK_FLOAT)
    shape = numpy_array.shape

    image_data.GetPointData().SetScalars(vtk_data)
    image_data.SetDimensions(shape[0], shape[1], shape[2])
    return image_data

  # Phase change enables later steps, but generally does not disable older steps unless those steps would cause problems
  def onPhaseChange(self, phase):
    if phase == "START":
      self.currentPhase = 'START'
      # self.casesPathBrowseButton.enabled = True
      # self.initializeButton.enabled = True
      # self.closeCaseButton.enabled = False

      if self.autoCheckBox.isChecked():
        self.initializeCollapsibleButton.collapsed = False
        self.imageListCollapsibleButton.collapsed = True
        self.registrationCollapsibleButton.collapsed = True
        self.planningCollapsibleButton.collapsed = True
    elif phase == "REGISTRATION":
      self.currentPhase = 'REGISTRATION'
      # self.casesPathBrowseButton.enabled = False
      # self.initializeButton.enabled = False
      # self.closeCaseButton.enabled = True

      if self.autoCheckBox.isChecked():
        self.initializeCollapsibleButton.collapsed = True
        self.imageListCollapsibleButton.collapsed = False
        self.registrationCollapsibleButton.collapsed = False
        self.planningCollapsibleButton.collapsed = True

        qt.QTimer.singleShot(1000, lambda: self.onRegister())
    elif phase == "PLANNING": 
      self.currentPhase = 'PLANNING'
      # self.casesPathBrowseButton.enabled = False
      # self.initializeButton.enabled = False
      # self.closeCaseButton.enabled = True

      if self.autoCheckBox.isChecked():
        self.initializeCollapsibleButton.collapsed = True
        self.imageListCollapsibleButton.collapsed = False
        self.registrationCollapsibleButton.collapsed = True
        self.planningCollapsibleButton.collapsed = False

        self.focusSliceWindowsOnVolume(self.getNodeFromImageRole("PLANNING"))
    print(f'Current phase: {self.currentPhase}')

  # ------------------------------------- Connection -------------------------------------

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

    self.initializeButton.enabled = False
    self.casesPathBrowseButton.enabled = False
    self.closeCaseButton.enabled = True

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

  def saveAndCloseCase(self):
    self.seriesList = []
    self.loadedFiles = []
    self.filesToBeLoaded = []
    self.observationTimer.stop()

    # TODO: Might be safer to restart Slicer
    sceneSaveFilename = f'{self.caseDirPath}/saved-scene-{time.strftime("%Y%m%d-%H%M%S")}.mrb'
    if slicer.util.saveScene(sceneSaveFilename):
      print("Scene saved to: {0}".format(sceneSaveFilename))
      slicer.mrmlScene.Clear(0)
      self.onReload()
    else:
      print("Scene saving failed")
        
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
    if (newNode.GetName() in self.ignoredVolumeNames):
      return
    self.addToImageList(newNode)

  def addToImageList(self, newNode):
    if not self.imageListTableWidget:
      return
    rowCount = self.imageListTableWidget.rowCount

    imageName = qt.QTableWidgetItem(newNode.GetName())
    imageName.setFlags(imageName.flags() & ~qt.Qt.ItemIsEditable)
    imageTime = qt.QTableWidgetItem(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    imageTime.setFlags(imageTime.flags() & ~qt.Qt.ItemIsEditable)
    imageRoleChoice = qt.QComboBox()
    imageRoleChoice.addItems(self.imageRoles) # add some options to the combo box
    imageRoleChoice.currentTextChanged.connect(lambda: self.updateImageListRoles(imageRoleChoice.currentText, rowCount))

    self.imageListTableWidget.insertRow(rowCount)
    self.imageListTableWidget.setItem(rowCount, 0, imageName)
    self.imageListTableWidget.setItem(rowCount, 1, imageTime)
    self.imageListTableWidget.setCellWidget(rowCount, 2, imageRoleChoice) # set the combo box as a cell widget

    self.autoImageRoleAssignment(newNode, imageRoleChoice, rowCount)

    # TODO: Maybe force the slice windows back to the image that should be showing depending on which Phase the module is in

  def updateImageListRoles(self, newRoleAssigned, rowCount):
    if newRoleAssigned in ["CALIBRATION", "PLANNING"]:
      for index in range(0, self.imageListTableWidget.rowCount):
        if (index != rowCount) and (newRoleAssigned == (self.imageListTableWidget.cellWidget(index, 2).currentText)):
          self.imageListTableWidget.cellWidget(index, 2).setCurrentIndex(self.imageRoles.index("N/A"))

  def autoImageRoleAssignment(self, newNode, imageRoleChoice, rowCount):
    name = newNode.GetName()
    if "template" in name.casefold():
      imageRoleChoice.setCurrentIndex(self.imageRoles.index("CALIBRATION"))
      self.updateImageListRoles(imageRoleChoice.currentText, rowCount)
      if self.currentPhase == "START":
        # self.onPhaseChange("REGISTRATION")
        pass
      elif self.currentPhase == "REGISTRATION":
        if self.autoCheckBox.isChecked():
          qt.QTimer.singleShot(1000, lambda: self.onRegister())

    elif "cover" in name.casefold():
      imageRoleChoice.setCurrentIndex(self.imageRoles.index("PLANNING"))
      self.updateImageListRoles(imageRoleChoice.currentText, rowCount)
      if self.validRegistration:
        self.onPhaseChange("PLANNING")

  def getNodeFromImageRole(self, imageRole):
    for index in range(0, self.imageListTableWidget.rowCount):
      if self.imageListTableWidget.cellWidget(index, 2).currentText == imageRole:
        return slicer.mrmlScene.GetFirstNodeByName(self.imageListTableWidget.item(index, 0).text())
    return None
  
  # ------------------------------------- Registration -----------------------------------

  def onRegister(self):
    if not self.getNodeFromImageRole("CALIBRATION"):
      return

    self.loadTemplateConfiguration()

    result = False
    result, outputTransform = self.registerZFrame()
    self.increaseThresholdForRetry = False

    if self.zFrameModelNode:
      self.zFrameModelNode.SetAndObserveTransformNodeID(outputTransform.GetID())
      self.zFrameModelNode.GetDisplayNode().SetVisibility2D(True)
      self.templateModelNode.SetDisplayVisibility(True)
    if self.templateModelNode:
      self.templateModelNode.SetAndObserveTransformNodeID(outputTransform.GetID())
      self.templateModelNode.GetDisplayNode().SetVisibility2D(False)
      self.templateModelNode.SetDisplayVisibility(True)
    if self.calibratorModelNode:
      self.calibratorModelNode.SetAndObserveTransformNodeID(outputTransform.GetID())
      self.calibratorModelNode.GetDisplayNode().SetVisibility2D(True)
      self.calibratorModelNode.GetDisplayNode().SetSliceIntersectionThickness(1)
      self.calibratorModelNode.SetDisplayVisibility(True)
    if self.guideHolesModelNode:
      self.guideHolesModelNode.SetAndObserveTransformNodeID(outputTransform.GetID())
      if self.toggleGuideButton.isChecked():
        self.guideHolesModelNode.GetDisplayNode().SetVisibility2D(True)
      else:
        self.guideHolesModelNode.GetDisplayNode().SetVisibility2D(False)
      self.guideHolesModelNode.GetDisplayNode().SetSliceIntersectionThickness(1)
      self.guideHolesModelNode.SetDisplayVisibility(True)
    if self.guideHoleLabelsModelNode:
      self.guideHoleLabelsModelNode.SetAndObserveTransformNodeID(outputTransform.GetID())
      self.guideHoleLabelsModelNode.GetDisplayNode().SetVisibility2D(True)
      self.guideHoleLabelsModelNode.GetDisplayNode().SetSliceIntersectionThickness(1)
      self.guideHoleLabelsModelNode.SetDisplayVisibility(True)

    if result:
      self.onRegistrationSuccess()
    else:
      self.onRegistrationFailure()

  def registerZFrame(self):
    # If there is a zFrame image selected, perform the calibration step to calculate the CLB matrix
    inputVolume = self.getNodeFromImageRole("CALIBRATION")

    if self.ZFrameCalibrationTransformNode:
      outputTransform = self.ZFrameCalibrationTransformNode
    else:
      self.removeNodeByName("ZFrameTransform")
      self.ZFrameCalibrationTransformNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLLinearTransformNode", "ZFrameTransform")
      outputTransform = self.ZFrameCalibrationTransformNode

    if not inputVolume:
      return False, outputTransform
    
    # First try without repair methods
    loopRegistration = True
    while loopRegistration:
      zFrameMaskedVolume = self.createMaskedVolumeBySize(inputVolume, False)
      if zFrameMaskedVolume.GetImageData().GetScalarRange()[1] > 0:
        # Crop if not 256x256
        zFrameMaskedVolumeDims = zFrameMaskedVolume.GetImageData().GetDimensions()
        if zFrameMaskedVolumeDims[0] != 256 and zFrameMaskedVolumeDims[1] != 256:
          self.cropVolume(zFrameMaskedVolume, 256, 256)
        
        centerOfMassSlice = int(self.findCentroidOfVolume(zFrameMaskedVolume)[2])
        # Run zFrameRegistration CLI module
        params = {'inputVolume': zFrameMaskedVolume, 'startSlice': centerOfMassSlice-3, 'endSlice': centerOfMassSlice+3,
                  'outputTransform': outputTransform, 'zframeConfig': self.zframeConfig, 'frameTopology': self.frameTopologyString, 
                  'zFrameFids': ''}
        cliNode = slicer.cli.run(slicer.modules.zframeregistration, None, params, wait_for_completion=True)
        if cliNode.GetStatus() & cliNode.ErrorsMask:
          print(cliNode.GetErrorText())
        if self.removeOrientationCheckBox.isChecked():
          self.removeOrientationComponent(outputTransform)
      else:
        print("Masked volume empty")
      regResult = self.checkRegistrationResult(outputTransform, zFrameMaskedVolume, self.zFrameFiducials)
      if not regResult:
        # Try to process at different thresholds
        if self.retryFailedRegistrationCheckBox.isChecked():
          if not self.increaseThresholdForRetry:
            if not (self.thresholdSliderWidget.value <= self.thresholdSliderWidget.minimum):
              self.thresholdSliderWidget.value = self.thresholdSliderWidget.value - 0.02
              print(f'Retrying; decreasing threshold percentage to {self.thresholdSliderWidget.value}')
              loopRegistration = True
            else:
              self.increaseThresholdForRetry = True
              self.thresholdSliderWidget.value = self.defaultThresholdPercentage + 0.04
              print(f'Retrying; increasing threshold percentage to {self.thresholdSliderWidget.value}')
              loopRegistration = True
          else:
            if not (self.thresholdSliderWidget.value >= (self.thresholdSliderWidget.maximum/5)):
              self.thresholdSliderWidget.value = self.thresholdSliderWidget.value + 0.04
              print(f'Retrying; increasing threshold percentage to {self.thresholdSliderWidget.value}')
              loopRegistration = True
            else:
              print("Retries failed; Moving on to repair attempt")
              self.thresholdSliderWidget.value = self.defaultThresholdPercentage
              loopRegistration = False
        else:
          loopRegistration = False
      else:
        loopRegistration = False
        return True, outputTransform
      
    if self.repairFiducialImageCheckBox.isChecked():
      zFrameMaskedVolume = self.createMaskedVolumeBySize(inputVolume, True)
      if zFrameMaskedVolume.GetImageData().GetScalarRange()[1] > 0:
        # Crop if not 256x256
        zFrameMaskedVolumeDims = zFrameMaskedVolume.GetImageData().GetDimensions()
        # TODO: Pad images smaller than 256 by 256
        if zFrameMaskedVolumeDims[0] != 256 and zFrameMaskedVolumeDims[1] != 256:
          self.cropVolume(zFrameMaskedVolume, 256, 256)
        
        centerOfMassSlice = int(self.findCentroidOfVolume(zFrameMaskedVolume)[2])
        # Run zFrameRegistration CLI module
        params = {'inputVolume': zFrameMaskedVolume, 'startSlice': centerOfMassSlice-3, 'endSlice': centerOfMassSlice+3,
                  'outputTransform': outputTransform, 'zframeConfig': self.zframeConfig, 'frameTopology': self.frameTopologyString, 
                  'zFrameFids': ''}
        cliNode = slicer.cli.run(slicer.modules.zframeregistration, None, params, wait_for_completion=True)
        if cliNode.GetStatus() & cliNode.ErrorsMask:
          print(cliNode.GetErrorText())
        if self.removeOrientationCheckBox.isChecked():
          self.removeOrientationComponent(outputTransform)
      else:
        print("Masked volume empty")

      regResult = self.checkRegistrationResult(outputTransform, zFrameMaskedVolume, self.zFrameFiducials)
      return regResult, outputTransform
    
    return False, outputTransform
  
  def removeOrientationComponent(self, transformNode):
    # Get the transformation matrix
    matrix = vtk.vtkMatrix4x4()
    transformNode.GetMatrixTransformToParent(matrix)

    # Set the orientation part of the matrix to identity
    for i in range(3):
        for j in range(3):
            if i == j:
                matrix.SetElement(i, j, 1)
            else:
                matrix.SetElement(i, j, 0)

    # Update the transform node
    transformNode.SetMatrixTransformToParent(matrix)

  def focusSliceWindowsOnVolume(self, volumeNode):
    red_logic = slicer.app.layoutManager().sliceWidget("Red").sliceLogic()
    compositeNodeR = red_logic.GetSliceCompositeNode()
    green_logic = slicer.app.layoutManager().sliceWidget("Green").sliceLogic()
    compositeNodeG = green_logic.GetSliceCompositeNode()
    yellow_logic = slicer.app.layoutManager().sliceWidget("Yellow").sliceLogic()
    compositeNodeY = yellow_logic.GetSliceCompositeNode()

    compositeNodeR.SetBackgroundVolumeID(volumeNode.GetID())
    compositeNodeG.SetBackgroundVolumeID(volumeNode.GetID())
    compositeNodeY.SetBackgroundVolumeID(volumeNode.GetID())

  def onRegistrationSuccess(self):
    print("Registration Successful")
    self.increaseThresholdForRepair = False
    inputVolume = self.getNodeFromImageRole("CALIBRATION")

    self.focusSliceWindowsOnVolume(inputVolume)
    self.validRegistration = True

    # with open('C:/w/data/ProstateBiopsyModuleTest/RegTest/successLog.txt', 'a') as f:
    #   f.write(f'{inputVolume.GetName()}\n')

    if self.getNodeFromImageRole("PLANNING"):
      self.onPhaseChange("PLANNING")

  def onRegistrationFailure(self):
    print("Registration Failure")
    self.increaseThresholdForRepair = False
    self.manualRegistrationGroupBox.collapsed = False
    inputVolume = self.getNodeFromImageRole("CALIBRATION")

    self.focusSliceWindowsOnVolume(inputVolume)
    self.validRegistration = False
    
    # with open('C:/w/data/ProstateBiopsyModuleTest/RegTest/failLog.txt', 'a') as f:
    #   f.write(f'{inputVolume.GetName()}\n')
  
  def onUseManualRegistration(self):
    self.validRegistration = True
    if self.getNodeFromImageRole("PLANNING"):
      self.onPhaseChange("PLANNING")
  
  def checkRegistrationResult(self, outputTransform, fiducialVolume, zFrameFiducials):
    # Check the midpoint of each ZFrame fiducial and some points around it for a detected fiducial
    zFrameMidpoints = []
    for zFrameFiducial in zFrameFiducials:
      zFrameMidpoints.append([(zFrameFiducial[0] + zFrameFiducial[3]) / 2, (zFrameFiducial[1] + zFrameFiducial[4]) / 2, (zFrameFiducial[2] + zFrameFiducial[5]) / 2])
    for zFrameMidpoint in zFrameMidpoints:
      zFrameMidpoint = zFrameMidpoint + [1]

      outputMatrix = vtk.vtkMatrix4x4()
      outputTransform.GetMatrixTransformToParent(outputMatrix)
      transformedMidpoint = [0, 0, 0, 1]
      outputMatrix.MultiplyPoint(zFrameMidpoint, transformedMidpoint)

      rasToIjkMatrix = vtk.vtkMatrix4x4()
      fiducialVolume.GetRASToIJKMatrix(rasToIjkMatrix)
      ijkMidpoint = [0, 0, 0, 1]
      rasToIjkMatrix.MultiplyPoint(transformedMidpoint, ijkMidpoint)

      fiducialImageData = fiducialVolume.GetImageData()
      # Check if point is in extent of volume
      extent = fiducialImageData.GetExtent()
      if not (extent[0] <= ijkMidpoint[0] <= extent[1] and extent[2] <= ijkMidpoint[1] <= extent[3] and extent[4] <= ijkMidpoint[2] <= extent[5]):
        return False
      # Check point and surrounding points
      fiducialFound = False
      if fiducialImageData.GetScalarComponentAsDouble(int(ijkMidpoint[0]), int(ijkMidpoint[1]), int(ijkMidpoint[2]), 0) > 0: fiducialFound = True
      if fiducialImageData.GetScalarComponentAsDouble(int(ijkMidpoint[0])-2, int(ijkMidpoint[1]), int(ijkMidpoint[2]), 0) > 0: fiducialFound = True
      if fiducialImageData.GetScalarComponentAsDouble(int(ijkMidpoint[0])+2, int(ijkMidpoint[1]), int(ijkMidpoint[2]), 0) > 0: fiducialFound = True
      if fiducialImageData.GetScalarComponentAsDouble(int(ijkMidpoint[0]), int(ijkMidpoint[1])-2, int(ijkMidpoint[2]), 0) > 0: fiducialFound = True
      if fiducialImageData.GetScalarComponentAsDouble(int(ijkMidpoint[0]), int(ijkMidpoint[1])+2, int(ijkMidpoint[2]), 0) > 0: fiducialFound = True
      if fiducialImageData.GetScalarComponentAsDouble(int(ijkMidpoint[0]), int(ijkMidpoint[1]), int(ijkMidpoint[2])-2, 0) > 0: fiducialFound = True
      if fiducialImageData.GetScalarComponentAsDouble(int(ijkMidpoint[0]), int(ijkMidpoint[1]), int(ijkMidpoint[2])+2, 0) > 0: fiducialFound = True
      if not fiducialFound:
        return False
    return True

  def loadTemplateConfiguration(self):
    currentFilePath = os.path.dirname(slicer.util.modulePath(self.__module__))
    self.zframeConfig = ""
    if self.configFileSelectionBox.currentIndex == 0:
      ZFRAME_MODEL_PATH = 'template001/zframe001-model.vtk'
      TEMPLATE_MODEL_PATH = 'template001/template001.vtk'
      CALIBRATOR_MODEL_PATH = 'template001/template001-Calibrator.vtk'
      GUIDEHOLES_MODEL_PATH = 'template001/template001-GuideHoles.vtk'
      GUIDEHOLELABELS_MODEL_PATH = 'template001/template001-GuideHoleLabels.vtk'
      self.templateWorksheetPath = 'template001/BiopsyWorksheet.pdf'
      self.templateWorksheetOverlayPath = 'template001/BiopsyWorksheet_Overlay.pdf'
      self.zframeConfig = 'z001'
      zframeConfigFilePath = os.path.join(currentFilePath, "Resources/Templates/template001/zframe001.txt")
    elif self.configFileSelectionBox.currentIndex == 1:
      ZFRAME_MODEL_PATH = 'template002/zframe002-model.vtk'
      TEMPLATE_MODEL_PATH = 'template002/template002.vtk'
      CALIBRATOR_MODEL_PATH = 'template002/template002-Calibrator.vtk'
      GUIDEHOLES_MODEL_PATH = 'template002/template002-GuideHoles.vtk'
      GUIDEHOLELABELS_MODEL_PATH = 'template002/template002-GuideHoleLabels.vtk'
      self.templateWorksheetPath = 'template002/BiopsyWorksheet.pdf'
      self.templateWorksheetOverlayPath = 'template002/BiopsyWorksheet_Overlay.pdf'
      self.zframeConfig = 'z002'
      zframeConfigFilePath = os.path.join(currentFilePath, "Resources/Templates/template002/zframe002.txt")
    elif self.configFileSelectionBox.currentIndex == 2:
      ZFRAME_MODEL_PATH = 'template003/zframe003-model.vtk'
      TEMPLATE_MODEL_PATH = 'template003/template003.vtk'
      CALIBRATOR_MODEL_PATH = 'template003/template003-Calibrator.vtk'
      GUIDEHOLES_MODEL_PATH = 'template003/template003-GuideHoles.vtk'
      GUIDEHOLELABELS_MODEL_PATH = 'template003/template003-GuideHoleLabels.vtk'
      self.templateWorksheetPath = 'template003/BiopsyWorksheet.pdf'
      self.templateWorksheetOverlayPath = 'template003/BiopsyWorksheet_Overlay.pdf'
      self.zframeConfig = 'z003'
      zframeConfigFilePath = os.path.join(currentFilePath, "Resources/Templates/template003/zframe003.txt")
    else: #self.configFileSelectionBox.currentIndex == 3:
      ZFRAME_MODEL_PATH = 'template004/zframe004-model.vtk'
      TEMPLATE_MODEL_PATH = 'template004/template004.vtk'
      CALIBRATOR_MODEL_PATH = 'template004/template004-Calibrator.vtk'
      GUIDEHOLES_MODEL_PATH = 'template004/template004-GuideHoles.vtk'
      GUIDEHOLELABELS_MODEL_PATH = 'template004/template004-GuideHoleLabels.vtk'
      self.templateWorksheetPath = 'template004/BiopsyWorksheet.pdf'
      self.templateWorksheetOverlayPath = 'template004/BiopsyWorksheet_Overlay.pdf'
      self.zframeConfig = 'z004'
      zframeConfigFilePath = os.path.join(currentFilePath, "Resources/Templates/template004/zframe004.txt")
    
    with open(zframeConfigFilePath,"r") as f:
      configFileLines = f.readlines()

    # Parse zFrame configuration file here to identify the dimensions and topology of the zframe
    # Save the origins and diagonal vectors of each of the 3 sides of the zframe in an array
    self.frameTopology = []
    self.zFrameFiducials = []
    self.templateHorizontalLabels = []
    self.templateVerticalLabels = []
    self.worksheetOrigin = []
    templateOriginFound = False
    for line in configFileLines:
      if line.startswith('Side 1') or line.startswith('Side 2'): 
        vec = [float(s) for s in re.findall(r'-?\d+\.?\d*', line)]
        vec.pop(0)
        self.frameTopology.append(vec)
      elif line.startswith('Base'):
        vec = [float(s) for s in re.findall(r'-?\d+\.?\d*', line)]
        self.frameTopology.append(vec)
      elif line.startswith('Fiducial'):
        vec = [float(s) for s in re.findall(r'(-?\d+)(?!:)', line)]
        self.zFrameFiducials.append(vec)
      elif line.startswith('Template origin'):
        numbers = re.findall(r"[-+]?\d*\.\d+|\d+", line)
        self.templateOrigin = [float(i) for i in numbers]
        templateOriginFound = True
      elif line.startswith('Horizontal offset'):
        matches = re.findall(r'\b(\d+\.\d+?)\b', line)
        self.templateHorizontalOffset = float(matches[0])
      elif line.startswith('Vertical offset'):
        matches = re.findall(r'\b(\d+\.\d+?)\b', line)
        self.templateVerticalOffset = float(matches[0])
      elif line.startswith('Horizontal labels'):
        labels = re.findall(r'\((.*?)\)', line)
        self.templateHorizontalLabels = [char for char in labels[0].split(',')]
      elif line.startswith('Vertical labels'):
        labels = re.findall(r'\((.*?)\)', line)
        self.templateVerticalLabels = [char for char in labels[0].split(',')]
      elif line.startswith('Worksheet origin'):
        matches = re.findall(r'[-+]?\d*\.\d+|\d+', line)
        self.worksheetOrigin.append([float(matches[1]), float(matches[2])])
      elif line.startswith('Worksheet horizontal offset'):
        match = re.findall(r'\b(\d+\.\d+?)\b', line)
        self.worksheetHorizontalOffset = float(match[0])
      elif line.startswith('Worksheet vertical offset'):
        match = re.findall(r'\b(\d+\.\d+?)\b', line)
        self.worksheetVerticalOffset = float(match[0])
      
    if not templateOriginFound:
      raise Exception("ZFrame configuration file is missing template origin")

    # Convert frameTopology points to a string, for the sake of passing it as a string argument to the ZframeRegistration CLI 
    self.frameTopologyString = ' '.join([str(elem) for elem in self.frameTopology])
    
    self.loadTemplateModels(ZFRAME_MODEL_PATH,'ZFrameModel',TEMPLATE_MODEL_PATH,'TemplateModel',CALIBRATOR_MODEL_PATH,'CalibratorModel',GUIDEHOLES_MODEL_PATH,'GuideHolesModel',GUIDEHOLELABELS_MODEL_PATH,'GuideHoleLabelsModel')

  def loadTemplateModels(self, ZFRAME_MODEL_PATH, ZFRAME_MODEL_NAME, TEMPLATE_MODEL_PATH, TEMPLATE_MODEL_NAME, CALIBRATOR_MODEL_PATH, CALIBRATOR_MODEL_NAME, GUIDEHOLES_MODEL_PATH, GUIDEHOLES_MODEL_NAME,  GUIDEHOLELABELS_MODEL_PATH, GUIDEHOLELABELS_MODEL_NAME):
    currentFilePath = os.path.dirname(slicer.util.modulePath(self.__module__))

    # All models must be created with the same origin and fit together. The module assumes that the models were created correctly.
    # Z-Frame
    self.removeNodeByName(ZFRAME_MODEL_NAME)
    modelPath = os.path.join(currentFilePath, "Resources", "Templates", ZFRAME_MODEL_PATH)
    try:
      self.zFrameModelNode = slicer.util.loadModel(modelPath)
    except:
      print(f'Failed to load model from {modelPath}')
    if self.zFrameModelNode:
      self.zFrameModelNode.SetName(ZFRAME_MODEL_NAME)
      modelDisplayNode = self.zFrameModelNode.GetDisplayNode()
      modelDisplayNode.SetColor(0.0,1.0,1.0)
      modelDisplayNode.SetSliceIntersectionThickness(2)
      self.zFrameModelNode.SetDisplayVisibility(True)

    # Template
    self.removeNodeByName(TEMPLATE_MODEL_NAME)
    modelPath = os.path.join(currentFilePath, "Resources", "Templates", TEMPLATE_MODEL_PATH)
    try:
      self.templateModelNode = slicer.util.loadModel(modelPath)
    except:
      print(f'Failed to load model from {modelPath}')  
    if self.templateModelNode:    
      self.templateModelNode.SetName(TEMPLATE_MODEL_NAME)
      modelDisplayNode = self.templateModelNode.GetDisplayNode()
      modelDisplayNode.SetColor(0.67,0.67,0.67)
      modelDisplayNode.SetOpacity(0.35)
      self.templateModelNode.SetDisplayVisibility(True)

    # Calibrator
    self.removeNodeByName(CALIBRATOR_MODEL_NAME)
    modelPath = os.path.join(currentFilePath, "Resources", "Templates", CALIBRATOR_MODEL_PATH)
    try:
      self.calibratorModelNode = slicer.util.loadModel(modelPath)
    except:
      print(f'Failed to load model from {modelPath}')
    if self.calibratorModelNode:  
      self.calibratorModelNode.SetName(CALIBRATOR_MODEL_NAME)
      modelDisplayNode = self.calibratorModelNode.GetDisplayNode()
      modelDisplayNode.SetColor(1.0,1.0,0.0)
      modelDisplayNode.SetOpacity(0.15)
      modelDisplayNode.SetSliceIntersectionOpacity(0.25)
      self.calibratorModelNode.SetDisplayVisibility(True)

    # Guide Holes
    self.removeNodeByName(GUIDEHOLES_MODEL_NAME)
    modelPath = os.path.join(currentFilePath, "Resources", "Templates", GUIDEHOLES_MODEL_PATH)
    try:
      self.guideHolesModelNode = slicer.util.loadModel(modelPath)
    except:
      print(f'Failed to load model from {modelPath}')
    if self.guideHolesModelNode:
      self.guideHolesModelNode.SetName(GUIDEHOLES_MODEL_NAME)
      modelDisplayNode = self.guideHolesModelNode.GetDisplayNode()
      modelDisplayNode.SetColor(0.59,0.88,0.64)
      modelDisplayNode.SetOpacity(0.05)
      modelDisplayNode.SetSliceIntersectionOpacity(0.65)
      self.guideHolesModelNode.SetDisplayVisibility(True)

    # Guide Hole Labels
    self.removeNodeByName(GUIDEHOLELABELS_MODEL_NAME)
    modelPath = os.path.join(currentFilePath, "Resources", "Templates", GUIDEHOLELABELS_MODEL_PATH)
    try:
      self.guideHoleLabelsModelNode = slicer.util.loadModel(modelPath)
    except:
      print(f'Failed to load model from {modelPath}')
    if self.guideHoleLabelsModelNode:
      self.guideHoleLabelsModelNode.SetName(GUIDEHOLELABELS_MODEL_NAME)
      modelDisplayNode = self.guideHoleLabelsModelNode.GetDisplayNode()
      modelDisplayNode.SetColor(0.49,0.78,0.54)
      modelDisplayNode.SetOpacity(0.00)
      modelDisplayNode.SetSliceIntersectionOpacity(0.75)
      self.guideHoleLabelsModelNode.SetDisplayVisibility(True)

  def createMaskedVolumeBySize(self, inputVolume, repair):
    loopRegistration = True
    while loopRegistration:
      thresholdPercent = self.thresholdSliderWidget.value 
      minimumSize = self.fiducialSizeSliderWidget.minimumValue
      maximumSize = self.fiducialSizeSliderWidget.maximumValue
      zframeConfig = self.zframeConfig

      # Create segmentation node
      segmentationNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLSegmentationNode")
      segmentationNode.CreateDefaultDisplayNodes()
      segmentationNode.SetReferenceImageGeometryParameterFromVolumeNode(inputVolume)

      # Create segment
      segmentId = segmentationNode.GetSegmentation().AddEmptySegment("base")

      # Get access to the segment editor effect
      segmentEditorWidget = slicer.qMRMLSegmentEditorWidget()
      segmentEditorWidget.setMRMLScene(slicer.mrmlScene)
      segmentEditorNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLSegmentEditorNode")
      segmentEditorNode.SetOverwriteMode(slicer.vtkMRMLSegmentEditorNode.OverwriteNone)
      segmentEditorWidget.setMRMLSegmentEditorNode(segmentEditorNode)
      segmentEditorWidget.setSegmentationNode(segmentationNode)
      segmentEditorWidget.setSourceVolumeNode(inputVolume)
      segmentEditorWidget.setCurrentSegmentID(segmentId)

      # Thresholding
      inputVolume.GetImageData().GetScalarTypeMax()
      segmentEditorWidget.setActiveEffectByName("Threshold")
      effect = segmentEditorWidget.activeEffect()
      segmentEditorWidget.setCurrentSegmentID(segmentId)
      effect.setParameter("MinimumThreshold", int((inputVolume.GetImageData().GetScalarRange()[1]-inputVolume.GetImageData().GetScalarRange()[0]) * thresholdPercent + inputVolume.GetImageData().GetScalarRange()[0]))
      effect.setParameter("MaximumThreshold", int(inputVolume.GetImageData().GetScalarRange()[1]))
      effect.self().onApply()

      # Islands removal
      segmentEditorWidget.setActiveEffectByName("Islands")
      effect = segmentEditorWidget.activeEffect()
      effect.setParameter("Operation", "REMOVE_SMALL_ISLANDS")
      effect.setParameter("MinimumSize", minimumSize)
      effect.self().onApply()

      # Copy
      clonedSegmentId = segmentationNode.GetSegmentation().AddEmptySegment("cloned")
      segmentEditorWidget.setActiveEffectByName("Logical operators")
      effect = segmentEditorWidget.activeEffect()
      segmentEditorWidget.setCurrentSegmentID(clonedSegmentId)
      effect.setParameter("Operation", "COPY")
      effect.setParameter("ModifierSegmentID", segmentId)
      effect.self().onApply()

      # Isolate largest islands
      segmentEditorWidget.setActiveEffectByName("Islands")
      effect = segmentEditorWidget.activeEffect()
      effect.setParameter("MinimumSize", maximumSize)
      effect.self().onApply()

      # Subtract from base segment to leave only islands in size range
      segmentEditorWidget.setActiveEffectByName("Logical operators")
      effect = segmentEditorWidget.activeEffect()
      segmentEditorWidget.setCurrentSegmentID(segmentId)
      effect.setParameter("Operation", "SUBTRACT")
      effect.setParameter("ModifierSegmentID", clonedSegmentId)
      effect.self().onApply()

      segmentationNode.GetSegmentation().RemoveSegment(clonedSegmentId)
      
      # Remove all islands on the edges of the image
      if self.removeBorderIslandsCheckBox.isChecked():
        borderVoxels = True
        while borderVoxels:
          tempLabelmapVolume = slicer.mrmlScene.AddNewNodeByClass('vtkMRMLLabelMapVolumeNode', 'TempLabelMapVolume')
          slicer.modules.segmentations.logic().ExportVisibleSegmentsToLabelmapNode(segmentationNode, tempLabelmapVolume, inputVolume)
          tempLabelMapArray = slicer.util.arrayFromVolume(tempLabelmapVolume)
          slicer.mrmlScene.RemoveNode(tempLabelmapVolume)

          mask = np.zeros_like(tempLabelMapArray, dtype=bool)
          margin = int(self.borderMarginSliderWidget.value)
          # mask[0, :, :] = True # Front
          # mask[-1, :, :] = True # Back
          mask[:, 0:margin, :] = True
          mask[:, (-margin-1):-1, :] = True
          mask[:, :, 0:margin] = True
          mask[:, :, (-margin-1):-1] = True

          tempLabelMapArray[~mask] = 0
          indices = np.argwhere(tempLabelMapArray == 1)
          if len(indices) <= 0:
            borderVoxels = False
          else:
            segmentEditorWidget.setActiveEffectByName("Islands")
            effect = segmentEditorWidget.activeEffect()
            effect.setParameter("Operation", "REMOVE_SELECTED_ISLAND")
            self.removeSelectedIsland(effect, [indices[0][2], indices[0][1], indices[0][0]])

     

      # Export segmentation to label map
      self.removeNodeByName('MaskedCalibrationLabelMapVolume')
      labelMapVolumeNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLLabelMapVolumeNode", "MaskedCalibrationLabelMapVolume")
      slicer.modules.segmentations.logic().ExportVisibleSegmentsToLabelmapNode(segmentationNode, labelMapVolumeNode, inputVolume)

       # Clean up
      segmentEditorWidget.setActiveEffectByName("No editing")
      segmentEditorWidget.deleteLater()
      segmentEditorWidget = None
      slicer.mrmlScene.RemoveNode(segmentationNode)
      slicer.mrmlScene.RemoveNode(segmentEditorNode)

      # Count number of Islands and attempt repair if one is missing
      # Does not support 9 fiducial frame
      if repair:
        if not zframeConfig == 'z003':
          loopRegistration = self.countAndRepairFiducials(labelMapVolumeNode)
        else: 
          loopRegistration = False
      else:
        loopRegistration = False

    # Convert label map to scalar volume
    scalarVolumeNode = None
    self.removeNodeByName('MaskedCalibrationVolume')
    scalarVolumeNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLScalarVolumeNode", "MaskedCalibrationVolume")
    slicer.modules.volumes.logic().CreateScalarVolumeFromVolume(slicer.mrmlScene, scalarVolumeNode, labelMapVolumeNode)
    self.increaseThresholdForRepair = False
    slicer.mrmlScene.RemoveNode(labelMapVolumeNode)

    return scalarVolumeNode
  
  def cropVolume(self, volumeNode, xSize, ySize):
    imageData = volumeNode.GetImageData()
    dims = imageData.GetDimensions()

    xMargin = (dims[0] - xSize, dims[0] - xSize)
    if xMargin[0] % 2 == 1:
      xMargin = (xMargin[0]//2+1, xMargin[1]//2)
    else:
      xMargin = (xMargin[0]//2, xMargin[1]//2)
    yMargin = (dims[1] - ySize, dims[1] - ySize)
    if yMargin[0] % 2 == 1:
      yMargin = (yMargin[0]//2+1, yMargin[1]//2)
    else:
      yMargin = (yMargin[0]//2, yMargin[1]//2)

    # Create a vtkExtractVOI filter to crop the image data
    extractVoi = vtk.vtkExtractVOI()
    extractVoi.SetInputData(imageData)
    extractVoi.SetVOI(xMargin[0], dims[0] - xMargin[1] - 1, yMargin[0], dims[1] - yMargin[1] - 1, 0, dims[2] - 1)
    extractVoi.Update()
    volumeNode.SetAndObserveImageData(extractVoi.GetOutput())

  def removeSelectedIsland(self, scriptedEffect, ijk):
    # Generate merged labelmap of all visible segments
    segmentationNode = scriptedEffect.parameterSetNode().GetSegmentationNode()

    selectedSegmentLabelmap = scriptedEffect.selectedSegmentLabelmap()
    # We need to know exactly the value of the segment voxels, apply threshold to make force the selected label value
    labelValue = 1
    backgroundValue = 0
    thresh = vtk.vtkImageThreshold()
    thresh.SetInputData(selectedSegmentLabelmap)
    thresh.ThresholdByLower(0)
    thresh.SetInValue(backgroundValue)
    thresh.SetOutValue(labelValue)
    thresh.SetOutputScalarType(selectedSegmentLabelmap.GetScalarType())
    thresh.Update()

    # Create oriented image data from output
    inputLabelImage = slicer.vtkOrientedImageData()
    inputLabelImage.ShallowCopy(thresh.GetOutput())
    selectedSegmentLabelmapImageToWorldMatrix = vtk.vtkMatrix4x4()
    selectedSegmentLabelmap.GetImageToWorldMatrix(selectedSegmentLabelmapImageToWorldMatrix)
    inputLabelImage.SetImageToWorldMatrix(selectedSegmentLabelmapImageToWorldMatrix)

    pixelValue = inputLabelImage.GetScalarComponentAsFloat(ijk[0], ijk[1], ijk[2], 0)

    try:
      floodFillingFilter = vtk.vtkImageThresholdConnectivity()
      floodFillingFilter.SetInputData(inputLabelImage)
      seedPoints = vtk.vtkPoints()
      origin = inputLabelImage.GetOrigin()
      spacing = inputLabelImage.GetSpacing()
      seedPoints.InsertNextPoint(origin[0] + ijk[0] * spacing[0], origin[1] + ijk[1] * spacing[1], origin[2] + ijk[2] * spacing[2])
      floodFillingFilter.SetSeedPoints(seedPoints)
      floodFillingFilter.ThresholdBetween(pixelValue, pixelValue)

      if pixelValue != 0:  # if clicked on empty part then there is nothing to remove or keep
        floodFillingFilter.SetInValue(1)
        floodFillingFilter.SetOutValue(0)

        floodFillingFilter.Update()
        modifierLabelmap = scriptedEffect.defaultModifierLabelmap()
        modifierLabelmap.DeepCopy(floodFillingFilter.GetOutput())

        scriptedEffect.modifySelectedSegmentByLabelmap(modifierLabelmap, slicer.qSlicerSegmentEditorAbstractEffect.ModificationModeRemove)
    except IndexError:
      print("Island processing failed")

  def countAndRepairFiducials(self, labelMapVolumeNode):
    # Returns False if redoing registration with different parameters
    if labelMapVolumeNode.GetImageData().GetScalarRange()[1] == 0:
      numberOfSegments = 0
    else:
      segmentationNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLSegmentationNode")
      segmentationNode.CreateDefaultDisplayNodes()
      slicer.modules.segmentations.logic().ImportLabelmapToSegmentationNode(labelMapVolumeNode, segmentationNode)

      segmentEditorWidget = slicer.qMRMLSegmentEditorWidget()
      segmentEditorWidget.setMRMLScene(slicer.mrmlScene)
      segmentEditorNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLSegmentEditorNode")
      segmentEditorNode.SetOverwriteMode(slicer.vtkMRMLSegmentEditorNode.OverwriteNone)
      segmentEditorWidget.setMRMLSegmentEditorNode(segmentEditorNode)
      segmentEditorWidget.setSegmentationNode(segmentationNode)
      segmentEditorWidget.setSourceVolumeNode(labelMapVolumeNode)

      segmentEditorWidget.setActiveEffectByName("Islands")
      effect = segmentEditorWidget.activeEffect()
      effect.setParameter("Operation", "SPLIT_ISLANDS_TO_SEGMENTS")
      effect.setParameter("MinimumSize", 0)
      effect.self().onApply()

      numberOfSegments = segmentationNode.GetSegmentation().GetNumberOfSegments()

      # Cleanup
      segmentEditorWidget.setActiveEffectByName("No editing")
      slicer.mrmlScene.RemoveNode(segmentationNode)
      slicer.mrmlScene.RemoveNode(segmentEditorNode)
      segmentEditorWidget.deleteLater()
      segmentEditorWidget = None

    # Attempt repair
    result = ""
    print(f'Segments detected: {numberOfSegments}')

    if numberOfSegments > 0:
      # Determine if the image is salvageable
      # Isolate middle slice
      imageData = labelMapVolumeNode.GetImageData()
      centroid = self.findCentroidOfVolume(labelMapVolumeNode)
      middleSlice = int(centroid[2])
      dims = imageData.GetDimensions()
      numpy_array = vtk.util.numpy_support.vtk_to_numpy(imageData.GetPointData().GetScalars())
      numpy_array = numpy_array.reshape(dims[2], dims[1], dims[0])
      numpy_array = numpy_array.transpose(2,1,0)
      slice_array = numpy_array[:, :, middleSlice]

      # Calculate bounding box to see if the dimensions are about right and that there are only 1-2 missing fiducials
      rangeNumber = 15 # how much wiggle room to allow for bounding box
      leftColumn, rightColumn, topRow, bottomRow = self.calculateBoundingBox(slice_array)
      width = (rightColumn - leftColumn) * labelMapVolumeNode.GetSpacing()[0]
      height = (bottomRow - topRow) * labelMapVolumeNode.GetSpacing()[1]
      # print(f'leftColumn {leftColumn}')
      # print(f'rightColumn {rightColumn}')
      # print(f'topRow {topRow}')
      # print(f'bottomRow {bottomRow}')
      expectedWidth = abs(self.frameTopology[0][0] - self.frameTopology[2][0])
      expectedHeight = abs(self.frameTopology[0][1] - self.frameTopology[2][1])
      widthCorrect = (expectedWidth - rangeNumber) <= width <= (expectedWidth + rangeNumber)
      heightCorrect = (expectedWidth - rangeNumber) <= width <= (expectedWidth + rangeNumber)

      if numberOfSegments == 7:
        if widthCorrect and heightCorrect:
          print("7 fiducials detected; bounding box dimensions correct")
          return False
        else:
          print("7 fiducials detected; however, bounding box dimensions are wrong")
      if (numberOfSegments == 6 or numberOfSegments == 5) and widthCorrect and heightCorrect:
        print("5-6 fiducials detected; bounding box dimensions correct; trying to repair")
        result = self.repairMissingFiducial(slice_array, numpy_array, leftColumn, rightColumn, topRow, bottomRow, middleSlice, labelMapVolumeNode)
        if result == "success":
          return False
    if not self.increaseThresholdForRepair:
      if not (self.thresholdSliderWidget.value <= self.thresholdSliderWidget.minimum):
        self.thresholdSliderWidget.value = self.thresholdSliderWidget.value - 0.01
        print(f'Less than 6 or more than 7 fiducials detected; decreasing threshold percentage to {self.thresholdSliderWidget.value}')
        return True
      else:
        self.increaseThresholdForRepair = True
        self.thresholdSliderWidget.value = self.defaultThresholdPercentage + 0.02
        print(f'Switching to increasing to {self.thresholdSliderWidget.value}')
        return True
    else:
      # Try again with the assumption that the thresholding was too lenient
      if not (self.thresholdSliderWidget.value >= (self.thresholdSliderWidget.maximum/5)):
        self.thresholdSliderWidget.value = self.thresholdSliderWidget.value + 0.02
        print(f'Less than 6 or more than 7 fiducials detected; increasing threshold percentage to {self.thresholdSliderWidget.value}')
        return True
      else:
        print("Fiducial repair failed")
        self.thresholdSliderWidget.value = self.defaultThresholdPercentage
        return False

  def calculateBoundingBox(self, slice_array):
    leftColumn = 0
    rightColumn = slice_array.shape[0]
    topRow = 0
    bottomRow = slice_array.shape[1]
    for row in range(0, slice_array.shape[0]):
      if np.any(slice_array[row,:] > 0):
        leftColumn = row
        break
    for row in range(slice_array.shape[0]-1, -1, -1):
      if np.any(slice_array[row,:] > 0):
        rightColumn = row
        break
    for column in range(0, slice_array.shape[1]):
      if np.any(slice_array[:,column] > 0):
        topRow = column
        break
    for column in range(slice_array.shape[1]-1, -1, -1):
      if np.any(slice_array[:,column] > 0):
        bottomRow = column
        break
    return leftColumn, rightColumn, topRow, bottomRow

  def repairMissingFiducial(self, slice_array, numpy_array, leftColumn, rightColumn, topRow, bottomRow, middleSlice, labelMapVolumeNode):
    # Identify missing fiducial
    # Shrink margins of image
    cropped = slice_array[leftColumn:rightColumn, topRow:bottomRow]

    # Probe array for values to look for missing value
    missingFiducial = 0
    r = 10
    thickness = 8
    adjust = 4
    length = 4
    diagLength = int(math.sqrt(length**2 + length**2))
    # Corners
    # Top Left
    if not np.any(cropped[0:r, 0:r] > 0):
      print("Attempting repair of top left fiducial")
      startLine = (leftColumn + adjust, topRow + adjust, middleSlice - (length//2))
      endLine = (leftColumn + adjust, topRow + adjust, middleSlice + (length//2))
      numpy_array = self.drawThickLine(startLine, endLine, thickness, numpy_array)
      missingFiducial += 1
    # Top Right
    if not np.any(cropped[cropped.shape[0]-r:cropped.shape[0],0:r] > 0):
      print("Attempting repair of top right fiducial")
      startLine = (leftColumn + cropped.shape[0] - adjust, topRow + adjust, middleSlice - (length//2))
      endLine = (leftColumn + cropped.shape[0] - adjust, topRow + adjust, middleSlice + (length//2))
      numpy_array = self.drawThickLine(startLine, endLine, thickness, numpy_array)
      missingFiducial += 1
    # Bottom Left
    if not np.any(cropped[0:r,cropped.shape[1]-r:cropped.shape[1]] > 0):
      print("Attempting repair of bottom left fiducial")
      startLine = (leftColumn + adjust, topRow + cropped.shape[1] - adjust, middleSlice - (length//2))
      endLine = (leftColumn + adjust, topRow + cropped.shape[1] - adjust, middleSlice + (length//2))
      numpy_array = self.drawThickLine(startLine, endLine, thickness, numpy_array)
      missingFiducial += 1
    # Bottom Right
    if not np.any(cropped[cropped.shape[0]-r:cropped.shape[0],cropped.shape[1]-r:cropped.shape[1]] > 0):
      print("Attempting repair of bottom right fiducial")
      startLine = (leftColumn + cropped.shape[0] - adjust, topRow + cropped.shape[1] - adjust, middleSlice - (length//2))
      endLine = (leftColumn + cropped.shape[0] - adjust, topRow + cropped.shape[1] - adjust, middleSlice + (length//2))
      numpy_array = self.drawThickLine(startLine, endLine, thickness, numpy_array)
      missingFiducial += 1

    # Sides
    # Middle Left
    if not np.any(cropped[0:r, cropped.shape[1]//2-r//2:cropped.shape[1]//2+r//2] > 0):
      print("Attempting repair of middle left fiducial")
      startLine = (leftColumn + adjust, topRow + cropped.shape[1]//2 - (diagLength//2), middleSlice - (diagLength//2))
      endLine = (leftColumn + adjust, topRow + cropped.shape[1]//2 + (diagLength//2), middleSlice + (diagLength//2))
      numpy_array = self.drawThickLine(startLine, endLine, thickness, numpy_array)
      missingFiducial += 1
    # Middle Top  
    if not np.any(cropped[cropped.shape[0]//2-r//2:cropped.shape[0]//2+r//2, 0:r] > 0):
      print("Attempting repair of middle top fiducial")
      startLine = (leftColumn + cropped.shape[0]//2 + (diagLength//2), topRow + adjust, middleSlice - (diagLength//2))
      endLine = (leftColumn + cropped.shape[0]//2 - (diagLength//2), topRow + adjust, middleSlice + (diagLength//2))
      numpy_array = self.drawThickLine(startLine, endLine, thickness, numpy_array)
      missingFiducial += 1
    # Middle Right  
    if not np.any(cropped[cropped.shape[0]-r:cropped.shape[0], cropped.shape[1]//2-r//2:cropped.shape[1]//2+r//2] > 0):
      print("Attempting repair of middle right fiducial")
      startLine = (leftColumn + cropped.shape[0] - adjust, topRow + cropped.shape[1]//2 + (diagLength//2), middleSlice - (diagLength//2))
      endLine = (leftColumn + cropped.shape[0] - adjust, topRow +  cropped.shape[1]//2 - (diagLength//2), middleSlice + (diagLength//2))
      numpy_array = self.drawThickLine(startLine, endLine, thickness, numpy_array)
      missingFiducial += 1

    if 2 >= missingFiducial >= 1:
      imageData = self.numpy_to_vtk_image_data(numpy_array)
      labelMapVolumeNode.SetAndObserveImageData(imageData)
      return "success"
    else:
      return "anomaly"

  def drawThickLine(self, start, end, thickness, numpy_array):
    for dx in range(-thickness//2, thickness//2+1):
      for dy in range(-thickness//2, thickness//2+1):
          for dz in range(-thickness//2, thickness//2+1):
            rr, cc, zz = line_nd([start[0]+dx, start[1]+dy, start[2]+dz], [end[0]+dx, end[1]+dy, end[2]+dz], endpoint=True)
            numpy_array[rr, cc, zz] = 1
    return numpy_array

  def findCentroidOfVolume(self, inputVolume):
    imageData = inputVolume.GetImageData()
    dimensions = imageData.GetDimensions()
    spacing = imageData.GetSpacing()
    origin = imageData.GetOrigin()
    # Convert vtkImageData to numpy array
    voxels = vtk.util.numpy_support.vtk_to_numpy(imageData.GetPointData().GetScalars())
    voxels = voxels.reshape(dimensions, order='F')
    # Calculate the center of mass
    indices = np.indices(dimensions).reshape(3, -1)
    center_of_mass = np.average(indices, axis=1, weights=voxels.ravel())
    return center_of_mass
    
  def onIdentity(self):
    if self.ZFrameCalibrationTransformNode:
      identityMatrix = vtk.vtkMatrix4x4()
      identityMatrix.Identity()
      self.ZFrameCalibrationTransformNode.SetMatrixTransformToParent(identityMatrix)

  # ------------------------------------- Planning -----------------------------------
  
  def onAddTarget(self):
    if not self.validRegistration:
      return
    
    if self.biopsyFiducialListNode:
      fiducialListNode = self.biopsyFiducialListNode
    else:
      self.removeNodeByName("Target")
      self.biopsyFiducialListNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLMarkupsFiducialNode", "Target")
      self.biopsyFiducialListNode.AddObserver(slicer.vtkMRMLMarkupsNode.PointPositionDefinedEvent, self.onTargetAdded)
      self.biopsyFiducialListNode.AddObserver(slicer.vtkMRMLMarkupsNode.PointModifiedEvent, self.onTargetMoved)
      fiducialListNode = self.biopsyFiducialListNode

    slicer.modules.markups.logic().StartPlaceMode(False)

    # Set the current node in the selection node
    selectionNode = slicer.mrmlScene.GetNodeByID("vtkMRMLSelectionNodeSingleton")
    selectionNode.SetReferenceActivePlaceNodeClassName(fiducialListNode.GetClassName())
    selectionNode.SetActivePlaceNodeID(fiducialListNode.GetID())
  
  def onTargetMoved(self, caller, event):
    # Cannot identify which changed, so change all of them
    numFiducials = caller.GetNumberOfControlPoints()

    # Grab old list of fiducials
    oldTargetRASList = []
    for i in range(self.targetListTableWidget.rowCount):
      oldTarget = [float(i) for i in ast.literal_eval(self.targetListTableWidget.item(i,3).text())]
      oldTargetRASList.append(oldTarget)

    newTargetRASList = []
    for i in range(numFiducials):
      # Need to recalculate all the points because of the many ways the fiducial list could've been modified
      # if the list item exists
      if self.targetListTableWidget.item(i, 0):
        targetName, targetGrid, targetDepth, targetRAS = self.calculateGridCoordinates(i)
        self.targetListTableWidget.item(i, 0).setText(targetName)
        self.targetListTableWidget.item(i, 1).setText(targetGrid)
        self.targetListTableWidget.item(i, 2).setText(f'{targetDepth:.2f} cm')
        targetRAS_string = f'[{targetRAS[0]:.2f}, {targetRAS[1]:.2f}, {targetRAS[2]:.2f}]'
        self.targetListTableWidget.item(i, 3).setText(targetRAS_string)
        newTargetRASList.append([float(i) for i in ast.literal_eval(targetRAS_string)])
    
    if self.targetListTableWidget.rowCount == numFiducials:
      for i in range(self.targetListTableWidget.rowCount):
        if newTargetRASList[i] != oldTargetRASList[i]:
          lm = slicer.app.layoutManager()
          for slice in ['Yellow', 'Green', 'Red']:
            sliceNode = lm.sliceWidget(slice).mrmlSliceNode()
            sliceNode.JumpSliceByOffsetting(newTargetRASList[i][0], newTargetRASList[i][1], newTargetRASList[i][2]) 

  
  def onTargetAdded(self, caller, event):
    newFiducialIndex = self.biopsyFiducialListNode.GetNumberOfControlPoints() - 1

    rowCount = self.targetListTableWidget.rowCount
    if not (rowCount == newFiducialIndex):
      print("Target list count does not match table row count")
      return
    self.targetListTableWidget.insertRow(newFiducialIndex)

    targetName, targetGrid, targetDepth, targetRAS = self.calculateGridCoordinates(newFiducialIndex)
    
    targetListFont = qt.QFont()
    targetListFont.setPointSize(18)
    targetListFont.setBold(False)

    # Update table with Target name item
    targetItem = qt.QTableWidgetItem(targetName)
    targetItem.setTextAlignment(qt.Qt.AlignVCenter | qt.Qt.AlignHCenter)
    targetItem.setFlags(targetItem.flags() | qt.Qt.ItemIsEditable)
    targetItem.setFont(targetListFont)
    self.targetListTableWidget.setItem(newFiducialIndex,  0, targetItem)

    # Update table with Grid coordinate
    gridItem = qt.QTableWidgetItem(targetGrid)
    gridItem.setTextAlignment(qt.Qt.AlignVCenter | qt.Qt.AlignHCenter)
    gridItem.setFlags(gridItem.flags() & ~qt.Qt.ItemIsEditable)
    gridItem.setFont(targetListFont)
    self.targetListTableWidget.setItem(newFiducialIndex,  1, gridItem)

    # Update table with Depth
    depthItem = qt.QTableWidgetItem(f'{targetDepth:.2f} cm')
    depthItem.setTextAlignment(qt.Qt.AlignVCenter | qt.Qt.AlignHCenter)
    depthItem.setFlags(depthItem.flags() & ~qt.Qt.ItemIsEditable)
    depthItem.setFont(targetListFont)
    self.targetListTableWidget.setItem(newFiducialIndex,  2, depthItem)

    # Update table with RAS coordinates
    rasItem = qt.QTableWidgetItem(f'[{targetRAS[0]:.2f}, {targetRAS[1]:.2f}, {targetRAS[2]:.2f}]')
    rasItem.setTextAlignment(qt.Qt.AlignVCenter | qt.Qt.AlignHCenter)
    rasItem.setFlags(rasItem.flags() & ~qt.Qt.ItemIsEditable)
    self.targetListTableWidget.setItem(newFiducialIndex, 3, rasItem)

    # Update table with Delete button
    # Delete by deleting the index in the fiducial list that matches the table item, easy
    deleteItem = qt.QPushButton()
    moduleDir = os.path.dirname(slicer.util.modulePath(self.__module__))
    deleteIconPath = os.path.join(moduleDir, 'Resources/Icons', 'MarkupsDelete.png')
    deleteIcon = qt.QIcon(deleteIconPath)
    deleteItem.setIcon(deleteIcon)
    deleteItem.connect('clicked()', lambda: self.onDeleteTarget(deleteItem))
    self.targetListTableWidget.setCellWidget(newFiducialIndex, 4, deleteItem)

  def onDeleteTarget(self, button):
    for index in range(self.targetListTableWidget.rowCount):
      if self.targetListTableWidget.cellWidget(index, 4) == button:
        self.biopsyFiducialListNode.RemoveNthControlPoint(index)
        self.targetListTableWidget.removeRow(index)

        # Delete trajectory model
        shNode = slicer.mrmlScene.GetSubjectHierarchyNode()
        sceneItemID = shNode.GetSceneItemID()
        trajectoryFolderItemID = shNode.GetItemChildWithName(sceneItemID, "TrajectoryModels")
        trajectoryChildren = vtk.vtkIdList()
        shNode.GetItemChildren(trajectoryFolderItemID, trajectoryChildren)
        shNode.RemoveItem(shNode.GetItemByPositionUnderParent(trajectoryFolderItemID,index))
        return

  def calculateGridCoordinates(self, index):
    fiducialNode = self.biopsyFiducialListNode

    # Target Name
    targetName = self.biopsyFiducialListNode.GetNthControlPointLabel(index)

    # Target RAS
    targetRAS = [0, 0, 0]
    fiducialNode.GetNthControlPointPosition(index, targetRAS)

    # Target Grid
    # TODO: Add a check for when it's outside of the target volume? But maybe not; sometimes done intentionally
    targetGrid = None
    closestHole = None
    inverseMatrix = vtk.vtkMatrix4x4()
    self.ZFrameCalibrationTransformNode.GetMatrixTransformToParent(inverseMatrix)
    inverseMatrix.Invert()
    targetIJK = inverseMatrix.MultiplyPoint(targetRAS + [1])[:3]
    min_distance = math.inf
    for x, horizontalLabel in enumerate(self.templateHorizontalLabels):  # number of holes horizontally
      for y, verticalLabel in enumerate(self.templateVerticalLabels):  # number of holes vertically
        holeCenter = [self.templateOrigin[0] - (x * self.templateHorizontalOffset), self.templateOrigin[1] - (y * self.templateVerticalOffset)]
        distance = math.sqrt((holeCenter[0] - targetIJK[0]) ** 2 + (holeCenter[1] - targetIJK[1]) ** 2)
        if distance < min_distance:
          min_distance = distance
          closestHole = holeCenter
          targetGrid = f'{horizontalLabel}, {verticalLabel}'

    # Target Depth (cm)
    targetDepth = abs(targetIJK[2] - self.templateOrigin[2]) / 10

    # Update trajectory model
    shNode = slicer.mrmlScene.GetSubjectHierarchyNode()
    sceneItemID = shNode.GetSceneItemID()
    trajectoryFolderItemID = shNode.GetItemChildWithName(sceneItemID, "TrajectoryModels")
    if trajectoryFolderItemID == 0:
      trajectoryFolderItemID = shNode.CreateFolderItem(sceneItemID , "TrajectoryModels")
    trajectoryChildren = vtk.vtkIdList()
    shNode.GetItemChildren(trajectoryFolderItemID, trajectoryChildren)

    numModelsToAdd = fiducialNode.GetNumberOfControlPoints() - trajectoryChildren.GetNumberOfIds()
    modelName = f'{targetName}_Trajectory'
    if numModelsToAdd > 0:
      modelNode = slicer.mrmlScene.AddNewNodeByClass('vtkMRMLModelNode', modelName)
      modelNode.CreateDefaultDisplayNodes()
      modelNode.SetAndObserveTransformNodeID(self.ZFrameCalibrationTransformNode.GetID())
      shNode.CreateItem(trajectoryFolderItemID, modelNode)
    childModelNode = shNode.GetItemDataNode(shNode.GetItemByPositionUnderParent(trajectoryFolderItemID, index))
    childModelNode.SetName(modelName)

    cylinder = vtk.vtkCylinderSource()
    cylinder.SetRadius(1.5)
    cylinder.SetHeight(250)
    cylinder.SetResolution(72)
    cylinder.Update()

    modelMatrix = vtk.vtkMatrix4x4()
    modelMatrix.Identity()
    modelMatrix.SetElement(0, 0, 1)
    modelMatrix.SetElement(1, 1, 0)
    modelMatrix.SetElement(2, 2, 0)
    modelMatrix.SetElement(1, 2, 1)
    modelMatrix.SetElement(2, 1, -1)
    modelMatrix.SetElement(0, 3, closestHole[0])
    modelMatrix.SetElement(1, 3, closestHole[1])
    modelMatrix.SetElement(2, 3, self.templateOrigin[2] + 125)
    modelTransform = vtk.vtkTransform()
    modelTransform.SetMatrix(modelMatrix)
    modelTransformFilter = vtk.vtkTransformPolyDataFilter()
    modelTransformFilter.SetInputConnection(cylinder.GetOutputPort())
    modelTransformFilter.SetTransform(modelTransform)
    modelTransformFilter.Update()

    childModelNode.SetAndObservePolyData(modelTransformFilter.GetOutput())
    childModelNode.GetDisplayNode().SetVisibility2D(True)
    childModelNode.GetDisplayNode().SetSliceIntersectionOpacity(0.6)
    childModelNode.GetDisplayNode().SetSliceIntersectionThickness(2)
    childModelNode.GetDisplayNode().SetColor(1,0,1)
    
    return targetName, targetGrid, targetDepth, targetRAS
  
  def onTargetListItemChanged(self, tableItem):
    # If target name item
    if tableItem.column() != 0:
      return
    self.biopsyFiducialListNode.SetNthControlPointLabel(tableItem.row(), tableItem.text())

  def onGenerateWorksheet(self):
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from PyPDF2 import PdfWriter, PdfReader
    from io import BytesIO

    numberOfSheets = math.ceil(self.targetListTableWidget.rowCount / 2)
    currentFilePath = os.path.dirname(slicer.util.modulePath(self.__module__))
    blankWorksheetPath = os.path.join(currentFilePath, "Resources", "Templates", self.templateWorksheetPath)
    blankOverlayWorksheetPath = os.path.join(currentFilePath, "Resources", "Templates", self.templateWorksheetOverlayPath)
    newWorksheetPath = f'{self.caseDirPath}/BiopsyWorksheet_{os.path.basename(os.path.normpath(self.caseDirPath))}.pdf'
    newWorksheetOverlayPath =f'{self.caseDirPath}/BiopsyWorksheetOverlay_{os.path.basename(os.path.normpath(self.caseDirPath))}.pdf'
    #newWorksheetPath = f'C:/w/data/ProstateBiopsyModuleTest/RegTest/BiopsyWorksheet_test.pdf'
    #newWorksheetOverlayPath = f'C:/w/data/ProstateBiopsyModuleTest/RegTest/BiopsyWorksheetOverlay_test.pdf'
    with open(blankWorksheetPath, "rb") as f, open(blankOverlayWorksheetPath, "rb") as f_overlay:
      reader = PdfReader(f)
      reader_overlay = PdfReader(f_overlay)
      writer = PdfWriter()
      writer_overlay = PdfWriter()
    
      for sheetIndex in range(numberOfSheets):
        page = reader.pages[0]
        page_overlay = reader_overlay.pages[0]
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)

        rowIndex = sheetIndex * 2
        targetName = self.targetListTableWidget.item(rowIndex, 0).text()
        targetGrid = self.targetListTableWidget.item(rowIndex, 1).text()
        targetDepth = self.targetListTableWidget.item(rowIndex, 2).text()
        
        #writer.update_page_form_field_values(page, {"TARGET_1": targetName})
        writer.update_page_form_field_values(page, {"GRID_1": targetGrid})
        writer.update_page_form_field_values(page, {"DEPTH_1": targetDepth})
        #writer_overlay.update_page_form_field_values(page_overlay, {"TARGET_1": targetName})
        writer_overlay.update_page_form_field_values(page_overlay, {"GRID_1": targetGrid})
        writer_overlay.update_page_form_field_values(page_overlay, {"DEPTH_1": targetDepth})

        gridMatches = re.findall(r'(\w+)', targetGrid)
        worksheetHole_1 = [self.worksheetOrigin[0][0] + (self.templateHorizontalLabels.index(gridMatches[0]) * self.worksheetHorizontalOffset), self.worksheetOrigin[0][1] - (self.templateVerticalLabels.index(gridMatches[1]) * self.worksheetVerticalOffset)]

        # Draw a cross
        can.line(worksheetHole_1[0]-7, worksheetHole_1[1]+7, worksheetHole_1[0]+7, worksheetHole_1[1]-7)
        can.line(worksheetHole_1[0]-7, worksheetHole_1[1]-7, worksheetHole_1[0]+7, worksheetHole_1[1]+7)
        
        rowIndex += 1
        if rowIndex < self.targetListTableWidget.rowCount:
          targetName = self.targetListTableWidget.item(rowIndex, 0).text()
          targetGrid = self.targetListTableWidget.item(rowIndex, 1).text()
          targetDepth = self.targetListTableWidget.item(rowIndex, 2).text()

          #writer.update_page_form_field_values(page, {"TARGET_2": targetName})
          writer.update_page_form_field_values(page, {"GRID_2": targetGrid})
          writer.update_page_form_field_values(page, {"DEPTH_2": targetDepth})
          #writer_overlay.update_page_form_field_values(page_overlay, {"TARGET_2": targetName})
          writer_overlay.update_page_form_field_values(page_overlay, {"GRID_2": targetGrid})
          writer_overlay.update_page_form_field_values(page_overlay, {"DEPTH_2": targetDepth})

          gridMatches = re.findall(r'(\w+)', targetGrid)
          worksheetHole_2 = [self.worksheetOrigin[1][0] + (self.templateHorizontalLabels.index(gridMatches[0]) * self.worksheetHorizontalOffset), self.worksheetOrigin[1][1] - (self.templateVerticalLabels.index(gridMatches[1]) * self.worksheetVerticalOffset)]

          # Draw a cross
          can.line(worksheetHole_2[0]-7, worksheetHole_2[1]+7, worksheetHole_2[0]+7, worksheetHole_2[1]-7)
          can.line(worksheetHole_2[0]-7, worksheetHole_2[1]-7, worksheetHole_2[0]+7, worksheetHole_2[1]+7)
        
        can.save()
        packet.seek(0)
        draw_pdf = PdfReader(packet)
        page.merge_page(draw_pdf.pages[0])
        page_overlay.merge_page(draw_pdf.pages[0])
        writer.add_page(page)
        writer_overlay.add_page(page_overlay)
    
    with open(newWorksheetPath, "wb") as outputStream, open(newWorksheetOverlayPath, "wb") as outputStream_overlay:
      writer.write(outputStream)
      writer_overlay.write(outputStream_overlay)
    
    os.startfile(newWorksheetPath)

  def addRuler(self):
    biopsyRulerListNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLMarkupsLineNode", "Ruler")
    slicer.modules.markups.logic().StartPlaceMode(False)
    selectionNode = slicer.mrmlScene.GetNodeByID("vtkMRMLSelectionNodeSingleton")
    selectionNode.SetReferenceActivePlaceNodeClassName(biopsyRulerListNode.GetClassName())
    selectionNode.SetActivePlaceNodeID(biopsyRulerListNode.GetID())

  def toggleGuideHoles(self):
    if self.toggleGuideButton.isChecked():
      if self.guideHolesModelNode:
        self.guideHolesModelNode.GetDisplayNode().SetVisibility2D(True)
    else:
      if self.guideHolesModelNode:
        self.guideHolesModelNode.GetDisplayNode().SetVisibility2D(False)