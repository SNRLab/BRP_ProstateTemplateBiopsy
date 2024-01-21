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
import re
import numpy as np
# pip_install('scikit-image')
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
    self.ignoredVolumeNames = ['MaskedCalibrationVolume', 'MaskedCalibrationLabelMapVolume']
    self.imageRoles = ['N/A', 'CALIBRATION', 'PLANNING', 'CONFIRMATION']
    self.zFrameModelNode = None
    self.removeNodeByName('ZFrameTransform')
    self.ZFrameCalibrationTransformNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLLinearTransformNode", "ZFrameTransform")

  def setup(self):
    ScriptedLoadableModuleWidget.setup(self)

    # ------------------------------------ Initialization UI ---------------------------------------
    # TODO: 
    # - Change port to 104 [Port can only be changed in DICOM module UI under Query and Retrieve]
    # - Set Default values using a config file
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
    initializeLayout.addRow(self.initializeButton)

    self.caseDirLabel = qt.QLabel()
    self.caseDirLabel.text = "Waiting to Initialize Case"
    initializeLayout.addRow("Case Directory: ", self.caseDirLabel)

    self.closeCaseButton = qt.QPushButton("Close Case")
    self.closeCaseButton.toolTip = "Close Case"
    self.closeCaseButton.enabled = False
    self.closeCaseButton.connect('clicked()', self.closeCase)
    initializeLayout.addRow(self.closeCaseButton)

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
    imageListCollapsibleButton = ctk.ctkCollapsibleButton()
    imageListCollapsibleButton.text = "Images"
    self.layout.addWidget(imageListCollapsibleButton)
    imageListLayout = qt.QVBoxLayout(imageListCollapsibleButton)

    # Image List Table with combo boxes to set role for images
    self.imageListTableWidget = qt.QTableWidget(0, 3)
    self.imageListTableWidget.setHorizontalHeaderLabels(["Image Description", "Acquisition Time", "Role"])
    self.imageListTableWidget.horizontalHeader().setSectionResizeMode(qt.QHeaderView.Stretch)
    imageListLayout.addWidget(self.imageListTableWidget)
    self.imageListTableWidget.setSizePolicy(qt.QSizePolicy.MinimumExpanding, qt.QSizePolicy.Minimum)
    # -------------------------------------- ----------  --------------------------------------

    # ---------------------------------- Registration UI --------------------------------------
    # TODO: 
    # - Handle a missing fiducial (paint it in? modify registration code?)
    # - Sometimes Transform sliders have a bug that cause one slider to change another when clicking directly towards it; couldn't reproduce but keep in mind
    registrationCollapsibleButton = ctk.ctkCollapsibleButton()
    registrationCollapsibleButton.text = "Registration"
    self.layout.addWidget(registrationCollapsibleButton)
    registrationLayout = qt.QFormLayout(registrationCollapsibleButton)

    registerFont = qt.QFont()
    registerFont.setPointSize(18)
    registerFont.setBold(False)
    self.registrationButton = qt.QPushButton("Register")
    self.registrationButton.setFont(registerFont)
    self.registrationButton.toolTip = "Start registration process for Z-Frame"
    self.registrationButton.enabled = True
    self.registrationButton.connect('clicked()', self.registerZFrame)
    registrationLayout.addRow(self.registrationButton)

    registrationParametersGroupBox = ctk.ctkCollapsibleGroupBox()
    registrationParametersGroupBox.title = "Automatic Registration Parameters"
    registrationParametersGroupBox.collapsed = True
    registrationParametersLayout = qt.QFormLayout(registrationParametersGroupBox)
    registrationLayout.addWidget(registrationParametersGroupBox)

    self.configFileSelectionBox = qt.QComboBox()
    self.configFileSelectionBox.addItems(['Z-frame z001', 'Z-frame z002', 'Z-frame z003', 'Z-frame z004'])
    self.configFileSelectionBox.setCurrentIndex(3)
    registrationParametersLayout.addRow('ZFrame Configuration:', self.configFileSelectionBox)

    self.thresholdSliderWidget = ctk.ctkSliderWidget()
    self.thresholdSliderWidget.setToolTip("Set range for threshold percentage for isolating registration fiducial markers")
    self.thresholdSliderWidget.setDecimals(2)
    self.thresholdSliderWidget.minimum = 0.00
    self.thresholdSliderWidget.maximum = 1.00
    self.thresholdSliderWidget.singleStep = 0.01
    self.thresholdSliderWidget.value = 0.06
    registrationParametersLayout.addRow("Threshold Percentage:", self.thresholdSliderWidget)

    self.fiducialSizeSliderWidget = ctk.ctkRangeWidget()
    self.fiducialSizeSliderWidget.setToolTip("Set range for fiducial size for isolating registration fiducial markers")
    self.fiducialSizeSliderWidget.setDecimals(0)
    self.fiducialSizeSliderWidget.maximum = 5000
    self.fiducialSizeSliderWidget.minimum = 0
    self.fiducialSizeSliderWidget.singleStep = 1
    self.fiducialSizeSliderWidget.maximumValue = 1000
    self.fiducialSizeSliderWidget.minimumValue = 300
    registrationParametersLayout.addRow("Fiducial Size Range:", self.fiducialSizeSliderWidget)

    manualRegistrationGroupBox = ctk.ctkCollapsibleGroupBox()
    manualRegistrationGroupBox.title = "Manual Registration"
    manualRegistrationGroupBox.collapsed = False
    registrationParametersLayout = qt.QVBoxLayout(manualRegistrationGroupBox)
    registrationLayout.addWidget(manualRegistrationGroupBox)

    self.manualRegistrationTransformSliders = slicer.qMRMLTransformSliders()
    self.manualRegistrationTransformSliders.setWindowTitle("Translation")
    self.manualRegistrationTransformSliders.TypeOfTransform = slicer.qMRMLTransformSliders.TRANSLATION
    self.manualRegistrationTransformSliders.setDecimals(3)
    registrationParametersLayout.addWidget(self.manualRegistrationTransformSliders)

    self.manualRegistrationRotationSliders = slicer.qMRMLTransformSliders()
    self.manualRegistrationTransformSliders.setWindowTitle("Rotation")
    self.manualRegistrationRotationSliders.TypeOfTransform = slicer.qMRMLTransformSliders.ROTATION
    self.manualRegistrationRotationSliders.setDecimals(3)
    registrationParametersLayout.addWidget(self.manualRegistrationRotationSliders)

    self.manualRegistrationTransformSliders.setMRMLTransformNode(self.ZFrameCalibrationTransformNode)
    self.manualRegistrationRotationSliders.setMRMLTransformNode(self.ZFrameCalibrationTransformNode)
    # ------------------------------------- ----------  --------------------------------------
    
    # ------------------------------------- Planning UI --------------------------------------
    planningCollapsibleButton = ctk.ctkCollapsibleButton()
    planningCollapsibleButton.text = "Planning"
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
    if (newNode.GetName() in self.ignoredVolumeNames):
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

  def getNodeFromImageRole(self, imageRole):
    for index in range(0, self.imageListTableWidget.rowCount):
      if self.imageListTableWidget.cellWidget(index, 2).currentText == imageRole:
        return slicer.mrmlScene.GetFirstNodeByName(self.imageListTableWidget.item(index, 0).text())

  def registerZFrame(self):
    # If there is a zFrame image selected, perform the calibration step to calculate the CLB matrix
    #inputVolume = self.tempTemplateSelector.currentNode()
    inputVolume = self.getNodeFromImageRole("CALIBRATION")

    if not inputVolume:
      return
    
    outputTransform = self.ZFrameCalibrationTransformNode
    currentFilePath = os.path.dirname(os.path.realpath(__file__))
    if self.configFileSelectionBox.currentText == "Z-frame z001":
      ZFRAME_MODEL_PATH = 'zframe001-model.vtk'
      zframeConfig = 'z001'
      zframeConfigFilePath = os.path.join(currentFilePath, "Resources/zframe/zframe001.txt")
    elif self.configFileSelectionBox.currentText == "Z-frame z002":
      ZFRAME_MODEL_PATH = 'zframe002-model.vtk'
      zframeConfig = 'z002'
      zframeConfigFilePath = os.path.join(currentFilePath, "Resources/zframe/zframe002.txt")
    elif self.configFileSelectionBox.currentText == "Z-frame z003":
      ZFRAME_MODEL_PATH = 'zframe003-model.vtk'
      zframeConfig = 'z003'
      zframeConfigFilePath = os.path.join(currentFilePath, "Resources/zframe/zframe003.txt")
    else: #self.configFileSelectionBox.currentText == "Z-frame z003":
      ZFRAME_MODEL_PATH = 'zframe004-model.vtk'
      zframeConfig = 'z004'
      zframeConfigFilePath = os.path.join(currentFilePath, "Resources/zframe/zframe004.txt")
    with open(zframeConfigFilePath,"r") as f:
      configFileLines = f.readlines()

    # Parse zFrame configuration file here to identify the dimensions and topology of the zframe
    # Save the origins and diagonal vectors of each of the 3 sides of the zframe in a 2D array
    frameTopology = []
    for line in configFileLines:
      if line.startswith('Side 1') or line.startswith('Side 2'): 
        vec = [float(s) for s in re.findall(r'-?\d+\.?\d*', line)]
        vec.pop(0)
        frameTopology.append(vec)
      elif line.startswith('Base'):
        vec = [float(s) for s in re.findall(r'-?\d+\.?\d*', line)]
        frameTopology.append(vec)
    # Convert frameTopology points to a string, for the sake of passing it as a string argument to the ZframeRegistration CLI 
    frameTopologyString = ' '.join([str(elem) for elem in frameTopology])

    self.removeNodeByName('ZFrameModel')
    self.loadZFrameModel(ZFRAME_MODEL_PATH,'ZFrameModel')
    zFrameMaskedVolume = self.createMaskedVolumeBySize(inputVolume, self.thresholdSliderWidget.value, self.fiducialSizeSliderWidget.minimumValue, self.fiducialSizeSliderWidget.maximumValue)
    centerOfMassSlice = int(self.findCentroidOfVolume(zFrameMaskedVolume)[2])
    # Run zFrameRegistration CLI module
    params = {'inputVolume': zFrameMaskedVolume, 'startSlice': centerOfMassSlice-3, 'endSlice': centerOfMassSlice+3,
              'outputTransform': outputTransform, 'zframeConfig': zframeConfig, 'frameTopology': frameTopologyString, 
              'zFrameFids': ''}
    cliNode = slicer.cli.run(slicer.modules.zframeregistration, None, params, wait_for_completion=True)
    if cliNode.GetStatus() & cliNode.ErrorsMask:
      print(cliNode.GetErrorText())


    self.zFrameModelNode.SetAndObserveTransformNodeID(outputTransform.GetID())
    self.zFrameModelNode.GetDisplayNode().SetVisibility2D(True)
    self.zFrameModelNode.GetDisplayNode().SetSliceIntersectionThickness(2)
    self.zFrameModelNode.SetDisplayVisibility(True)
  
  def loadZFrameModel(self, ZFRAME_MODEL_PATH, ZFRAME_MODEL_NAME):
    if self.zFrameModelNode:
      slicer.mrmlScene.RemoveNode(self.zFrameModelNode)
      self.zFrameModelNode = None
    currentFilePath = os.path.dirname(os.path.realpath(__file__))
    zFrameModelPath = os.path.join(currentFilePath, "Resources", "zframe", ZFRAME_MODEL_PATH)
    self.zFrameModelNode = slicer.util.loadModel(zFrameModelPath)
    self.zFrameModelNode.SetName(ZFRAME_MODEL_NAME)
    modelDisplayNode = self.zFrameModelNode.GetDisplayNode()
    modelDisplayNode.SetColor(0.0,1.0,1.0)
    self.zFrameModelNode.SetDisplayVisibility(True)

  def createMaskedVolumeBySize(self, inputVolume, thresholdPercent, minimumSize, maximumSize):
    # Create segmentation node
    segmentationNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLSegmentationNode")
    segmentationNode.SetReferenceImageGeometryParameterFromVolumeNode(inputVolume)

    # Create segment
    segmentId = segmentationNode.GetSegmentation().AddEmptySegment("base")

    # Get access to the segment editor effect
    segmentEditorWidget = slicer.qMRMLSegmentEditorWidget()
    segmentEditorWidget.setMRMLScene(slicer.mrmlScene)
    segmentEditorNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLSegmentEditorNode")
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

    segmentEditorWidget.setActiveEffectByName("No editing")

    # Clean up
    segmentationNode.GetSegmentation().RemoveSegment(clonedSegmentId)
    # TODO: Does this cause problems with closing Slicer? Should delete widget?
    segmentEditorWidget = None

    # Export segmentation to label map
    self.removeNodeByName('MaskedCalibrationLabelMapVolume')
    labelMapVolumeNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLLabelMapVolumeNode", "MaskedCalibrationLabelMapVolume")
    slicer.modules.segmentations.logic().ExportVisibleSegmentsToLabelmapNode(segmentationNode, labelMapVolumeNode, inputVolume)

    # Count number of Islands and attempt repair if one is missing
    self.countAndRepairFiducials(labelMapVolumeNode)

    # Convert label map to scalar volume
    self.removeNodeByName('MaskedCalibrationVolume')
    scalarVolumeNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLScalarVolumeNode", "MaskedCalibrationVolume")
    slicer.modules.volumes.logic().CreateScalarVolumeFromVolume(slicer.mrmlScene, scalarVolumeNode, labelMapVolumeNode)

    slicer.mrmlScene.RemoveNode(segmentationNode)
    slicer.mrmlScene.RemoveNode(segmentEditorNode)
    slicer.mrmlScene.RemoveNode(labelMapVolumeNode)

    return scalarVolumeNode

  def countAndRepairFiducials(self, labelMapVolumeNode):
    segNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLSegmentationNode")
    slicer.modules.segmentations.logic().ImportLabelmapToSegmentationNode(labelMapVolumeNode, segNode)

    segmentEditorWidget = slicer.qMRMLSegmentEditorWidget()
    segmentEditorWidget.setMRMLScene(slicer.mrmlScene)
    segmentEditorNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLSegmentEditorNode")
    segmentEditorWidget.setMRMLSegmentEditorNode(segmentEditorNode)
    segmentEditorWidget.setSegmentationNode(segNode)
    segmentEditorWidget.setSourceVolumeNode(labelMapVolumeNode)

    segmentEditorWidget.setActiveEffectByName("Islands")
    effect = segmentEditorWidget.activeEffect()
    effect.setParameter("Operation", "SPLIT_ISLANDS_TO_SEGMENTS")
    effect.setParameter("MinimumSize", 0)
    effect.self().onApply()

    numberOfSegments = segNode.GetSegmentation().GetNumberOfSegments()
    # Attempt repair
    if numberOfSegments == 6:
      self.repairMissingFiducial(labelMapVolumeNode)
    
    # Cleanup
    segmentEditorWidget.setActiveEffectByName("No editing")
    slicer.mrmlScene.RemoveNode(segNode)
    slicer.mrmlScene.RemoveNode(segmentEditorNode)
    segmentEditorWidget = None

  def repairMissingFiducial(self, labelMapVolumeNode):
    # Identify missing fiducial
    # Isolate middle slice
    imageData = labelMapVolumeNode.GetImageData()
    centroid = self.findCentroidOfVolume(labelMapVolumeNode)
    middleSlice = int(centroid[2])
    dims = imageData.GetDimensions()
    numpy_array = vtk.util.numpy_support.vtk_to_numpy(imageData.GetPointData().GetScalars())
    numpy_array = numpy_array.reshape(dims[2], dims[1], dims[0])
    numpy_array = numpy_array.transpose(2,1,0)
    slice_array = numpy_array[:, :, middleSlice]

    imageData = self.numpy_to_vtk_image_data(numpy_array)
    labelMapVolumeNode.SetAndObserveImageData(imageData)

    # Shrink margins of image
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
    cropped = slice_array[leftColumn:rightColumn, topRow:bottomRow]

    # print(f'leftColumn {leftColumn}')
    # print(f'rightColumn {rightColumn}')
    # print(f'topRow {topRow}')
    # print(f'bottomRow {bottomRow}')

    # Probe array for values to look for missing value
    missingFiducial = False
    r = 10
    thickness = 5
    adjust = 4
    # Corners
    # Top Left
    if not np.any(cropped[0:r, 0:r] > 0):
      print("Attempting repair of top left fiducial")
      startLine = (leftColumn + adjust, topRow + adjust, middleSlice - 5)
      endLine = (leftColumn + adjust, topRow + adjust, middleSlice + 5)
      numpy_array = self.drawThickLine(startLine, endLine, thickness, numpy_array)
      missingFiducial = True
    # Top Right
    if not np.any(cropped[cropped.shape[0]-r:cropped.shape[0],0:r] > 0):
      print("Attempting repair of top right fiducial")
      startLine = (leftColumn + cropped.shape[0] - adjust, topRow + adjust, middleSlice - 5)
      endLine = (leftColumn + cropped.shape[0] - adjust, topRow + adjust, middleSlice + 5)
      numpy_array = self.drawThickLine(startLine, endLine, thickness, numpy_array)
      missingFiducial = True
    # Bottom Left
    if not np.any(cropped[0:r,cropped.shape[1]-r:cropped.shape[1]] > 0):
      print("Attempting repair of bottom left fiducial")
      startLine = (leftColumn + adjust, topRow + cropped.shape[1] - adjust, middleSlice - 5)
      endLine = (leftColumn + adjust, topRow + cropped.shape[1] - adjust, middleSlice + 5)
      numpy_array = self.drawThickLine(startLine, endLine, thickness, numpy_array)
      missingFiducial = True
    # Bottom Right
    if not np.any(cropped[cropped.shape[0]-r:cropped.shape[0],cropped.shape[1]-r:cropped.shape[1]] > 0):
      print("Attempting repair of bottom right fiducial")
      startLine = (leftColumn + cropped.shape[0] - adjust, topRow + cropped.shape[1] - adjust, middleSlice - 5)
      endLine = (leftColumn + cropped.shape[0] - adjust, topRow + cropped.shape[1] - adjust, middleSlice + 5)
      numpy_array = self.drawThickLine(startLine, endLine, thickness, numpy_array)
      missingFiducial = True

    # Sides
    # Middle Left
    if not np.any(cropped[0:r, cropped.shape[1]//2-r//2:cropped.shape[1]//2+r//2] > 0):
      print("Attempting repair of middle left fiducial")
      startLine = (leftColumn + adjust, topRow + cropped.shape[1]//2 - 5, middleSlice - 5)
      endLine = (leftColumn + adjust, topRow + cropped.shape[1]//2 + 5, middleSlice + 5)
      numpy_array = self.drawThickLine(startLine, endLine, thickness, numpy_array)
      missingFiducial = True
    # Middle Top  
    if not np.any(cropped[cropped.shape[0]//2-r//2:cropped.shape[0]//2+r//2, 0:r] > 0):
      print("Attempting repair of middle top fiducial")
      startLine = (leftColumn + cropped.shape[0]//2 + 5, topRow + adjust, middleSlice - 5)
      endLine = (leftColumn + cropped.shape[0]//2 - 5, topRow + adjust, middleSlice + 5)
      numpy_array = self.drawThickLine(startLine, endLine, thickness, numpy_array)
      missingFiducial = True
    # Middle Right  
    if not np.any(cropped[cropped.shape[0]-r:cropped.shape[0], cropped.shape[1]//2-r//2:cropped.shape[1]//2+r//2] > 0):
      print("Attempting repair of middle right fiducial")
      startLine = (leftColumn + cropped.shape[0] - adjust, topRow + cropped.shape[1]//2 + 5, middleSlice - 5)
      endLine = (leftColumn + cropped.shape[0] - adjust, topRow +  cropped.shape[1]//2 - 5, middleSlice + 5)
      numpy_array = self.drawThickLine(startLine, endLine, thickness, numpy_array)
      missingFiducial = True

    if missingFiducial:
      imageData = self.numpy_to_vtk_image_data(numpy_array)
      labelMapVolumeNode.SetAndObserveImageData(imageData)

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

    # Convert the center of mass from voxel coordinates to world coordinates
    # center_of_mass_world = origin + np.array(spacing) * center_of_mass

    # print("The center of mass in voxel coordinates is: ", center_of_mass)
    # print("The center of mass in world coordinates is: ", center_of_mass_world)

    return center_of_mass
