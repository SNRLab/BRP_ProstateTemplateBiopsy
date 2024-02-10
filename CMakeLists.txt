cmake_minimum_required(VERSION 3.13.4)

project(BRP_ProstateTemplateBiopsy)

#-----------------------------------------------------------------------------
# Extension meta-information
set(EXTENSION_HOMEPAGE "")
set(EXTENSION_CATEGORY "IGT")
set(EXTENSION_CONTRIBUTORS "Franklin King (SNR)")
set(EXTENSION_DESCRIPTION "Module supporting MRI-guided prostate biopsy with template")
set(EXTENSION_ICONURL "")
set(EXTENSION_SCREENSHOTURLS "")
set(EXTENSION_DEPENDS "NA") # Specified as a list or "NA" if no dependencies

#-----------------------------------------------------------------------------
# Extension dependencies
find_package(Slicer REQUIRED)
include(${Slicer_USE_FILE})

#-----------------------------------------------------------------------------
# Extension modules
add_subdirectory(ProstateTemplateBiopsy)

#-----------------------------------------------------------------------------
include(${Slicer_EXTENSION_GENERATE_CONFIG})
include(${Slicer_EXTENSION_CPACK})
