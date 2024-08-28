# ********************************************************************
# Ericsson Inc.                                                 SCRIPT
# ********************************************************************
#
#
# (c) Ericsson Inc. 2020 - All rights reserved.
#
# The copyright to the computer program(s) herein is the property
# of Ericsson Inc. The programs may be used and/or copied only with
# the written permission from Ericsson Inc. or in accordance with the
# terms and conditions stipulated in the agreement/contract under
# which the program(s) have been supplied.
#
# ********************************************************************
# Name    : CreateDynamicCalculatedColumns.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : CreateDynamicCalculatedColumns
#
# Usage   : NetAn RNO MRR
#

import clr
import re
clr.AddReference('System.Data')  # needs to be imported before System.Data is called
import System
from System.Data import DataSet, DataTable, XmlReadMode
from Spotfire.Dxp.Data import DataType, DataTableSaveSettings
from System.IO import StringReader, StreamReader, StreamWriter, MemoryStream, SeekOrigin
from System.Collections.Generic import List
from Spotfire.Dxp.Data import Import
from Spotfire.Dxp.Data.Import import TextFileDataSource, TextDataReaderSettings
from Spotfire.Dxp.Data.Import import DatabaseDataSource, DatabaseDataSourceSettings
from Spotfire.Dxp.Data import AddRowsSettings, DataColumn
from Spotfire.Dxp.Data.Import import DataTableDataSource
from Spotfire.Dxp.Data import *
from System import Array
import operator
import collections
from collections import OrderedDict
from Spotfire.Dxp.Application.Visuals import *


# =============================================
# Dictionary sources (globals)
# ==============================================

measureMapping = {
 'PLOSSUL': 'IL_DIM_E_NETOP_MRR_PATHLOSS_RANGE_MAPPING',
 'BSPOWER': 'IL_DIM_E_NETOP_MRR_BSPOWERLEVEL_RANGE_MAPPING',
 'TAVAL': 'IL_DIM_E_NETOP_MRR_TIMEADV_RANGE_MAPPING',
 'PLDIFF': 'IL_DIM_E_NETOP_MRR_PATHLOSSDIFF_RANGE_MAPPING', 
 'RXQUALUL':'IL_DIM_E_NETOP_MRR_SIGNAL_QUAL_RANGE_MAPPING', 
 'RXLEVUL': 'IL_DIM_E_NETOP_MRR_SIGNAL_STRENGTH_RANGE_MAPPING', 
 'RXLEVDL': 'IL_DIM_E_NETOP_MRR_SIGNAL_STRENGTH_RANGE_MAPPING', 
 'PLOSSDL': 'IL_DIM_E_NETOP_MRR_PATHLOSS_RANGE_MAPPING', 
 'FERDL': 'IL_DIM_E_NETOP_MRR_FER_RANGE_MAPPING', 
 'FERUL':'IL_DIM_E_NETOP_MRR_FER_RANGE_MAPPING', 
 'MSPOWER':'IL_DIM_E_NETOP_MRR_MSPOWERLEVEL_RANGE_MAPPING', 
 'RXQUALDL': 'IL_DIM_E_NETOP_MRR_SIGNAL_QUAL_RANGE_MAPPING'
}

rangeValues = {}

measureParameters = {}

KPIRangeMappingColumn = {
  'MSPOWER': 'dBMaxVal',
  'BSPOWER': 'dBMaxVal',
  'PLDIFF': 'dBMaxVal',
  'PLOSSUL': 'avgdB',
  'PLOSSDL': 'avgdB',
  'RXLEVUL': 'avgdB',
  'RXLEVDL': 'avgdB',
  'FERUL': 'gsmMaxVal',
  'FERDL': 'gsmMaxVal',
  'RXQUALUL': 'gsmMaxVal',
  'RXQUALDL': 'gsmMaxVal',
  'TAVAL': 'avgGsm'
}

measure_thresholds = {
    'FERUL': ['feruldd1','ferulst1','feruldd2','ferulst2', 'ferulpercentile' ],
    'FERDL': ['ferdldd1','ferdlst1','ferdldd2','ferdlst2', 'ferdlpercentile'],
    'RXQUALUL': ['rxqualuldd1','rxqualulst1','rxqualuldd2','rxqualulst2', 'rxqualulpercentile'],
    'RXQUALDL': ['rxqualdldd1','rxqualdlst1','rxqualdldd2','rxqualdlst2', 'rxqualdlpercentile'],
    'RXLEVUL': ['rxlevuldd1','rxlevulst1','rxlevuldd2','rxlevulst2', 'rxlevulpercentile'],
    'RXLEVDL': ['rxlevdldd1','rxlevdlst1','rxlevdldd2','rxlevdlst2', 'rxlevdlpercentile'],
    'PLOSSUL': ['pluldd1','plulst1','pluldd2','plulst2', 'plulpercentile'],
    'PLOSSDL': ['pldldd1','pldlst1','pldldd2','pldlst2', 'pldlpercentile'],
    'PLDIFF': ['pldiffdd1','pldiffst1','pldiffdd2','pldiffst2', 'pldiffpercentile'],
    'TAVAL': ['timeadvdd1','timeadvst1','timeadvdd2','timeadvst2', 'timeadvpercentile'],
    'BSPOWER': ['powerredbsdd1','powerredbsst1','powerredbsdd2','powerredbsst2','powerredbspercentile'],
    'MSPOWER': ['powerlevelmsdd1','powerlevelmsst1','powerlevelmsdd2','powerlevelmsst2','powerlevelmspercentile']
    }

measLimits = {
    'FERUL': 96,
    'FERDL': 96,
    'RXQUALUL': 7,
    'RXQUALDL': 7,
    'RXLEVUL': 63,
    'RXLEVDL': 63,
    'PLOSSUL': 59,
    'PLOSSDL': 64,
    'PLDIFF': 50,
    'TAVAL': 75,
    'BSPOWER': 15,
    'MSPOWER': 31}

convertNames = {
  'FERUL': 'FER UL',
  'FERDL': 'FER DL',
  'RXQUALUL': 'RXQUAL UL',
  'RXQUALDL': 'RXQUAL DL',
  'RXLEVUL': 'RXLEV UL',
  'RXLEVDL': 'RXLEV DL',
  'PLOSSUL': 'Path Loss UL',
  'PLOSSDL': 'Path Loss DL',
  'PLDIFF': 'Path Loss Diff.',
  'TAVAL': 'Timing Adv.',
  'BSPOWER': 'Power Red. BS',
  'MSPOWER': 'Power Lev. MS'
}

reverseLoopDict = {
  'PLOSSUL': False,
  'BSPOWER': True,
  'RXLEVDL': False,
  'TAVAL': False,
  'PLDIFF': False,
  'RXQUALUL': False,
  'RXLEVUL': False,
  'PLOSSDL': False,
  'FERDL': False,
  'FERUL': False,
  'MSPOWER': False,
  'RXQUALDL': False
}

# general columns and their calculations (used for mainly overview table, but also contains some for msmt export)
calculatedColumnList = OrderedDict([
    ("No. of Meas. Rep Passed Filter" , "sum(sum([BSPOWER0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([BSPOWER1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([BSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([BSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([BSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([BSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([BSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([BSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([BSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([BSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([BSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([BSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([BSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([BSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([BSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([BSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))"),
    ("Meas. Rep. Passed Filter (%)" , "100 * (sum(sum([BSPOWER0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),sum([BSPOWER1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),sum([BSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),sum([BSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),sum([BSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),sum([BSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),sum([BSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),sum([BSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),sum([BSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),sum([BSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),sum([BSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),sum([BSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),sum([BSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),sum([BSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),sum([BSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),sum([BSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])) / sum([REP]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))"),
    ("FER UL Passed Filter (%)" , "(sum(sum([FERUL0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL20]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL21]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL22]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL23]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL24]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL25]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL26]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL27]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL28]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL32]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL33]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL34]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL35]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL36]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL37]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL38]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL39]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL40]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL41]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL42]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL43]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL44]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL45]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL46]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL47]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL48]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL49]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL50]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL51]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL52]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL53]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL54]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL55]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL56]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL57]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL58]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL59]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL60]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL61]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL62]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL63]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL64]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL65]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL66]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL67]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL68]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL69]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL70]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL71]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL72]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL73]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL74]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL75]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL76]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL77]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL78]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL79]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL80]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL81]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL82]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL83]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL84]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL85]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL86]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL87]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL88]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL89]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL90]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL91]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL92]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL93]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL94]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL95]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL96]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])) *100)/sum([REPFERUL]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])"),
    ("FER DL Passed Filter (%)" , "(sum(sum([FERDL0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL20]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL21]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL22]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL23]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL24]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL25]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL26]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL27]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL28]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL32]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL33]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL34]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL35]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL36]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL37]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL38]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL39]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL40]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL41]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL42]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL43]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL44]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL45]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL46]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL47]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL48]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL49]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL50]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL51]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL52]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL53]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL54]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL55]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL56]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL57]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL58]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL59]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL60]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL61]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL62]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL63]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL64]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL65]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL66]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL67]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL68]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL69]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL70]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL71]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL72]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL73]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL74]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL75]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL76]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL77]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL78]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL79]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL80]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL81]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL82]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL83]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL84]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL85]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL86]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL87]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL88]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL89]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL90]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL91]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL92]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL93]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL94]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL95]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL96]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))*100)/sum([REPFERDL]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])"),
    ("No. of FER UL Passed Filter" , "sum(sum([FERUL0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL20]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL21]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL22]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL23]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL24]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL25]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL26]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL27]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL28]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL32]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL33]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL34]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL35]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL36]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL37]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL38]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL39]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL40]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL41]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL42]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL43]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL44]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL45]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL46]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL47]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL48]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL49]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL50]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL51]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL52]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL53]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL54]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL55]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL56]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL57]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL58]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL59]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL60]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL61]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL62]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL63]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL64]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL65]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL66]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL67]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL68]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL69]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL70]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL71]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL72]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL73]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL74]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL75]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL76]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL77]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL78]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL79]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL80]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL81]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL82]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL83]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL84]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL85]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL86]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL87]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL88]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL89]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL90]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL91]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL92]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL93]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL94]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL95]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERUL96]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))"),
    ("No. of FER DL Passed Filter" , "sum(sum([FERDL0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL20]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL21]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL22]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL23]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL24]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL25]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL26]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL27]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL28]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL32]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL33]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL34]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL35]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL36]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL37]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL38]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL39]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL40]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL41]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL42]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL43]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL44]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL45]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL46]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL47]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL48]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL49]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL50]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL51]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL52]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL53]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL54]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL55]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL56]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL57]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL58]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL59]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL60]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL61]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL62]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL63]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL64]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL65]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL66]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL67]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL68]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL69]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL70]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL71]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL72]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL73]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL74]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL75]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL76]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL77]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL78]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL79]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL80]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL81]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL82]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL83]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL84]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL85]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL86]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL87]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL88]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL89]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL90]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL91]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL92]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL93]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL94]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL95]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]),	sum([FERDL96]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))"),
    ("FER UL&DL Passed Filter (%)" , "(sum(REPFERTHL) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 100) / SUM([REPFERBL]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) "),
    ("No. of FER UL&DL Passed Filter" , "SUM([REPFERTHL]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])"),
    ("NoReportsUnfiltered" , "SUM([REP]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])"),
    ("NoFERULUnfiltered" , "SUM([REPFERUL]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])"),
    ("NoFERDLUnfiltered" , "SUM([REPFERDL]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])"),
    ("NoFERULDLUnfiltered" , "SUM([REPFERBL]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])" )
])

calculatedColumnsTrafficLevelSingle = OrderedDict([
    ("Traffic Level Average (E)" ,"0.48 * Sum([REP]) over ([BSC_Node_Name],[CELL_BAND],[CellName], [FILE_ID],[ChannelGroup]) / (60 * Sum([TTIME]) over ([CellName], [FILE_ID], [ChannelGroup])) "),
    ("TrafficLevelMaxCS(E)" , "Max([Traffic Level Average (E)]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])"),
    ("TrafficLevelMinCS(E)" , "Min([Traffic Level Average (E)]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])"),
    ("Traffic Level Min/Max.(%)" , "[TrafficLevelMinCS(E)] / [TrafficLevelMaxCS(E)] * 100")
])


calculatedColumnsTrafficLevelGroup = OrderedDict([
    ("Traffic Level Average (E)_temp" ,"0.48 * Sum([REP]) over ([BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup],[FILE_ID], [GroupingName]) / (60 * Sum([TTIME]) over ([CellName], [GroupingName], [ChannelGroup])) "),
    ("Traffic Level Average (E)" ,"0.48 * Sum([REP]) over ([BSC_Node_Name],[CELL_BAND],[CellName], [GroupingName],[ChannelGroup]) / (60 * Sum([TTIME]) over ([CellName], [GroupingName], [ChannelGroup])) "),
    ("TrafficLevelMaxCS(E)" , "Max([Traffic Level Average (E)_temp]) over ([GroupingName], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])"),
    ("TrafficLevelMinCS(E)" , "Min([Traffic Level Average (E)_temp]) over ([GroupingName], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])"),
    ("Traffic Level Min/Max.(%)" , "[TrafficLevelMinCS(E)] / [TrafficLevelMaxCS(E)] * 100")
])


hardCodedCols = [
    "No. of Meas. Rep Passed Filter",
    "Meas. Rep. Passed Filter (%)",
    "FER UL Passed Filter (%)",
    "No. of FER UL Passed Filter",
    "FER UL&DL Passed Filter (%)",
    "FER DL Passed Filter (%)",
    "No. of FER UL&DL Passed Filter",
    "No. of FER DL Passed Filter",
    "Traffic Level Min/Max.(%)",
    "Traffic Level Average (E)",
    'Power Lev. MS Average (dBm)',
    'RXLEV DL Average (dBm)',
    'Power Red. BS Average (dB)',
    'Timing Adv. Average (GSM)',
    'Path Loss Diff. Average (dB)',
    'FER DL Average (GSM)',
    'Path Loss DL Average (dB)',
    'Path Loss UL Average (dB)',
    'RXLEV UL Average (dBm)',
    'RXQUAL UL Average (GSM)',
    'FER UL Average (GSM)',
    'RXQUAL DL Average (GSM)'
]

gsm800_900 = OrderedDict([
    ('MSPOWER2',39),
    ('MSPOWER3',37),
    ('MSPOWER4',35),
    ('MSPOWER5',33),
    ('MSPOWER6',31),
    ('MSPOWER7',29),
    ('MSPOWER8',27),
    ('MSPOWER9',25),
    ('MSPOWER10',23),
    ('MSPOWER11',21),
    ('MSPOWER12',19),
    ('MSPOWER13',17),
    ('MSPOWER14',15),
    ('MSPOWER15',13),
    ('MSPOWER16',11),
    ('MSPOWER17',9),
    ('MSPOWER18',7),
    ('MSPOWER19',5)
])

gsm1800 = OrderedDict([
    ('MSPOWER29',36),
    ('MSPOWER30',34),
    ('MSPOWER31',32),
    ('MSPOWER0',30),
    ('MSPOWER1',28),
    ('MSPOWER2',26),
    ('MSPOWER3',24),
    ('MSPOWER4',22),
    ('MSPOWER5',20),
    ('MSPOWER6',18),
    ('MSPOWER7',16),
    ('MSPOWER8',14),
    ('MSPOWER9',12),
    ('MSPOWER10',10),
    ('MSPOWER11',8),
    ('MSPOWER12',6),
    ('MSPOWER13',4),
    ('MSPOWER14',2),
    ('MSPOWER15',0)

])

gsm1900 = OrderedDict([
    ('MSPOWER31',32),
    ('MSPOWER0',30),
    ('MSPOWER1',28),
    ('MSPOWER2',26),
    ('MSPOWER3',24),
    ('MSPOWER4',22),
    ('MSPOWER5',20),
    ('MSPOWER6',18),
    ('MSPOWER7',16),
    ('MSPOWER8',14),
    ('MSPOWER9',12),
    ('MSPOWER10',10),
    ('MSPOWER11',8),
    ('MSPOWER12',6),
    ('MSPOWER13',4),
    ('MSPOWER14',2),
    ('MSPOWER15',0)
])

differColumns = {
    'Path Loss DL Average (dB) DifferComp': 0,
    'FER UL Average (GSM) DifferComp': 0,
    'RXLEV UL Average (dBm) DifferComp': 0,
    'RXQUAL DL Average (GSM) DifferComp': 0,
    'Power Red. BS Average (dB) DifferComp': 0,
    'Power Lev. MS Average (dBm) DifferComp': 0,
    'Path Loss UL Average (dB) DifferComp': 0,
    'RXQUAL UL Average (GSM) DifferComp': 0,
    'FER DL Average (GSM) DifferComp': 0,
    'Path Loss Diff. Average (dB) DifferComp': 0,
    'RXLEV DL Average (dBm) DifferComp': 0,
    'Timing Adv. Average (GSM) DifferComp' : 0
}

# =============================================
# Utility Functions
# ==============================================

def createCursor(eTable):
    """Create cursors for a given table, these are used to loop through columns"""
    cursList = []
    colList = []
    for eColumn in eTable.Columns:
        cursList.append(DataValueCursor.CreateFormatted(eTable.Columns[eColumn.Name]))
        colList.append(eTable.Columns[eColumn.Name].ToString())
    cusrDict = dict(zip(colList, cursList))
    return cusrDict


def findMeasurementValue():
    """ Gets range mapping values"""
    values = {}
    columnName = ""
    count = 0
    for key, value in measureMapping.items():
        columnName = KPIRangeMappingColumn.get(key)
        mappingDataTable = Document.Data.Tables[value]
        mappingDataCur = createCursor(mappingDataTable)
        for row in mappingDataTable.GetRows(mappingDataCur[columnName], mappingDataCur['Measurement']):
            rangeValues[mappingDataCur['Measurement'].CurrentValue] = mappingDataCur[columnName].CurrentValue
            count += 1
    return rangeValues


def thresholds():
    """ creates a dict to hold threshold infor from settings page"""
    thresholdsVals = []
    for key, value in measure_thresholds.items():
            temp = {}
            temp[key] = []
            for each in value:
                temp[key].append(Document.Properties[each])
            measureParameters.update(temp)
    return measureParameters


def getNumFromStr(s):
    """ used in threshold functions to split num from counter"""
    m = re.search(r'\d+$', s)
    return int(m.group())


def removeColumns(tableCols, formulaType):
    """ takes a list of columns and deletes them. Used to remove old versions of formulas if settings change etc."""
    columnsToDelete = List[DataColumn]()
    calcColumnNameList = [calcColName for calcColName in convertNames.values()]

    if formulaType == "Average":
        val = ' Average'
    elif formulaType == "PercentileSum":
        val = "PERCENTILE_SUM"
        calcColumnNameList = [counter for counter in measLimits.keys()]
    elif formulaType == "Percentile":
        val = " Percentile"
    elif formulaType == "MSMT":
        calcColumnNameList = combined_col_list
        val = ""

    try:
        if formulaType == 'Threshold':
            for column in tableCols:
                for checkName in calcColumnNameList:
                    if checkName+'<' in column.Name or checkName+'>' in column.Name or checkName+'=' in column.Name or checkName+'None' in column.Name:
                        columnsToDelete.Add(column)
        elif formulaType == 'General':
            generalList = calculatedColumnList.keys()
            for column in tableCols:
                if column.Name in generalList:
                    columnsToDelete.Add(column)
        elif formulaType == 'TrafficLevelSingle':
            generalList = calculatedColumnsTrafficLevelSingle.keys()
            for column in tableCols:
                if column.Name in generalList:
                    columnsToDelete.Add(column)
        elif formulaType == 'TrafficLevelGroup':
            generalList = calculatedColumnsTrafficLevelGroup.keys()
            for column in tableCols:
                if column.Name in generalList:
                    columnsToDelete.Add(column)
        else:
            for column in tableCols:
                for checkName in calcColumnNameList:
                    if checkName + val in column.Name and 'Trend_' not in column.Name:
                        columnsToDelete.Add(column)

        tableCols.Remove(columnsToDelete)
    except Exception as e:
        print e.message


def addColumns(tableCols, formulaDict, tableName, change = True):
    """ genereic function to add calc columns to table"""
    for key, val in formulaDict:
        if tableName == 'combinedRecordingDataGrouping':
            if change:
                val = val.replace("[FILE_ID]", "[GroupingName]")
        if len(str(key)) > 0 and len(str(val)) > 0:
            tableCols.AddCalculatedColumn(str(key),str(val))


def getUnit(counter):
    """ differenct characteristics have different units. This function returns that unit"""
    unit = ""
    if counter in ['FERUL','FERDL','RXQUALDL','RXQUALUL','TAVAL']:
        unit = "GSM"
    elif counter in ['MSPOWER','RXLEVUL','RXLEVDL']:
        unit = "dBm"
    else:
        unit = "dB"

    return unit


def createDataTable(textData):
    """"Used to create the measure names table"""
    dataTableName = 'Measure Names'

    stream = MemoryStream()
    writer = StreamWriter(stream)
    writer.WriteLine('Measures\r\n')
    writer.Flush()

    for line in textData:
        writer.WriteLine(line)

    writer.Flush()
    settings = TextDataReaderSettings()
    settings.AddColumnNameRow(0)
    settings.ClearDataTypes(False)
    stream.Seek(0, SeekOrigin.Begin)
    fs = TextFileDataSource(stream, settings)

    if Document.Data.Tables.Contains(dataTableName):
        dataTable = Document.Data.Tables[dataTableName]
        dataTable.ReplaceData(fs)
    else:
        # Create Table if not already present
        print "Adding Data Table"
        dataTable = Document.Data.Tables.Add(dataTableName, fs)
        print 'Done'

# =============================================
# Threshold Functions
# ==============================================

def swapThresholdOperator(operator):
    """ Swaps operator for BSPOWER in threshold func"""
    if operator == "<":
        operator =">"
    elif operator == ">":
        operator = "<"
    else:
        operator = "="

    return operator


def buildString(dictValues, threshold, operatorCheck):
    """ takes a dictionary of MSPOWER values by band, and then calulates from key values which MSPOWER to use"""

    finalString = ""

    # use operator to simplify if check
    if operatorCheck == "<":
        comp = operator.lt
    elif operatorCheck == ">":
        comp = operator.gt
    else:
        comp = operator.eq

    buildString = []

    # check if value less than/greater and then create list of strings
    for k,v in dictValues.items():
        if comp(v, threshold):
            buildString.append("sum(["+str(k)+"]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup])")

    # if the string list is empty, then set to 0.0
    if buildString:
        finalString = ' + '.join(buildString)
    else:
        finalString = "0.0"

    return finalString


def mspowerThresholdCalc(operatorCheck, threshold):
    """ Build case statement for mspower threshold"""

    # outer case statements
    outerCase_gsm800_900 = "when [CELL_BAND] = 'GSM800' or [CELL_BAND] = 'GSM900' then "
    outerCase_gsm1800 = "when [CELL_BAND] = 'GSM1800' then "
    outerCase_gsm1900 = "when [CELL_BAND] = 'GSM1900' then "

    # inner sums built up using a function checking against a dict for start/stop
    innerSum_800_900 = buildString(gsm800_900, threshold, operatorCheck)
    innerSum_1800 = buildString(gsm1800, threshold, operatorCheck)
    innerSum_1900 = buildString(gsm1900, threshold, operatorCheck)

    #final case string
    caseString = "(CASE "+ outerCase_gsm800_900 + "(" + innerSum_800_900 + ") " \
                    + outerCase_gsm1800 + "(" + innerSum_1800 + ") " \
                    + outerCase_gsm1900 + "(" + innerSum_1900 +") END )"

    return caseString


def thresholdCalc(operatorIndex, thresholdIndex):
    """ Generates all the threshold values based on op/thresh index (i.e first column/second column threshold values)"""

    calcColumns = {}
    for key, val in measLimits.items():

        threshold = 0
        name = ""
        unit = ""
        unit = getUnit(key)
        operator = thresholdVals[key][operatorIndex]
        tempOperator = operator

        # dont add columns with none in threshold settings
        if str(operator) != 'None':
            # BSPOWER need to be summed in reverse, so temporarily change the directon of how the counters are summed
            if key == 'BSPOWER':
                tempOperator = swapThresholdOperator(operator)

            # get threshold value
            if key == "FERUL" or key == "FERDL":
                threshold = thresholdVals[key][thresholdIndex]
            elif key != "MSPOWER":
                threshold = int("".join([str(getNumFromStr(i)) for i, a in rangeValues.items() if key in i if a == str(thresholdVals[key][thresholdIndex])]))

            # create formula strings
            A = "("
            B = "("

            if key == "MSPOWER":
                mspowerThreshold = int(thresholdVals[key][thresholdIndex])
                A = "(case when [CELL_BAND] = 'GSM800' or [CELL_BAND] = 'GSM900' then sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) when [CELL_BAND] = 'GSM1800' then sum([MSPOWER0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) when [CELL_BAND] = 'GSM1900' then sum([MSPOWER0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) + sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) end)"
                B = mspowerThresholdCalc(tempOperator, mspowerThreshold)
            else:
                # formula for A part
                for i in range(0,val+1):
                    if i != val:
                        A += "sum(["+key+""+str(i)+"]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup])+"
                    else:
                        A += "sum(["+key+""+str(i)+"]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup]))"

                # formula for B Part
                if tempOperator == ">":
                    if threshold < val:
                        for i in range(threshold+1,val+1):
                            if i != val:
                                B += "sum(["+key+""+str(i)+"]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup])+"
                            else:
                                B+="sum(["+key+""+str(i)+"]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup]))"
                elif tempOperator == "<":
                    if threshold > 0:
                        for i in range(0, threshold):
                            if i != threshold-1:
                                B += "sum(["+key+""+str(i)+"]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup])+"
                            else:
                                B += "sum(["+key+""+str(i)+"]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup]))"
                elif tempOperator == "=":
                    B = "(sum(["+key+""+str(threshold)+"]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup]))"

            # if threshold is less than 0 or greater than max, then set to 0.0, else use normal formula
            if B == "" or B == "(":
                formula = "0.0"
            else:
                # check data resolution
                if dataResolution == "Cell":
                    A = A.replace(',[ChannelGroup]', '')
                    B = B.replace(',[ChannelGroup]', '')

                formula = "100*"+B+"/"+A

            # assign column name and final formula
            name = convertNames[key] + str(operator) + " " + str(thresholdVals[key][thresholdIndex]) + " " + unit + " (%)"
            calcColumns[name] = "IF(" + A +" !=0,"+ formula + ",NULL)"

    return calcColumns


# =============================================
# Percentile Functions
# ==============================================

def percentileSum(counterName):
    """Creates a calculated column (using a built up case string) for sum of counters for a particualr counter and multiply by perecentile doc prop"""

    calcColumns = {}

    # MSPOWER has band info so cant be done the same as the others
    if counterName == 'MSPOWER':
        percentileProp  = str(measureParameters[counterName][4])
        sumString = """
                    CASE WHEN ([CELL_BAND]="GSM800") OR ([CELL_BAND]="GSM900") THEN (sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))* ({0} / 100)WHEN ([CELL_BAND]="GSM1800") THEN (sum([MSPOWER0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])) * ({0} / 100)WHEN ([CELL_BAND]="GSM1900") THEN (sum([MSPOWER0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+ sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])) * ({0} / 100)ELSE 0 END
                    """.format(percentileProp)
    else:
        sumString = "("

        for i in range(0,measLimits[counterName] + 1):
            if i != measLimits[counterName]:
                sumString += "sum([" + counterName +str(i)+"]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup])+"
            else:
                sumString +="sum([" + counterName +str(i)+"]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup]))"

        percentileProp  = str(measureParameters[counterName][4])
        sumString += "*(" + percentileProp +"/100)"
    columnName = counterName + "PERCENTILE_SUM"

    if dataResolution == "Cell":
        calcColumns[columnName] = sumString.replace(',[ChannelGroup]', '')
    else:
        calcColumns[columnName] = sumString

    return calcColumns


def percentileFullCalc(counterName):
    """Creates a calculated column (using a built up case string) for a counter """

    whenStatement = ""
    calcColumns = {}
    percentileSumColumn = "[" + counterName + "PERCENTILE_SUM]"
    percentileValue = str(measureParameters[counterName][4])

    if percentileValue != 'None':

        # MSPOWER required a hard coded statement as it is by band and diff than the rest.
        if counterName == 'MSPOWER':
            caseString = """
            If({0}!=0,If([CELL_BAND] = 'GSM800' OR [CELL_BAND] = 'GSM900' ,CASE  WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])) >=[MSPOWERPERCENTILE_SUM] THEN 5 WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 7                                 WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 9                                 WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 11                                 WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 13                                WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 15                                WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 17                                WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 19                                WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 21                                WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 23                                WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 25                                WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 27                                WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 29                                WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 31                                WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 33                                WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 35                                WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))                        >=[MSPOWERPERCENTILE_SUM] THEN 37                                 WHEN (sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) )                        >=[MSPOWERPERCENTILE_SUM] THEN 39                        ELSE 0                        END,If([CELL_BAND] = 'GSM1900',CASE                          WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        32                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        0                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        2                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        4                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        6                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        8                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        10                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        12                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        14                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        16                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        18                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        20                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        22                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        24                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        26                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        28                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN                        30                        ELSE 0                        END,If([CELL_BAND] = 'GSM1800',CASE                          WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))>=[MSPOWERPERCENTILE_SUM] THEN 32                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 34                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 36                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 0                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 2                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 4                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 6                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 8                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 10                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 12                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 14                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 16                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 18                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 20                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 22                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 24                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 26                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 28                        WHEN (sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))+sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])+sum([MSPOWER0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup])>=[MSPOWERPERCENTILE_SUM] THEN 30 ELSE 0                        END, 0                        )                        )                        ) 
            ,NULL)
            """.format(percentileSumColumn)
        else:
            # check if counters need to be summed in reverse, i.e from 30 to 0, or 0 to 30
            if reverseLoopDict[counterName]:
                direction = -1
                start = measLimits[counterName]
                stop = -1
            else:
                direction = 1
                start = 0
                stop = measLimits[counterName] + 1

            for i in range(start, stop, direction):
                subSum = "("
                if direction == -1:
                    # only for bspower, mspower is handled by the above hard coded case statement
                    for j in range(start,i -1, direction):
                        subSum += "(sum([" + counterName + str(j) + "]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup])) +"
                else:
                    for j in range(start,i + 1):
                        subSum += "(sum([" + counterName + str(j) + "]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup])) +"
                subSum = subSum[:-1] # removing + at end of string


                whenStatement += "WHEN "+ subSum +")  >= " + percentileSumColumn + " THEN " + rangeValues[counterName+str(i)] + " " #using i, not j as we want the current column val

            # final case statement for the calculated column
            caseString = """If({0}!=0,
                                CASE
                                {1}
                                ELSE 0
                                END,
                            NULL)""".format(percentileSumColumn, whenStatement)

        convertedName = convertNames[counterName]
        unit = getUnit(counterName)

        columnName = convertedName + " Percentile " + percentileValue + " (" + unit + ")"

        if dataResolution == "Cell":
            calcColumns[columnName] = caseString.replace(',[ChannelGroup]', '')
        else:
            calcColumns[columnName] = caseString

    return calcColumns

# =============================================
# Average Functions
# =============================================

def averageCalc():
    """ Calculates all the average formulas for all given charateristics (except for MSPOWER)"""
    calcColumns = {}
    for key, val in measLimits.items():
        if key != "MSPOWER":
            A = "("
            B = "("
            name = ""
            unit = ""
            unit = getUnit(key)
            for i in range(0,val+1):
                if i != val:
                    A += "sum(["+key+""+str(i)+"]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup])+"
                    B += "sum(["+key+""+str(i)+"]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup])*"+rangeValues[""+key+""+str(i)]+"+"
                else:
                    A+="sum(["+key+""+str(i)+"]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup]))"
                    B+="sum(["+key+""+str(i)+"]) over ([FILE_ID], [BSC_Node_Name], [CELL_BAND], [CellName],[ChannelGroup]) *"+rangeValues[""+key+""+str(i)]+")"
            if dataResolution == "Cell":
                A = A.replace(",[ChannelGroup]","")
                B = B.replace(",[ChannelGroup]","")
            name = convertNames[key] +" Average ("+ unit +")"
            formula = B +"/"+A
            calcColumns[name] = "IF(" + A +" !=0,"+ formula + ",NULL)"

    #mspower calculated differently for average
    msPowerCalc = "case  when [CELL_BAND]='GSM1800' then ((    sum([MSPOWER0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 30) + (   sum([MSPOWER1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 28) + (   sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 26) + (   sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 24) + (   sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 22) + (   sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 20) + (   sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 18) + (   sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 16) + (   sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 14) + (   sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 12) + (   sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 10) + (   sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 8) + (   sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 6) + (   sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 4) + (   sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 2) + (   sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 0) + (   sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 36) + (   sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 34) + (   sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 32))  / (sum([MSPOWER0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER29]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER30]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))    when [CELL_BAND]='GSM1900' then ((    sum([MSPOWER0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 30) + (   sum([MSPOWER1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 28) + (   sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 26) + (   sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 24) + (   sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 22) + (   sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 20) + (   sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 18) + (   sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 16) + (   sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 14) + (   sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 12) + (   sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 10) + (   sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 8) + (   sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 6) + (   sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 4) + (   sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 2) + (   sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 0) + (   sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 32))  / (sum([MSPOWER0]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER1]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER31]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))    when ([CELL_BAND]='GSM800') or ([CELL_BAND]='GSM900') then ((    sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 39) + (   sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 37) + (   sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 35) + (   sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 33) + (   sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 31) + (   sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 29) + (   sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 27) + (   sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 25) + (   sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 23) + (   sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 21) + (   sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 19) + (   sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 17) + (   sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 15) + (   sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 13) + (   sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 11) + (   sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 9) + (   sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 7) + (   sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) * 5))   / (   sum([MSPOWER2]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER3]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER4]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER5]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER6]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER7]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER8]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER9]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER10]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER11]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER12]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER13]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER14]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER15]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER16]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER17]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER18]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]) +   sum([MSPOWER19]) over ([FILE_ID], [BSC_Node_Name],[CELL_BAND],[CellName],[ChannelGroup]))  end"
    
    if dataResolution == "Cell":
        msPowerCalc = msPowerCalc.replace(",[ChannelGroup]","")

    #add mspower average column to dict
    calcColumns["Power Lev. MS Average (dBm)"] = msPowerCalc



    return calcColumns


# =============================================
# General Formula Functions (Meas. Report filter etc.)
# =============================================

def generalCalc(calculatedColumnListModify):
    """ mainly removes the channel group field for cell name agg, otherwise just returns the calc column dict"""
    if dataResolution == "Cell":
        for key,val in calculatedColumnListModify.items():
            calculatedColumnListModify[key] = calculatedColumnListModify[key].replace(',[ChannelGroup]', '')

    return calculatedColumnListModify


# msmt formulas

def msmt_modified_dict(original_dict):
    modified_dict = OrderedDict()
    column_translate = OrderedDict()
    for column_name, formula in original_dict.items():
        new_column_name = column_name + "_MSMT"
        modified_dict[new_column_name] = formula
        column_translate["["+column_name+"]"] = "["+new_column_name+"]"

    for old_col, new_col in column_translate.items():
        for modified_col, original_formula in modified_dict.items():
            modified_dict[modified_col] = original_formula.replace(old_col,new_col)

    return modified_dict

#Change Chart filter and axis for Data Resolution
def changeDataResolution(pageTitle, visualName):
	for page in Document.Pages:
		if page.Title == pageTitle:
			for visual in page.Visuals:
				if visual.Title == visualName:
					print visual.Title
					html = visual.As[HtmlTextArea]().HtmlContent
					if Document.Properties['DataResolution'] == "Channel Group":
						visual.As[HtmlTextArea]().HtmlContent = html.Replace('hidden','visible')
					else:
						visual.As[HtmlTextArea]().HtmlContent = html.Replace('visible','hidden')
				if visual.TypeId == VisualTypeIdentifiers.CombinationChart or visual.TypeId == VisualTypeIdentifiers.BarChart:
					chart = visual.As[VisualContent]()
					if Document.Properties['DataResolution'] == "Channel Group":
						chart.XAxis.Expression = "<[ChannelGroup] NEST [SubCell_String] NEST [CellName] NEST [BSC_Node_Name] NEST [CELL_BAND]>"
					else:
						chart.XAxis.Expression = "<[SubCell_String] NEST [CellName] NEST [BSC_Node_Name] NEST [CELL_BAND]>"


# =============================================
# Main loop - Variable declartion and dict settings
# ==============================================
Document.Properties["errorMessage"] = ''
Document.Properties["settingsSavedMessage"] = ''

#data tables
combinedRecordingData = 'combinedRecordingData'
combinedRecordingDataGrouping = 'combinedRecordingDataGrouping'

# data resolution from settings page - channel group or cellname. Chgrp is lowest level
dataResolution = Document.Properties["DataResolution"]

# set table names and columns
combinedRecordingDataColumnsGrouping = Document.Data.Tables['combinedRecordingDataGrouping'].Columns
combinedRecordingDataColumns = Document.Data.Tables['combinedRecordingData'].Columns

# generate dicts to hold range mapping and threshold values from doc props
findMeasurementValue()
thresholdVals = thresholds()

# lists used for writing to measure table for drop down list
textData = []
percentileCols = []

# =============================================
# Main loop - Remove Cols
# ==============================================

# remove from single data
removeColumns(combinedRecordingDataColumns,'General')
removeColumns(combinedRecordingDataColumns,'Average')
removeColumns(combinedRecordingDataColumns,'PercentileSum')
removeColumns(combinedRecordingDataColumns,'Percentile')
removeColumns(combinedRecordingDataColumns,'Threshold')
removeColumns(combinedRecordingDataColumns,'TrafficLevelSingle')


# remove from grouped data
removeColumns(combinedRecordingDataColumnsGrouping,'General')
removeColumns(combinedRecordingDataColumnsGrouping,'Average')
removeColumns(combinedRecordingDataColumnsGrouping,'PercentileSum')
removeColumns(combinedRecordingDataColumnsGrouping,'Percentile')
removeColumns(combinedRecordingDataColumnsGrouping,'Threshold')
removeColumns(combinedRecordingDataColumnsGrouping,'TrafficLevelGroup')


# =============================================
# Main loop - Add Cols
# ==============================================


# # MSMT columns
generalFormulasMSMT = msmt_modified_dict(calculatedColumnList)
trafficLevelFormulaSingleMSMT = msmt_modified_dict(calculatedColumnsTrafficLevelSingle)
trafficLevelFormulaGroupMSMT = msmt_modified_dict(calculatedColumnsTrafficLevelGroup)

# create dict to get all column names to remove for msmt
combined_col_list = list(generalFormulasMSMT) + list(trafficLevelFormulaSingleMSMT) + list(trafficLevelFormulaGroupMSMT)

removeColumns(combinedRecordingDataColumns,'MSMT')
removeColumns(combinedRecordingDataColumnsGrouping,'MSMT')

addColumns(combinedRecordingDataColumns,generalFormulasMSMT.items(), combinedRecordingData)
addColumns(combinedRecordingDataColumnsGrouping,generalFormulasMSMT.items(), combinedRecordingDataGrouping)
addColumns(combinedRecordingDataColumns,trafficLevelFormulaSingleMSMT.items(), combinedRecordingData,False)
addColumns(combinedRecordingDataColumnsGrouping,trafficLevelFormulaGroupMSMT.items(), combinedRecordingDataGrouping, False)


# General Formulas
generalFormulas = generalCalc(calculatedColumnList).items()

addColumns(combinedRecordingDataColumns,generalFormulas, combinedRecordingData)
addColumns(combinedRecordingDataColumnsGrouping,generalFormulas, combinedRecordingDataGrouping)

#Traffic Level Formulas

trafficLevelFormulaSingle = generalCalc(calculatedColumnsTrafficLevelSingle).items()
addColumns(combinedRecordingDataColumns,trafficLevelFormulaSingle, combinedRecordingData,False)


trafficLevelFormulaGroup = generalCalc(calculatedColumnsTrafficLevelGroup).items()
addColumns(combinedRecordingDataColumnsGrouping,trafficLevelFormulaGroup, combinedRecordingDataGrouping, False)

# Average formulas
averageCalcFormulas =  averageCalc().items()

addColumns(combinedRecordingDataColumns,averageCalcFormulas, combinedRecordingData)
addColumns(combinedRecordingDataColumnsGrouping,averageCalcFormulas, combinedRecordingDataGrouping)

# Theshold formulas

thresholdCalcFormulasFirstCol = thresholdCalc(0,1).items()
thresholdCalcFormulasSecondCol = thresholdCalc(2,3).items()

thresholdCalcFormulasCombined = list(set(thresholdCalcFormulasFirstCol) | set(thresholdCalcFormulasSecondCol))

addColumns(combinedRecordingDataColumns,thresholdCalcFormulasCombined, combinedRecordingData)
addColumns(combinedRecordingDataColumnsGrouping,thresholdCalcFormulasCombined, combinedRecordingDataGrouping)

# Percentile Formulas

for counter in measLimits.keys():

    counterVal = measureParameters[counter][4]

    if counterVal != None:
        percentileSumFormulas = percentileSum(counter).items()

        addColumns(combinedRecordingDataColumns,percentileSumFormulas, combinedRecordingData)
        addColumns(combinedRecordingDataColumnsGrouping,percentileSumFormulas, combinedRecordingDataGrouping)

        percentileFullCalcFormulas = percentileFullCalc(counter).items()
        # create list of col names for percentile - different than rest of formulas as it generates one at a time
        for key, val in percentileFullCalcFormulas:
            percentileCols.append(str(key))

        addColumns(combinedRecordingDataColumns,percentileFullCalcFormulas, combinedRecordingData)
        addColumns(combinedRecordingDataColumnsGrouping,percentileFullCalcFormulas, combinedRecordingDataGrouping)

# add dummy columsn for comparison report as removing these causes linked data errors when saving settings,
# and then the user saves. Reopening the workbook throws error in unpivot table as variables dont exsist yet
addColumns(combinedRecordingDataColumns,differColumns.items(), combinedRecordingData)

# =============================================
# Main loop -  Clean up overview table to avoid having errors with mismatched columns
# ==============================================

OverviewTable = Document.Data.Tables["OverviewTable"]

columnsToDelete = List[DataColumn]()

for column in OverviewTable.Columns:

    if column.Name != 'IsDuplicate': # isuplicate is a calculated column to limit data based on grouping etc.
        columnsToDelete.Add(column)

OverviewTable.Columns.Remove(columnsToDelete)

sourceDataTable = Document.Data.Tables["combinedRecordingData"]

dataSource=DataTableDataSource(sourceDataTable)

OverviewTable.ReplaceData(dataSource)

# =============================================
# Main loop -  Create items for dropdowns - measures data table
# ==============================================

for col in hardCodedCols:
    textData.append(col)

for key, val in thresholdCalcFormulasCombined:
    textData.append(str(key))

for colPercentile in percentileCols:
    textData.append(colPercentile)

createDataTable(textData)

changeDataResolution("MRR Top Ten Chart", "FilterTextArea")
changeDataResolution("MRR Comparison Chart", "CompFilterTextArea")

Document.Properties["settingsSavedMessage"] = "Settings Saved"
