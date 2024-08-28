from Spotfire.Dxp.Data import *
from Spotfire.Dxp.Application.Filters import *

def resetAllFilteringSchemes(dataTable):
     # Loop through all filtering schemes
     for filteringScheme in Document.FilteringSchemes:
          filteringScheme[dataTable].ResetAllFilters()

#Reset Filters
resetAllFilteringSchemes(Document.Data.Tables['OverviewTableUnpivot'])

#navigate to overview hist. page
for page in Document.Pages:
    if (page.Title == 'MRR Top Ten Chart'):
        Document.ActivePageReference=page