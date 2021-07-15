# NPMRDS Tool

This tool processes NPMRDS data downloaded from RITIS' NPMRDS Trend Map tool to use in ArcGIS. It converts the downloaded .XML files to tables and imports that data into a geodatabase.

## Prerequisites

* XML files downloaded from the [NPMRDS Trend Map Tool](https://npmrds.ritis.org/analytics/)
* Latest version of the pywin32 package added to the [ArcGIS Pro Python Package Manager](https://pro.arcgis.com/en/pro-app/arcpy/get-started/what-is-conda.htm) or to your machine.

## Installing

1. Non-ArcGIS - [ProcessNPMRDSTool.py](ProcessNPMRDSTool.py) - This script will only process the .XML files.
2. UNDER DEVELOPMENT - For ArcGIS - [ProcessNPMRDSTool_Extended.py](ProcessNPMRDSTool_Extended.py) - This script will process the .XML files and import them.

## Running the Tool

1. Non-ArcGIS - [ProcessNPMRDSTool.py](ProcessNPMRDSTool.py) - This script will only process the .XML files.
    1. Change the file path of Line 17 to the folder where your .XML files are located

## Author

[Christian Matthews](https://github.com/csmatthews)

## License

This project is licensed under the [MIT License](LICENSE.md)

## Acknowledgments

* README Template = [PurpleBooth](https://github.com/PurpleBooth)
* Script = [dilbert](https://stackoverflow.com/users/2507539/dilbert) , [ExtendOffice](https://www.extendoffice.com/)
