# NPMRDS Tool

This tool processes NPMRDS data downloaded from RITIS' NPMRDS Trend Map tool to use in ArcGIS. It converts the downloaded .XML files to tables and imports that data into a geodatabase.

## Prerequisites

* XML files downloaded from the [NPMRDS Trend Map Tool](https://npmrds.ritis.org/analytics/)
* Latest version of the pywin32 package added to the [ArcGIS Pro Python Package Manager](https://pro.arcgis.com/en/pro-app/arcpy/get-started/what-is-conda.htm)

## Installing

Download [ProcessNPMRDSTool.py](ProcessNPMRDSTool.py) for non-GIS import or [ProcessNPMRDSTool_Extended.py](ProcessNPMRDSTool_Extended.py) for GIS import and place the file in a desired location

## Running the Tool

1. Change the file path of Line 18 to where your .XML files are located
2. Rename the gdb on Line 19 to what you would like

## Author

[Christian Matthews](https://github.com/csmatthews)

## License

This project is licensed under the [MIT License](LICENSE.md)

## Acknowledgments

* README Template = [PurpleBooth](https://github.com/PurpleBooth)
* Script = [dilbert](https://stackoverflow.com/users/2507539/dilbert) , [ExtendOffice](https://www.extendoffice.com/)
