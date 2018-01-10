# NPMRDS Tool

This tool processes NPMRDS data downloaded from RITIS' NPMRDS Trend Map tool to use in ArcGIS Pro. It converts the downloaded .XML files to tables, imports that data into a geodatabase, and creates new fields.

## Prerequisites

* XML files downloaded from the [NPMRDS Trend Map Tool](https://npmrds.ritis.org/analytics/)
* Latest version of the pywin32 package added to the [ArcGIS Pro Python Package Manager](https://pro.arcgis.com/en/pro-app/arcpy/get-started/what-is-conda.htm)

## Installing

Download [NPMRDSTool.py](NPMRDSTool.py) and place the file in a desired location

## Running the Tool

1. Change the file path of Line 60 to where your .XML files are located
2. Run the ProcessExcel.py file using the ArcGIS Pro IDLE

## Author

[Christian Matthews](https://github.com/csmatthews)

## License

This project is licensed under the [MIT License](LICENSE.md)

## Acknowledgments

* README Template = [PurpleBooth](https://github.com/PurpleBooth)
* Script = [dilbert](https://stackoverflow.com/users/2507539/dilbert) , [ExtendOffice](https://www.extendoffice.com/)
