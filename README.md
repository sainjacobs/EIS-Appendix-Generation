# EIS Appendix Generation

## Name
EIS Appendix Generation

## Description
Automatically generate appendices for EIS reports containing VI/508 Compliant tables and figures.


## Installation
You will need local copies of this repo and the CalSim DSS Reader: https://gitlab.bor.doi.net/usbr-cvp-modeling/calsim_dss_reader

Use appendix_gen.yml to create a conda environment with necessary packages for running the scripts.

conda env create -f apendix_gen.yml

## Usage
1. Edit the desired runs and location codes in the CalSim DSS Reader dssReader.py and then run according to instructions in DSSReader repo. 
2. When the DSS Reader has finished running, copy the DSSContents_CFS.xlsx output from the DSSReader 
directory to eis-appendix-generation/inputs. 
3. Edit the desired runs and field codes in the EIS Appendix Generation EISAppendixGen.py. They must match the strings that were input
to the DSS Reader. 
4. Edit the file paths in EISAppendixGen.py.
5. Run!

## Support
Please contact emadonna@usbr.gov for support