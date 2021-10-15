                 Parametric study of the influence of wall thickness on pipeline expansion design.

Tools used:
 -  Abaqus
 -  Python
 - xlsxwriter

Description:
Pipeline resting on seabed and subjected to pressure and temperature will expand. The forces causing the expansion are the resultant effect of  endcap force, poisson force, thermal force and frictional resistance. 

This tool has been prepared to verify the behaviour of pipeline expansion due to effect of  wall thickness variation.
A basecase abaqus input file of the pipeline expansion was first prepared . The wall thickness in the input file was parameterized in such a way as to allow variation of the thickness value .
Six thickness values has been evaluated in this study using the python parametric study method for Abaqus to generate 6 input files and automatically batch runs to obtain 6odbs which is called LOADCASES in this analysis. The thickness values are :[15.9mm,19.1mm,22.3mm,25.1mm,27.1mm, 30.2mm]. The parametric study script  used here is called "study.psf".


Upon successful completion of the preprocessing, the 6odbs are generated in the project folder. Each odbs contains 6 steps and 1000 pipeline nodes. Thus, the total loadcase data  is 6*6*1000 = 36000 loadcases.
The postprocessing is done in the Python scripts called "Python postprocessor". The script open the odbs one after the other ,  loop through each loading steps and extract the results for nodal displacements U1,U2 and U3,
then write the result to excel sheet2 with help of xlsxwriter. open sheet1 and write the maximum, minimum values of U1,U2 and U3 and the corrsponding node, loadstep were it occurs.
Thus with a table of values of worst thermal expansion for each loadcase and their corresponding thickness values, a plot is made as shown in the attached excel.
From the plot, one can see that as the wall thickness increases, the expansion also increases which depicts a  linear relationship.

1.Preprocessing
- prepare a working basecase abaqus model of the pipeline
- parametrize the variable for study in the input file
- prepare parametric study file (.psf) that contains the values to be varied 
- run the the.psf file until successful completion 

2. Postprocessing
- prepare the python script for postprocessibg
- open all the odbs, loop through each of them
- loop through the steps in the odb
- extract results for U1,U2,U3,etc
- determine the worst expansion vaues and worstloadcases
- create excel workbook and worksheets using xlsxwriter library
- write results to the sheets 
- make a plot of expansion versus wall thickness
- parametric study completed





