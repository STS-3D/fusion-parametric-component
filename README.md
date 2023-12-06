
# Overview
 Fusion 360 Add-In to parametricly modify components with spreadsheet input. 
 This is a VERY early beta version, I have many improvements planned, both to the Add-In functinality and code quality.
 Please let me know about any bugs/ sugested features.

### purpose
 I created this Add-In so I can easily create components that have many minor varations
 that would be otherwise be too time consuming to model out individually
 This Add-In is especialy usefull when components are designed to be 3d printed

# Instalation
## Download
1. Download this extension onto you computer, either by downloading the zip file, or running: `git clone https://github.com/STS-3D/fusion-parametric-component.git` in terminal


## Setup
1. Open Fusion 360
2. Go to the "Utilities" tab
3. Click on the "Add-Ins" Icon (you can also press "Shift S" to open Add-Ins), this opens the Add-Ins window
4. There are two tabs, "Scripts" and "Add-Ins", click on the "Add-Ins" tab
5. On the top of the Add-Ins list you will see "My Add-Ins", Click on the green "+" to add a new Add-In, a file dilog will appear
6. Navigate to the extension, SELECT THE FOLDER INSIDE FOLDER, (Parametric-Component), and click "Open", this will load the Add-In under "My Add-Ins"
7. Select the Add-In (Parametric-Component) from the list of Add-Ins and click "Run", you should see a gray circle of dots next to the add in
8. The Add-In will now appear in the "Utilities" Tab
9. To restart the Add-In, go back to the "Addins" window and click "Stop" then "Start"

# Work flow
The basic work flow is as follows:
1. Design the master component in Fusion 360. This component will be modified with values from the spreadsheet.
    you should create create parameter names for dimensions/values that you want substituted for each new componenet
2. Create a CSV Comma Seperated Values) spreadsheet in any program (Microsoft Excel, Apple Numbers, Text Edit, VIm, etc..)
    each row in the spreadsheet represents a new component configuration that will be created in Fusion 360.
    the columns header represet the object to be modified. They should be in the form of objectName.attribute, for example 
    all model parameters should be in the form of paramterName.expression (length.expression). If all columns are in this form
    the Add-In will automaticly match the column to the object in fusion. If the names don't match, you have to manually select
    which object the column header is mapped
3. Once the spreadsheet and master-component are finished run the Add-In, you will be prompted to select the spreadsheet,  master component and 
    sketch texts. Each attribute that can be modifed will have a dropdown menu where you can manually map parameters.
    if you want to export the new components as an STL, select the check box on the bottom of the menu, this will prompt you to select an 
    output directory.
4. When you click "OK", the Add-In will create the new components with values substituted from the spreadsheet.
    after, you can delete these new components by clicking the Add-In dropdown and clicking "DeleteComponents"


# Examples
## Example 1
Imagine you want to create 10 cubes, each with a sidle length 5mm longer than the last, it will also have extruded text in the form of "Box Number N"
In fusion you will first create a new design.
1.  Create a new Component, and rename it "Master-Cube"
2.  Create a sketch, in the sketch create a square, dimension one side to 100mm, use the euality constraint to set the other sides equal to the 100m side
3.  Create text in the sketch, set the text to "text1"
4.  Extrude the sketch to 100mm
5.  Extrude the Text to 105mm, so the text is extends through the cube by 5mm, make sure you use a join operation with the extrude
6.  In the "Solid" tab, click on "Modify" then click "Change Parameters" expand the Skecth and Extrude features, set the first 100mm dimension 
 .  parameter name to "width", set the other 100mm dimension parameters equal to the parameter name "=width"
 .  set the 105mm extruseion param name (the text) to "text_depth"
7.  create a new CSV preadsheet, with column headers mapped to the parameters

### Spreadsheet:
```
    Master-Cube.name	width.expression	text_depth.expression	text1.text
    Cube-Comp-1	        105	                110	                    Cube 1
    Cube-Comp-2	        110	                115	                    Cube 2
    Cube-Comp-3	        115	                120	                    Cube 3
    ...                 ...                 ...                     ...
    Cube-Comp-10        150	                155	                    Cube 10
```
    The headers are in the form of Object.attribute. column 1 sets the name of the new component object.
    column 2 sets the expression (value) of the width model parameter. Column 4 sets the text value
    of the SketchText object.

8.  In Fusion 360, run the Add-In, click "Select Spreadsheet" and select the CSV spreadsheet you created
9.  Select the component when prompted, then select the sketch text
10. You should see 4 tables now, there are dropdowns in the first 3 tables where you can manually map object attributes to
    column headers
11. The bottom table displays selected atributes to will be substituted for each new component
12. Click okay, you should see the components being created



