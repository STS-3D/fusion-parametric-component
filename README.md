
# Overview
 This is a Fusion 360 addin to parametricly modify components with spreadsheet input. 
 This is a VERY early beta version, I have many imporovments planned, both to the Add-In functinality and code quality

### purpose
 I created this Add-In so I can easily create components that have many minor varations
 that would be otherwise too time consuming to model out induviduly
 This Add-In is especialy usfull when components are designed to be 3d printed


# Instalation

## download
1. Download this extension onto you computer, you may want to create a parent directy called "Fusion-Addins", but this is not required


## setup
1. Open Fusion 360
2. Go to the "Utilities" tab
3. Click on the "Add-Ins" Icon (you can also press "Shift S" to open Add-Ins), this opens the Add-Ins window
4. There are two tabs, "Scripts" and "Add-Ins", click on the "Add-Ins" tab
5. On the top of the Add-Ins list you will see "My Add-Ins", Click on the green "+" to add a new Add-In, a file dilog will apear
6. Select the the extension you downloaded from GitHub, and click "Open", this will load the Add-In under "My Add-Ins"
7. Select the Add-In (Parametric-Component) from the list of Add-Ins and click "Run", you should see a gray circle of dots next to the add in
8. The Add-In will now apear in the "Utilities" Tab
9. To restart the Add-In go back to the "Addins" window and click "Stop" then "Start"

# work flow
The basic work flow is as follows:
1. Design the master component in Fusion 360. This component will be modified with values from the spreadsheet.
    you should create create parameter names for dimensions/values that you want to have substituted for each new componenet
2. Create a CSV (Comma Seperated Values) spreadsheet in any program (Microsoft Excel, Apple Numbers, Text Edit, VIm, etc..)
    each row in the spreadsheet represents a new component configuration that will be created in Fusion 360.
    the columns header represet the object to be modified. They should be in the form of objectName.attribute, for example 
    all model parameters should be in the form of paramterName.expression (length.expression). if all columns are in this form
    the Add-In will automaticly match the column to the object in fusion. If the names don't match, you have to manuly select
    which object the column header should be mapped to
3. Once the spreadsheet and component are finished. run the Add-In, you will be prompted to select the master component and 
    sketch texts. Each attribute that can be modifed will have a dropdown menu where you can manuly map parameters.
    if you want to export as STL, select the check box on the bottom of the menu, this will propt you to select an 
    output location.
4. When you click "OK", the Add-In will create new components with values substituted from the spreadsheet.
    after you can delete these new components by click the Add-In drop down and clicking DeleteComponents
5.  In the Add-In drop down you can delete the new components by clicking DeleteComponents


# usage
1. In the "Utilities" tab click on the Add-In (Parametric-Component), this the Add-In window
2. Click on the "Select Spreadsheet" text box, this open a file dialog, select the CSV spreadsheet you created



## example 1
Imagine you want to create 10 cubes, each with a sidle length 5mm longer than the last, it will also have extruded text in the form of "Box Number N"
In fusion you will create a new design.
1.  Create a new Component, and rename it "Master-Cube"
2.  Create a sketch, in the sketch create a square, dimension one side to 100mm, use euality constraint to for the other sides
3.  Create text in the sketch, set the text to "text1"
4.  Extrude the sketch to 100mm
5.  Extrude the Text to 105mm, so the text is extends through the cube by 5mm
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
    the headers are in the form of Object.attribute. column 1 sets the name of the new component object.
    column 2 sets the expression (value) of the width model parameter. column 4 sets the text value
    of the SketchText object

8.  In Fusion 360, run the Add-In, click select spreadsheet and select the CSV spreadsheet you just created
9.  Select the componet when propted, then select the sketch text.
10. You should see 4 tables now, the there are dropdowns in the first 3 tables where you can manualy map object attributes 
    column headers
11. The bottom table display selected atributes
12. Click okay, you should see the components being created





## example 2
imagine you want to design electrical wire labels for a 3d printer
you have 2 different sizes of wire ( 16ga/1.2mm stepper motor wire, 18ga/1mm signal wire)
you have 6 motors, 4 wires per motor (x, y0, y1, z0, z1, extruder)
you have 4 limit switches, 3 wires each
The first motor wire lable would look like "M0 A+", wire 2 would be "M0 A-", etc... 
this means there will be a total of 36 different lables
```
The CSV spreadsheet would look like this:
Label-Comp.name    wire_diameter.Expression    wireText.text
Motor_0_A+         2                           M0 A+ 
Motor_1_A-         2                           M1 A-
Limit_x_s          1                           LX0 Sig
...                ...                         ...


```

