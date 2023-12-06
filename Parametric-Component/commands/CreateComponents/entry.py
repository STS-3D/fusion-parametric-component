#Author STS Innovations LLC #@sts_3dp
# STSI Select Component
import adsk.core
import os
from ...lib import fusion360utils as futil
from ... import config

import csv #for reading in local config data from spread sheet
import json
import math
from adsk.fusion import SketchText, Occurrence, ModelParameters, Component

app = adsk.core.Application.get()
ui = app.userInterface


# from default extension
CMD_NAME = os.path.basename(os.path.dirname(__file__))
CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_{CMD_NAME}'
CMD_Description = 'Create new components with spreadsheet input'
IS_PROMOTED = False

# Global variables by referencing values from /config.py
WORKSPACE_ID = config.design_workspace
TAB_ID = config.tools_tab_id
TAB_NAME = config.my_tab_name

PANEL_ID = config.my_panel_id
PANEL_NAME = config.my_panel_name
PANEL_AFTER = config.my_panel_after

CONFIG = config.MODEL_CONFIG_DATA

SPREADSHEET = config.SPREADSHEET
INPUT_TABLE_JSON= config.INPUT_TABLE_JSON
OUTPUT_TABLE_JSON = config.OUTPUT_TABLE_JSON
BASE_OCC= config.BASE_OCC

# Resource location for command icons, here we assume a sub folder in this directory named "resources".
ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')

# Holds references to event handlers
local_handlers = []


# table layout configs
table_configs = {

    'comp_table': {
        'table_id': 'comp_table',
        'row_attrs' : {'entityToken': 'entityToken', 'objectType': 'objectType', 'name': 'name'},
        # columns to display in input table
        'cols': [
            {'col_name': 'comp_name', 'attr':'name', 'selectable': 0 },
            {'col_name': 'name', 'attr': 'name', 'selectable': 1 },
            {'col_name': 'object_type', 'attr': 'objectType', 'selectable': 0 },
        ]
    },
    'param_table': {
        'table_id': 'param_table',
        'row_attrs' : {'entityToken': 'entityToken', 'objectType': 'objectType', 'name': 'name'},
        # columns to display in input table
        'cols': [
            {'col_name': 'name', 'attr':'name'},
            {'col_name': 'expression', 'attr':'expression','selectable': 1 },
            {'col_name': 'value', 'attr': 'value' },
            {'col_name': 'role', 'attr': 'role'},
        ]
    },
    'text_table': {
        'table_id': 'text_table',
        'row_attrs' : {'entityToken': 'entityToken', 'objectType': 'objectType', 'name': 'text'},
        # columns to display in input table
        'cols': [
            {'col_name': 'name', 'attr': 'text', 'selectable': 0 },
            {'col_name': 'text','attr': 'text', 'selectable': 1 },
            {'col_name': 'fontName', 'attr': 'fontName', 'selectable': 1 },
            {'col_name': 'height', 'attr': 'height' },
        ]
    },
    'output_table': {
        'table_id': 'output_table',
        'row_attrs' : {'entityToken': 'entityToken', 'objectType': 'objectType'},
        'cols': [
            {'col_name': 'object_name', 'display': 1, 'selectable': 0 },
            {'col_name': 'var_val', 'display': 1, 'selectable': 0 },
            {'col_name': 'object_type', 'display': 1, 'selectable': 0 },
        ]

    }


} # end table configs

comp_table_config = table_configs['comp_table']
param_table_config = table_configs['param_table']
text_table_config = table_configs['text_table']
output_table_config = table_configs['output_table']



# setting for component creation
OUTPUT_SETTINGS = {
    'save_stl':False,
    'save_dir':'',
    'create_new_components': True
}


def print(string):
    '''redefine print function'''
    if len(str(string)) == 0:
        return
    futil.log(str(string))

# Executed when add-in is run.
def start():
    # ******************************** Create Command Definition ********************************
    cmd_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_Description, ICON_FOLDER)

    # Add command created handler. The function passed here will be executed when the command is executed.
    futil.add_handler(cmd_def.commandCreated, command_created)

    # ******************************** Create Command Control ********************************
    # Get target workspace for the command.
    workspace = ui.workspaces.itemById(WORKSPACE_ID)

    # Get target toolbar tab for the command and create the tab if necessary.
    toolbar_tab = workspace.toolbarTabs.itemById(TAB_ID)
    if toolbar_tab is None:
        toolbar_tab = workspace.toolbarTabs.add(TAB_ID, TAB_NAME)

    # Get target panel for the command and and create the panel if necessary.
    panel = toolbar_tab.toolbarPanels.itemById(PANEL_ID)
    if panel is None:
        panel = toolbar_tab.toolbarPanels.add(PANEL_ID, PANEL_NAME, PANEL_AFTER, False)

    # Create the command control, i.e. a button in the UI.
    control = panel.controls.addCommand(cmd_def)

    # Now you can set various options on the control such as promoting it to always be shown.
    control.isPromoted = IS_PROMOTED


def export_to_stl(exportMgr, comp, file_name):
    # directory where stl files are saved
    stl_base_dir = OUTPUT_SETTINGS['save_dir']

    save_path = os.path.join(stl_base_dir, f'{file_name}.stl')
    # export the root compoennt

    stlExportOptions = exportMgr.createSTLExportOptions(comp, save_path)

    exportMgr.execute(stlExportOptions)


# Function called when a user clicks the corresponding button in the UI.
def command_created(args: adsk.core.CommandCreatedEventArgs):
    # Connect to the events that are needed by this command.
    futil.add_handler(args.command.execute, command_execute, local_handlers=local_handlers)
    futil.add_handler(args.command.inputChanged, command_input_changed, local_handlers=local_handlers)
    futil.add_handler(args.command.destroy, command_destroy, local_handlers=local_handlers)

    # inputs in dialog box
    inputs = args.command.commandInputs

    # spreadsheet checkbox
    select_spreadsheet = inputs.addBoolValueInput('select_spreadsheet', 'Select Spreadsheet', True)
    # name of selected spreadsheet
    spreadsheet_name_box = inputs.addTextBoxCommandInput('selected_spreadsheet', 'Spreadsheet', 'None', 1, True)


    # component selection
    comp_input = inputs.addSelectionInput('component_input', 'Component Selection', 'Select Component')
    comp_input.addSelectionFilter('Occurrences')
    comp_input.setSelectionLimits(0, 1)

    comp_text_box_val = f'<b style=color:"black";>Component Attributes:</b>'
    comp_name_box = inputs.addTextBoxCommandInput('comp_name_box', '', comp_text_box_val, 1, True)
    # Create param table input
    comp_table_input = create_table(inputs, table_configs['comp_table'])

    # param
    param_text_box_val = f'<b style=color:"black";>Model Parameters:</b>'
    param_name_text_box = inputs.addTextBoxCommandInput('param_name_box', '', param_text_box_val, 1, True)
    # Create param table input
    param_table_input = create_table(inputs, table_configs['param_table'])

    # Create a selection input, apply filters and set the selection limits
    sketch_text_input = inputs.addSelectionInput('sketch_text_input', 'Sketch Text Selection', 'Select Sketch Text')
    sketch_text_input.addSelectionFilter('Texts')
    sketch_text_input.setSelectionLimits(0, 0)

    sketch_text_box_val = f'<b style=color:"black";>Sketch Text Attributes:</b>'
    # sketch name box
    text_name_box = inputs.addTextBoxCommandInput('text_name_box', '', sketch_text_box_val, 1, True)

    # Create sketch text table input
    text_table_input = create_table(inputs, table_configs['text_table'])
    output_text_box_val = f'<b style=color:"black";>Selected Attributes:</b>'
    output_text_box = inputs.addTextBoxCommandInput('output_text_box', '', output_text_box_val, 1, True)
    # Create out_puttable
    output_table_input = create_table(inputs, output_table_config)

    # component creation control
    # spreadsheet checkbox, weather or not to make actual new components
    make_new_components = inputs.addBoolValueInput('create_new_components','Create New Components', True, '', True)

    # export to csv or not
    export_csv = inputs.addBoolValueInput('export_csv', 'Export as CSV', True)

    # number of components that will be created
    n_comps_text_box = inputs.addTextBoxCommandInput('n_comps_text_box', 'Components To Create:', '0', 1, True)


def select_output_dir():
    '''seclect local save location for spreadsheet'''

    # folder system dialog
    folderDialog = ui.createFolderDialog()

    folderDialog.title = "Select Output Directory"

    dialogResult = folderDialog.showDialog()

    if dialogResult == adsk.core.DialogResults.DialogOK:
        OUTPUT_SETTINGS['save_stl'] = True
        OUTPUT_SETTINGS['save_dir'] = folderDialog.folder

# prompt user to select spreadsheet
def select_spreadsheet():
    '''seclect local config spreadsheet'''
    fileDialog = ui.createFileDialog()
    fileDialog.isMultiSelectEnabled = False
    fileDialog.title = "Select Config CSV"
    fileDialog.filter = 'CSV files (*.csv)'
    fileDialog.filterIndex = 0
    dialogResult = fileDialog.showOpen()

    # okay clicked
    if dialogResult == adsk.core.DialogResults.DialogOK:
        filename = fileDialog.filename
        # store spreadsheet path
        SPREADSHEET['path'] = filename
    else:
        return 'None'

    # if 20 blank cells in a row are encountered assume row is finished (rest are blank)
    max_col_check = 20
    #max_col_check = 40
    with open(filename, 'r', newline='', encoding='utf-8-sig') as csvfile:
        csv_data = csv.reader(csvfile, delimiter=',', quotechar='|')

        # list col that contains any data
        start_col = 0
        spreadsheet_data = []
        for index, row in enumerate(csv_data):
            if index == 0:

                # iterate over row cells, prevent reading in blank cells
                for cell_index, c in enumerate(row):
                    if c != '':
                        end_col = cell_index
                    if cell_index > max_col_check:
                        break

                headers = row[:end_col]
                SPREADSHEET['col_headers'] = headers
            else:
                # prevent reading in empty rows
                row_sum = sum([len(c) for i,c in enumerate(row)])

                if row_sum == 0:
                    continue

                row_dict = {header: row[i] for i, header in enumerate(headers)}
                print(row_dict)
                spreadsheet_data.append(row_dict)

        SPREADSHEET['data'] = spreadsheet_data
        # used to set text box val
    return filename


def create_table(inputs, table_config):
    '''create initial table_input, values filled in once spreadsheet is selected'''

    table_id = table_config['table_id']

    # n visable columns
    n_vis_cols = len([c for c in table_config['cols'] if c.get('display') == 1 ])
    col_spacing = f"{'1:'*n_vis_cols}".strip(':')

    table_input = inputs.addTableCommandInput(table_id, table_id, n_vis_cols, col_spacing)

    return table_input


def update_base_occ(occ_entity):
    '''update global object which stores component entityToken'''
    BASE_OCC['entityToken'] = occ_entity.entityToken


def create_table_json(selections: list, table_config: dict, spreadsheet=SPREADSHEET):
    '''
    selection: selected objects (component occurence, sketchTexts)
    create json dict that window tables will read values from

    '''
    # sub entity attrs
    sub_ent_config = table_config['cols']

    # local spreadsheet headers
    spreadsheet_headers = SPREADSHEET.get('col_headers')
    if spreadsheet_headers == None:
        ui.messageBox('Please Select Spreadsheet First', 'Select Spreadsheet First')
        return None

    ent_iter = []
    # iterate over entities passed in, handle for sever different types of objects
    for ent_index, ent in enumerate(selections, 1):
        if isinstance(ent, ModelParameters):
            params = ent
            n_params = params.count
            # iterable of sub entities 
            ent_iter += [params.item(i) for i in range(n_params)]
        elif isinstance(ent, SketchText):
            # iterable of sub entities 
            ent_iter += [ent]
        elif isinstance(ent, Component):
            # iterable of sub entities 
            ent_iter += [ent]

    # table rows
    row_list = []
    # sub entities (modelParameter, sketchText) 
    for sub_ent in ent_iter:
        # column dict, contains object attr key, info about display and visablity for gui table

        # cells in row
        cells = []
        # dict containing both cells and meta info for object
        row_dict = {}
        # index/id values for row (entiry)
        for k, v in table_config['row_attrs'].items():
            row_dict[k] = getattr(sub_ent, v)

        # val info for displayed cells
        for col_dict in sub_ent_config:
            # "cell"
            cell_dict = {}

            for k, v in col_dict.items():
                if k == 'attr':
                    attr_val = getattr(sub_ent, v)
                    cell_dict[k] = attr_val
                elif (k == 'selectable') and (v==1):
                    cell_dict[f'selection'] = spreadsheet_headers
                    cell_dict[f'selectable'] = 1
                #val_dict[['col_name'] = col_dict['col_name']
                else:
                    cell_dict[k] = v

            # add to list of cells
            cells.append(cell_dict)

        row_dict['cells'] = cells

        # add row dict (inluding cells) to list of rows
        row_list.append(row_dict)

    table_json = {'table_id': table_config['table_id'], 'table_rows':row_list}

    #print(json.dumps(row_list, indent=4))
    INPUT_TABLE_JSON[table_config['table_id']] = table_json

    return table_json

def update_output_json(changed_input, input_table_data=INPUT_TABLE_JSON) -> None:
    '''
    updates output table
    get input_json table when config is dropdown value is changed
    called from input_change, and input table update
    '''

    design = app.activeProduct #stsi
    #print(f'id: {changed_input.id}')
    cell_type, table_id, col_name, row, col = changed_input.id.split('__')

    row, col = int(row)-1, int(col)

    # value of selected item
    current_cell_val = changed_input.selectedItem.name

    input_row = input_table_data[table_id]['table_rows'][row]

    # row base object entity token/ name
    entityToken = input_row['entityToken']
    objectType = input_row['objectType']

    name = input_row['name']

    # cells in row, list
    cell = input_row['cells'][col]

    # inital cell vall
    initial_cell_val = cell['attr']

    output_row_key = f'{col_name}_{row}'

    if current_cell_val != initial_cell_val:
        # base entity of changed val
        entity = design.findEntityByToken(entityToken)[0]

        output_dict  = {'name': name, 'objectType': objectType, 'attr': col_name, 'cell_val': current_cell_val, 'entityToken': entityToken}
        OUTPUT_TABLE_JSON[output_row_key] = output_dict


    # if current cell val has not changed dont 
    elif current_cell_val == initial_cell_val:
        if OUTPUT_TABLE_JSON.get(output_row_key):
            del OUTPUT_TABLE_JSON[output_row_key]


def fill_input_table(table_input: adsk.core.TableCommandInput, table_json={}):
    '''
    fill in table data
    '''
    # Get the CommandInputs object associated with the parent command.
    inputs = adsk.core.CommandInputs.cast(table_input.commandInputs)
    drop_down_style = adsk.core.DropDownStyles.LabeledIconDropDownStyle

    table_id = table_json['table_id']
    table_rows = table_json['table_rows']

    # create header row if no rows in table
    if table_input.rowCount == 0: # create headers of no rows in table
        index = 0
        for cell in table_rows[0]['cells']:
            color = 'black'
            selectable = cell.get('selectable')
            if selectable == 1:
                color = 'red'
            column_name = f'<b style=color:"{color}";>{cell["col_name"]}</b>'
            text_input = inputs.addTextBoxCommandInput(f'header_{index}', '', column_name, 1, True)
            table_input.addCommandInput(text_input, 0, index)
            index +=1


    # create by rows
    for row_index, row in enumerate(table_rows):
        row_index +=1

        # param name/ initial sketch text val etc
        row_name = row['name']
        for col_index, cell in enumerate(row['cells']):
            # expression, fontName, etc
            col_name = cell['col_name']

            # used to identify json table when dropdown input is changed
            cell_id = f'{table_id}__{col_name}__{row_index}__{col_index}'
            cell_val = cell['attr']

            if cell.get('selectable') == 1:
                # add spreadsheet param data
                cell_input = inputs.addDropDownCommandInput(f'dropdown__{cell_id}', f'{cell_val}', drop_down_style)

                dropdown_items = cell_input.listItems

                # if row name found in spreadsheet headers set to True, dont select inital val then
                preselected = False

                # spreadsheet_header
                for header in cell['selection']:
                    # spreadsheet_header
                    if header == f'{row_name}.{col_name}':
                        selected = True
                        preselected = True
                    else:
                        selected = False
                    dropdown_items.add(header, selected, '')

                #if row name was not found in drodown/ spreadsheet_headers
                if preselected == False:
                    selected = True
                else:
                    selected = False
                # add in default val to list
                dropdown_items.add(cell_val, selected,'', 0)

                # Add the inputs to the table.
                table_input.addCommandInput(cell_input, row_index, col_index, 0, 0)
                # update output table
                update_output_json(cell_input)

            if cell.get('selectable') != 1:
                # create text box, this the "cell"
                cell_input = inputs.addTextBoxCommandInput(f'{cell_id}', formattedText=f'{cell_val}', name=f'{cell_val}', numRows=1, isReadOnly=True)

                # Add the inputs to the table.
                table_input.addCommandInput(cell_input, row_index, col_index, 0, 0)



    # expand table so no scroll bar
    if table_input.rowCount > table_input.maximumVisibleRows:
        table_input.maximumVisibleRows = table_input.rowCount

def fill_output_table(table_input: adsk.core.TableCommandInput, table_json={}):
    '''
    Add rows to output table
    table json, is the row data from the 3 input tables
    '''

    inputs = adsk.core.CommandInputs.cast(table_input.commandInputs)

    if len(table_json.keys()) == 0:
           return
    table_rows = [v for k, v in table_json.items()]

    headers = [k for k in table_rows[0].keys()]

    # make headers
    for i, h in enumerate(headers):
        if h == 'entityToken':
            continue
        column_name = f'<b style=color:"black";>{h}</b>'
        cell_input = inputs.addTextBoxCommandInput(f'{h}', formattedText=column_name, name=f'{h}', numRows=1, isReadOnly=True)
        # Add the inputs to the table.
        table_input.addCommandInput(cell_input, 0, i, 0, 0)

    for row_index, (row_k, row) in enumerate(table_json.items()):
        # off set for header row
        row_index += 1

        for col_index, (k,v) in enumerate(row.items()):
            if k == 'entityToken':
                continue
            cell_id = f'{row_index}_{col_index}'
            cell_val = f'{v}'
            cell_input = inputs.addTextBoxCommandInput(f'{cell_id}', formattedText=f'{cell_val}', name=f'{cell_val}', numRows=1, isReadOnly=True)

            # Add the inputs to the table.
            table_input.addCommandInput(cell_input, row_index, col_index, 0, 0)


    # expand table so no scroll bar
    if table_input.rowCount > table_input.maximumVisibleRows:
        table_input.maximumVisibleRows = table_input.rowCount




# This function will be called when the user changes anything in the command dialog.
def command_input_changed(args: adsk.core.InputChangedEventArgs):
    changed_input = args.input
    inputs = args.input.parentCommand.commandInputs

    comp_input: adsk.core.SelectionCommandInput = inputs.itemById('component_input')
    sketch_text_input: adsk.core.SelectionCommandInput = inputs.itemById('sketch_text_input')

    # model parameters
    param_table_input = inputs.itemById(param_table_config['table_id'])
    comp_table_input = inputs.itemById(comp_table_config['table_id'])
    # sketch text 
    text_table_input = inputs.itemById(text_table_config['table_id'])

    output_table_input = inputs.itemById('output_table')

    # number of components to modify text value, updated when spreadsheed is selected
    n_comps_text_box = inputs.itemById('n_comps_text_box')

    # select spreadsheet
    if changed_input.id == 'select_spreadsheet':
        spreadsheet_select = inputs.itemById('select_spreadsheet')
        spreadsheet_text_box = inputs.itemById('selected_spreadsheet')
        select_val = spreadsheet_select.value

        if select_val == True:
            # fills in spreadsheet data
            file_path = select_spreadsheet()
            # spreadsheet text box name
            spreadsheet_text_box.formattedText = file_path

            # set text box val for number of rows in spreadsheet
            n_comps_text_box.formattedText = str(len(SPREADSHEET['data']))

    # clear data from tables
    if changed_input.id == 'param_table_clear':
        param_table_input.clear()

    if changed_input.id == 'text_table_clear':
        text_table_input.clear()

    # save dir checkbox
    if changed_input.id == 'export_csv':
        # sets output dir location for stls
        if changed_input.value == True:
            select_output_dir()
            OUTPUT_SETTINGS['save_stl'] = True
        else:
            OUTPUT_SETTINGS['save_stl'] = False
            OUTPUT_SETTINGS['stl_dir'] = ''

    # create components checkbox
    if changed_input.id == 'create_new_components':
        if changed_input.value == True:
            OUTPUT_SETTINGS['create_new_components'] = True
        else:
            OUTPUT_SETTINGS['create_new_components'] = False

    # component
    if changed_input.id == 'component_input':
        if comp_input.selectionCount > 0:

            comp_selections = [comp_input.selection(s).entity.component for s in range(comp_input.selectionCount)]
            comp_table_json = create_table_json(comp_selections, comp_table_config)

            # handle for no spreadsheet dialog box
            if comp_table_json is None:
                return

            fill_input_table(comp_table_input, comp_table_json )
            # param table
            param_selections = [comp_input.selection(s).entity.component.modelParameters for s in range(comp_input.selectionCount)]
            # after component selected, update global entityToken, used to identify component when deleteing
            # base occ
            update_base_occ(comp_input.selection(0).entity)
            param_table_input.clear()

            # json date to fill in input param table
            param_table_json = create_table_json(param_selections, param_table_config)
            # handle for no spreadsheet dialog box
            if param_table_json is None:
                return
            fill_input_table(param_table_input, param_table_json )
            output_table_input.clear()

            # call function here to fill table with initialy matched cells
            fill_output_table(output_table_input, OUTPUT_TABLE_JSON)

    # when sketch text is selected
    if changed_input.id == 'sketch_text_input':
        if sketch_text_input.selectionCount > 0:
            selections = [sketch_text_input.selection(s).entity for s in range(0, sketch_text_input.selectionCount)]
            text_table_input.clear()

            text_table_json = create_table_json(selections, text_table_config)

            if text_table_json is None:
                return
            fill_input_table(text_table_input, text_table_json)
            output_table_input.clear()
            # call function here to fill table with initialy matched cells
            fill_output_table(output_table_input, OUTPUT_TABLE_JSON)

    # when dropdown input is changed
    if changed_input.id.split('_')[0] == 'dropdown':
        update_output_json(changed_input)
        output_table_input.clear()
        fill_output_table(output_table_input, OUTPUT_TABLE_JSON)



def make_components(SPREADSHEET=SPREADSHEET):
    '''make new components'''

    # current active design
    design = app.activeProduct

    # root component of the active design.
    root_comp = design.rootComponent

    # master occurrence to be copied
    master_occ = design.findEntityByToken(BASE_OCC['entityToken'])[0]
    master_comp = master_occ.component

    # initial param values storred untill end to reset master comp
    initial_vals ={}

    # spreadsheet list of unique component configs
    config_list = SPREADSHEET['data']

    # number of new components to make
    n_confs = len(config_list)

    # used rounded sqrt to place components in a grid
    output_row_len = round(math.sqrt(n_confs))

    # list of created occurences
    occ_list = []
    occ_list.append(master_occ)

    # make new components
    for index, row in enumerate(config_list):
        # set new attributes on master comp before copy

        comp_name = None
        for k, v in OUTPUT_TABLE_JSON.items():
            # sketchText object, modelParameter, etc
            attr_object = design.findEntityByToken(v['entityToken'])[0]

            attr_name = v['attr']

            # save initial component vals, reset to these later
            if index == 0:
                initial_vals[v['cell_val']] = getattr(attr_object, attr_name)
            # new object val
            new_val = row[v['cell_val']]

            if attr_name == 'name':
                comp_name = new_val
                continue

            if hasattr(attr_object, attr_name):
                try:
                    setattr(attr_object, attr_name, new_val)
                except Exception as e:
                    print(f'ERROR: {e}, NAME: {attr_name}, VAL: {new_val}')
            else:
                print(f'ERROR: ATTR {attr_name} NOT FOUND: {new_val}')


        new_occ = root_comp.occurrences.addNewComponentCopy(master_comp, adsk.core.Matrix3D.create())

        # set name after so fusion doesnt auto number as duplicate
        if comp_name:
            new_occ.component.name = comp_name

        # add new comp occurrence to list, used to arrange and export components
        occ_list.append(new_occ)



    # reset comp to initial values
    for k, v in OUTPUT_TABLE_JSON.items():
        # sketchText object, modelParameter, etc
        attr_object = design.findEntityByToken(v['entityToken'])[0]
        attr_name = v['attr']
        new_val = initial_vals[v['cell_val']]
        setattr(attr_object, attr_name, new_val)

    # recompute new parametric vals
    design.computeAll()

    # list of newly created componenet occurences
    return occ_list


def arange_comps(occ_list):
    '''arange_components'''

    # active design
    design = app.activeProduct #stsi

    # bounds for master occ
    row_max_comp_height = 0
    bodies = occ_list[0].bRepBodies
    n_bodies = bodies.count
    body_boxes = [bodies.item(i).boundingBox for i in range(n_bodies)]
    master_body_min_x = min([b.minPoint.x for b in body_boxes])
    master_body_max_x = max([b.maxPoint.x for b in body_boxes])
    master_body_min_y= min([b.minPoint.y for b in body_boxes])
    master_body_max_y = max([b.maxPoint.y for b in body_boxes])

    master_body_x_len = abs(master_body_max_x - master_body_min_x)
    master_body_y_len = abs(master_body_max_y - master_body_min_x)
    # spreadsheet list of unique component configs

    n_comps = len(occ_list)
    # used rounded sqrt to place components in a grid
    output_row_len = round(math.sqrt(n_comps))
    x_trans = 0
    #initial_y_trans = abs(comp_y_max)
    y_bottom = master_body_max_y
    x_gap = 0
    y_gap = 0

    for index, occ in enumerate(occ_list[1:]):

        # bodies in occurrence
        bodies = occ.bRepBodies
        n_bodies = bodies.count

        body_boxes = [bodies.item(i).boundingBox for i in range(n_bodies)]
        body_min_x = min([b.minPoint.x for b in body_boxes])
        body_max_x = max([b.maxPoint.x for b in body_boxes])
        body_min_y= min([b.minPoint.y for b in body_boxes])
        body_max_y = max([b.maxPoint.y for b in body_boxes])
        body_len_x = abs(body_max_x - body_min_x)
        body_len_y = abs(body_max_y - body_min_y)

        # new row creation, above current row
        if index % output_row_len == 0:
            # bottom edge of master comp
            y_bottom += row_max_comp_height
            # start each new row aligned to master_body_min_x left edge
            x_trans = master_body_min_x

        ## move half the new comp width 
        comp_x_trans = x_trans - (body_min_x)

        ## get tallest y of comps in row
        row_max_comp_height = max(body_len_y, row_max_comp_height)
        comp_y_trans = y_bottom + (-body_min_y )

        vector = adsk.core.Vector3D.create(comp_x_trans, comp_y_trans, 0.0)
        transform = adsk.core.Matrix3D.create()
        transform.translation = vector
        occ.transform2 = transform

        # set max row height for first element in each row
        if index % output_row_len == 0:
            row_max_comp_height = body_len_y

        # adjust x transition for next occurrence
        x_trans += ( (body_len_x) + x_gap/2)

        # capture position
        try:
           design.snapshots.add()
        except Exception as e:
           print(str(e))


def save_stls(occ_list):
    '''save new components as stl'''

    design = app.activeProduct

    # create a single exportManager instance, for STL file export
    exportMgr = design.exportManager

    for index, occ in enumerate(occ_list[1:]):
        occ_name = occ.component.name

        export_to_stl(exportMgr, occ, occ_name)

        print(f'SAVED STL: {occ_name}')




# This function will be called when the user clicks the OK button in the command dialog.
def command_execute(args: adsk.core.CommandEventArgs):
    # run function to create components with substituted parameters/ save stl

    new_occs = make_components()

    arange_comps(new_occs)

    # save st;
    if OUTPUT_SETTINGS['save_stl'] == True:
        save_stls(new_occs)

    config.MODEL_CONFIG_DATA = {}
    config.SPREADSHEET = {}
    # data used to create output table of seleted parameters
    config.INPUT_TABLE_JSON = {}
    config.OUTPUT_TABLE_JSON = {}


# Executed when add-in is stopped.
def stop():
    # Get the various UI elements for this command
    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    panel = workspace.toolbarPanels.itemById(PANEL_ID)
    toolbar_tab = workspace.toolbarTabs.itemById(TAB_ID)
    command_control = panel.controls.itemById(CMD_ID)
    command_definition = ui.commandDefinitions.itemById(CMD_ID)

    # Delete the button command control
    if command_control:
        command_control.deleteMe()

    # Delete the command definition
    if command_definition:
        command_definition.deleteMe()

    # Delete the panel if it is empty
    if panel.controls.count == 0:
        panel.deleteMe()

    # Delete the tab if it is empty
    if toolbar_tab.toolbarPanels.count == 0:
        toolbar_tab.deleteMe()



# This function will be called when the user completes the command.
def command_destroy(args: adsk.core.CommandEventArgs):
    global local_handlers
    local_handlers = []
    futil.log(f'{CMD_NAME} Command Destroy Event')

































