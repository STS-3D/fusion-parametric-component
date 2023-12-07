# STSI DeleteComponents
import adsk.core
import os
from ...lib import fusion360utils as futil
from ... import config

import csv #for reading in local config data from spread sheet
import json
import random
from adsk.fusion import SketchText, Occurrence
from adsk.core import MessageBoxButtonTypes, ObjectCollection

app = adsk.core.Application.get()
ui = app.userInterface

# default 

# coippied from default extension
CMD_NAME = os.path.basename(os.path.dirname(__file__))
CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_{CMD_NAME}'
CMD_Description = 'Delete previously created components'
IS_PROMOTED = False

# Global variables by referencing values from /config.py
WORKSPACE_ID = config.design_workspace
TAB_ID = config.tools_tab_id
TAB_NAME = config.my_tab_name

PANEL_ID = config.my_panel_id
PANEL_NAME = config.my_panel_name
PANEL_AFTER = config.my_panel_after


# Resource location for command icons, here we assume a sub folder in this directory named "resources".
ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')

# Holds references to event handlers
local_handlers = []


def print(string):
    '''redefine print function'''
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




# Function to be called when a user clicks the corresponding button in the UI.
def command_created(args: adsk.core.CommandCreatedEventArgs):
    # Connect to the events that are needed by this command.
    futil.add_handler(args.command.execute, command_execute, local_handlers=local_handlers)
    futil.add_handler(args.command.inputChanged, command_input_changed, local_handlers=local_handlers)
    futil.add_handler(args.command.destroy, command_destroy, local_handlers=local_handlers)

    # inputs in dialog box
    inputs = args.command.commandInputs

def delete_components():
    '''make new components'''

    design = app.activeProduct #stsi
    # root component of the active design.

    new_occs = config.SHARED_REFS.get('new_occs')
    if new_occs:
        # object collection to delete
        to_delete = ObjectCollection.createWithArray(new_occs)
        design.deleteEntities(to_delete)
        config.SHARED_REFS['new_occs'] == None

    # recompute new parametric vals
    design.computeAll()


# This function will be called when the user changes anything in the command dialog.
def command_input_changed(args: adsk.core.InputChangedEventArgs):
    changed_input = args.input
    inputs = args.input.parentCommand.commandInputs



# This function will be called when the user clicks the OK button in the command dialog.
def command_execute(args: adsk.core.CommandEventArgs):

    # prompt user to confim delete
    button_val = ui.messageBox('Delete New Components?', 'Warning', MessageBoxButtonTypes.OKCancelButtonType)

    if button_val == 0:
        delete_components()

# This function will be called when the user completes the command.
def command_destroy(args: adsk.core.CommandEventArgs):
    global local_handlers
    local_handlers = []
    futil.log(f'{CMD_NAME} Command Destroy Event')

































