

# STSI
# Assuming you have not changed the general structure of the template no modification is needed in this file.
from . import commands
from .lib import fusion360utils as futil
import adsk.core


def run(context):
    try:
        # Display a message when the add-in is manually run.
        if not context['IsApplicationStartup']:
            app = adsk.core.Application.get()
            ui = app.userInterface

            futil.log('A new Parametric Component Addin panel containing several commands has been added to the "UTILITIES" tab.')

            #ui.messageBox('A new Parametric Component Addin panel containing several commands has been added to the "UTILITIES" tab.', 'STSI')

        # This will run the start function in each of your commands as defined in commands/__init__.py
        commands.start()

    except:
        futil.handle_error('run')


def stop(context):
    try:
        # Remove all of the event handlers your app has created
        futil.clear_handlers()

        # This will run the start function in each of your commands as defined in commands/__init__.py
        commands.stop()

    except:
        futil.handle_error('stop')
