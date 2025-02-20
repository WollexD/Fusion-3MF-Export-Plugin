#Author-Rau
"""Description-This Fusion 360 plugin exports
   each component as a separate .3mf file with high mesh refinement, 
   allowing users to select a destination folder for the exports. 
   It ensures efficient organization by keeping all bodies within their respective components.
"""

import adsk.core, adsk.fusion, adsk.cam, traceback
import os

def select_folder():
    ui = adsk.core.Application.get().userInterface
    folder_dlg = ui.createFolderDialog()
    folder_dlg.title = "Wähle den Exportordner"
    if folder_dlg.showDialog() == adsk.core.DialogResults.DialogOK:
        return folder_dlg.folder
    return None

def export_3mf():
    app = adsk.core.Application.get()
    ui = app.userInterface
    try:
        doc = app.activeDocument
        design = adsk.fusion.Design.cast(doc.products.itemByProductType('DesignProductType'))
        if not design:
            ui.messageBox('Kein aktives Design gefunden.')
            return

        exportMgr = design.exportManager
        if not hasattr(exportMgr, 'createC3MFExportOptions'):
            ui.messageBox('3MF-Export wird von dieser Fusion 360 API-Version nicht unterstützt.')
            return

        export_folder = select_folder()
        if not export_folder:
            ui.messageBox('Kein Ordner ausgewählt.')
            return

        for occ in design.rootComponent.occurrences:
            comp = occ.component
            file_path = os.path.join(export_folder, f"{comp.name}.3mf")
            export_options = exportMgr.createC3MFExportOptions(comp, file_path)
            exportMgr.execute(export_options)

        ui.messageBox('Export abgeschlossen.')
    except Exception as e:
        ui.messageBox(f'Fehler: {traceback.format_exc()}')

def run(context):
    export_3mf()
