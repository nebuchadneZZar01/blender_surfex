# author: Michele Ferro
# mail: michele.ferro1998@libero.it
# github: github.com/nebuchadneZZar01

bl_info = {
    "name": "SurfEx",
    "author": "nebuchadnezzar",
    "version": (1, 1),
    "blender": (2, 90, 0),
    "location": "Scene > SurfEx",
    "description": "Calculates the total surface of a mesh and puts it in a CSV file.",
    "warning": "",
    "doc_url": "",
    "category": "Add Mesh",
}

# Importazioni delle librerie necessarie
# itertools contiene utili operatori di iterazione e ritaglio dei dizionari

import bpy
import itertools
import csv

#======================================================================
# GESTIONE DEL FILE CSV
# Funzione che gestisce la creazione del file CSV
csvFile = None
writer = None
meshesData = dict()

def createCSV(filepath): 
    global csvFile
    global writer
    
    csvFile = open(filepath, 'w', newline='')

    # compatibilità con tabulazione di excel per dati separati da virgola
    csvFile.write("sep=,\n")

    # lista dei campi del foglio csv
    fieldnames = ['oggetto', 'collection', 'nVert', 'nFace', 'totSurface', 'objType']
    
    writer = csv.DictWriter(csvFile, fieldnames=fieldnames)
    writer.writeheader()

    return csvFile, writer

def writeCSV(filepath):
    csvFile, writer = createCSV(filepath)

    try:
        for obj in meshesData:
            writer.writerow(meshesData[obj])
    except:
        print("Error in writer.writerow(): csv file was not created!")

    try:
        csvFile.close()
    except:
        print("Error in csv.close(): csv file was not created!")
    
#========================================================

#======================================================================
#               SCRIPT PRINCIPALE da eseguire
#        (tutte le funzioni create sopra servono qui!)
#======================================================================
def main(context):
    global csvFile
    global writer
    global meshesData

    #========================================================
    # CALCOLI!
    # carico nella lista objs tutte le mesh visibili presenti nella scena     
    objs = bpy.context.visible_objects.copy()
    
    # I calcoli vengono eseguiti per ogni oggetto visibile in scena
    # Inizialmente è necessario assicurarsi che i cambiamenti effettuati in
    # scena siano applicati, così che il data object della mesh sia aggiornato
    for obj in objs:
        if obj.type == "MESH":
            bpy.ops.object.transform_apply(location=True, rotation=True, scale=True)
            colName = obj.users_collection[0].name_full
            obj['Obj Type'] = ""
            obj_data = obj.data
            
            obj_verts = obj_data.vertices
            nVerts = len(obj_verts)

            # prendo il riferimento delle facce del data object (la mesh) in viewport
            obj_polygons = obj_data.polygons
            nFaces = len(obj_polygons)
            totSurface = 0
            
            # calcolo la superficie totale della mesh
            for p in obj_polygons:
                totSurface += p.area

            # --- RIEMPIMENTO DIZIONARIO ---
            meshInfo = {'oggetto': obj.name, 'collection': colName, 'nVert': nVerts, 'nFace': nFaces, 'totSurface': round(totSurface, 3), 'objType': obj['Obj Type']}

            meshesData[obj.name] = meshInfo


#======================================================================
#                 CLASSI DI GESTIONE DELLA UI BLENDER
# (tutto il codice utilizzato nel main verrà chiamati in queste classi
#  necessarie all'aspetto grafico dell'add-on)
#======================================================================

# --- CLASSI OPERATORE ---

# Incapsula l'intero main all'interno di un operatore Blender
class MainFlowOperator(bpy.types.Operator):
    bl_idname = "scene.main_flow"
    bl_label = "Calcola!"
    bl_description = "Calcola le principali caratteristiche geometriche delle mesh"

    def execute(self, context):
        main(context)
        return {'FINISHED'}

# Incapsula la creazione del file CSV all'interno di un operatore Blender
class CreateCSVOperator(bpy.types.Operator):
    bl_idname = "scene.select_dir"
    bl_label = "Esporta su file CSV"
    bl_description = "Esporta i file analizzati su un file CSV"
    filepath = bpy.props.StringProperty(subtype="DIR_PATH")
 
    def execute(self, context):
        writeCSV(self.filepath)
        return {'FINISHED'}
 
    def invoke(self, context, event):
        context.window_manager.fileselect_add(self)
        return {'RUNNING_MODAL'}

# --- CLASSI UI --- 

# Crea un elemento UI, in questo caso un pannello, che a sua volta
# contiene un pulsante che richiama l'operatore sopra definito
class LayoutPanel(bpy.types.Panel):
    """Creates a Panel in the scene context of the properties editor"""
    bl_label = "Total Surface Extractor"
    bl_idname = "SCENE_PT_layout"
    bl_space_type = 'PROPERTIES'
    bl_region_type = 'WINDOW'
    bl_context = "scene"              

    def draw(self, context):
        layout = self.layout

        scene = context.scene

        row = layout.row()
        column = layout.column()
        column.scale_y = 2.0
        row.operator("scene.main_flow")
        row.operator("scene.select_dir")

# Crea un elemento UI: visibile direttamente sulla scena 3D
# esso conterrà tutte le informazioni geometriche sulle rocce
class View3DPanel:
    bl_space_type = 'VIEW_3D'
    bl_region_type = 'UI'
    bl_category = "Total Surface Extractor"

    @classmethod
    def poll(cls, context):
        return (context.object is not None)

# Crea un elemento UI: visibile direttamente sulla scena 3D
# schede contenenti le informazioni geometriche
class InfoTab(View3DPanel, bpy.types.Panel):
    bl_idname = "VIEW3D_PT_test_1"
    bl_label = "Info"

    def draw(self, context):
        global meshesData
        obj = bpy.context.active_object
        if obj.name in meshesData:
            meshDict = meshesData[obj.name]
            meshDict = dict(itertools.islice(meshDict.items(), 5))
            for k, v in meshDict.items():
                self.layout.label(text = k + ": " + str(v))
            row = self.layout.row()
            row.prop(obj, '["Obj Type"]')
            meshesData[obj.name]['objType'] = obj['Obj Type']


# --- REGISTER/UNREGISTER ---
# La funzione register rende le due classi sopra definite disponibili
# alle funzionalità già presenti in Blender, così che possano essere
# integrate nel software.
# La funzione unregister fa l'esatto contrario.

def register():
    bpy.utils.register_class(MainFlowOperator)
    bpy.utils.register_class(CreateCSVOperator)
    bpy.utils.register_class(LayoutPanel)
    bpy.utils.register_class(InfoTab)

def unregister():
    bpy.utils.unregister_class(MainFlowOperator)
    bpy.utils.unregister_class(CreateCSVOperator)
    bpy.utils.unregister_class(LayoutPanel)
    bpy.utils.unregister_class(InfoTab)

if __name__ == "__main__":
    register()
