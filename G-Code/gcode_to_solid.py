import win32com.client as wc
# install win32com by using pip install pywin32
import time
import numpy as np
from gcodeparser import GcodeParser

from collections import defaultdict
from uuid import uuid1

import json
import os

import concurrent.futures as cf
import win32api,win32process,win32con
import psutil

import sys

# inventor assingn a value to these Methods
PART_OBJ_VALUE = 12290
ASM_OBJ_VALUE = 12291
METRIC = 8962
XY = 3
XZ = 2
YZ = 1
JOIN = 20481
NEW_BODY = 20485
PATH_SWEEP_TYPE = 104449
OPTIMIZED_COMPUTE = 47363
EXTRUDE_POSITIVE_DIRECTION = 20993
EXTRUDE_NEGATIVE_DIRECTION = 20994
EXTRUDE_SYMETRIC = 20995 


class Inventor:
    def __init__(self):
        self.application = wc.DispatchEx("Inventor.Application")
        # create a part document
        self.part_doc = self.application.Documents.Add(
            PART_OBJ_VALUE, 
            self.application.FileManager.GetTemplateFile(PART_OBJ_VALUE, METRIC)
        )
        # inventor api object model see 
        # https://damassets.autodesk.net/content/dam/autodesk/www/pdfs/Inventor2022ObjectModel.pdf
        # decreases runtime by activating silent mode
        self.application.silentoperation = True
        self.application.screenupdating = False
        self.application.Visible = False

        self.part_comp_def = self.part_doc.ComponentDefinition
        self.transient_geometry = self.application.TransientGeometry

        self.sketches = self.part_comp_def.Sketches
        self.sketches3d = self.part_comp_def.Sketches3D             
        self.workplanes = self.part_comp_def.Workplanes
        self.xy_plane = self.part_comp_def.WorkPlanes.Item(XY)

        self.features = self.part_comp_def.Features
        self.extrude_features = self.features.ExtrudeFeatures
        self.sweep_features = self.features.SweepFeatures
        self.body = None


        try:
            for i in psutil.pids():
                if psutil.Process(i).name() == "Inventor.exe":
                    self.setpriority(psutil.Process(i).pid, 5)
        except:
            # priorty setting might decrease runtime
            pass

    
    def derive_part(self, file_path):
        derived_asm_comp = self.part_comp_def.ReferenceComponents.DerivedAssemblyComponents
        derived_def = derived_asm_comp.CreateDefinition(file_path)
        derived_def.IndependentSolidsOnFailedBoolean = True
        derived_part = derived_asm_comp.Add(derived_def)
        derived_part.BreakLinkToFile()
        
    def cleanup(self, file_name, inventor_part_folder):
        list_class_hide = [self.sketches, self.sketches3d, self.workplanes]
        for class_hide in list_class_hide:
            for item in class_hide:
                item.visible = False

        self.part_doc.SaveAs(inventor_part_folder + "/" + file_name + ".ipt", False)
        self.part_doc.Close(SkipSave = True) 
        self.application.Quit()

    def setpriority(self, pid=None, priority=1):
        priorityclasses = [win32process.IDLE_PRIORITY_CLASS,
                        win32process.BELOW_NORMAL_PRIORITY_CLASS,
                        win32process.NORMAL_PRIORITY_CLASS,
                        win32process.ABOVE_NORMAL_PRIORITY_CLASS,
                        win32process.HIGH_PRIORITY_CLASS,
                        win32process.REALTIME_PRIORITY_CLASS]
        if pid == None:
            pid = win32api.GetCurrentProcessId()
        handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, True, pid)
        win32process.SetPriorityClass(handle, priorityclasses[priority])


    def make_3_point_workplane(self, origin, normal_vector):  
        point_workplane = [origin, 
                        origin + np.cross(normal_vector, [0, 0, 1]),
                        origin + np.array([[0,0,1]])]
        

        sketch3d = self.sketches3d.Add()
        sketch3d.DeferUpdates = True

        workpoints = []
        for i in range(3):
            transient_point = self.transient_geometry.CreatePoint(*point_workplane[i][0])
            workpoints.append(sketch3d.SketchPoints3D.Add(transient_point))

        sketch3d.DeferUpdates = False

        workplane = self.workplanes.AddByThreePoints(*workpoints) 

        return workplane

    def create_profile(self, normal_vector, origin, width):
        points_profile = [[0, layer_height/2], [width/2, 0]]

        workplane = self.make_3_point_workplane(origin, normal_vector)

        # Add sketch and connecting points to profile
        sketch = self.sketches.Add(workplane)
        sketch.DeferUpdates = True

        points = []
        for point in points_profile:
            tg = self.transient_geometry.CreatePoint2d(*point)
            points.append(sketch.SketchPoints.Add(tg))

        sketch.SketchLines.AddAsTwoPointCenteredRectangle(*points)

        sketch.DeferUpdates = False

        return sketch

    def generate_path(self, points):
        transient_points = []
        for point in points:
            transient_points.append(self.transient_geometry.CreatePoint2d(*point))

        sketch = self.sketches.Add(self.xy_plane)
        sketch.DeferUpdates = True

        for i in range(len(points)-1):
            sketchline = sketch.SketchLines.AddByTwoPoints(transient_points[i], transient_points[i+1])
        
        sketch.DeferUpdates = False

        path = self.features.CreatePath(sketchline)
        return path


    def sweep(self, profile, path,
            sweeptype = PATH_SWEEP_TYPE, boolean_operation=NEW_BODY):
        solid_profile = profile.Profiles.AddForSolid()

        sweep_def = self.sweep_features.CreateSweepDefinition(sweeptype, solid_profile, path, boolean_operation)
        solid = self.sweep_features.Add(sweep_def)

        return solid

    def extrude(self, profile, length, adjustment_amount=0.02, boolean_operation=NEW_BODY):
        solid_profile = profile.Profiles.AddForSolid()

        extrude_def = self.extrude_features.CreateExtrudeDefinition(
            solid_profile,
            boolean_operation
        )
        extrude_def.SetDistanceExtent(length + adjustment_amount, EXTRUDE_NEGATIVE_DIRECTION)
        extrude_def.SetDistanceExtentTwo(adjustment_amount)
        solid = self.extrude_features.Add(extrude_def)

        return solid
    
    def create_feature(self, added_to_edges, list_points, width, boolean_operation=JOIN):
        for points in list_points:
            normal_vector = np.array(points[1]+[0])-np.array(points[0]+[0])
            origin = np.array([points[0]+[0]])
            profile = self.create_profile(normal_vector, origin, width)
            try:
                path = self.generate_path(points)
                self.sweep(profile, path, sweeptype=PATH_SWEEP_TYPE, 
                    boolean_operation=boolean_operation)
            except:
                # if sweep doesn't work then extrude is used it is a rough approximation 
                # but is enough because this exeption occurs mostly inside the part 
                try:
                    for i in range(len(points)-1):
                        normal_vector = np.array(points[i+1]+[0])-np.array(points[i]+[0])
                        length = np.linalg.norm(normal_vector)
                        origin = np.array([points[i]+[0]])
                        profile = self.create_profile(normal_vector, origin, width)
                        try:
                            self.extrude(profile, length, 
                                adjustment_amount=added_to_edges, 
                                boolean_operation=boolean_operation)
                        except:
                            self.extrude(profile, length, 
                                adjustment_amount=added_to_edges, 
                                boolean_operation=NEW_BODY)
                except:
                    print("failed on extrude", self.application)
                    
        if self.body == None:
            self.body = self.part_comp_def.SurfaceBodies.Item(1)
    
    def get_edges(self):
        self.edges_xy_collection = self.application.TransientObjects.CreateEdgeCollection()

        self.edges_xy_0_obj = []
        self.edges_xy_0_points = []

        for edge in self.part_comp_def.SurfaceBodies.Item(1).Edges:
            p1, p2 = edge.Evaluator.GetEndPoints()
            # rounding because of floatingpoint error
            if round(p1[2],8) == round(p2[2],8):
                self.edges_xy_collection.Add(edge)
                if round(p1[2],8) == 0:
                    p1 = list(p1)
                    p2 = list(p2)
                    for p in [p1,p2]:
                        for i in range(len(p)):
                            p[i] = round(p[i],8)
                    self.edges_xy_0_obj.append(edge)
                    self.edges_xy_0_points.append([p1, p2])


    def fill_small_gaps(self, max_length):
        def calc_connected_edges_indices(edges):
            nodes = []
            adj = defaultdict(list)
            val_to_id = defaultdict(list)
            id_to_val = {}
            edges_to_val = {}

            for i, (a, b) in enumerate(edges):
                id_a = uuid1()
                id_b = uuid1()
                adj[id_a].append(id_b)
                adj[id_b].append(id_a)
                nodes.extend([id_a, id_b])
                val_to_id[str(a)].append(id_a)
                val_to_id[str(b)].append(id_b)
                id_to_val[id_a] = str(a)
                id_to_val[id_b] = str(b)
                edges_to_val[id_a, id_b] = i
                edges_to_val[id_b, id_a] = i

            for v1, v2 in val_to_id.values():
                adj[v1].append(v2)
                adj[v2].append(v1)

            vis = {}
            for n in nodes:
                vis[n] = False
            def dfa(n, last, start):
                if vis[n]:
                    return []

                vis[n] = True
                a, b = adj[n]
                if a == last:
                    next_node = b
                else:
                    next_node = a

                if start == next_node:
                    return [n]
                else:
                    return [n]+dfa(next_node, n, start)

            def simplify_path(path):
                ret = []
                last = None
                for i in path:
                    if i != last:
                        ret.append(list(i))
                    last = i
                return ret

            def simplify_path_to_edge_index(path):
                ret = []
                aux = [path[i:i+2] for i in range(0, len(path), 2)]
                for a,b in aux:
                    ret.append(edges_to_val[a,b])
                return ret

            ll = []
            ll_edge = []
            for n in nodes:
                if vis[n]:
                    continue
                result = dfa(n, None, n)
                path = [id_to_val[i] for i in result]
                ll.append(simplify_path(path))
                ll_edge.append(simplify_path_to_edge_index(result))

            return ll_edge
            
        def calc_len_connected_edges(edges_point, indices):
            lengths = []
            for i in indices:
                lengths.append(0)
                for j in i:
                    vector = np.array(edges_point[j][1]) - np.array(edges_point[j][0])
                    lengths[-1] += np.linalg.norm(vector)
            
            return lengths

        connected_edges_indices = calc_connected_edges_indices(self.edges_xy_0_points)
        self.lengths = calc_len_connected_edges(self.edges_xy_0_points, connected_edges_indices)

        for indices, length in zip(connected_edges_indices, self.lengths):
            if length < max_length:
                sketch = self.sketches.Add(self.xy_plane)
                sketch.DeferUpdates = True

                points = []
                sketch_points = []
                if self.edges_xy_0_points[indices[0]][0] in self.edges_xy_0_points[indices[-1]]:
                    point = self.edges_xy_0_points[indices[0]][0]
                else:
                    point = self.edges_xy_0_points[indices[0]][1]

                points.append(point)
                tg = self.transient_geometry.CreatePoint2d(*point[:2])
                sketch_points.append(sketch.SketchPoints.Add(tg))
                for j in indices[:-1]:
                    # Add sketch and connecting points to profile
                    if self.edges_xy_0_points[j][0] in points:
                        point = self.edges_xy_0_points[j][1]
                    else:
                        point = self.edges_xy_0_points[j][0]

                    points.append(point)
                    tg = self.transient_geometry.CreatePoint2d(*point[:2])
                    sketch_points.append(sketch.SketchPoints.Add(tg))

                points.append(points[0])
                sketch_points.append(sketch_points[0])

                for i in range(len(sketch_points)-1):
                    sketch.SketchLines.AddByTwoPoints(sketch_points[i],sketch_points[i+1])

                sketch.DeferUpdates = False

                solid_profile = sketch.Profiles.AddForSolid()

                extrude_def = self.extrude_features.CreateExtrudeDefinition(
                    solid_profile,
                    JOIN
                )
                extrude_def.SetDistanceExtent(layer_height, EXTRUDE_POSITIVE_DIRECTION)
                self.extrude_features.Add(extrude_def)
                
    def pattern_bodies(self, layer_count):
        # patterns the first layer to make the full print
        bodies = self.application.TransientObjects.CreateObjectCollection()
        bodies.Add(self.part_comp_def.SurfaceBodies.Item(1))    

        pattern_def = self.features.RectangularPatternFeatures.CreateDefinition(
            bodies,
            self.part_comp_def.WorkAxes.Item(3),
            True, 
            layer_count, 
            layer_height
        )
        pattern_def.ComputeType = OPTIMIZED_COMPUTE
        self.features.RectangularPatternFeatures.AddByDefinition(pattern_def)


class InventorAssembly():
    def __init__(self):
        self.application = wc.DispatchEx("Inventor.Application") 
        # create a assembly document
        self.part_doc = self.application.Documents.Add(
            ASM_OBJ_VALUE, 
            self.application.FileManager.GetTemplateFile(ASM_OBJ_VALUE, METRIC)
        )

        # decreases runtime by activating silent mode
        self.application.silentoperation = True
        self.application.screenupdating = False
        self.application.Visible = False

        self.part_comp_def = self.part_doc.ComponentDefinition
        self.transient_geometry = self.application.TransientGeometry
        self.matrix = self.transient_geometry.CreateMatrix()

    def cleanup(self, file_name, inventor_assembly_folder):
        self.part_doc.SaveAs(inventor_assembly_folder + "/" + file_name + ".iam", False)
        self.part_doc.Close(SkipSave = True) 
        self.application.Quit()

    def import_layers(self, layer_folder,n):
        for i in range(n):
            self.import_part(layer_height*i, layer_folder + f"\layer{i}.ipt")
        
    def import_part(self, z, part_path):
        self.matrix.SetTranslation(self.transient_geometry.CreateVector(0,0,z))
        self.part_comp_def.Occurrences.Add(part_path, self.matrix)


def init_globals(global_vars, func, *args, **kwargs):
    for k, v in global_vars.items():
        exec(f"global {k}; {k} = {v}")
    result = func(*args, **kwargs) 
    return result 

def read_config():
    with open("config.json", "r") as jsonfile:
        data = json.load(jsonfile)    
    
    global d_filament
    global max_workers
    global inventor_kill_on_start
    global current_folder
    global same_in_z_direction

    d_filament = data["d_filament"]["value"]
    max_workers = data["max_workers"]["value"]
    inventor_kill_on_start = data["inventor_kill_on_start"]["value"]
    current_folder = fr'{data["current_folder"]["value"]}'

    if not os.path.isdir(current_folder):
        current_folder = os.getcwd()

    same_in_z_direction = data["same_in_z_direction"]["value"]


def read_gcode_variables(file_path):
    # Opens Gcode and extracts layer_count and layer_height
    # default Error for griffin code
    with open(file_path, 'r') as file:
        for line in file:
            line = line.strip()
            if line.startswith(';LAYER_COUNT:'):
                layer_count = int(line.split(':')[1])
            elif line.startswith(';Layer height:'):
                layer_height = float(line.split(':')[1])/10

    return layer_count, layer_height

def extract_layer_coordinates(file_path, start_layer):
    def calc_initialcoords(parsed):
        keywords = ["MINX", "MINY", "MAXX", "MAXY"]
        borders = [0, 0, 0, 0]
        c = 0
        for line in parsed:
            p = line.comment.split(":")
            if p[0] in keywords:
                borders[keywords.index(p[0])] = float(p[1])*0.1
                c+=1
            if c==len(borders):
                break

        initialcoords = [
            round((borders[2]+borders[0])/2, 4),
            round((borders[3]+borders[1])/2, 4),
        ]
        return initialcoords

    with open(file_path, "r") as f:
        gcode = f.read()
    parsed = GcodeParser(gcode, include_comments=True).lines
    
    start_key = f"LAYER:{start_layer}"
    end_key = f"LAYER:{start_layer+1}"
    collecting = False
    extruding = False
    list_of_points = []
    current_path = []
    extrusion_list = []
    width_list = []
    extruded_before = None

    initialcoords = calc_initialcoords(parsed)
    
    for i, line in enumerate(parsed):
        if start_key == parsed[min(i+3,len(parsed)-1)].comment: # because of Coordinates before extruding
            collecting = True
        
        if end_key == line.comment:
            if len(current_path) != 1:
                list_of_points.append(current_path)
            break

        if collecting:
            if extruded_before == None or line.command_str == "G92":#or line.command_str == "G92" griffin 
                extruded_before = line.get_param("E")

            if line.get_param("E") != None and line.get_param("X") != None and line.get_param("Y") != None:
                extruding = True
                current_path.append([
                    round(line.get_param("X")*0.1 - initialcoords[0],4),
                    round(line.get_param("Y")*0.1 - initialcoords[1],4),
                ])
                extruded_last = line.get_param("E")
            elif line.get_param("X") != None and line.get_param("Y") != None:
                if extruding:   
                    # filletfeature is not clean otherwise
                    length = 0
                    for i in range(len(current_path)-1):
                        vector = np.array(current_path[i+1]) - np.array(current_path[i])
                        length += np.linalg.norm(vector)
                        
                    extrusion_list.append((extruded_last - extruded_before)/10)
                    # width if rectangle
                    width = extrusion_list[-1]*d_filament**2*np.pi/length/layer_height/4
                    width = round(width, 3)
                    width_list.append(width)
                    list_of_points.append(current_path)

                    extruding = False
                    extruded_before = extruded_last
                        
                current_path = [[
                    round(line.get_param("X")*0.1 - initialcoords[0],4),
                    round(line.get_param("Y")*0.1 - initialcoords[1],4),
                ]]
    
    return list_of_points, width_list
   

def split_path(points, closed):
    if len(points) == 2:
        return [points]

    if closed:
        for i in range(2):
            points[0][i] = (points[0][i]+points[1][i])/2
        points.append(points[0])

    if points[-1] == points[-2]:
        points.pop(-1)    

    points_array = np.array(points)
    list_normal_vec = []
    for i in range(len(points)-1):
        normal_vec = points_array[i+1]-points_array[i]
        normal_vec = normal_vec / np.linalg.norm(normal_vec) # normalize vec
        list_normal_vec.append(normal_vec)

    index_new_path = 0
    new_points = []
    i = 1
    fill_point = None
    while i < len(list_normal_vec):
        if np.dot(list_normal_vec[index_new_path], list_normal_vec[i]) < 0:
            # if angle > 90 then new path because of self intersection 
            fill_point_last = [fill_point]
            if len(points[index_new_path:i+1]) <= 2:
                fill_point = [0, 0]
                for j in range(2):
                    fill_point[j] = (points[i][j]+points[i+1][j])/2
                new_points.append(fill_point_last+points[index_new_path:i+1]+[fill_point])
                i+=1
            else:
                fill_point = [0, 0]
                for j in range(2):
                    fill_point[j] = (points[i-1][j]+points[i][j])/2
                new_points.append(fill_point_last+points[index_new_path:i]+[fill_point])

            i-=1
            index_new_path = i 
            
        i+=1
    if len(points[index_new_path:])>=2:
        new_points.append([fill_point]+points[index_new_path:])

    new_points[0].pop(0)

    return new_points

def adjust_start_end_points(points, closed, adjustment_amount=0.02):
    if closed:
        return points

    start_normal_vector = np.array(points[1])-np.array(points[0])
    normalized_start_normal_vector = start_normal_vector/np.linalg.norm(start_normal_vector)
    end_normal_vector = np.array(points[-1])-np.array(points[-2])
    normalized_end_normal_vector = end_normal_vector/np.linalg.norm(end_normal_vector)

    points[0][0] = points[0][0] - normalized_start_normal_vector[0]*adjustment_amount
    points[0][1] = points[0][1] - normalized_start_normal_vector[1]*adjustment_amount
    points[-1][0] = points[-1][0] + normalized_end_normal_vector[0]*adjustment_amount
    points[-1][1] = points[-1][1] + normalized_end_normal_vector[1]*adjustment_amount

    return points


def generate_one_layer(file_path, start_layer,
                        inventor_file_name, inventor_part_folder, same_in_z = False,
                        layer_count = 1):
    layer_coordinates, width_list = extract_layer_coordinates(file_path, start_layer) 

    inventor = Inventor()

    for points, width_rectangle in zip(layer_coordinates, width_list):
        closed = (points[0] == points[-1])
        width = width_rectangle + layer_height*0.4
        points = adjust_start_end_points(points, closed, adjustment_amount = width/2)
        # split path because of self-intersection problems of sweep in inventor
        list_split_path_points = split_path(points, closed) 
        inventor.create_feature(width/2, list_split_path_points, width)

    try:
        inventor.get_edges()
        inventor.fill_small_gaps(6*np.mean(width_list))
        print(np.median(width_list))
    except:
        pass
        print("failed on fill_small_gaps")

    if same_in_z:
        inventor.pattern_bodies(layer_count)

    inventor.cleanup(inventor_file_name, inventor_part_folder)

def combine_layers(layer_part_folder, file_name, asm_folder, part_folder, n):
    inventor_assembly = InventorAssembly()
    inventor_assembly.import_layers(layer_part_folder,n)
    inventor_assembly.cleanup(file_name, asm_folder)

    try:
        inventor = Inventor()
        inventor.derive_part(fr"{asm_folder}\{file_name}.iam")
        inventor.cleanup(file_name, part_folder)
    except:
        # derive can fail on big files due to inventors limitations
        inventor.cleanup(file_name + "failed", asm_folder)

def generate_geometry(file_name, current_folder):
    global layer_height

    gcode_file_path = "gcode/" + file_name + ".gcode"
    layer_count, layer_height = read_gcode_variables(gcode_file_path)

    if same_in_z_direction == "y":
        inventor_part_folder = current_folder + "/inventor_part"
        generate_one_layer(gcode_file_path, 1,
            file_name, inventor_part_folder , same_in_z = True, layer_count = layer_count)
    else:
        inventor_part_folder = current_folder + "/inventor_part"
        inventor_layer_folder = current_folder + "/inventor_assembly/layers_temp_" + file_name
        inventor_asm_folder = current_folder + "/inventor_assembly"

        if not os.path.exists(inventor_layer_folder):
            os.makedirs(inventor_layer_folder)

        with cf.ProcessPoolExecutor(max_workers=max_workers) as executor:
            layers_exec = {}
            global_vars = {}

            for k, v in globals().copy().items():
                if type(v) in (int, float) :
                    global_vars[k] = v
            
            for i in range(layer_count):
                layer_exec = executor.submit(init_globals, global_vars, 
                        generate_one_layer, gcode_file_path, i, f"layer{i}", inventor_layer_folder)
                layers_exec[layer_exec] = i

            for f in cf.as_completed(layers_exec):
                try:
                    result = f.result()
                except Exception as exc:
                    # should happen but sometimes inventor crashes
                    print(f"failed on layer {layers_exec[f]}")
                    print(exc)

        start_time_combine = time.perf_counter()
        combine_layers(inventor_layer_folder, file_name, 
                    inventor_asm_folder, inventor_part_folder, layer_count)
        print(time.perf_counter()-start_time_combine)
           

def main():
    read_config()
    if inventor_kill_on_start == "y":
        os.system("taskkill /f /im inventor.exe")

    file_names = []
    file_name = input("""Enter a file name of gcode file you want to convert to 
                    a solid(has to be in gcode folder without .gcode ending):""")

    while True:
        gcode_file_path = "gcode/" + file_name + ".gcode"
        if os.path.exists(gcode_file_path):
            file_names.append(file_name)
        else:
            print("file name doesn't exist")
        file_name = input("""another file you want to convert
                        (if there are no other files hit enter):""")
        if file_name == "":
            break

    for file_name in file_names:
        start_time = time.perf_counter()
        generate_geometry(file_name, current_folder)
        print(file_name, time.perf_counter() - start_time)


if __name__ == "__main__":
    main()
