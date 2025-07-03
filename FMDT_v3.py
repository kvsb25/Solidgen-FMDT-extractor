import win32com.client
import os
import sys
import json

class SolidWorksTreeTraverser:
    def __init__(self):
        self.sw_app = None
        self.sw_model = None
        self.feature_data = []
        
    def connect_to_solidworks(self):
        """Connect to SolidWorks application"""
        try:
            # Try to connect to existing SolidWorks instance
            self.sw_app = win32com.client.GetActiveObject("SldWorks.Application")
            print("Connected to existing SolidWorks instance")
        except:
            try:
                # If no existing instance, create new one
                self.sw_app = win32com.client.Dispatch("SldWorks.Application")
                self.sw_app.Visible = True
                print("Started new SolidWorks instance")
            except Exception as e:
                print(f"Failed to connect to SolidWorks: {e}")
                return False
        return True
    
    def open_document(self, file_path):
        """Open a SolidWorks document (Part, Assembly, or Drawing)"""
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return False
        
        try:
            # Determine document type from extension
            ext = os.path.splitext(file_path)[1].lower()
            doc_type_map = {
                '.sldprt': 1,  # Part
                '.sldasm': 2,  # Assembly
                '.slddrw': 3   # Drawing
            }
            
            doc_type = doc_type_map.get(ext, 1)
            options = 0
            configuration = ""
            errors = 0
            warnings = 0
            
            self.sw_model = self.sw_app.OpenDoc6(
                file_path, doc_type, options, configuration, errors, warnings
            )
            
            if self.sw_model is None:
                print("Failed to open the document")
                return False
            
            print(f"Successfully opened: {file_path}")
            return True
            
        except Exception as e:
            print(f"Error opening file: {e}")
            return False
    
    def traverse_feature_tree(self, indent_level=0, export_json=False):
        """Traverse and analyze the FeatureManager tree with detailed information"""
        if self.sw_model is None:
            print("No model loaded")
            return
        
        try:
            doc_type = self.sw_model.GetType()
            doc_type_names = {1: "Part", 2: "Assembly", 3: "Drawing"}
            
            print(f"Document Type: {doc_type_names.get(doc_type, 'Unknown')}")
            print("=" * 80)
            print("DETAILED FEATUREMANAGER TREE ANALYSIS")
            print("=" * 80)
            
            # Get document properties
            self._print_document_properties()
            
            # Get the first feature
            feature = self.sw_model.FirstFeature()
            feature_index = 0
            
            # Traverse all features
            while feature is not None:
                feature_index += 1
                feature_info = self._analyze_feature_comprehensive(feature, indent_level, feature_index)
                self.feature_data.append(feature_info)
                
                # Check for sub-features
                sub_feature = feature.GetFirstSubFeature()
                if sub_feature is not None:
                    self._traverse_sub_features(sub_feature, indent_level + 1, feature_index)
                
                # Move to next feature
                feature = feature.GetNextFeature()
            
            # Export to JSON if requested
            if export_json:
                self._export_to_json()
                
        except Exception as e:
            print(f"Error traversing feature tree: {e}")
    
    def _print_document_properties(self):
        """Print document-level properties"""
        try:
            print("\nDOCUMENT PROPERTIES:")
            print("-" * 40)
            
            # Basic document info
            title = self.sw_model.GetTitle()
            path_name = self.sw_model.GetPathName()
            
            print(f"Title: {title}")
            print(f"Path: {path_name}")
            
            # Get custom properties
            custom_prop_mgr = self.sw_model.Extension.CustomPropertyManager("")
            if custom_prop_mgr:
                prop_names = custom_prop_mgr.GetNames()
                if prop_names:
                    print("\nCustom Properties:")
                    for prop_name in prop_names:
                        val = custom_prop_mgr.Get(prop_name)
                        if val[0]:  # If successful
                            print(f"  {prop_name}: {val[1]}")
            
            # Material information (for parts)
            if self.sw_model.GetType() == 1:  # Part
                try:
                    part_doc = self.sw_model
                    material = part_doc.GetMaterialPropertyName2("", "")
                    if material:
                        print(f"Material: {material}")
                except:
                    pass
            
            print()
            
        except Exception as e:
            print(f"Error getting document properties: {e}")
    
    def _traverse_sub_features(self, feature, indent_level, parent_index):
        """Recursively traverse sub-features"""
        sub_index = 0
        while feature is not None:
            sub_index += 1
            feature_info = self._analyze_feature_comprehensive(
                feature, indent_level, f"{parent_index}.{sub_index}"
            )
            self.feature_data.append(feature_info)
            
            # Check for nested sub-features
            sub_feature = feature.GetFirstSubFeature()
            if sub_feature is not None:
                self._traverse_sub_features(sub_feature, indent_level + 1, f"{parent_index}.{sub_index}")
            
            # Move to next sub-feature
            feature = feature.GetNextSubFeature()
    
    def _analyze_feature_comprehensive(self, feature, indent_level, feature_index):
        """Comprehensive analysis of a single feature"""
        indent = "  " * indent_level
        feature_info = {
            'index': feature_index,
            'indent_level': indent_level,
            'name': '',
            'type': '',
            'state': '',
            'definition': {},
            'parameters': {},
            'geometry_info': {},
            'references': [],
            'creation_info': {}
        }
        
        try:
            # Basic feature information
            name = feature.Name
            feature_type = feature.GetTypeName2()
            is_suppressed = feature.IsSuppressed2(0, 0)[0]
            
            feature_info['name'] = name
            feature_info['type'] = feature_type
            feature_info['state'] = "SUPPRESSED" if is_suppressed else "ACTIVE"
            
            print(f"\n{indent}[{feature_index}] {name}")
            print(f"{indent}{'═' * (len(str(feature_index)) + len(name) + 5)}")
            print(f"{indent}Type: {feature_type}")
            print(f"{indent}State: {feature_info['state']}")
            
            # Get detailed information based on feature type
            self._analyze_feature_by_type(feature, feature_info, indent)
            
            # Get parameters
            self._get_feature_parameters(feature, feature_info, indent)
            
            # Get feature definition
            self._get_feature_definition(feature, feature_info, indent)
            
            # Get references
            self._get_feature_references(feature, feature_info, indent)
            
        except Exception as e:
            print(f"{indent}[Error analyzing feature: {e}]")
            feature_info['error'] = str(e)
        
        return feature_info
    
    def _analyze_feature_by_type(self, feature, feature_info, indent):
        """Analyze feature based on its specific type"""
        try:
            feature_type = feature.GetTypeName2().lower()
            specific_feature = feature.GetSpecificFeature2()
            
            if 'sketch' in feature_type or 'profilefeature' in feature_type:
                self._analyze_sketch_feature(specific_feature, feature_info, indent)
            
            elif 'extrude' in feature_type or 'boss' in feature_type:
                self._analyze_extrude_feature(specific_feature, feature_info, indent)
            
            elif 'cut' in feature_type:
                self._analyze_cut_feature(specific_feature, feature_info, indent)
            
            elif 'revolve' in feature_type:
                self._analyze_revolve_feature(specific_feature, feature_info, indent)
            
            elif 'fillet' in feature_type:
                self._analyze_fillet_feature(specific_feature, feature_info, indent)
            
            elif 'chamfer' in feature_type:
                self._analyze_chamfer_feature(specific_feature, feature_info, indent)
            
            elif 'hole' in feature_type:
                self._analyze_hole_feature(specific_feature, feature_info, indent)
            
            elif 'pattern' in feature_type:
                self._analyze_pattern_feature(specific_feature, feature_info, indent)
            
            elif 'mirror' in feature_type:
                self._analyze_mirror_feature(specific_feature, feature_info, indent)
            
            elif 'shell' in feature_type:
                self._analyze_shell_feature(specific_feature, feature_info, indent)
            
            elif 'draft' in feature_type:
                self._analyze_draft_feature(specific_feature, feature_info, indent)
            
            elif 'rib' in feature_type:
                self._analyze_rib_feature(specific_feature, feature_info, indent)
            
            elif 'loft' in feature_type:
                self._analyze_loft_feature(specific_feature, feature_info, indent)
            
            elif 'sweep' in feature_type:
                self._analyze_sweep_feature(specific_feature, feature_info, indent)
            
            elif 'plane' in feature_type or 'refplane' in feature_type:
                self._analyze_plane_feature(specific_feature, feature_info, indent)
            
            elif 'axis' in feature_type:
                self._analyze_axis_feature(specific_feature, feature_info, indent)
            
            elif 'mate' in feature_type:
                self._analyze_mate_feature(specific_feature, feature_info, indent)
            
            else:
                print(f"{indent}Feature Type: Generic ({feature_type})")
                feature_info['geometry_info']['type'] = 'generic'
                
        except Exception as e:
            print(f"{indent}Error in type-specific analysis: {e}")
    
    def _analyze_sketch_feature(self, sketch, feature_info, indent):
        """Analyze sketch features"""
        try:
            if sketch:
                print(f"{indent}Sketch Analysis:")
                
                # Get sketch info
                sketch_name = sketch.Name
                feature_info['geometry_info'].update({
                    'sketch_name': sketch_name,
                    'type': 'sketch'
                })
                
                print(f"{indent}  Sketch Name: {sketch_name}")
                
                # Get sketch entities
                try:
                    sketch_mgr = self.sw_model.SketchManager
                    if sketch_mgr:
                        # This would require selecting the sketch first
                        print(f"{indent}  Contains sketch entities")
                        feature_info['geometry_info']['has_entities'] = True
                except:
                    pass
                    
        except Exception as e:
            print(f"{indent}Error analyzing sketch: {e}")
    
    def _analyze_extrude_feature(self, extrude, feature_info, indent):
        """Analyze extrude/boss features"""
        try:
            if extrude:
                print(f"{indent}Extrude Analysis:")
                extrude_info = {}
                
                # Get extrude direction
                try:
                    direction = extrude.Direction
                    if direction == 0:
                        dir_text = "Blind"
                    elif direction == 1:
                        dir_text = "Through All"
                    elif direction == 2:
                        dir_text = "Up To Next"
                    elif direction == 3:
                        dir_text = "Up To Vertex"
                    elif direction == 4:
                        dir_text = "Up To Surface"
                    elif direction == 5:
                        dir_text = "Offset From Surface"
                    else:
                        dir_text = f"Direction Type: {direction}"
                    
                    print(f"{indent}  Direction: {dir_text}")
                    extrude_info['direction'] = dir_text
                except:
                    pass
                
                # Get depth values
                try:
                    depth1 = extrude.GetDepth(True)  # True for first direction
                    depth2 = extrude.GetDepth(False)  # False for second direction
                    print(f"{indent}  Depth 1: {depth1:.6f}")
                    if depth2 != 0:
                        print(f"{indent}  Depth 2: {depth2:.6f}")
                    
                    extrude_info['depth1'] = depth1
                    extrude_info['depth2'] = depth2
                except:
                    pass
                
                # Get draft angle
                try:
                    draft_angle = extrude.DraftAngle
                    if draft_angle != 0:
                        print(f"{indent}  Draft Angle: {draft_angle:.6f} radians")
                        extrude_info['draft_angle'] = draft_angle
                except:
                    pass
                
                feature_info['geometry_info'].update({
                    'type': 'extrude',
                    'details': extrude_info
                })
                
        except Exception as e:
            print(f"{indent}Error analyzing extrude: {e}")
    
    def _analyze_cut_feature(self, cut, feature_info, indent):
        """Analyze cut features"""
        try:
            if cut:
                print(f"{indent}Cut Analysis:")
                cut_info = {}
                
                # Similar to extrude but for cuts
                try:
                    direction = cut.Direction
                    depth1 = cut.GetDepth(True)
                    depth2 = cut.GetDepth(False)
                    
                    print(f"{indent}  Cut Direction: {direction}")
                    print(f"{indent}  Cut Depth 1: {depth1:.6f}")
                    if depth2 != 0:
                        print(f"{indent}  Cut Depth 2: {depth2:.6f}")
                    
                    cut_info.update({
                        'direction': direction,
                        'depth1': depth1,
                        'depth2': depth2
                    })
                except:
                    pass
                
                feature_info['geometry_info'].update({
                    'type': 'cut',
                    'details': cut_info
                })
                
        except Exception as e:
            print(f"{indent}Error analyzing cut: {e}")
    
    def _analyze_revolve_feature(self, revolve, feature_info, indent):
        """Analyze revolve features"""
        try:
            if revolve:
                print(f"{indent}Revolve Analysis:")
                revolve_info = {}
                
                try:
                    angle = revolve.Angle
                    direction = revolve.Direction
                    
                    print(f"{indent}  Revolve Angle: {angle:.6f} radians ({angle * 180 / 3.14159:.2f}°)")
                    print(f"{indent}  Direction: {direction}")
                    
                    revolve_info.update({
                        'angle_radians': angle,
                        'angle_degrees': angle * 180 / 3.14159,
                        'direction': direction
                    })
                except:
                    pass
                
                feature_info['geometry_info'].update({
                    'type': 'revolve',
                    'details': revolve_info
                })
                
        except Exception as e:
            print(f"{indent}Error analyzing revolve: {e}")
    
    def _analyze_fillet_feature(self, fillet, feature_info, indent):
        """Analyze fillet features"""
        try:
            if fillet:
                print(f"{indent}Fillet Analysis:")
                fillet_info = {}
                
                try:
                    # Get fillet radius
                    radius = fillet.Radius
                    print(f"{indent}  Radius: {radius:.6f}")
                    fillet_info['radius'] = radius
                    
                    # Get number of edges
                    edge_count = fillet.GetEdgeCount()
                    print(f"{indent}  Number of Edges: {edge_count}")
                    fillet_info['edge_count'] = edge_count
                    
                except:
                    pass
                
                feature_info['geometry_info'].update({
                    'type': 'fillet',
                    'details': fillet_info
                })
                
        except Exception as e:
            print(f"{indent}Error analyzing fillet: {e}")
    
    def _analyze_chamfer_feature(self, chamfer, feature_info, indent):
        """Analyze chamfer features"""
        try:
            if chamfer:
                print(f"{indent}Chamfer Analysis:")
                chamfer_info = {}
                
                try:
                    distance = chamfer.Distance
                    print(f"{indent}  Distance: {distance:.6f}")
                    chamfer_info['distance'] = distance
                    
                    # Get chamfer type
                    chamfer_type = chamfer.Type
                    print(f"{indent}  Type: {chamfer_type}")
                    chamfer_info['type'] = chamfer_type
                    
                except:
                    pass
                
                feature_info['geometry_info'].update({
                    'type': 'chamfer',
                    'details': chamfer_info
                })
                
        except Exception as e:
            print(f"{indent}Error analyzing chamfer: {e}")
    
    def _analyze_hole_feature(self, hole, feature_info, indent):
        """Analyze hole features"""
        try:
            if hole:
                print(f"{indent}Hole Analysis:")
                hole_info = {}
                
                try:
                    hole_type = hole.Type
                    diameter = hole.Diameter
                    depth = hole.Depth
                    
                    print(f"{indent}  Hole Type: {hole_type}")
                    print(f"{indent}  Diameter: {diameter:.6f}")
                    print(f"{indent}  Depth: {depth:.6f}")
                    
                    hole_info.update({
                        'hole_type': hole_type,
                        'diameter': diameter,
                        'depth': depth
                    })
                except:
                    pass
                
                feature_info['geometry_info'].update({
                    'type': 'hole',
                    'details': hole_info
                })
                
        except Exception as e:
            print(f"{indent}Error analyzing hole: {e}")
    
    def _analyze_pattern_feature(self, pattern, feature_info, indent):
        """Analyze pattern features"""
        try:
            if pattern:
                print(f"{indent}Pattern Analysis:")
                pattern_info = {}
                
                try:
                    pattern_type = pattern.Type
                    total_instances = pattern.TotalInstances
                    
                    print(f"{indent}  Pattern Type: {pattern_type}")
                    print(f"{indent}  Total Instances: {total_instances}")
                    
                    pattern_info.update({
                        'pattern_type': pattern_type,
                        'total_instances': total_instances
                    })
                except:
                    pass
                
                feature_info['geometry_info'].update({
                    'type': 'pattern',
                    'details': pattern_info
                })
                
        except Exception as e:
            print(f"{indent}Error analyzing pattern: {e}")
    
    def _analyze_mirror_feature(self, mirror, feature_info, indent):
        """Analyze mirror features"""
        try:
            print(f"{indent}Mirror Analysis:")
            feature_info['geometry_info']['type'] = 'mirror'
        except Exception as e:
            print(f"{indent}Error analyzing mirror: {e}")
    
    def _analyze_shell_feature(self, shell, feature_info, indent):
        """Analyze shell features"""
        try:
            if shell:
                print(f"{indent}Shell Analysis:")
                shell_info = {}
                
                try:
                    thickness = shell.Thickness
                    print(f"{indent}  Thickness: {thickness:.6f}")
                    shell_info['thickness'] = thickness
                except:
                    pass
                
                feature_info['geometry_info'].update({
                    'type': 'shell',
                    'details': shell_info
                })
                
        except Exception as e:
            print(f"{indent}Error analyzing shell: {e}")
    
    def _analyze_draft_feature(self, draft, feature_info, indent):
        """Analyze draft features"""
        try:
            if draft:
                print(f"{indent}Draft Analysis:")
                draft_info = {}
                
                try:
                    angle = draft.Angle
                    print(f"{indent}  Draft Angle: {angle:.6f} radians ({angle * 180 / 3.14159:.2f}°)")
                    draft_info['angle'] = angle
                except:
                    pass
                
                feature_info['geometry_info'].update({
                    'type': 'draft',
                    'details': draft_info
                })
                
        except Exception as e:
            print(f"{indent}Error analyzing draft: {e}")
    
    def _analyze_rib_feature(self, rib, feature_info, indent):
        """Analyze rib features"""
        try:
            if rib:
                print(f"{indent}Rib Analysis:")
                rib_info = {}
                
                try:
                    thickness = rib.Thickness
                    print(f"{indent}  Thickness: {thickness:.6f}")
                    rib_info['thickness'] = thickness
                except:
                    pass
                
                feature_info['geometry_info'].update({
                    'type': 'rib',
                    'details': rib_info
                })
                
        except Exception as e:
            print(f"{indent}Error analyzing rib: {e}")
    
    def _analyze_loft_feature(self, loft, feature_info, indent):
        """Analyze loft features"""
        try:
            print(f"{indent}Loft Analysis:")
            feature_info['geometry_info']['type'] = 'loft'
        except Exception as e:
            print(f"{indent}Error analyzing loft: {e}")
    
    def _analyze_sweep_feature(self, sweep, feature_info, indent):
        """Analyze sweep features"""
        try:
            print(f"{indent}Sweep Analysis:")
            feature_info['geometry_info']['type'] = 'sweep'
        except Exception as e:
            print(f"{indent}Error analyzing sweep: {e}")
    
    def _analyze_plane_feature(self, plane, feature_info, indent):
        """Analyze reference plane features"""
        try:
            print(f"{indent}Reference Plane Analysis:")
            feature_info['geometry_info']['type'] = 'reference_plane'
        except Exception as e:
            print(f"{indent}Error analyzing plane: {e}")
    
    def _analyze_axis_feature(self, axis, feature_info, indent):
        """Analyze axis features"""
        try:
            print(f"{indent}Axis Analysis:")
            feature_info['geometry_info']['type'] = 'axis'
        except Exception as e:
            print(f"{indent}Error analyzing axis: {e}")
    
    def _analyze_mate_feature(self, mate, feature_info, indent):
        """Analyze mate features (for assemblies)"""
        try:
            if mate:
                print(f"{indent}Mate Analysis:")
                mate_info = {}
                
                try:
                    mate_type = mate.Type
                    print(f"{indent}  Mate Type: {mate_type}")
                    mate_info['mate_type'] = mate_type
                except:
                    pass
                
                feature_info['geometry_info'].update({
                    'type': 'mate',
                    'details': mate_info
                })
                
        except Exception as e:
            print(f"{indent}Error analyzing mate: {e}")
    
    def _get_feature_parameters(self, feature, feature_info, indent):
        """Get all parameters for a feature"""
        try:
            param_count = feature.GetParameterCount()
            if param_count > 0:
                print(f"{indent}Parameters ({param_count}):")
                parameters = {}
                
                for i in range(param_count):
                    try:
                        param = feature.Parameter(i)
                        if param:
                            param_name = param.Name
                            param_value = param.Value
                            param_units = param.Units
                            
                            print(f"{indent}  {param_name}: {param_value} {param_units}")
                            parameters[param_name] = {
                                'value': param_value,
                                'units': param_units
                            }
                    except:
                        continue
                
                feature_info['parameters'] = parameters
                
        except Exception as e:
            if param_count > 0:  # Only print error if we expected parameters
                print(f"{indent}Error getting parameters: {e}")
    
    def _get_feature_definition(self, feature, feature_info, indent):
        """Get feature definition details"""
        try:
            definition = feature.GetDefinition()
            if definition:
                print(f"{indent}Definition: Available")
                
                # Try to get definition-specific information
                def_info = {}
                
                try:
                    # Common definition properties
                    if hasattr(definition, 'AccessSelections'):
                        result = definition.AccessSelections(self.sw_model, None)
                        if result:
                            def_info['has_selections'] = True
                except:
                    pass
                
                feature_info['definition'] = def_info
                
        except Exception as e:
            print(f"{indent}Error getting definition: {e}")
    
    def _get_feature_references(self, feature, feature_info, indent):
        """Get feature references (what this feature depends on)"""
        try:
            depends_count = feature.GetDependentCount()
            children_count = feature.GetChildrenCount()
            parents = feature.GetParents()
            
            references = {}
            
            if depends_count > 0:
                references['dependent_count'] = depends_count
                print(f"{indent}Dependents: {depends_count}")
            
            if children_count > 0:
                references['children_count'] = children_count
                print(f"{indent}Children: {children_count}")
            
            if parents:
                references['has_parents'] = True
                print(f"{indent}Has Parent Features: Yes")
            
            if references:
                feature_info['references'] = references
                
        except Exception as e:
            print(f"{indent}Error getting references: {e}")
    
    def _export_to_json(self):
        """Export feature data to JSON file"""
        try:
            filename = "solidworks_feature_analysis.json"
            with open(filename, 'w') as f:
                json.dump(self.feature_data, f, indent=2, default=str)
            print(f"\nFeature analysis exported to: {filename}")
        except Exception as e:
            print(f"Error exporting to JSON: {e}")
    
    def get_comprehensive_statistics(self):
        """Get comprehensive statistics about the features"""
        if self.sw_model is None:
            print("No model loaded")
            return
        
        try:
            stats = {
                'total_features': 0,
                'feature_types': {},
                'suppressed_features': 0,
                'features_with_parameters': 0,
                'total_parameters': 0,
                'sketch_count': 0,
                'solid_features': 0,
                'reference_features': 0
            }
            
            for feature_data in self.feature_data:
                stats['total_features'] += 1
                
                # Count by type
                feature_type = feature_data.get('type', 'Unknown')
                stats['feature_types'][feature_type] = stats['feature_types'].get(feature_type, 0) + 1
                
                # Count suppressed
                if feature_data.get('state') == 'SUPPRESSED':
                    stats['suppressed_features'] += 1
                
                # Count parameters
                params = feature_data.get('parameters', {})
                if params:
                    stats['features_with_parameters'] += 1
                    stats['total_parameters'] += len(params)
                
                # Categorize features
                if 'sketch' in feature_type.lower():
                    stats['sketch_count'] += 1
                elif any(x in feature_type.lower() for x in ['extrude', 'cut', 'revolve', 'fillet', 'chamfer']):
                    stats['solid_features'] += 1
                elif any(x in feature_type.lower() for x in ['plane', 'axis', 'origin']):
                    stats['reference_features'] += 1
            
            print("\n" + "="*60)
            print("COMPREHENSIVE FEATURE STATISTICS")
            print("="*60)
            
            for key, value in stats.items():
                if key == 'feature_types':
                    print(f"\nFeature Types:")
                    for ftype, count in sorted(value.items()):
                        print(f"  {ftype}: {count}")
                else:
                    print(f"{key.replace('_', ' ').title()}: {value}")
            
            return stats
            
        except Exception as e:
            print(f"Error getting statistics: {e}")
    
    def close_model(self):
        """Close the current model"""
        if self.sw_model:
            self.sw_app.CloseDoc(self.sw_model.GetTitle())
            self.sw_model = None
            self.feature_data = []
            print("Model closed")

def main():
    """Example usage with comprehensive analysis"""
    traverser = SolidWorksTreeTraverser()
    
    # Connect to SolidWorks
    if not traverser.connect_to_solidworks():
        print("Failed to connect to SolidWorks")
        return
    
    # Specify your file path here (Part, Assembly, or Drawing)
    file_path = r"C:\Path\To\Your\File.SLDPRT"  # Change this to your file
    
    # You can also use the currently active document by commenting out the open call
    # traverser.sw_model = traverser.sw_app.ActiveDoc
    
    if traverser.open_document(file_path):
        print("Starting comprehensive feature analysis...")
        
        # Perform comprehensive tree traversal
        traverser.traverse_feature_tree(export_json=True)
        
        # Get detailed statistics
        traverser.get_comprehensive_statistics()
        
        # Close the model when done
        traverser.close_model()
    
    print("\n" + "="*60)
    print("ANALYSIS COMPLETE!")
    print("="*60)
    print("The detailed analysis has been saved to 'solidworks_feature_analysis.json'")
    print("This data can be used to recreate the part/assembly/drawing.")

def analyze_current_document():
    """Analyze the currently active SolidWorks document"""
    traverser = SolidWorksTreeTraverser()
    
    if not traverser.connect_to_solidworks():
        print("Failed to connect to SolidWorks")
        return
    
    # Use currently active document
    try:
        traverser.sw_model = traverser.sw_app.ActiveDoc
        if traverser.sw_model is None:
            print("No active document found. Please open a SolidWorks document first.")
            return
        
        print("Analyzing currently active document...")
        traverser.traverse_feature_tree(export_json=True)
        traverser.get_comprehensive_statistics()
        
    except Exception as e:
        print(f"Error analyzing current document: {e}")

def batch_analyze_files(file_list):
    """Analyze multiple SolidWorks files in batch"""
    traverser = SolidWorksTreeTraverser()
    
    if not traverser.connect_to_solidworks():
        print("Failed to connect to SolidWorks")
        return
    
    results = {}
    
    for file_path in file_list:
        print(f"\n{'='*80}")
        print(f"ANALYZING: {os.path.basename(file_path)}")
        print(f"{'='*80}")
        
        if traverser.open_document(file_path):
            traverser.feature_data = []  # Reset for each file
            traverser.traverse_feature_tree()
            stats = traverser.get_comprehensive_statistics()
            
            results[file_path] = {
                'features': traverser.feature_data.copy(),
                'statistics': stats
            }
            
            # Export individual file analysis
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            json_filename = f"{base_name}_analysis.json"
            
            try:
                with open(json_filename, 'w') as f:
                    json.dump({
                        'file_path': file_path,
                        'features': traverser.feature_data,
                        'statistics': stats
                    }, f, indent=2, default=str)
                print(f"Analysis saved to: {json_filename}")
            except Exception as e:
                print(f"Error saving analysis: {e}")
            
            traverser.close_model()
        else:
            print(f"Failed to open: {file_path}")
    
    # Export batch summary
    try:
        with open("batch_analysis_summary.json", 'w') as f:
            json.dump(results, f, indent=2, default=str)
        print(f"\nBatch analysis summary saved to: batch_analysis_summary.json")
    except Exception as e:
        print(f"Error saving batch summary: {e}")
    
    return results

def create_feature_recreation_guide(json_file_path):
    """Create a step-by-step guide to recreate the part from the analysis"""
    try:
        with open(json_file_path, 'r') as f:
            data = json.load(f)
        
        features = data.get('features', [])
        if isinstance(data, list):
            features = data
        
        guide_content = []
        guide_content.append("# SolidWorks Part Recreation Guide")
        guide_content.append("=" * 50)
        guide_content.append("")
        guide_content.append("This guide provides step-by-step instructions to recreate the analyzed part.")
        guide_content.append("")
        
        step_counter = 1
        
        for feature in features:
            if feature.get('indent_level', 0) == 0:  # Only top-level features
                feature_name = feature.get('name', 'Unknown')
                feature_type = feature.get('type', 'Unknown')
                geometry_info = feature.get('geometry_info', {})
                parameters = feature.get('parameters', {})
                
                guide_content.append(f"## Step {step_counter}: {feature_name}")
                guide_content.append(f"**Feature Type:** {feature_type}")
                guide_content.append("")
                
                # Add specific instructions based on feature type
                if 'sketch' in feature_type.lower():
                    guide_content.append("### Instructions:")
                    guide_content.append("1. Create a new sketch on the appropriate plane")
                    guide_content.append("2. Draw the required geometry")
                    guide_content.append("3. Add dimensions and constraints")
                    guide_content.append("4. Exit the sketch")
                    
                elif 'extrude' in feature_type.lower():
                    guide_content.append("### Instructions:")
                    guide_content.append("1. Select the sketch profile")
                    guide_content.append("2. Use the Extruded Boss/Base feature")
                    
                    details = geometry_info.get('details', {})
                    if details:
                        if 'direction' in details:
                            guide_content.append(f"3. Set direction to: {details['direction']}")
                        if 'depth1' in details:
                            guide_content.append(f"4. Set depth to: {details['depth1']:.6f}")
                        if 'draft_angle' in details and details['draft_angle'] != 0:
                            guide_content.append(f"5. Set draft angle to: {details['draft_angle']:.6f} radians")
                    
                elif 'cut' in feature_type.lower():
                    guide_content.append("### Instructions:")
                    guide_content.append("1. Select the sketch profile")
                    guide_content.append("2. Use the Extruded Cut feature")
                    
                    details = geometry_info.get('details', {})
                    if details and 'depth1' in details:
                        guide_content.append(f"3. Set cut depth to: {details['depth1']:.6f}")
                
                elif 'fillet' in feature_type.lower():
                    guide_content.append("### Instructions:")
                    guide_content.append("1. Select the Fillet feature")
                    guide_content.append("2. Select the edges to fillet")
                    
                    details = geometry_info.get('details', {})
                    if details and 'radius' in details:
                        guide_content.append(f"3. Set radius to: {details['radius']:.6f}")
                    if details and 'edge_count' in details:
                        guide_content.append(f"4. Total edges to select: {details['edge_count']}")
                
                elif 'chamfer' in feature_type.lower():
                    guide_content.append("### Instructions:")
                    guide_content.append("1. Select the Chamfer feature")
                    guide_content.append("2. Select the edges to chamfer")
                    
                    details = geometry_info.get('details', {})
                    if details and 'distance' in details:
                        guide_content.append(f"3. Set distance to: {details['distance']:.6f}")
                
                elif 'hole' in feature_type.lower():
                    guide_content.append("### Instructions:")
                    guide_content.append("1. Use the Hole Wizard")
                    
                    details = geometry_info.get('details', {})
                    if details:
                        if 'diameter' in details:
                            guide_content.append(f"2. Set diameter to: {details['diameter']:.6f}")
                        if 'depth' in details:
                            guide_content.append(f"3. Set depth to: {details['depth']:.6f}")
                
                elif 'revolve' in feature_type.lower():
                    guide_content.append("### Instructions:")
                    guide_content.append("1. Select the sketch profile")
                    guide_content.append("2. Use the Revolved Boss/Base feature")
                    guide_content.append("3. Select the axis of revolution")
                    
                    details = geometry_info.get('details', {})
                    if details and 'angle_degrees' in details:
                        guide_content.append(f"4. Set angle to: {details['angle_degrees']:.2f} degrees")
                
                # Add parameters if available
                if parameters:
                    guide_content.append("")
                    guide_content.append("### Parameters:")
                    for param_name, param_info in parameters.items():
                        value = param_info.get('value', 'N/A')
                        units = param_info.get('units', '')
                        guide_content.append(f"- {param_name}: {value} {units}")
                
                guide_content.append("")
                guide_content.append("---")
                guide_content.append("")
                step_counter += 1
        
        # Add material and properties section
        guide_content.append("## Final Steps")
        guide_content.append("")
        guide_content.append("### Material Assignment:")
        guide_content.append("1. Right-click on the part name in FeatureManager")
        guide_content.append("2. Select 'Material' > 'Edit Material'")
        guide_content.append("3. Assign the appropriate material")
        guide_content.append("")
        guide_content.append("### Custom Properties:")
        guide_content.append("1. Go to File > Properties")
        guide_content.append("2. Add any custom properties as needed")
        guide_content.append("")
        guide_content.append("### Save the Part:")
        guide_content.append("1. Save the part with an appropriate name")
        guide_content.append("2. Consider saving in different formats if needed")
        
        # Write the guide to a file
        guide_filename = "part_recreation_guide.md"
        with open(guide_filename, 'w') as f:
            f.write('\n'.join(guide_content))
        
        print(f"Part recreation guide saved to: {guide_filename}")
        return guide_filename
        
    except Exception as e:
        print(f"Error creating recreation guide: {e}")
        return None

def compare_parts(json_file1, json_file2):
    """Compare two part analyses to find differences"""
    try:
        with open(json_file1, 'r') as f:
            data1 = json.load(f)
        with open(json_file2, 'r') as f:
            data2 = json.load(f)
        
        features1 = data1.get('features', []) if isinstance(data1, dict) else data1
        features2 = data2.get('features', []) if isinstance(data2, dict) else data2
        
        print("PART COMPARISON ANALYSIS")
        print("=" * 50)
        print(f"Part 1: {json_file1}")
        print(f"Part 2: {json_file2}")
        print()
        
        # Compare feature counts
        print(f"Feature Count - Part 1: {len(features1)}, Part 2: {len(features2)}")
        
        # Compare feature types
        types1 = [f.get('type', 'Unknown') for f in features1]
        types2 = [f.get('type', 'Unknown') for f in features2]
        
        unique_to_1 = set(types1) - set(types2)
        unique_to_2 = set(types2) - set(types1)
        common_types = set(types1) & set(types2)
        
        if unique_to_1:
            print(f"Features unique to Part 1: {', '.join(unique_to_1)}")
        if unique_to_2:
            print(f"Features unique to Part 2: {', '.join(unique_to_2)}")
        if common_types:
            print(f"Common feature types: {', '.join(common_types)}")
        
        # Detailed feature comparison
        print("\nDETAILED FEATURE COMPARISON:")
        print("-" * 30)
        
        max_features = max(len(features1), len(features2))
        for i in range(max_features):
            f1 = features1[i] if i < len(features1) else None
            f2 = features2[i] if i < len(features2) else None
            
            if f1 and f2:
                if f1.get('type') != f2.get('type'):
                    print(f"Feature {i+1}: Type mismatch - {f1.get('type')} vs {f2.get('type')}")
                elif f1.get('name') != f2.get('name'):
                    print(f"Feature {i+1}: Name mismatch - {f1.get('name')} vs {f2.get('name')}")
            elif f1 and not f2:
                print(f"Feature {i+1}: Only in Part 1 - {f1.get('name')} ({f1.get('type')})")
            elif f2 and not f1:
                print(f"Feature {i+1}: Only in Part 2 - {f2.get('name')} ({f2.get('type')})")
        
        return {
            'features_1': len(features1),
            'features_2': len(features2),
            'unique_to_1': list(unique_to_1),
            'unique_to_2': list(unique_to_2),
            'common_types': list(common_types)
        }
        
    except Exception as e:
        print(f"Error comparing parts: {e}")
        return None

if __name__ == "__main__":
    # Example usage scenarios
    
    print("SolidWorks Feature Tree Analyzer")
    print("=" * 40)
    print("Choose an option:")
    print("1. Analyze specific file")
    print("2. Analyze current active document")
    print("3. Batch analyze multiple files")
    print("4. Create recreation guide from JSON")
    print("5. Compare two parts")
    
    choice = input("Enter your choice (1-5): ").strip()
    
    if choice == "1":
        main()
    elif choice == "2":
        analyze_current_document()
    elif choice == "3":
        # Example file list - modify as needed
        files = [
            r"C:\Path\To\Part1.SLDPRT",
            r"C:\Path\To\Part2.SLDPRT",
            r"C:\Path\To\Assembly1.SLDASM"
        ]
        batch_analyze_files(files)
    elif choice == "4":
        json_file = input("Enter path to JSON analysis file: ").strip()
        create_feature_recreation_guide(json_file)
    elif choice == "5":
        file1 = input("Enter path to first JSON file: ").strip()
        file2 = input("Enter path to second JSON file: ").strip()
        compare_parts(file1, file2)
    else:
        print("Invalid choice. Running default analysis...")
        main()