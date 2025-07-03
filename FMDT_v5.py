import win32com.client
import os
import json
import sys
from datetime import datetime
from typing import Dict, List, Any, Optional

class SolidWorksTreeExtractor:
    def __init__(self):
        self.sw_app = None
        self.sw_model = None
        self.feature_data = {}
    
    def connect_to_solidworks(self) -> bool:
        """Connect to active SolidWorks application"""
        try:
            self.sw_app = win32com.client.Dispatch("SldWorks.Application")
            self.sw_model = self.sw_app.ActiveDoc
            
            if self.sw_model is None:
                print("Error: No active SolidWorks document found!")
                return False
                
            print(f"Connected to SolidWorks. Active document: {self.sw_model.GetTitle()}")
            return True
            
        except Exception as e:
            print(f"Error connecting to SolidWorks: {str(e)}")
            return False
    
    def extract_feature_tree(self) -> Dict[str, Any]:
        """Extract optimized FeatureManager Design Tree for LLM-based recreation"""
        
        tree_data = {
            'document_type': self._get_document_type(),
            'creation_sequence': [],
            'feature_relationships': {},
            'sketch_data': {},
            'reference_geometry': [],
            'material_properties': {},
            'configurations': []
        }
        
        try:
            feature = self.sw_model.FirstFeature()
            creation_order = 0
            
            while feature is not None:
                feature_info = self._extract_comprehensive_feature_data(feature, creation_order)
                
                if feature_info:
                    # Add to creation sequence (main workflow)
                    tree_data['creation_sequence'].append(feature_info)
                    
                    # Build relationship map for dependencies
                    if feature_info['dependencies']:
                        tree_data['feature_relationships'][feature_info['name']] = feature_info['dependencies']
                    
                    # Extract detailed sketch information if it's a sketch
                    if feature_info['type'] in ['ProfileFeature', 'Sketch']:
                        sketch_details = self._extract_sketch_details(feature)
                        if sketch_details:
                            tree_data['sketch_data'][feature_info['name']] = sketch_details
                   
                    creation_order += 1
                
                feature = feature.GetNextFeature()
            
            # Extract additional context
            tree_data['reference_geometry'] = self._extract_reference_geometry()
            tree_data['material_properties'] = self._extract_material_info()
            tree_data['configurations'] = self._extract_configurations()
            
        except Exception as e:
            print(f"Error extracting feature tree: {str(e)}")
        
        return tree_data


    def _extract_comprehensive_feature_data(self, feature, creation_order: int) -> Dict[str, Any]:
        """Extract all relevant data for feature recreation"""
        
        try:
            feature_data = {
                'creation_order': creation_order,
                'name': feature.Name,
                'type': feature.GetTypeName(),
                'suppressed': feature.IsSuppressed(),
                'feature_definition': self._get_feature_definition(feature),
                'parameters': self._extract_feature_parameters(feature),
                'dependencies': self._get_feature_dependencies(feature),
                'geometric_constraints': self._extract_constraints(feature),
                'reference_planes': self._get_reference_planes(feature),
                'selection_data': self._get_selection_references(feature)
            }
            
            # Add type-specific data
            feature_type = feature.GetTypeName()
            
            if 'Extrude' in feature_type:
                feature_data.update(self._extract_extrude_data(feature))
            elif 'Cut' in feature_type:
                feature_data.update(self._extract_cut_data(feature))
            elif 'Fillet' in feature_type:
                feature_data.update(self._extract_fillet_data(feature))
            elif 'Pattern' in feature_type:
                feature_data.update(self._extract_pattern_data(feature))
            elif 'Hole' in feature_type:
                feature_data.update(self._extract_hole_data(feature))
            elif 'Sketch' in feature_type:
                feature_data.update(self._extract_sketch_metadata(feature))
            
            return feature_data
            
        except Exception as e:
            print(f"Error extracting feature {feature.Name}: {str(e)}")
            return None

    def _get_feature_definition(self, feature) -> Dict[str, Any]:
        """Extract feature definition for recreation"""
        try:
            feat_def = feature.GetDefinition()
            if feat_def:
                definition_data = {
                    'end_condition': getattr(feat_def, 'EndCondition', None),
                    'direction': getattr(feat_def, 'Direction', None),
                    'depth': getattr(feat_def, 'Depth', None),
                    'draft_angle': getattr(feat_def, 'DraftAngle', None),
                    'reverse_direction': getattr(feat_def, 'ReverseDirection', False)
                }
                return {k: v for k, v in definition_data.items() if v is not None}
        except:
            pass
        return {}

    def _extract_feature_parameters(self, feature) -> Dict[str, Any]:
        """Extract dimensional parameters and equations"""
        parameters = {}
        try:
            # Get feature dimensions
            dimensions = feature.Parameter
            if dimensions:
                for i in range(dimensions.GetCount()):
                    param = dimensions.Item(i)
                    parameters[param.Name] = {
                        'value': param.Value,
                        'equation': getattr(param, 'Equation', ''),
                        'units': getattr(param, 'Units', ''),
                        'linked': getattr(param, 'Linked', False)
                    }
        except:
            pass
        return parameters

    def _get_feature_dependencies(self, feature) -> List[str]:
        """Get what this feature depends on (parents)"""
        dependencies = []
        try:
            parents = feature.GetParents()
            if parents:
                for parent in parents:
                    dependencies.append(parent.Name)
        except:
            pass
        return dependencies

    def _extract_sketch_details(self, feature) -> Dict[str, Any]:
        """Extract detailed sketch information for recreation"""
        sketch_data = {}
        try:
            sketch = feature.GetSpecificFeature2()
            if sketch:
                sketch_data = {
                    'sketch_plane': self._get_sketch_plane(sketch),
                    'sketch_entities': self._extract_sketch_entities(sketch),
                    'dimensions': self._extract_sketch_dimensions(sketch),
                    'relations': self._extract_sketch_relations(sketch),
                    'external_references': self._get_sketch_external_refs(sketch)
                }
        except Exception as e:
            print(f"Error extracting sketch details: {str(e)}")
        return sketch_data

    def _extract_sketch_entities(self, sketch) -> List[Dict[str, Any]]:
        """Extract all sketch entities (lines, arcs, circles, etc.)"""
        entities = []
        try:
            sketch_segments = sketch.GetSketchSegments()
            if sketch_segments:
                for segment in sketch_segments:
                    entity_data = {
                        'type': segment.GetType(),
                        'start_point': getattr(segment, 'GetStartPoint2', lambda: None)(),
                        'end_point': getattr(segment, 'GetEndPoint2', lambda: None)(),
                        'center_point': getattr(segment, 'GetCenterPoint2', lambda: None)(),
                        'radius': getattr(segment, 'GetRadius', lambda: None)(),
                        'construction': getattr(segment, 'ConstructionGeometry', False)
                    }
                    entities.append({k: v for k, v in entity_data.items() if v is not None})
        except:
            pass
        return entities

    def _extract_extrude_data(self, feature) -> Dict[str, Any]:
        """Extract extrude-specific parameters"""
        try:
            extrude_feat = feature.GetSpecificFeature2()
            return {
                'extrude_type': 'boss' if 'Boss' in feature.GetTypeName() else 'cut',
                'end_condition_type': getattr(extrude_feat, 'EndCondition', ''),
                'depth_value': getattr(extrude_feat, 'Depth', 0),
                'reverse_direction': getattr(extrude_feat, 'ReverseDirection', False),
                'merge_result': getattr(extrude_feat, 'MergeResult', True)
            }
        except:
            return {}

    def _extract_fillet_data(self, feature) -> Dict[str, Any]:
        """Extract fillet-specific parameters"""
        try:
            fillet_feat = feature.GetSpecificFeature2()
            return {
                'fillet_type': getattr(fillet_feat, 'Type', ''),
                'radius_value': getattr(fillet_feat, 'Radius', 0),
                'selected_edges': self._get_fillet_edges(fillet_feat),
                'propagate_to_tangent': getattr(fillet_feat, 'PropagateToTangent', False)
            }
        except:
            return {}

    def _get_document_type(self) -> str:
        """Determine if it's part, assembly, or drawing"""
        try:
            doc_type = self.sw_model.GetType()
            type_map = {1: 'part', 2: 'assembly', 3: 'drawing'}
            return type_map.get(doc_type, 'unknown')
        except:
            return 'unknown'

    def _extract_reference_geometry(self) -> List[Dict[str, Any]]:
        """Extract reference planes, axes, coordinate systems"""
        ref_geometry = []
        try:
            # Extract reference planes
            ref_planes = self.sw_model.GetRefPlanes()
            for plane in ref_planes:
                ref_geometry.append({
                    'type': 'reference_plane',
                    'name': plane.Name,
                    'definition': self._get_plane_definition(plane)
                })
        except:
            pass
        return ref_geometry

    def _extract_material_info(self) -> Dict[str, Any]:
        """Extract material properties"""
        try:
            material = self.sw_model.MaterialPropertyName
            return {
                'material_name': material,
                'material_properties': self._get_material_properties()
            }
        except:
            return {}

    def _extract_configurations(self) -> List[Dict[str, Any]]:
        """Extract configuration information"""
        configurations = []
        try:
            config_manager = self.sw_model.ConfigurationManager
            config_names = self.sw_model.GetConfigurationNames()
            
            for config_name in config_names:
                config = config_manager.Item(config_name)
                configurations.append({
                    'name': config_name,
                    'active': config_name == config_manager.ActiveConfiguration.Name,
                    'suppressed_features': self._get_suppressed_features_in_config(config)
                })
        except:
            pass
        return configurations
    
    def _extract_constraints(self, feature) -> List[Dict[str, Any]]:
        """Extract geometric constraints and relations applied to the feature."""
        constraints = []
        try:
            # Get feature definition
            feat_def = feature.GetDefinition()
            if not feat_def:
                return constraints
                
            # Access constraint information
            constraint_count = getattr(feat_def, 'GetConstraintCount', lambda: 0)()
            for i in range(constraint_count):
                constraint = feat_def.GetConstraint(i)
                if constraint:
                    constraint_data = {
                        'type': getattr(constraint, 'GetType', lambda: '')(),
                        'name': getattr(constraint, 'Name', ''),
                        'entities': [],
                        'parameters': {}
                    }
                    
                    # Get constraint entities
                    entities = getattr(constraint, 'GetConstraintEntities', lambda: [])()
                    for entity in entities:
                        constraint_data['entities'].append(getattr(entity, 'Name', str(entity)))
                    
                    constraints.append(constraint_data)
                    
        except Exception as e:
            print(f"Error extracting constraints: {str(e)}")
        return constraints

    def _get_reference_planes(self, feature) -> List[str]:
        """Get reference planes referenced by the feature."""
        ref_planes = []
        try:
            # Check if feature has a sketch
            if hasattr(feature, 'GetSketch'):
                sketch = feature.GetSketch()
                if sketch:
                    sketch_plane = sketch.GetReferenceEntity()
                    if sketch_plane:
                        ref_planes.append(getattr(sketch_plane, 'Name', 'Unknown Plane'))
            
            # Get feature definition and check for plane references
            feat_def = feature.GetDefinition()
            if feat_def:
                # Check for direction references (often planes)
                if hasattr(feat_def, 'Direction'):
                    direction_ref = feat_def.Direction
                    if direction_ref:
                        ref_planes.append(getattr(direction_ref, 'Name', 'Direction Reference'))
                        
        except Exception as e:
            print(f"Error getting reference planes: {str(e)}")
        return ref_planes

    def _get_selection_references(self, feature) -> Dict[str, Any]:
        """Get selection data and entity references used when creating the feature."""
        selection_refs = {
            'faces': [],
            'edges': [],
            'vertices': [],
            'sketches': [],
            'bodies': []
        }
        try:
            feat_def = feature.GetDefinition()
            if not feat_def:
                return selection_refs
                
            # Get selection sets from feature definition
            if hasattr(feat_def, 'GetSelections'):
                selections = feat_def.GetSelections()
                for selection in selections:
                    sel_type = getattr(selection, 'GetType', lambda: 0)()
                    sel_name = getattr(selection, 'Name', '')
                    
                    # Map selection types to categories
                    if sel_type == 2:  # Face
                        selection_refs['faces'].append(sel_name)
                    elif sel_type == 1:  # Edge
                        selection_refs['edges'].append(sel_name)
                    elif sel_type == 3:  # Vertex
                        selection_refs['vertices'].append(sel_name)
                    elif sel_type == 7:  # Sketch
                        selection_refs['sketches'].append(sel_name)
                    elif sel_type == 6:  # Body
                        selection_refs['bodies'].append(sel_name)
                        
        except Exception as e:
            print(f"Error getting selection references: {str(e)}")
        return selection_refs

    def _extract_cut_data(self, feature) -> Dict[str, Any]:
        """Extract cut-specific parameters and properties."""
        cut_data = {}
        try:
            cut_feat = feature.GetSpecificFeature2()
            if cut_feat:
                cut_data = {
                    'cut_type': 'extrude_cut' if 'Cut' in feature.GetTypeName() else 'other',
                    'end_condition': getattr(cut_feat, 'EndCondition', 0),
                    'depth': getattr(cut_feat, 'Depth', 0.0),
                    'reverse_direction': getattr(cut_feat, 'ReverseDirection', False),
                    'flip_side_to_cut': getattr(cut_feat, 'FlipSideToCut', False),
                    'draft_angle': getattr(cut_feat, 'DraftAngle', 0.0),
                    'draft_outward': getattr(cut_feat, 'DraftOutward', True)
                }
                
        except Exception as e:
            print(f"Error extracting cut data: {str(e)}")
        return cut_data

    def _extract_pattern_data(self, feature) -> Dict[str, Any]:
        """Extract pattern-specific parameters (linear, circular, etc.)."""
        pattern_data = {}
        try:
            pattern_feat = feature.GetSpecificFeature2()
            if pattern_feat:
                feature_type = feature.GetTypeName()
                
                if 'LinearPattern' in feature_type:
                    pattern_data = {
                        'pattern_type': 'linear',
                        'direction_1_count': getattr(pattern_feat, 'D1TotalInstances', 1),
                        'direction_1_spacing': getattr(pattern_feat, 'D1Spacing', 0.0),
                        'direction_2_count': getattr(pattern_feat, 'D2TotalInstances', 1),
                        'direction_2_spacing': getattr(pattern_feat, 'D2Spacing', 0.0),
                        'direction_1_reverse': getattr(pattern_feat, 'D1ReverseDirection', False),
                        'direction_2_reverse': getattr(pattern_feat, 'D2ReverseDirection', False)
                    }
                elif 'CircularPattern' in feature_type:
                    pattern_data = {
                        'pattern_type': 'circular',
                        'total_instances': getattr(pattern_feat, 'TotalInstances', 1),
                        'angle_spacing': getattr(pattern_feat, 'Spacing', 0.0),
                        'equal_spacing': getattr(pattern_feat, 'EqualSpacing', True),
                        'reverse_direction': getattr(pattern_feat, 'ReverseDirection', False)
                    }
                    
        except Exception as e:
            print(f"Error extracting pattern data: {str(e)}")
        return pattern_data

    def _extract_hole_data(self, feature) -> Dict[str, Any]:
        """Extract hole feature parameters (hole wizard, simple hole, etc.)."""
        hole_data = {}
        try:
            hole_feat = feature.GetSpecificFeature2()
            if hole_feat:
                hole_data = {
                    'hole_type': getattr(hole_feat, 'Type', 0),
                    'diameter': getattr(hole_feat, 'Diameter', 0.0),
                    'depth': getattr(hole_feat, 'Depth', 0.0),
                    'end_condition': getattr(hole_feat, 'EndCondition', 0),
                    'countersink_diameter': getattr(hole_feat, 'CsinkDiameter', 0.0),
                    'countersink_angle': getattr(hole_feat, 'CsinkAngle', 0.0),
                    'counterbore_diameter': getattr(hole_feat, 'CboreDiameter', 0.0),
                    'counterbore_depth': getattr(hole_feat, 'CboreDepth', 0.0),
                    'thread_designation': getattr(hole_feat, 'ThreadDesignation', ''),
                    'thread_pitch': getattr(hole_feat, 'ThreadPitch', 0.0)
                }
                
        except Exception as e:
            print(f"Error extracting hole data: {str(e)}")
        return hole_data

    def _extract_sketch_metadata(self, feature) -> Dict[str, Any]:
        """Extract sketch metadata and general properties."""
        sketch_metadata = {}
        try:
            sketch = feature.GetSpecificFeature2()
            if sketch:
                sketch_metadata = {
                    'sketch_name': getattr(sketch, 'Name', ''),
                    'fully_defined': getattr(sketch, 'FullyDefined', False),
                    'visible': getattr(sketch, 'Visible', True),
                    'construction_geometry': getattr(sketch, 'ConstructionGeometry', False),
                    'sketch_picture': getattr(sketch, 'SketchPicture', None),
                    'relation_count': getattr(sketch, 'GetSketchRelationsCount', lambda: 0)(),
                    'dimension_count': getattr(sketch, 'GetSketchDimensionsCount', lambda: 0)()
                }
                
        except Exception as e:
            print(f"Error extracting sketch metadata: {str(e)}")
        return sketch_metadata

    def _get_sketch_plane(self, sketch) -> Dict[str, Any]:
        """Get the sketch plane information and orientation."""
        plane_data = {}
        try:
            if sketch:
                # Get sketch transformation
                transform = sketch.ModelToSketchTransform
                if transform:
                    plane_data['transformation_matrix'] = transform.ArrayData
                
                # Get reference entity (sketch plane)
                ref_entity = sketch.GetReferenceEntity()
                if ref_entity:
                    plane_data['plane_name'] = getattr(ref_entity, 'Name', '')
                    plane_data['plane_type'] = getattr(ref_entity, 'GetType', lambda: 0)()
                
                # Get sketch origin
                origin = getattr(sketch, 'GetOrigin', lambda: None)()
                if origin:
                    plane_data['origin'] = [origin.X, origin.Y, origin.Z]
                    
        except Exception as e:
            print(f"Error getting sketch plane: {str(e)}")
        return plane_data

    def _extract_sketch_dimensions(self, sketch) -> List[Dict[str, Any]]:
        """Extract all dimensions applied to the sketch."""
        dimensions = []
        try:
            if sketch:
                # Get display dimensions
                display_dims = sketch.GetDisplayDimensions()
                if display_dims:
                    for dim in display_dims:
                        dim_data = {
                            'name': getattr(dim, 'Name', ''),
                            'value': getattr(dim, 'Value', 0.0),
                            'type': getattr(dim, 'GetType', lambda: 0)(),
                            'dimension_text': getattr(dim, 'DimensionText', ''),
                            'tolerance_type': getattr(dim, 'ToleranceType', 0),
                            'driven': getattr(dim, 'DrivenState', False)
                        }
                        dimensions.append(dim_data)
                        
        except Exception as e:
            print(f"Error extracting sketch dimensions: {str(e)}")
        return dimensions

    def _extract_sketch_relations(self, sketch) -> List[Dict[str, Any]]:
        """Extract geometric relations between sketch entities."""
        relations = []
        try:
            if sketch:
                # Get sketch relations
                sketch_relations = sketch.GetSketchRelations()
                if sketch_relations:
                    for relation in sketch_relations:
                        relation_data = {
                            'type': getattr(relation, 'GetType', lambda: 0)(),
                            'name': getattr(relation, 'Name', ''),
                            'entities': [],
                            'status': getattr(relation, 'Status', 0)
                        }
                        
                        # Get entities involved in relation
                        entities = getattr(relation, 'GetSketchSegments', lambda: [])()
                        for entity in entities:
                            relation_data['entities'].append(getattr(entity, 'GetID', lambda: 0)())
                        
                        relations.append(relation_data)
                        
        except Exception as e:
            print(f"Error extracting sketch relations: {str(e)}")
        return relations

    def _get_sketch_external_refs(self, sketch) -> List[Dict[str, Any]]:
        """Get external references used by the sketch."""
        external_refs = []
        try:
            if sketch:
                # Get external sketch entities
                external_entities = getattr(sketch, 'GetExternalSketchEntities', lambda: [])()
                for entity in external_entities:
                    ref_data = {
                        'entity_type': getattr(entity, 'GetType', lambda: 0)(),
                        'entity_id': getattr(entity, 'GetID', lambda: 0)(),
                        'reference_name': getattr(entity, 'Name', ''),
                        'is_construction': getattr(entity, 'ConstructionGeometry', False)
                    }
                    external_refs.append(ref_data)
                    
        except Exception as e:
            print(f"Error getting sketch external references: {str(e)}")
        return external_refs

    def _get_fillet_edges(self, fillet_feature) -> List[Dict[str, Any]]:
        """Get edges selected for the fillet feature."""
        edges = []
        try:
            if fillet_feature:
                # Get fillet edges
                edge_array = getattr(fillet_feature, 'GetEdges', lambda: [])()
                for i, edge in enumerate(edge_array):
                    edge_data = {
                        'edge_index': i,
                        'edge_id': getattr(edge, 'GetID', lambda: 0)(),
                        'edge_type': getattr(edge, 'GetType', lambda: 0)(),
                        'radius': getattr(fillet_feature, f'GetRadius({i})', lambda: 0.0)() if hasattr(fillet_feature, f'GetRadius') else 0.0
                    }
                    edges.append(edge_data)
                    
        except Exception as e:
            print(f"Error getting fillet edges: {str(e)}")
        return edges

    def _get_plane_definition(self, plane) -> Dict[str, Any]:
        """Get the mathematical definition of a reference plane."""
        plane_def = {}
        try:
            if plane:
                # Get plane parameters
                plane_params = getattr(plane, 'PlaneParams', None)
                if plane_params:
                    plane_def = {
                        'normal_vector': plane_params[:3] if len(plane_params) >= 3 else [0, 0, 1],
                        'origin_point': plane_params[3:6] if len(plane_params) >= 6 else [0, 0, 0],
                        'plane_constant': plane_params[6] if len(plane_params) > 6 else 0.0
                    }
                
                # Get plane name and type
                plane_def['name'] = getattr(plane, 'Name', '')
                plane_def['type'] = getattr(plane, 'GetType', lambda: 0)()
                
        except Exception as e:
            print(f"Error getting plane definition: {str(e)}")
        return plane_def

    def _get_material_properties(self) -> Dict[str, Any]:
        """Get detailed material properties of the model."""
        material_props = {}
        try:
            # Get material property extension
            material_ext = self.sw_model.Extension.GetMaterialPropertyExtension()
            if material_ext:
                material_props = {
                    'density': getattr(material_ext, 'Density', 0.0),
                    'elastic_modulus': getattr(material_ext, 'ElasticModulus', 0.0),
                    'poisson_ratio': getattr(material_ext, 'PoissonRatio', 0.0),
                    'yield_strength': getattr(material_ext, 'YieldStrength', 0.0),
                    'tensile_strength': getattr(material_ext, 'TensileStrength', 0.0),
                    'thermal_expansion': getattr(material_ext, 'ThermalExpansionCoefficient', 0.0),
                    'thermal_conductivity': getattr(material_ext, 'ThermalConductivity', 0.0),
                    'specific_heat': getattr(material_ext, 'SpecificHeat', 0.0)
                }
                
            # Get material name
            material_props['name'] = getattr(self.sw_model, 'MaterialPropertyName', '')
            
        except Exception as e:
            print(f"Error getting material properties: {str(e)}")
        return material_props

    def _get_suppressed_features_in_config(self, config) -> List[str]:
        """Get list of features suppressed in the given configuration."""
        suppressed_features = []
        try:
            if config:
                # Get suppression states
                feature = self.sw_model.FirstFeature()
                while feature:
                    # Check if feature is suppressed in this configuration
                    suppression_state = config.GetSuppressionState(feature)
                    if suppression_state == 0:  # 0 = Suppressed
                        suppressed_features.append(feature.Name)
                    feature = feature.GetNextFeature()
                    
        except Exception as e:
            print(f"Error getting suppressed features: {str(e)}")
        return suppressed_features