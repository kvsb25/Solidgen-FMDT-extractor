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