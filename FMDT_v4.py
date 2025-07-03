"""
SolidWorks FeatureManager Design Tree to JSON Extractor

This script connects to SolidWorks, traverses the FeatureManager Design Tree
of an active document, and exports the complete tree structure with all
parameters, definitions, and metrics to a JSON file.

Requirements:
- SolidWorks installed
- pywin32 (pip install pywin32)
- Active SolidWorks document
"""

import win32com.client
import os
import json
import sys
from datetime import datetime
from typing import Dict, List, Any, Optional

class SolidWorksTreeExtractor:
    def __init__(self, sw_app = None, sw_model = None):
        self.sw_app = None
        self.sw_model = None
        self.feature_data = {}
        
    def connect_to_solidworks(self, file_path = None) -> bool:
        """Connect to active SolidWorks application"""
        try:
            self.sw_app = win32com.client.Dispatch("SldWorks.Application")
            if not file_path:
                self.sw_model = self.sw_app.ActiveDoc
            else:
                if not os.path.exists(file_path):
                    print(f"Error: File not found: {file_path}")
                    return False

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
            
            if self.sw_model is None:
                print("Error: No active SolidWorks document found!")
                return False
                
            print(f"Connected to SolidWorks. Active document: {self.sw_model.GetTitle()}")
            return True
            
        except Exception as e:
            print(f"Error connecting to SolidWorks: {str(e)}")
            return False
        
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
    
    def get_document_info(self) -> Dict[str, Any]:
        """Extract basic document information"""
        doc_info = {
            "document_type": self._get_document_type(),
            "title": self.sw_model.GetTitle(),
            "path": self.sw_model.GetPathName(),
            "units": self._get_document_units(),
            "creation_date": str(datetime.now()),
            "solidworks_version": self.sw_app.RevisionNumber(),
        }
        return doc_info
    
    def _get_document_type(self) -> str:
        """Determine document type"""
        doc_type = self.sw_model.GetType()
        type_map = {
            1: "part",
            2: "assembly", 
            3: "drawing"
        }
        return type_map.get(doc_type, "unknown")
    
    def _get_document_units(self) -> Dict[str, str]:
        """Get document unit system"""
        try:
            user_units = self.sw_model.GetUserUnit(1)  # Length units
            mass_units = self.sw_model.GetUserUnit(6)  # Mass units
            
            unit_map = {
                0: "meters", 1: "millimeters", 2: "centimeters", 3: "inches", 
                4: "feet", 5: "micrometers", 6: "nanometers", 7: "angstroms"
            }
            
            return {
                "length": unit_map.get(user_units, "unknown"),
                "mass": "grams" if mass_units == 1 else "pounds" if mass_units == 2 else "unknown"
            }
        except:
            return {"length": "millimeters", "mass": "grams"}  # Default
    
    def extract_feature_tree(self) -> List[Dict[str, Any]]:
        """Extract complete FeatureManager Design Tree"""
        features = []
        
        try:
            feature = self.sw_model.FirstFeature()
            
            while feature is not None:
                feature_data = self._extract_feature_data(feature)
                if feature_data:
                    features.append(feature_data)
                feature = feature.GetNextFeature()
                
        except Exception as e:
            print(f"Error extracting feature tree: {str(e)}")
            
        return features
    
    def _extract_feature_data(self, feature) -> Optional[Dict[str, Any]]:
        """Extract detailed data from a single feature"""
        try:
            feature_data = {
                "name": feature.Name,
                "type": feature.GetTypeName2(),
                "visible": feature.Visible,
                "suppressed": feature.IsSuppressed(),
                "parameters": self._extract_feature_parameters(feature),
                "dimensions": self._extract_feature_dimensions(feature),
                "sketch_info": self._extract_sketch_info(feature),
                "children": []
            }
            
            # Handle specific feature types
            feature_type = feature.GetTypeName2()
            
            if "Boss" in feature_type or "Cut" in feature_type:
                feature_data.update(self._extract_extrude_data(feature))
            elif "Revolve" in feature_type:
                feature_data.update(self._extract_revolve_data(feature))
            elif "Fillet" in feature_type:
                feature_data.update(self._extract_fillet_data(feature))
            elif "Chamfer" in feature_type:
                feature_data.update(self._extract_chamfer_data(feature))
            elif "Hole" in feature_type:
                feature_data.update(self._extract_hole_data(feature))
            elif "Pattern" in feature_type:
                feature_data.update(self._extract_pattern_data(feature))
            elif "Mirror" in feature_type:
                feature_data.update(self._extract_mirror_data(feature))
            elif "Shell" in feature_type:
                feature_data.update(self._extract_shell_data(feature))
            elif "Rib" in feature_type:
                feature_data.update(self._extract_rib_data(feature))
            
            # Extract child features
            sub_feature = feature.GetFirstSubFeature()
            while sub_feature is not None:
                child_data = self._extract_feature_data(sub_feature)
                if child_data:
                    feature_data["children"].append(child_data)
                sub_feature = sub_feature.GetNextSubFeature()
                
            return feature_data
            
        except Exception as e:
            print(f"Error extracting feature '{feature.Name}': {str(e)}")
            return None
    
    def _extract_feature_parameters(self, feature) -> Dict[str, Any]:
        """Extract feature parameters"""
        parameters = {}
        try:
            # Try to get feature definition
            feat_def = feature.GetDefinition()
            if feat_def:
                # This varies by feature type - would need specific handling
                pass
        except:
            pass
        return parameters
    
    def _extract_feature_dimensions(self, feature) -> List[Dict[str, Any]]:
        """Extract dimensions from feature"""
        dimensions = []
        try:
            display_dims = feature.GetDisplayDimensions()
            if display_dims:
                for i in range(len(display_dims)):
                    dim = display_dims[i]
                    dim_data = {
                        "name": dim.Name,
                        "value": dim.Value,
                        "tolerance_type": dim.GetToleranceType(),
                        "read_only": dim.ReadOnly
                    }
                    dimensions.append(dim_data)
        except:
            pass
        return dimensions
    
    def _extract_sketch_info(self, feature) -> Optional[Dict[str, Any]]:
        """Extract sketch information if feature is sketch-based"""
        try:
            sketch = feature.GetSketch()
            if sketch:
                return {
                    "sketch_name": sketch.Name,
                    "relations": self._extract_sketch_relations(sketch),
                    "dimensions": self._extract_sketch_dimensions(sketch),
                    "entities": self._extract_sketch_entities(sketch)
                }
        except:
            pass
        return None
    
    def _extract_sketch_relations(self, sketch) -> List[str]:
        """Extract sketch relations"""
        relations = []
        try:
            sketch_relations = sketch.GetSketchRelations()
            if sketch_relations:
                for relation in sketch_relations:
                    relations.append(relation.GetType())
        except:
            pass
        return relations
    
    def _extract_sketch_dimensions(self, sketch) -> List[Dict[str, Any]]:
        """Extract sketch dimensions"""
        dimensions = []
        try:
            sketch_dims = sketch.GetSketchDimensions()
            if sketch_dims:
                for dim in sketch_dims:
                    dimensions.append({
                        "name": dim.Name,
                        "value": dim.Value,
                        "type": dim.GetType()
                    })
        except:
            pass
        return dimensions
    
    def _extract_sketch_entities(self, sketch) -> List[Dict[str, Any]]:
        """Extract sketch entities (lines, arcs, etc.)"""
        entities = []
        try:
            sketch_segs = sketch.GetSketchSegments()
            if sketch_segs:
                for seg in sketch_segs:
                    entities.append({
                        "type": seg.GetType(),
                        "length": getattr(seg, 'GetLength', lambda: 0)(),
                        "construction": seg.ConstructionGeometry
                    })
        except:
            pass
        return entities
    
    # Specific feature extraction methods
    def _extract_extrude_data(self, feature) -> Dict[str, Any]:
        """Extract extrude/boss/cut specific data"""
        try:
            feat_def = feature.GetDefinition()
            return {
                "direction": "both" if feat_def.BothDirections else "one",
                "end_condition": feat_def.GetEndCondition(True),
                "depth": feat_def.GetDepth(True),
                "draft_angle": feat_def.GetDraftAngle(True),
                "reverse_direction": feat_def.ReverseDirection
            }
        except:
            return {}
    
    def _extract_revolve_data(self, feature) -> Dict[str, Any]:
        """Extract revolve specific data"""
        try:
            feat_def = feature.GetDefinition()
            return {
                "angle": feat_def.GetRevolveAngle(),
                "direction": "both" if feat_def.BothDirections else "one",
                "reverse_direction": feat_def.ReverseDirection
            }
        except:
            return {}
    
    def _extract_fillet_data(self, feature) -> Dict[str, Any]:
        """Extract fillet specific data"""
        try:
            feat_def = feature.GetDefinition()
            return {
                "radius": feat_def.GetRadius(),
                "full_round": feat_def.FullRound,
                "edge_count": feat_def.GetEdgeCount()
            }
        except:
            return {}
    
    def _extract_chamfer_data(self, feature) -> Dict[str, Any]:
        """Extract chamfer specific data"""
        try:
            feat_def = feature.GetDefinition()
            return {
                "distance": feat_def.GetDistance(),
                "angle": feat_def.GetAngle(),
                "type": feat_def.Type
            }
        except:
            return {}
    
    def _extract_hole_data(self, feature) -> Dict[str, Any]:
        """Extract hole wizard specific data"""
        try:
            feat_def = feature.GetDefinition()
            return {
                "hole_type": feat_def.Type,
                "diameter": feat_def.GetDiameter(),
                "depth": feat_def.GetDepth(),
                "countersink": feat_def.Countersink,
                "counterbore": feat_def.Counterbore
            }
        except:
            return {}
    
    def _extract_pattern_data(self, feature) -> Dict[str, Any]:
        """Extract pattern specific data"""
        try:
            feat_def = feature.GetDefinition()
            return {
                "pattern_type": feat_def.GetType(),
                "instance_count": feat_def.GetPatternBodyCount(),
                "spacing": feat_def.GetSpacing() if hasattr(feat_def, 'GetSpacing') else 0
            }
        except:
            return {}
    
    def _extract_mirror_data(self, feature) -> Dict[str, Any]:
        """Extract mirror specific data"""
        try:
            feat_def = feature.GetDefinition()
            return {
                "mirror_about_plane": True,
                "feature_count": feat_def.GetPatternBodyCount()
            }
        except:
            return {}
    
    def _extract_shell_data(self, feature) -> Dict[str, Any]:
        """Extract shell specific data"""
        try:
            feat_def = feature.GetDefinition()
            return {
                "thickness": feat_def.Thickness,
                "outward": feat_def.Outward
            }
        except:
            return {}
    
    def _extract_rib_data(self, feature) -> Dict[str, Any]:
        """Extract rib specific data"""
        try:
            feat_def = feature.GetDefinition()
            return {
                "thickness": feat_def.Thickness,
                "extrude_from": feat_def.ExtrudeFrom,
                "flip_material_side": feat_def.FlipMaterialSide
            }
        except:
            return {}
    
    def extract_materials(self) -> List[Dict[str, Any]]:
        """Extract material information"""
        materials = []
        try:
            if self.sw_model.GetType() == 1:  # Part document
                material_prop = self.sw_model.GetMaterialPropertyName2("", "")
                if material_prop[0]:
                    materials.append({
                        "name": material_prop[1],
                        "database": material_prop[0]
                    })
        except:
            pass
        return materials
    
    def extract_custom_properties(self) -> Dict[str, str]:
        """Extract custom properties"""
        properties = {}
        try:
            custom_prop_mgr = self.sw_model.Extension.CustomPropertyManager("")
            names = custom_prop_mgr.GetNames()
            if names:
                for name in names:
                    val, resolved_val = custom_prop_mgr.Get(name)
                    properties[name] = resolved_val
        except:
            pass
        return properties
    
    def extract_complete_tree(self) -> Dict[str, Any]:
        """Extract complete FeatureManager Design Tree with all data"""
        print("Extracting document information...")
        doc_info = self.get_document_info()
        
        print("Extracting feature tree...")
        features = self.extract_feature_tree()
        
        print("Extracting materials...")
        materials = self.extract_materials()
        
        print("Extracting custom properties...")
        custom_props = self.extract_custom_properties()
        
        complete_data = {
            "metadata": {
                "extraction_timestamp": str(datetime.now()),
                "extractor_version": "1.0",
                "source": "SolidWorks FeatureManager Design Tree"
            },
            "document_info": doc_info,
            "feature_tree": features,
            "materials": materials,
            "custom_properties": custom_props,
            "feature_count": len(features)
        }
        
        return complete_data
    
    def save_to_json(self, data: Dict[str, Any], output_file: str = None) -> str:
        """Save extracted data to JSON file"""
        if output_file is None:
            doc_name = self.sw_model.GetTitle().replace('.', '_')
            output_file = f"{doc_name}_feature_tree.json"
        
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False, default=str)
            
            print(f"Feature tree data saved to: {output_file}")
            return output_file
            
        except Exception as e:
            print(f"Error saving JSON file: {str(e)}")
            return ""


def main():
    """Main execution function"""
    print("SolidWorks FeatureManager Design Tree Extractor")
    print("=" * 50)
    
    extractor = SolidWorksTreeExtractor()
    
    # Connect to SolidWorks
    if not extractor.connect_to_solidworks():
        print("Failed to connect to SolidWorks. Make sure:")
        print("1. SolidWorks is running")
        print("2. A document is open and active")
        sys.exit(1)
    
    try:
        # Extract complete tree data
        tree_data = extractor.extract_complete_tree()
        
        # Save to JSON
        output_file = extractor.save_to_json(tree_data)
        
        if output_file:
            print(f"\nExtraction completed successfully!")
            print(f"Document type: {tree_data['document_info']['document_type']}")
            print(f"Features extracted: {tree_data['feature_count']}")
            print(f"Output file: {output_file}")
            
            # Display summary
            print("\nFeature Summary:")
            for feature in tree_data['feature_tree'][:10]:  # Show first 10
                print(f"  - {feature['name']} ({feature['type']})")
            
            if len(tree_data['feature_tree']) > 10:
                print(f"  ... and {len(tree_data['feature_tree']) - 10} more features")
        
    except Exception as e:
        print(f"Error during extraction: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()