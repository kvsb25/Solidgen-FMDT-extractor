import win32com.client
import os
import sys

class SolidWorksTreeTraverser:
    def __init__(self):
        self.sw_app = None
        self.sw_model = None
        
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
    
    def open_part_file(self, file_path):
        """Open a SolidWorks part file"""
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return False
        
        try:
            # Document types: 1=Part, 2=Assembly, 3=Drawing
            doc_type = 1  # Part document
            options = 0
            configuration = ""
            errors = 0
            warnings = 0
            
            self.sw_model = self.sw_app.OpenDoc6(
                file_path, doc_type, options, configuration, errors, warnings
            )
            
            if self.sw_model is None:
                print("Failed to open the part file")
                return False
            
            print(f"Successfully opened: {file_path}")
            return True
            
        except Exception as e:
            print(f"Error opening file: {e}")
            return False
    
    def traverse_feature_tree(self, indent_level=0):
        """Traverse and print the FeatureManager tree"""
        if self.sw_model is None:
            print("No model loaded")
            return
        
        try:
            # Get the feature manager
            feature_mgr = self.sw_model.FeatureManager
            
            # Get the first feature
            feature = self.sw_model.FirstFeature()
            
            print("FeatureManager Tree:")
            print("-" * 50)
            
            # Traverse all features
            while feature is not None:
                self._print_feature_info(feature, indent_level)
                
                # Check for sub-features
                sub_feature = feature.GetFirstSubFeature()
                if sub_feature is not None:
                    self._traverse_sub_features(sub_feature, indent_level + 1)
                
                # Move to next feature
                feature = feature.GetNextFeature()
                
        except Exception as e:
            print(f"Error traversing feature tree: {e}")
    
    def _traverse_sub_features(self, feature, indent_level):
        """Recursively traverse sub-features"""
        while feature is not None:
            self._print_feature_info(feature, indent_level)
            
            # Check for nested sub-features
            sub_feature = feature.GetFirstSubFeature()
            if sub_feature is not None:
                self._traverse_sub_features(sub_feature, indent_level + 1)
            
            # Move to next sub-feature
            feature = feature.GetNextSubFeature()
    
    def _print_feature_info(self, feature, indent_level):
        """Print detailed information about a feature"""
        indent = "  " * indent_level
        
        try:
            # Basic feature information
            name = feature.Name
            feature_type = feature.GetTypeName2()
            is_suppressed = feature.IsSuppressed2(0, 0)[0]  # Returns tuple
            
            # Feature state
            state = "SUPPRESSED" if is_suppressed else "ACTIVE"
            
            print(f"{indent}├─ {name}")
            print(f"{indent}│  Type: {feature_type}")
            print(f"{indent}│  State: {state}")
            
            # Additional properties based on feature type
            self._print_additional_properties(feature, indent)
            
        except Exception as e:
            print(f"{indent}├─ [Error reading feature: {e}]")
    
    def _print_additional_properties(self, feature, indent):
        """Print additional properties based on feature type"""
        try:
            feature_type = feature.GetTypeName2()
            
            # Handle different feature types
            if "Sketch" in feature_type:
                # Sketch specific properties
                sketch = feature.GetSpecificFeature2()
                if sketch:
                    sketch_name = sketch.Name
                    print(f"{indent}│  Sketch Name: {sketch_name}")
            
            elif "Extrude" in feature_type or "Cut" in feature_type:
                # Extrude/Cut specific properties
                extrude_feature = feature.GetSpecificFeature2()
                if extrude_feature:
                    print(f"{indent}│  Feature Details: Available")
            
            elif "Fillet" in feature_type or "Chamfer" in feature_type:
                # Fillet/Chamfer properties
                print(f"{indent}│  Edge Feature: Yes")
            
            # Try to get feature parameters
            try:
                param_count = feature.GetParameterCount()
                if param_count > 0:
                    print(f"{indent}│  Parameters: {param_count}")
            except:
                pass
                
        except Exception as e:
            # Silently continue if we can't get additional properties
            pass
    
    def get_feature_statistics(self):
        """Get statistics about the features in the model"""
        if self.sw_model is None:
            print("No model loaded")
            return
        
        try:
            stats = {
                'total_features': 0,
                'sketches': 0,
                'extrudes': 0,
                'cuts': 0,
                'fillets': 0,
                'suppressed': 0
            }
            
            feature = self.sw_model.FirstFeature()
            
            while feature is not None:
                self._count_feature_stats(feature, stats)
                
                # Count sub-features
                sub_feature = feature.GetFirstSubFeature()
                if sub_feature:
                    self._count_sub_feature_stats(sub_feature, stats)
                
                feature = feature.GetNextFeature()
            
            print("\nFeature Statistics:")
            print("-" * 30)
            for key, value in stats.items():
                print(f"{key.replace('_', ' ').title()}: {value}")
                
        except Exception as e:
            print(f"Error getting statistics: {e}")
    
    def _count_feature_stats(self, feature, stats):
        """Count statistics for a single feature"""
        try:
            stats['total_features'] += 1
            
            feature_type = feature.GetTypeName2().lower()
            
            if 'sketch' in feature_type:
                stats['sketches'] += 1
            elif 'extrude' in feature_type:
                stats['extrudes'] += 1
            elif 'cut' in feature_type:
                stats['cuts'] += 1
            elif 'fillet' in feature_type:
                stats['fillets'] += 1
            
            if feature.IsSuppressed2(0, 0)[0]:
                stats['suppressed'] += 1
                
        except:
            pass
    
    def _count_sub_feature_stats(self, feature, stats):
        """Recursively count sub-feature statistics"""
        while feature is not None:
            self._count_feature_stats(feature, stats)
            
            sub_feature = feature.GetFirstSubFeature()
            if sub_feature:
                self._count_sub_feature_stats(sub_feature, stats)
            
            feature = feature.GetNextSubFeature()
    
    def close_model(self):
        """Close the current model"""
        if self.sw_model:
            self.sw_app.CloseDoc(self.sw_model.GetTitle())
            self.sw_model = None
            print("Model closed")

def main():
    # Example usage
    traverser = SolidWorksTreeTraverser()
    
    # Connect to SolidWorks
    if not traverser.connect_to_solidworks():
        print("Failed to connect to SolidWorks")
        return
    
    # Specify your .SLDPRT file path here
    part_file_path = r"C:\Path\To\Your\File.SLDPRT"
    
    # You can also use the currently active document
    # by commenting out the open_part_file call
    
    if traverser.open_part_file(part_file_path):
        # Traverse the feature tree
        traverser.traverse_feature_tree()
        
        # Get feature statistics
        traverser.get_feature_statistics()
        
        # Close the model when done
        traverser.close_model()
    
    print("\nTraversal complete!")

if __name__ == "__main__":
    main()