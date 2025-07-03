import os
import win32com.client

class ApplicationSession:

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
            return {
                'app': self.sw_app,
                'model': self.sw_model
            }
            
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
            # return True
            return {self.sw_app, self.sw_model}
            
        except Exception as e:
            print(f"Error opening file: {e}")
            return False
        
    def getInstance(self):
        return {
            'app': self.sw_app,
            'model': self.sw_model
        }