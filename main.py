from FMDT_v4 import SolidWorksTreeExtractor as STEv4
from FMDT_v5 import SolidWorksTreeExtractor as STEv5
import sys
import os
import win32com.client

def main():

    extractor = STEv4()

    try:
        if not extractor.connect_to_solidworks():
            print("Failed to connect to SolidWorks. Make sure:")
            sys.exit(1)
        
        print(extractor.extract_complete_tree())
        
    except Exception as e:
        print(f'Error in main function: {e}')

    
if __name__ == "__main__":
    main()