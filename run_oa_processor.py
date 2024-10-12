import os
import sys
from oa_processor import main

def get_script_directory():
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, use the directory of the executable
        return os.path.dirname(sys.executable)
    else:
        # If it's run as a normal Python script, use the script's directory
        return os.path.dirname(os.path.abspath(__file__))

if __name__ == "__main__":
    print("OA Processor starting...")
    
    # Get the directory where the executable or script is located
    exe_directory = get_script_directory()
    print(f"Executable directory: {exe_directory}")
    
    # Change the current working directory to the executable's directory
    os.chdir(exe_directory)
    print(f"Current working directory set to: {os.getcwd()}")
    
    try:
        main()
        print("Processing complete.")
    except Exception as e:
        print(f"An error occurred: {e}")
    
    input("Press Enter to exit...")