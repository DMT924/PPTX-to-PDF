import os
import sys
import subprocess

def check_python_version():
    if sys.version_info < (3, 6):
        print("This program requires Python 3.6 or higher")
        sys.exit(1)

def install_requirements():
    required_packages = ["pywin32", "comtypes"]
    try:
        import tkinter
        import win32com.client
        import comtypes
        print("All required packages are already installed!")
    except ImportError as e:
        print(f"Missing package detected: {str(e)}")
        print("Installing required packages...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + required_packages)
            print("Packages installed successfully!")
        except subprocess.CalledProcessError as e:
            print(f"Error installing packages: {str(e)}")
            sys.exit(1)

def check_powerpoint_installed():
    try:
        import win32com.client
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Quit()
        return True
    except:
        return False

def main():
    # Check Python version
    check_python_version()
    
    # Install requirements
    install_requirements()
    
    # Check for PowerPoint
    if not check_powerpoint_installed():
        print("Microsoft PowerPoint is not installed on this system.")
        print("Please install PowerPoint to use this converter.")
        sys.exit(1)
    
    # Get the directory where main.py is located
    current_dir = os.path.dirname(os.path.abspath(__file__))
    converter_path = os.path.join(current_dir, "2pdf.py")
    
    # Check if converter exists
    if not os.path.exists(converter_path):
        print(f"Error: Could not find converter at {converter_path}")
        sys.exit(1)
    
    # Run the converter
    try:
        subprocess.call([sys.executable, converter_path])
    except Exception as e:
        print(f"Error running converter: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
