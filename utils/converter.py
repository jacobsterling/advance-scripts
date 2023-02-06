
from pathlib import Path
import webbrowser
import webarchive

def convert(path: Path):
    for file in path.glob("*"):
        if file.is_file() & file.suffix == ".webarchive":
              
            OUTPUT_path = path / "converted" / file.name 
            
            print(file)
            
            with webarchive.open(file) as archive:
                # Extract the archive, and assert that it succeeded
                archive.extract(OUTPUT_path )
                
            webbrowser.open(OUTPUT_path / "converted")

if __name__ == "__main__":
    convert(Path().absolute())