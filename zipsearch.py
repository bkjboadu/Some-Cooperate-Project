from zip_processor import ZipProcessor
from pathlib import Path
import sys

class ZipReplace(ZipProcessor):
    def __init__(self,zipname,search_content,replace_content):
        super().__init__(zipname)
        self.search_content = search_content
        self.replace_content = replace_content

    def process_files(self):
        for file in Path(self.temp_directory).iterdir():
            file_loc = Path(file).absolute()
            with file_loc.open('r') as f:
                content = f.read()
            new_content = content.replace(self.search_content,self.replace_content)

            with file_loc.open('w') as f:
                f.write(new_content)


if __name__ == "__main__":
    ZipReplace(*sys.argv[1:4]).process_zip()