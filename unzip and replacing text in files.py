import sys,os
import shutil
import zipfile
from pathlib import Path


class ZipReplace:
    def __init__(self,filename,search_content,replace_content):
        self.filename = filename
        self.search_content = search_content
        self.replace_content = replace_content
        self.temp_directory = (Path(str(self.filename).removesuffix('.zip')).name)

    def unzip_find_replace(self):
        os.makedirs(self.temp_directory,exist_ok=True)
        self.unzip()
        self.find_replace()
        self.zip()

    def unzip(self):
        with zipfile.ZipFile(self.filename) as zip:
            zip.extractall(self.temp_directory)

    def find_replace(self):
        for file in Path(self.temp_directory).iterdir():
            file_loc = Path(file).absolute()
            with file_loc.open('r') as f:
                content = f.read()
            new_content = content.replace(self.search_content,self.replace_content)

            with file_loc.open('w') as f:
                f.write(new_content)

    def zip(self):
        with zipfile.ZipFile(self.filename,'w') as f:
            for file in Path(self.temp_directory).iterdir():
                if not file.glob("*.txt"):
                    continue
                f.write(file,file.name,compress_type=zipfile.ZIP_DEFLATED)
        shutil.rmtree(self.temp_directory)

if __name__ == "__main__":
    ZipReplace(*sys.argv[1:4]).unzip_find_replace()



