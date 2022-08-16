import sys,os
import shutil
import zipfile
from pathlib import Path


class ZipProcessor:
    def __init__(self,zipname):
        self.zipname = zipname
        self.temp_directory = (Path(str(self.zipname).removesuffix('.zip')).name)

    def process_zip(self):
        # os.makedirs(self.temp_directory,exist_ok=True)
        self.unzip_files()
        self.process_files()
        self.zip_files()

    def unzip_files(self):
        with zipfile.ZipFile(self.zipname) as zip:
            zip.extractall(self.temp_directory)


    def zip_files(self):
        with zipfile.ZipFile(self.zipname,'w') as f:
            for file in Path(self.temp_directory).iterdir():
                if not file.glob("*.txt"):
                    continue
                f.write(file,file.name,compress_type=zipfile.ZIP_DEFLATED)
        shutil.rmtree(self.temp_directory)

if __name__ == "__main__":
    ZipProcessor(*sys.argv[1:4]).unzip_find_replace()



