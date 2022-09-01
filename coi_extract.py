import glob
import os
import time
import traceback
import sys
from accord_extractor import AccordExtractor

if __name__ == '__main__':
    try:
        print()
        files_in_downloads = glob. glob('C:/Users/driggerst/Downloads/*.pdf')
        most_recent_file = max(files_in_downloads, key=os. path. getctime)
        extractor = AccordExtractor(most_recent_file)
        extractor.access_riskonnect()
        operation_type = extractor.determine_type()
        xpath_map = extractor.create_map(operation_type)
        extractor.extract_text(xpath_map)
        print()
        time.sleep(1000000)
    except:
        print(traceback.format_exc())
        x = input('\nException detected, press Enter to close.')
