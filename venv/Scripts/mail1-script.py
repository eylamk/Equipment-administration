#!"C:\Users\user\PycharmProjects\Learning skills\venv\Scripts\python.exe"
# EASY-INSTALL-ENTRY-SCRIPT: 'mail1==3.1.1','console_scripts','mail1'
__requires__ = 'mail1==3.1.1'
import re
import sys
from pkg_resources import load_entry_point

if __name__ == '__main__':
    sys.argv[0] = re.sub(r'(-script\.pyw?|\.exe)?$', '', sys.argv[0])
    sys.exit(
        load_entry_point('mail1==3.1.1', 'console_scripts', 'mail1')()
    )
