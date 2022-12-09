# %%
import os
import shutil
import subprocess
import time

from loguru import logger
from win32com.client import GetObject

from utils import basic_info

EXTENSION_FOLDER = r'C:\Users\vhphan\AppData\Local\InfoVista\Planet\7.7\Planet Extensions'


# %%
@basic_info
def close_mapinfo():
    try:
        # %%
        WMI = GetObject('winmgmts:')
        process_info = {}
        for process in WMI.ExecQuery('select * from Win32_Process'):
            print(process.Name, process.ProcessID)
            process_info[process.Name] = process.ProcessID

        # %%
        os.system("taskkill /F /pid " + str(process_info['MapInfoPro.exe']))
    except KeyError as e:
        logger.info('Mapinfo/Planet not opened.')
        logger.debug(e)


# %%
@basic_info
def remove_pex(pex_name=None):
    if pex_name is None:
        pex_name = os.getcwd().split('\\')[-1]
    shutil.rmtree(fr'{EXTENSION_FOLDER}\{pex_name}')


# %%
@basic_info
def copy_new_pex(project_folder=None, pex_name=None):
    if pex_name is None:
        pex_name = os.getcwd().split('\\')[-1]
        logger.info(f'{pex_name=}')
    if project_folder is None:
        project_folder = os.getcwd()
        logger.info(f'{project_folder}')

    shutil.copytree(f'{project_folder}\\{pex_name}', f'{EXTENSION_FOLDER}\\{pex_name}')


# %%
@basic_info
def open_planet():
    subprocess.call(r"C:\Program Files\InfoVista\Planet 7.7\Planet.exe")


# %%
@basic_info
def main():
    close_mapinfo()


if __name__ == '__main__':
    # %%
    close_mapinfo()
    time.sleep(3)
    remove_pex('HelloPlanet')
    copy_new_pex(r'C:\Users\vhphan\source\repos\HelloPlanet\HelloPlanet', 'HelloPlanet')
    open_planet()
    quit()
