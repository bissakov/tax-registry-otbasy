import sys
from shutil import copyfile

system_paths = ['C:\\Users\\robot.ad\\Desktop\\tax-registry-otbasy\\src', 'C:\\Users\\robot.ad\\Desktop\\tax-registry-otbasy', 'C:\\Program Files\\Python310\\python310.zip', 'C:\\Program Files\\Python310\\DLLs', 'C:\\Program Files\\Python310\\lib', 'C:\\Program Files\\Python310', 'C:\\Users\\robot.ad\\Desktop\\tax-registry-otbasy\\venv', 'C:\\Users\\robot.ad\\Desktop\\tax-registry-otbasy\\venv\\lib\\site-packages', 'C:\\Users\\robot.ad\\Desktop\\tax-registry-otbasy\\venv\\lib\\site-packages\\win32', 'C:\\Users\\robot.ad\\Desktop\\tax-registry-otbasy\\venv\\lib\\site-packages\\win32\\lib', 'C:\\Users\\robot.ad\\Desktop\\tax-registry-otbasy\\venv\\lib\\site-packages\\Pythonwin']


def get_lines() -> list[str]:
    with open(__file__) as f:
        lines = f.readlines()
    return lines


def get_corrected_lines() -> list[str]:
    lines = get_lines()
    for i in range(len(lines)):
        if 'system_paths =' in lines[i][0:15]:
            lines[i] = f'system_paths = {str(sys.path)}\n'
    return lines


def write_lines() -> None:
    lines = get_corrected_lines()
    with open(file=__file__, mode='w') as f:
        f.write(''.join(lines))


if __name__ == '__main__':
    write_lines()
    for system_path in system_paths:
        if system_path.endswith('\\site-packages'):
            copyfile(rf'{system_path}\pywin32_system32\pythoncom310.dll',
                     rf'{system_path}\win32\lib\pythoncom310.dll')
            copyfile(rf'{system_path}\pywin32_system32\pywintypes310.dll',
                     rf'{system_path}\win32\lib\pywintypes310.dll')
