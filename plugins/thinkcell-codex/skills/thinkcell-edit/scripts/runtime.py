"""Portable runtime discovery; no install, network access, or Office launch."""
from pathlib import Path
import importlib.metadata
import os
import platform
import shutil
import sys

REQUIRED = {'lxml': 'lxml', 'olefile': 'olefile', 'xlrd': 'xlrd', 'pyxlsb': 'pyxlsb', 'Pillow': 'PIL'}

def find_ppttc():
    override = os.environ.get('THINKCELL_PPTTC')
    if override:
        p = Path(override).expanduser()
        return p.resolve() if p.is_file() else None
    found = shutil.which('ppttc.exe')
    if found:
        return Path(found).resolve()
    roots = [Path(os.environ[k]) for k in ('ProgramFiles(x86)', 'ProgramFiles', 'LOCALAPPDATA') if os.environ.get(k)]
    for root in roots:
        for rel in ('think-cell/ppttc.exe', 'think-cell/ppttc/ppttc.exe', 'think-cell/think-cell/ppttc.exe'):
            p = root/rel
            if p.is_file():
                return p.resolve()
    return None

def powershell():
    root = os.environ.get('SystemRoot', r'C:\Windows')
    p = Path(root)/'System32/WindowsPowerShell/v1.0/powershell.exe'
    if not p.is_file():
        raise RuntimeError('Windows PowerShell 5.1 is required for native verification.')
    return str(p)

def powershell_env():
    """Do not inherit a PowerShell 7-only module path into Windows PowerShell."""
    env = os.environ.copy()
    root = Path(os.environ.get('SystemRoot', r'C:\Windows'))
    module_paths = [root/'System32/WindowsPowerShell/v1.0/Modules']
    if os.environ.get('ProgramFiles'):
        module_paths.append(Path(os.environ['ProgramFiles'])/'WindowsPowerShell/Modules')
    env['PSModulePath'] = os.pathsep.join(str(p) for p in module_paths)
    return env

def doctor():
    deps = {}
    for package in REQUIRED:
        try:
            deps[package] = importlib.metadata.version(package)
        except importlib.metadata.PackageNotFoundError:
            deps[package] = None
    exe = find_ppttc()
    return {'platform':platform.system(), 'python':platform.python_version(), 'python_executable':sys.executable,
            'dependencies':deps, 'ppttc':str(exe) if exe else None,
            'static_requirements_met':platform.system() == 'Windows' and sys.version_info >= (3,10) and all(deps.values()) and exe is not None,
            'native_office_and_license':'Checked during a real update; doctor does not launch PowerPoint.',
            'experimental_namer_default':True, 'network_requests':False}
