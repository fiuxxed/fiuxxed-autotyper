import subprocess, sys, os

here = os.path.dirname(os.path.abspath(__file__))
main_py = os.path.join(here, "main.py")

subprocess.Popen(
    [sys.executable, main_py],
    cwd=here,
    creationflags=0x00000008 | 0x08000000,  # DETACHED_PROCESS | CREATE_NO_WINDOW
    stdin=subprocess.DEVNULL,
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL,
)
