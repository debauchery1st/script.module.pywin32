import sys
from os.path import join
# import struct
import imp
import xbmc

home = xbmc.translatePath("special://home").decode("utf-8")
package_path = join(home, "addons", "script.module.pywin32")

assert xbmc.getCondVisibility("system.platform.windows")

# ARCH = "x%d" % (8 * struct.calcsize("P"))
ARCH = "x32"
LIB_PATHS = [
    (join(package_path, "lib", ARCH, "win32"), "win32"),
    (join(package_path, "lib", ARCH, "win32", "lib"), "win32/lib"),
    (join(package_path, "lib", ARCH), "win32com"),
    (join(package_path, "lib", ARCH, "win32comext"), "win32comext")
]

for lib in LIB_PATHS:
    if lib[0] not in sys.path:
        sys.path.append(lib[0])
        # print("Pywin32: Added '%s'" % lib[1])

imp.load_dynamic("pythoncom", join(package_path, "lib", ARCH, "pywin32_system32", "pythoncom27.dll"))
