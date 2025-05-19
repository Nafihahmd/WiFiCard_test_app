# _version.py
import sys,platform

# Version components
MAJOR = 1
MINOR = 0
PATCH = 1
# — Detect OS — 
if sys.platform.startswith("win"):
    OS_TAG = "windows"
elif sys.platform.startswith("linux"):
    OS_TAG = "linux"
elif sys.platform.startswith("darwin"):
    OS_TAG = "macos"
else:
    OS_TAG = sys.platform  # fallback to raw string

# — Detect architecture — 
ARCH_TAG = platform.machine().lower()    # e.g. x86_64, aarch64, armv7l

# — Combine into version string — 
__version__ = f"{MAJOR}.{MINOR}.{PATCH}-{OS_TAG}-{ARCH_TAG}"

# — Optional: expose a function for clarity —
def get_version():
    """Return the full version string with platform tag."""
    return __version__
