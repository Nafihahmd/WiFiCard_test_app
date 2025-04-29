"""Engine for scanning Wi-Fi adapters, and connecting to networks.

This module provides functions to:
  * Scan for USB-attached MT7601U Wi-Fi interfaces.
  * Connect a given interface to a configured SSID/password.
  * Verify that an interface has been assigned an IPv4 address.

Public API:
    scan_interfaces
    connect_wifi
    check_ip
"""
import pyudev
import subprocess
import netifaces

def scan_interfaces(vendor="148f", product="7601"):
    """Return a list of sysfs names for MT7601U Wi-Fi adapters.

    Uses pyudev to enumerate all network-subsystem devices whose
    ID_VENDOR_ID and ID_MODEL_ID match the given values.

    Args:
        vendor (str): USB vendor ID to match (hex string).
        product (str): USB product ID to match (hex string).

    Returns:
        List[str]: List of interface names (e.g. ["wlan0", "wlan1"]).
    """
    ctx = pyudev.Context()
    devices = ctx.list_devices(subsystem="net",
                               ID_VENDOR_ID=vendor,
                               ID_MODEL_ID=product)  # uses match_property under hood :contentReference[oaicite:4]{index=4}
    return [d.sys_name for d in devices]

def connect_wifi(iface, ssid, password, timeout=20):
    try:
        subprocess.run(["nmcli","radio","wifi","on"], check=True)
        subprocess.run(["nmcli","device","wifi","rescan"], check=True)
        subprocess.run(
            ["nmcli","device","wifi","connect", ssid, "password", password, "ifname", iface],
            check=True, timeout=timeout
        )
        return True
    except subprocess.SubprocessError:
        return False

def check_ip(iface):
    addrs = netifaces.ifaddresses(iface)
    return netifaces.AF_INET in addrs and bool(addrs[netifaces.AF_INET])

def get_mac(iface):
    addrs = netifaces.ifaddresses(iface)
    link = addrs.get(netifaces.AF_LINK)
    return link[0]['addr'] if link else "Unknown"


import random
def random_mac():
    """Generate a random MAC address for testing purposes."""
    # Locally administered unicast: set top-byte to 0x02
    mac = [
        0x02,
        0x00,
        0x00,
        random.randint(0x00, 0xFF),
        random.randint(0x00, 0xFF),
        random.randint(0x00, 0xFF)
    ]
    return ":".join(f"{b:02x}" for b in mac)
