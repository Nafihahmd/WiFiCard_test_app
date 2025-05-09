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
import subprocess
import netifaces
import sys

if sys.platform.startswith("linux"):
    # Linux: use pyudev
    import pyudev
    def scan_interfaces(vendor="148f", product="7601"):
        ctx = pyudev.Context()
        devices = ctx.list_devices(subsystem="net",
                                   ID_VENDOR_ID=vendor,
                                   ID_MODEL_ID=product)
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
        print(f"Checking IP for {iface}")
        addrs = netifaces.ifaddresses(iface)
        print(addrs)
        return netifaces.AF_INET in addrs and bool(addrs[netifaces.AF_INET])

    def get_mac(iface):
        addrs = netifaces.ifaddresses(iface)
        link = addrs.get(netifaces.AF_LINK)
        return link[0]['addr'] if link else "Unknown"

else:
    # Windows: attempt WMI import, else stub
    import subprocess, json, sys, tempfile, os
    def scan_interfaces(vendor="148f", product="7601"):
        ps_cmd = [
             "powershell.exe", "-NoProfile", "-Command",
             "Get-NetAdapter -IncludeHidden | "
             "Where-Object { $_.Name -like 'Wi-Fi*' } | "
             "Select-Object Name,InterfaceDescription,Status | ConvertTo-Json -Compress"]

        # run and capture stdout
        try:
            result = subprocess.run(
                ps_cmd,
                capture_output=True,
                text=True,
                check=True,
                timeout=10
            )
        except subprocess.SubprocessError as e:
            raise RuntimeError(f"PowerShell scan failed: {e}") from e

        # parse JSON
        try:
            data = json.loads(result.stdout)
        except json.JSONDecodeError as e:
            raise RuntimeError(f"Failed to parse JSON: {e}\nOutput was: {result.stdout!r}")

        # if only one adapter, PowerShell returns an object, not a list
        if isinstance(data, dict):
            data = [data]
        # print(data) # Debugging
        return [nic["Name"] for nic in data if "usb wireless lan card" in nic["InterfaceDescription"].lower()]
    
    def connect_wifi(iface, ssid, password, timeout=20):
        
        try:
            # Check if profile already exists
            print(f"checking for profile {ssid} on {iface}")
            res = subprocess.run(
            ["netsh", "wlan", "show", "profiles", f"interface={iface}"],
            capture_output=True, text=True, check=True
        )
            profiles_output = res.stdout.lower()
        except subprocess.SubprocessError:
            print("Error checking profiles")
            profiles_output = ""
            
        # Check if the profile already exists
        if ssid.lower() not in profiles_output:
            try:
                # Create the XML profile
                xml = f"""<?xml version="1.0"?>
                <WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
                <name>{ssid}</name>
                <SSIDConfig><SSID><name>{ssid}</name></SSID></SSIDConfig>
                <connectionType>ESS</connectionType>
                <connectionMode>auto</connectionMode>
                <MSM><security>
                    <authEncryption>
                    <authentication>WPA2PSK</authentication>
                    <encryption>AES</encryption>
                    <useOneX>false</useOneX>
                    </authEncryption>
                    <sharedKey>
                    <keyType>passPhrase</keyType>
                    <protected>false</protected>
                    <keyMaterial>{password}</keyMaterial>
                    </sharedKey>
                </security></MSM>
                </WLANProfile>"""

                # write to temp file
                with tempfile.NamedTemporaryFile("w", suffix=".xml", delete=False) as f:
                    f.write(xml)
                    xml_path = f.name
                # Add the profile to the system
                print(f"Adding profile {ssid} to {iface}")
                subprocess.run([
                    "netsh", "wlan", "add", "profile",
                    f'filename={xml_path}', f'interface={iface}', "user=current"
                ], check=True)
            except subprocess.SubprocessError:
                print("Error adding profile")
                return False
            finally:
                print("Removing temporary XML file")
                os.remove(xml_path)
        else:
            print(f"Profile {ssid} already exists on {iface}")
        
        # Connect to the network
        try:
            print(f"Connecting to {ssid} on {iface}")
            res = subprocess.run([
                "netsh", "wlan", "connect",
                f'ssid={ssid}', f'name={ssid}', f'interface={iface}'
            ], check=True, timeout=timeout)
            print(res)
            return True
        except subprocess.SubprocessError:
            print("Error connecting to network")
            return False

    import psutil, socket
    def check_ip(iface):
        addrs = psutil.net_if_addrs().get(iface, [])
        return any(a.family == socket.AF_INET for a in addrs)

    def get_mac(iface: str) -> str:
        """
        Return the MAC address for the given interface using psutil.
        If the interface does not exist or has no MAC, returns "Unknown".
        """
        # psutil.net_if_addrs() → { iface_name: [snic(addr, netmask, broadcast, ptp), ...], ... }
        addrs = psutil.net_if_addrs().get(iface, [])  # dict lookup returns [] if iface missing :contentReference[oaicite:2]{index=2}
        for snic in addrs:
            # Each snic has .family, .address, .netmask, .broadcast, .ptp
            if snic.family == psutil.AF_LINK:           # AF_LINK identifies MAC entries :contentReference[oaicite:3]{index=3}
                return snic.address                     # .address holds the MAC string :contentReference[oaicite:4]{index=4}
        return "Unknown" 


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
