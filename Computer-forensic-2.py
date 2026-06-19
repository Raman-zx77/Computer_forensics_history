import os
import subprocess
import winreg
import sys
import subprocess
import importlib

def install_and_import(package):
    try:
        importlib.import_module(package)
    except ImportError:
        print(f"Package '{package}' not found. Installing...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"Package '{package}' installed successfully.")
            # Restart script after installation to ensure imports are resolved
            print("Restarting script to apply changes...")
            subprocess.check_call([sys.executable] + sys.argv)
            sys.exit(0)
        except subprocess.CalledProcessError as e:
            print(f"Failed to install package '{package}'. Please install it manually.")
            print(f"Error details: {e}")
            sys.exit(1)
    globals()[package] = importlib.import_module(package)

def main():
    # Original main function code here (omitted for brevity)
    # The full main function code from the original script should be here
    pass

if __name__ == "__main__":
    # Ensure required packages are installed
    install_and_import('wmi')
    install_and_import('reportlab')

    # Run main function
    main()

import wmi
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
import datetime

def get_recently_installed_software():
    """
    Note: This function returns all installed software names found in registry keys.
    It does not filter by recent installation date.
    """
    installed_software = []
    uninstall_key_paths = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]
    for uninstall_key_path in uninstall_key_paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, uninstall_key_path) as uninstall_key:
                num_subkeys = winreg.QueryInfoKey(uninstall_key)[0]
                for i in range(num_subkeys):
                    try:
                        subkey_name = winreg.EnumKey(uninstall_key, i)
                        with winreg.OpenKey(uninstall_key, subkey_name) as subkey:
                            try:
                                display_name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                                install_date = "Unknown"
                                for date_key in ["InstallDate", "InstallTime", "InstallDateUTC", "InstallTimestamp"]:
                                    try:
                                        install_date, _ = winreg.QueryValueEx(subkey, date_key)
                                        if install_date:
                                            break
                                    except FileNotFoundError:
                                        continue
                                installed_software.append({'Name': display_name, 'InstallDate': install_date})
                            except FileNotFoundError:
                                pass
                    except (FileNotFoundError, OSError):
                        pass
        except (FileNotFoundError, OSError):
            pass
    return installed_software

def get_recently_uninstalled_software():
    """
    Note: This function attempts to find uninstalled software by checking for missing UninstallString.
    This method is not reliable and may not reflect recent uninstallations.
    """
    uninstalled_software = []
    uninstall_key_paths = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]
    for uninstall_key_path in uninstall_key_paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, uninstall_key_path) as uninstall_key:
                num_subkeys = winreg.QueryInfoKey(uninstall_key)[0]
                for i in range(num_subkeys):
                    try:
                        subkey_name = winreg.EnumKey(uninstall_key, i)
                        with winreg.OpenKey(uninstall_key, subkey_name) as subkey:
                            try:
                                display_name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                                uninstall_date = "Unknown"
                                try:
                                    uninstall_string, _ = winreg.QueryValueEx(subkey, "UninstallString")
                                except FileNotFoundError:
                                    for date_key in ["InstallDate", "InstallTime", "InstallDateUTC", "InstallTimestamp"]:
                                        try:
                                            uninstall_date, _ = winreg.QueryValueEx(subkey, date_key)
                                            if uninstall_date:
                                                break
                                        except FileNotFoundError:
                                            continue
                                    uninstalled_software.append({'Name': display_name, 'UninstallDate': uninstall_date})
                            except FileNotFoundError:
                                pass
                    except (FileNotFoundError, OSError):
                        pass
        except (FileNotFoundError, OSError):
            pass
    return uninstalled_software

import glob
import pythoncom
import win32com.client

def get_recently_opened_documents_from_recent_folder():
    """
    Retrieves recently opened documents by reading .lnk files in the user's Recent folder.
    Filters for PDF, Excel, Word, image files, and HTML files.
    """
    recent_docs = []
    try:
        recent_folder = os.path.join(os.environ['APPDATA'], r'Microsoft\Windows\Recent')
        if not os.path.exists(recent_folder):
            print(f"Recent folder path does not exist: {recent_folder}")
            return recent_docs
        shell = win32com.client.Dispatch("WScript.Shell")
        lnk_files = glob.glob(os.path.join(recent_folder, '*.lnk'))
        for lnk_path in lnk_files:
            try:
                shortcut = shell.CreateShortcut(lnk_path)
                target_path = shortcut.Targetpath
                ext = os.path.splitext(target_path)[1].lower()
                if ext in ['.pdf', '.xls', '.xlsx', '.doc', '.docx', '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.html', '.htm']:
                    recent_docs.append(target_path)
            except Exception as e:
                print(f"Error processing shortcut {lnk_path}: {e}")
    except Exception as e:
        print(f"Error accessing Recent folder: {e}")
    return recent_docs

def get_recent_files_powershell_recent_items():
    """
    Retrieves recent files from the Windows Recent Items folder using PowerShell,
    sorted by LastAccessTime descending.
    """
    recent_files = []
    try:
        import subprocess
        ps_command = '''
        $recentPath = [Environment]::GetFolderPath("Recent")
        Get-ChildItem -Path $recentPath -File | Sort-Object LastAccessTime -Descending | ForEach-Object {
            $_.FullName
        }
        '''
        result = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True)
        if result.returncode == 0:
            for line in result.stdout.splitlines():
                ext = os.path.splitext(line.strip())[-1].lower()
                if ext in ['.pdf', '.xls', '.xlsx', '.doc', '.docx', '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.html', '.htm']:
                    recent_files.append(line.strip())
    except Exception as e:
        print(f"Error running PowerShell command for recent items: {e}")
    return recent_files

def get_recently_opened_documents():
    """
    Combines recently opened documents from registry, Recent folder, PowerShell Shell.Application Namespace 0x0008,
    PowerShell Recent Items folder, recent files by modification date in common user folders,
    and UserAssist registry keys.
    """
    def get_recent_files_powershell():
        recent_files = []
        try:
            import subprocess
            ps_command = '''
            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.Namespace(0x0008)
            $items = $folder.Items()
            $items | ForEach-Object {
                $_.Path
            }
            '''
            result = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True)
            if result.returncode == 0:
                for line in result.stdout.splitlines():
                    ext = os.path.splitext(line.strip())[-1].lower()
                    if ext in ['.pdf', '.xls', '.xlsx', '.doc', '.docx', '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.html', '.htm']:
                        recent_files.append(line.strip())
        except Exception as e:
            print(f"Error running PowerShell command for recent files: {e}")
        return recent_files

    def get_recent_files_by_modification_date(days=7):
        import os
        import datetime
        recent_files = []
        try:
            user_dirs = [os.path.expandvars(r"C:\Users\%USERNAME%\Documents"),
                         os.path.expandvars(r"C:\Users\%USERNAME%\Downloads"),
                         os.path.expandvars(r"C:\Users\%USERNAME%\Desktop")]
            cutoff_date = datetime.datetime.now() - datetime.timedelta(days=days)
            extensions = ['.pdf', '.xls', '.xlsx', '.doc', '.docx', '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.html', '.htm']
            for user_dir in user_dirs:
                if os.path.exists(user_dir):
                    for root, dirs, files in os.walk(user_dir):
                        for file in files:
                            ext = os.path.splitext(file)[1].lower()
                            if ext in extensions:
                                file_path = os.path.join(root, file)
                                try:
                                    mtime = datetime.datetime.fromtimestamp(os.path.getmtime(file_path))
                                    if mtime >= cutoff_date:
                                        recent_files.append(file_path)
                                except Exception:
                                    pass
        except Exception as e:
            print(f"Error scanning user directories for recent files: {e}")
        return recent_files

    def rot13(s):
        result = []
        for c in s:
            if 'a' <= c <= 'z':
                result.append(chr((ord(c) - ord('a') + 13) % 26 + ord('a')))
            elif 'A' <= c <= 'Z':
                result.append(chr((ord(c) - ord('A') + 13) % 26 + ord('A')))
            else:
                result.append(c)
        return ''.join(result)

    def get_recent_files_userassist():
        recent_files = []
        try:
            import winreg
            userassist_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, userassist_path) as ua_key:
                num_subkeys = winreg.QueryInfoKey(ua_key)[0]
                for i in range(num_subkeys):
                    try:
                        guid_subkey_name = winreg.EnumKey(ua_key, i)
                        with winreg.OpenKey(ua_key, guid_subkey_name + r"\Count") as count_key:
                            num_values = winreg.QueryInfoKey(count_key)[1]
                            for j in range(num_values):
                                try:
                                    value_name, value_data, value_type = winreg.EnumValue(count_key, j)
                                    decoded_name = rot13(value_name)
                                    ext = os.path.splitext(decoded_name)[1].lower()
                                    if ext in ['.pdf', '.xls', '.xlsx', '.doc', '.docx', '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.html', '.htm']:
                                        recent_files.append(decoded_name)
                                except Exception:
                                    pass
                    except Exception:
                        pass
        except Exception as e:
            print(f"Error reading UserAssist registry keys: {e}")
        return recent_files

    docs_from_registry = []
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs") as recent_docs_key:
            num_subkeys = winreg.QueryInfoKey(recent_docs_key)[0]
            for i in range(num_subkeys):
                try:
                    subkey_name = winreg.EnumKey(recent_docs_key, i)
                    with winreg.OpenKey(recent_docs_key, subkey_name) as subkey:
                        num_values = winreg.QueryInfoKey(subkey)[1]
                        for j in range(num_values):
                            try:
                                value_name, value_data, value_type = winreg.EnumValue(subkey, j)
                                if isinstance(value_data, bytes):
                                    try:
                                        file_path = value_data.decode('utf-16le').rstrip('\x00')
                                        ext = os.path.splitext(file_path)[1].lower()
                                        if ext in ['.pdf', '.xls', '.xlsx', '.doc', '.docx', '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.html', '.htm']:
                                            docs_from_registry.append(file_path)
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                except Exception:
                    pass
    except Exception:
        pass

    docs_from_recent_folder = get_recently_opened_documents_from_recent_folder()
    docs_from_powershell = get_recent_files_powershell()
    docs_from_powershell_recent_items = get_recent_files_powershell_recent_items()
    docs_from_modification_date = get_recent_files_by_modification_date()
    docs_from_userassist = get_recent_files_userassist()

    combined_docs = list(set(docs_from_registry + docs_from_recent_folder + docs_from_powershell + docs_from_powershell_recent_items + docs_from_modification_date + docs_from_userassist))
    return combined_docs

def get_recently_opened_documents():
    """
    Combines recently opened documents from registry, Recent folder, and PowerShell command.
    """
    def get_recent_files_powershell():
        recent_files = []
        try:
            import subprocess
            ps_command = '''
            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.Namespace(0x0008)
            $items = $folder.Items()
            $items | ForEach-Object {
                $_.Path
            }
            '''
            result = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True)
            if result.returncode == 0:
                for line in result.stdout.splitlines():
                    ext = os.path.splitext(line.strip())[-1].lower()
                    if ext in ['.pdf', '.xls', '.xlsx', '.doc', '.docx', '.jpg', '.jpeg', '.png', '.gif', '.bmp']:
                        recent_files.append(line.strip())
        except Exception as e:
            print(f"Error running PowerShell command for recent files: {e}")
        return recent_files

    docs_from_registry = []
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs") as recent_docs_key:
            num_subkeys = winreg.QueryInfoKey(recent_docs_key)[0]
            for i in range(num_subkeys):
                try:
                    subkey_name = winreg.EnumKey(recent_docs_key, i)
                    with winreg.OpenKey(recent_docs_key, subkey_name) as subkey:
                        num_values = winreg.QueryInfoKey(subkey)[1]
                        for j in range(num_values):
                            try:
                                value_name, value_data, value_type = winreg.EnumValue(subkey, j)
                                if isinstance(value_data, bytes):
                                    try:
                                        file_path = value_data.decode('utf-16le').rstrip('\x00')
                                        ext = os.path.splitext(file_path)[1].lower()
                                        if ext in ['.pdf', '.xls', '.xlsx', '.doc', '.docx', '.jpg', '.jpeg', '.png', '.gif', '.bmp']:
                                            docs_from_registry.append(file_path)
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                except Exception:
                    pass
    except Exception:
        pass

    docs_from_recent_folder = get_recently_opened_documents_from_recent_folder()
    docs_from_powershell = get_recent_files_powershell()

    combined_docs = list(set(docs_from_registry + docs_from_recent_folder + docs_from_powershell))
    return combined_docs

def get_wifi_passwords():
    wifi_data = []
    try:
        profiles_result = subprocess.run(
            ['netsh', 'wlan', 'show', 'profiles'],
            capture_output=True, text=True, encoding='utf-8', errors='ignore')
        profiles_output = profiles_result.stdout
        if not profiles_output:
            print("No WiFi profiles found or unable to retrieve profiles.")
            return wifi_data
        profile_names = [line.split(":")[1].strip() for line in profiles_output.split('\n') if "All User Profile" in line]
        for profile_name in profile_names:
            try:
                profile_result = subprocess.run(
                    ['netsh', 'wlan', 'show', 'profile', profile_name, 'key=clear'],
                    capture_output=True, text=True, encoding='utf-8', errors='ignore')
                profile_info = profile_result.stdout
                if not profile_info:
                    print(f"No profile info found for {profile_name}")
                    password = "Not available"
                else:
                    password_lines = [line for line in profile_info.split('\n') if 'Key Content' in line]
                    password = password_lines[0].split(':')[1].strip() if password_lines else "Not available"
                wifi_data.append({
                    'SSID': profile_name,
                    'Password': password
                    # 'Date Created' and 'Last Connected' are not available via netsh command
                })
            except Exception as e:
                print(f"Error retrieving password for profile {profile_name}: {e}")
    except Exception as e:
        print(f"Error retrieving WiFi profiles: {e}")
    return wifi_data

def read_usb_registry_key(base_key, path, status):
    devices = []
    try:
        key = winreg.OpenKey(base_key, path)
        num_devices = winreg.QueryInfoKey(key)[0]
        for i in range(num_devices):
            try:
                device_id = winreg.EnumKey(key, i)
                device_key = winreg.OpenKey(key, device_id)
                num_instances = winreg.QueryInfoKey(device_key)[0]
                for j in range(num_instances):
                    try:
                        instance_id = winreg.EnumKey(device_key, j)
                        instance_key = winreg.OpenKey(device_key, instance_id)
                        try:
                            device_desc = winreg.QueryValueEx(instance_key, "DeviceDesc")[0].split(';')[-1]
                        except Exception:
                            device_desc = "Unknown"
                        try:
                            friendly_name = winreg.QueryValueEx(instance_key, "FriendlyName")[0]
                        except Exception:
                            friendly_name = "Unknown"
                        try:
                            mfg = winreg.QueryValueEx(instance_key, "Mfg")[0]
                        except Exception:
                            mfg = "Unknown"
                        devices.append({
                            'Description': device_desc,
                            'Friendly Name': friendly_name,
                            'Manufacturer': mfg,
                            'Status': status
                        })
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception as e:
        print(f"Error reading USB registry key {path}: {e}")
    return devices

def get_usb_devices():
    usb_data = []
    current_connected = []
    previously_connected = []
    try:
        c = wmi.WMI()
        for usb in c.Win32_PnPEntity():
            if usb.PNPClass == "USB" or (usb.DeviceID and "USB" in usb.DeviceID):
                device_desc = usb.Description if usb.Description else "Unknown"
                friendly_name = usb.Name if usb.Name else "Unknown"
                mfg = usb.Manufacturer if hasattr(usb, 'Manufacturer') and usb.Manufacturer else "Unknown"
                current_connected.append({
                    'Description': device_desc,
                    'Friendly Name': friendly_name,
                    'Manufacturer': mfg,
                    'Status': 'Currently connected'
                })
    except Exception as e:
        print(f"Error retrieving currently connected USB devices: {e}")
    previously_connected.extend(read_usb_registry_key(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Enum\USBSTOR", 'Previously connected'))
    previously_connected.extend(read_usb_registry_key(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Enum\USB", 'Previously connected'))
    seen = set()
    deduped_usb_data = []
    for device in current_connected + previously_connected:
        identifier = (device.get('Description'), device.get('Friendly Name'), device.get('Manufacturer'))
        if identifier not in seen:
            seen.add(identifier)
            deduped_usb_data.append(device)
    return deduped_usb_data

def get_system_usernames():
    users_dir = os.path.expandvars(r"C:\Users")
    try:
        usernames = [name for name in os.listdir(users_dir) if os.path.isdir(os.path.join(users_dir, name))]
    except Exception:
        usernames = []
    return usernames

def write_to_pdf(wifi_data, usb_data, usernames, recently_used_apps, recently_installed_software=None, recently_uninstalled_software=None, recently_opened_docs=None, filename="report/forensic_report.pdf"):
    import os
    base_dir = os.path.dirname(os.path.abspath(__file__))
    abs_filename = os.path.join(base_dir, filename)
    os.makedirs(os.path.dirname(abs_filename), exist_ok=True)
    doc = SimpleDocTemplate(abs_filename, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    title_style = styles['Title']
    heading_style = styles['Heading2']
    normal_style = styles['BodyText']

    def format_date(date_str):
        # Convert YYYYMMDD to M/D/YYYY format
        if isinstance(date_str, str) and len(date_str) == 8 and date_str.isdigit():
            year = date_str[0:4]
            month = str(int(date_str[4:6]))
            day = str(int(date_str[6:8]))
            return f"{month}/{day}/{year}"
        return date_str

    def add_unique_paragraphs(items):
        seen = set()
        for item in items:
            if isinstance(item, dict):
                # Format dict items with name and date if available
                name = item.get('Name', 'Unknown')
                # Remove usage count display as per user request
                raw_date = item.get('InstallDate') or item.get('UninstallDate') or item.get('LastUsed') or 'Unknown'
                date = format_date(raw_date)
                if date == 'Unknown':
                    date = 'Date: Unknown'
                else:
                    date = f"Date: {date}"
                text = f"{name} ({date})"
            else:
                text = item
            if text not in seen:
                seen.add(text)
                story.append(Paragraph(text, normal_style))

    story.append(Paragraph("Device Information Report", title_style))
    story.append(Spacer(1, 12))
    # Add current date and time below the title
    current_datetime = datetime.datetime.now().strftime("%m/%d/%Y %I:%M:%S %p")
    story.append(Paragraph(f"Report generated on: {current_datetime}", styles['Normal']))
    story.append(Spacer(1, 24))
    story.append(Paragraph("WiFi and Passwords", heading_style))
    story.append(Spacer(1, 12))
    if wifi_data:
        wifi_table_data = [['SSID', 'Password']]
        seen_wifi = set()
        for wifi in wifi_data:
            ssid = wifi.get('SSID', 'N/A')
            password = wifi.get('Password', 'N/A')
            if (ssid, password) not in seen_wifi:
                seen_wifi.add((ssid, password))
                wifi_table_data.append([ssid, password])
        wifi_table = Table(wifi_table_data, hAlign='LEFT', colWidths=[200, 200])
        wifi_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.darkblue),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 8),
            ('BACKGROUND', (0,1), (-1,-1), colors.whitesmoke),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ]))
        story.append(wifi_table)
    else:
        story.append(Paragraph("No WiFi data available or error occurred", normal_style))
    story.append(Spacer(1, 24))
    story.append(Paragraph("USB Devices", heading_style))
    story.append(Spacer(1, 12))
    if usb_data:
        usb_table_data = [['Device Name', 'Status']]
        seen_usb = set()
        for usb in usb_data:
            desc = usb.get('Description', 'Unknown')
            status = usb.get('Status', 'Unknown')
            if (desc, status) not in seen_usb:
                seen_usb.add((desc, status))
                usb_table_data.append([
                    Paragraph(desc, normal_style),
                    Paragraph(status, normal_style)
                ])
        usb_table = Table(usb_table_data, hAlign='LEFT', colWidths=[350, 150])
        usb_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.darkblue),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 8),
            ('BACKGROUND', (0,1), (-1,-1), colors.whitesmoke),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ]))
        story.append(usb_table)
    else:
        story.append(Paragraph("No USB data available or error occurred", normal_style))
    story.append(Spacer(1, 24))
    story.append(Paragraph("System Usernames", heading_style))
    story.append(Spacer(1, 12))
    if usernames:
        add_unique_paragraphs(usernames)
    else:
        story.append(Paragraph("No usernames available or error occurred", normal_style))
    story.append(Spacer(1, 24))
    story.append(Paragraph("Recently Used Applications", heading_style))
    story.append(Spacer(1, 12))
    if recently_used_apps:
        add_unique_paragraphs(recently_used_apps)
    else:
        story.append(Paragraph("No recently used applications available or error occurred", normal_style))
    story.append(Spacer(1, 24))
    story.append(Paragraph("Recently Installed Software", heading_style))
    story.append(Spacer(1, 12))
    if recently_installed_software:
        add_unique_paragraphs(recently_installed_software)
    else:
        story.append(Paragraph("No recently installed software available or error occurred", normal_style))
    story.append(Spacer(1, 24))
    story.append(Paragraph("Recently Uninstalled Software", heading_style))
    story.append(Spacer(1, 12))
    if recently_uninstalled_software:
        add_unique_paragraphs(recently_uninstalled_software)
    else:
        story.append(Paragraph("No recently uninstalled software available or error occurred", normal_style))
    story.append(Spacer(1, 24))
    story.append(Paragraph("Recently Opened Documents", heading_style))
    story.append(Spacer(1, 12))
    if recently_opened_docs:
        seen_docs = set()
        for document in recently_opened_docs:
            if document not in seen_docs:
                seen_docs.add(document)
                story.append(Paragraph(document, normal_style))
    else:
        # Show a text message instead of an image indicating no data available
        story.append(Paragraph("No recently opened documents available or error occurred", normal_style))
    doc.build(story)

def rot13(s):
    result = []
    for c in s:
        if 'a' <= c <= 'z':
            result.append(chr((ord(c) - ord('a') + 13) % 26 + ord('a')))
        elif 'A' <= c <= 'Z':
            result.append(chr((ord(c) - ord('A') + 13) % 26 + ord('A')))
        else:
            result.append(c)
    return ''.join(result)

import struct
import datetime

def filetime_to_dt(filetime):
    # Convert Windows FILETIME to datetime
    if filetime == 0 or filetime == 116444736000000000:
        # FILETIME zero or epoch start, treat as invalid
        return None
    us = filetime / 10
    return datetime.datetime(1601, 1, 1) + datetime.timedelta(microseconds=us)

def get_recently_used_apps():
    recent_apps = []
    userassist_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, userassist_path) as ua_key:
            num_subkeys = winreg.QueryInfoKey(ua_key)[0]
            for i in range(num_subkeys):
                try:
                    guid_subkey_name = winreg.EnumKey(ua_key, i)
                    with winreg.OpenKey(ua_key, guid_subkey_name + r"\Count") as count_key:
                        num_values = winreg.QueryInfoKey(count_key)[1]
                        for j in range(num_values):
                            try:
                                value_name, value_data, value_type = winreg.EnumValue(count_key, j)
                                decoded_name = rot13(value_name)
                                last_used = "Unknown"
                                if isinstance(value_data, bytes) and len(value_data) >= 72:
                                    # FILETIME is at offset 60, 8 bytes
                                    filetime_bytes = value_data[60:68]
                                    filetime = struct.unpack('<Q', filetime_bytes)[0]
                                    dt = filetime_to_dt(filetime)
                                    if dt is None:
                                        last_used = "Unknown"
                                    else:
                                        last_used = dt.strftime("%Y%m%d")
                                recent_apps.append({'Name': decoded_name, 'LastUsed': last_used})
                            except Exception:
                                pass
                except Exception:
                    pass
    except Exception:
        pass
    return recent_apps

import sys
import subprocess
import importlib

def install_and_import(package):
    try:
        importlib.import_module(package)
    except ImportError:
        print(f"Package '{package}' not found. Installing...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"Package '{package}' installed successfully.")
            # Restart script after installation to ensure imports are resolved
            print("Restarting script to apply changes...")
            subprocess.check_call([sys.executable] + sys.argv)
            sys.exit(0)
        except subprocess.CalledProcessError as e:
            print(f"Failed to install package '{package}'. Please install it manually.")
            print(f"Error details: {e}")
            sys.exit(1)
    globals()[package] = importlib.import_module(package)

def main():
    def print_with_datetime(message):
        current_datetime = datetime.datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        print(f"[{current_datetime}] {message}")

    print_with_datetime("Collecting WiFi information...")
    wifi_data = get_wifi_passwords()
    print_with_datetime("Collecting USB device information...")
    usb_data = get_usb_devices()
    print_with_datetime("Collecting system usernames...")
    usernames = get_system_usernames()
    print_with_datetime("Collecting recently used applications...")
    recently_used_apps = get_recently_used_apps()
    print_with_datetime("Collecting recently installed software...")
    recently_installed_software = get_recently_installed_software()
    print_with_datetime("Collecting recently uninstalled software...")
    recently_uninstalled_software = get_recently_uninstalled_software()
    print_with_datetime("Collecting recently opened documents...")
    recently_opened_docs = get_recently_opened_documents()
    print_with_datetime("Writing report to PDF file...")
    write_to_pdf(wifi_data, usb_data, usernames, recently_used_apps, recently_installed_software, recently_uninstalled_software, recently_opened_docs, filename="report/forensic_report.pdf")
    print_with_datetime("Report generated successfully: report/forensic_report.pdf")
    base_dir = os.path.dirname(os.path.abspath(__file__))
    abs_path = os.path.join(base_dir, "report", "forensic_report.pdf")
    print(f"File location: {abs_path}")

if __name__ == "__main__":
    # Ensure required packages are installed
    install_and_import('wmi')
    install_and_import('reportlab')

    # Run main function
    main()

print("System compromised!")
