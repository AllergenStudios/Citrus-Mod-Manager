import requests
import colorama
import shutil
import os
import winshell
from win32com.client import Dispatch

version = "0.0.1"
exit_installer = False

def download_all_assets():
    print(f"{colorama.Fore.GREEN}Installing BepInEx...{colorama.Fore.BLUE}")
    download_specific_github_asset("BepInEx", "BepInEx", "BepInEx_win_x86_5.4.23.2.zip",
                                   download_path="CitrusInstallTemp")
    print(f"{colorama.Fore.GREEN}Installing Bepinject-Auros...{colorama.Fore.BLUE}")
    download_specific_github_asset("Auros", "Bepinject", "Bepinject-Auros.zip", download_path="CitrusInstallTemp")
    print(f"{colorama.Fore.GREEN}Installing Extenject...{colorama.Fore.BLUE}")
    download_specific_github_asset("Auros", "Bepinject", "Extenject.zip", download_path="CitrusInstallTemp")
    print(f"{colorama.Fore.GREEN}Installing Newtonsoft.Json...{colorama.Fore.BLUE}")
    download_specific_github_asset("legoandmars", "Newtonsoft.Json", "Newtonsoft.Json-12.0.3.zip",
                                   download_path="CitrusInstallTemp")
    print(f"{colorama.Fore.GREEN}Installing Newtonsoft.Json...{colorama.Fore.BLUE}")
    download_specific_github_asset("AHauntedArmy", "TMPLoader", "TMPLoader-v1.0.2.zip",
                                   download_path="CitrusInstallTemp")
    print(f"{colorama.Fore.GREEN}Dependencies installed!...{colorama.Fore.BLUE}")

    print(f"{colorama.Fore.GREEN}Unpacking BepInEx...{colorama.Fore.BLUE}")
    shutil.unpack_archive('CitrusInstallTemp/BepInEx_win_x86_5.4.23.2.zip', 'CitrusInstallTemp/Unpacked/BepInEx')
    print(f"{colorama.Fore.GREEN}Unpacking Bepinject-Auros...{colorama.Fore.BLUE}")
    shutil.unpack_archive('CitrusInstallTemp/Bepinject-Auros.zip', 'CitrusInstallTemp/Unpacked/Bepinject-Auros')
    print(f"{colorama.Fore.GREEN}Unpacking Extenject...{colorama.Fore.BLUE}")
    shutil.unpack_archive('CitrusInstallTemp/Extenject.zip', 'CitrusInstallTemp/Unpacked/Extenject')
    print(f"{colorama.Fore.GREEN}Unpacking Newtonsoft.Json...{colorama.Fore.BLUE}")
    shutil.unpack_archive('CitrusInstallTemp/Newtonsoft.Json-12.0.3.zip', 'CitrusInstallTemp/Unpacked/Newtonsoft.Json')
    print(f"{colorama.Fore.GREEN}Unpacking TMPLoader...{colorama.Fore.BLUE}")
    shutil.unpack_archive('CitrusInstallTemp/TMPLoader-v1.0.2.zip', 'CitrusInstallTemp/Unpacked/TMPLoader')

    print(f"{colorama.Fore.GREEN}Dependencies unpacked!...{colorama.Fore.BLUE}")

    print(f"{colorama.Fore.GREEN}Moving BepInEx...{colorama.Fore.BLUE}")
    shutil.move("CitrusInstallTemp/Unpacked/BepInEx/BepInEx",
                "C:/Program Files (x86)/Steam/steamapps/common/Gorilla Tag/BepInEx")
    print(f"{colorama.Fore.GREEN}Creating plugins folder...{colorama.Fore.BLUE}")
    if not os.path.exists("C:/Program Files (x86)/Steam/steamapps/common/Gorilla Tag/BepInEx/plugins"):
        os.makedirs("C:/Program Files (x86)/Steam/steamapps/common/Gorilla Tag/BepInEx/plugins")
    print(f"{colorama.Fore.GREEN}Moving Bepinject-Auros...{colorama.Fore.BLUE}")
    shutil.move("CitrusInstallTemp/Unpacked/Bepinject-Auros/Bepinject-Auros",
                "C:/Program Files (x86)/Steam/steamapps/common/Gorilla Tag/BepInEx/plugins/Bepinject-Auros")
    print(f"{colorama.Fore.GREEN}Moving Extenject...{colorama.Fore.BLUE}")
    shutil.move("CitrusInstallTemp/Unpacked/Extenject/Extenject",
                "C:/Program Files (x86)/Steam/steamapps/common/Gorilla Tag/BepInEx/plugins/Extenject")
    print(f"{colorama.Fore.GREEN}Moving Newtonsoft.Json...{colorama.Fore.BLUE}")
    shutil.move("CitrusInstallTemp/Unpacked/Newtonsoft.Json/BepInEx/core/Newtonsoft.Json.dll",
                "C:/Program Files (x86)/Steam/steamapps/common/Gorilla Tag/BepInEx/core/Newtonsoft.Json.dll")
    print(f"{colorama.Fore.GREEN}Moving TMPLoader...{colorama.Fore.BLUE}")
    shutil.move("CitrusInstallTemp/Unpacked/TMPLoader/BepInEx/plugins/TMPLoader",
                "C:/Program Files (x86)/Steam/steamapps/common/Gorilla Tag/BepInEx/plugins/TMPLoader")

    print(f"{colorama.Fore.GREEN}Dependencies moved!...{colorama.Fore.BLUE}")
    print(f"{colorama.Fore.GREEN}Installing main application...{colorama.Fore.BLUE}")

    download_specific_github_asset("AllergenStudios", "Citrus-Mod-Manager",
                                   "CitrusModManager-v0.0.1-win.exe",
                                   download_path=f"{os.path.expanduser('~')}/CitrusModManager")

    pin_to_start_menu(f"{os.path.expanduser('~')}/CitrusModManager/CitrusModManager-v0.0.1-win.exe")

def pin_to_start_menu(exe_path):
    start_menu = winshell.startup()
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortcut(f'{start_menu}\\CitrusModManager.lnk')
    shortcut.TargetPath = exe_path
    shortcut.WorkingDirectory = exe_path
    shortcut.Save()

def download_specific_github_asset(owner, repo, asset_name, download_path="downloads"):
    """
    Downloads a specific asset from the latest release of a GitHub repository.

    Args:
        owner (str): The owner of the GitHub repository.
        repo (str): The name of the GitHub repository.
        asset_name (str): The exact name of the asset to download.
        download_path (str): Path to save the downloaded asset.
    """
    api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"

    try:
        response = requests.get(api_url)
        response.raise_for_status()
        release_data = response.json()

        tag_name = release_data.get("tag_name", "latest")
        assets = release_data.get("assets", [])

        print(f"Latest Release: {tag_name}")

        for asset in assets:
            if asset["name"] == asset_name:
                print(f"Found asset: {asset_name}")
                asset_url = asset["browser_download_url"]

                os.makedirs(download_path, exist_ok=True)

                print(f"Downloading {asset_name}...")
                asset_response = requests.get(asset_url, stream=True)
                asset_response.raise_for_status()

                asset_path = os.path.join(download_path, asset_name)
                with open(asset_path, "wb") as file:
                    for chunk in asset_response.iter_content(chunk_size=8192):
                        file.write(chunk)

                print(f"Asset downloaded successfully: {asset_path}")
                return

        print(f"Asset '{asset_name}' not found in the latest release.")

    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {e}")

print(f"{colorama.Fore.YELLOW}Citrus Mod Manager {version} Installer")
print(f"Would you like to install Citrus Mod Manager {version}?")
print(f"(Please note that Citrus Mod Manager only supports Windows.)")

while True:
    yn = input(f"{colorama.Fore.BLUE}[Y/N]: ")
    if yn.lower() == "y":
        break
    elif yn.lower() == "n":
        exit_installer = True
        break
    else:
        print(f"{colorama.Fore.RED}Invalid response")
if exit_installer:
    print(f"{colorama.Fore.RED}Exiting installer...")
else:
    if os.path.exists("C:/Program Files (x86)/Steam/steamapps/common/Gorilla Tag/BepInEx"):
        print(f"{colorama.Fore.YELLOW}All previously installed mods will be wiped, do you wish to continue?")
        while True:
            yn = input(f"{colorama.Fore.BLUE}[Y/N]: ")
            if yn.lower() == "y":
                break
            elif yn.lower() == "n":
                exit_installer = True
                break
            else:
                print(f"{colorama.Fore.RED}Invalid response")
        if exit_installer:
            print(f"{colorama.Fore.RED}Exiting installer...")
        else:
            print(f"{colorama.Fore.GREEN}Deleting original BepInEx...{colorama.Fore.BLUE}")
            shutil.rmtree("C:/Program Files (x86)/Steam/steamapps/common/Gorilla Tag/BepInEx")
            download_all_assets()
    else:
        download_all_assets()