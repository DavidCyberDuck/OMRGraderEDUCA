import subprocess
import sys
import platform

def run(cmd):
    subprocess.run(cmd, check=True)

os_name = platform.system()

print(f"Detected OS: {os_name}")

print("\nInstalling system dependency: poppler")
if os_name == "Darwin":
    brew = subprocess.run(["which", "brew"], capture_output=True, text=True).stdout.strip()
    if not brew:
        print("Homebrew not found. Installing it now...")
        run(["/bin/bash", "-c",
             'curl -fsSL -o /tmp/brew_install.sh https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh && bash /tmp/brew_install.sh'])
        # Homebrew on Apple Silicon installs to /opt/homebrew, Intel to /usr/local
        for candidate in ["/opt/homebrew/bin/brew", "/usr/local/bin/brew"]:
            if subprocess.run(["test", "-f", candidate], capture_output=True).returncode == 0:
                brew = candidate
                break
        if not brew:
            print("Homebrew installed but brew binary not found. Please restart your terminal and re-run.")
            sys.exit(1)
    run([brew, "install", "poppler"])
elif os_name == "Windows":
    if subprocess.run(["where", "winget"], capture_output=True).returncode == 0:
        run(["winget", "install", "--id", "oscarblancartesarabia.poppler", "-e", "--silent"])
    elif subprocess.run(["where", "choco"], capture_output=True).returncode == 0:
        run(["choco", "install", "poppler", "-y"])
    else:
        print("No package manager found (winget or choco).")
        print("Download poppler for Windows from: https://github.com/oscarblancartesarabia/poppler-windows/releases")
        print("Extract it and add the bin/ folder to your PATH, then re-run this script.")
        sys.exit(1)
else:
    print(f"Unsupported OS: {os_name}. Install poppler manually and re-run.")
    sys.exit(1)

print("\nInstalling Python dependencies")
run([sys.executable, "-m", "pip", "install", "-r", "omr_app/requirements.txt"])

print("\nDone. Run the app with: python omr_app/main.py")
