import customtkinter as ctk
import os
import requests
import subprocess

# Define paths and URLs
odt_download_url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe"
odt_folder_path = "C:\\OfficeSetup"
odt_exe_path = os.path.join(odt_folder_path, "officedeploymenttool.exe")
config_xml_path = os.path.join(odt_folder_path, "config.xml")


class OfficeInstallerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Microsoft Office 2024 Installer")
        self.geometry("500x400")

        # Main Frame
        self.main_frame = ctk.CTkFrame(self, width=460, height=360)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Title Label
        self.title_label = ctk.CTkLabel(self.main_frame, text="Install Microsoft Office 2024 Professional Plus",
                                        font=("Helvetica", 16))
        self.title_label.pack(pady=10)

        # Install Button
        self.install_button = ctk.CTkButton(self.main_frame, text="Download & Install",
                                            command=self.download_and_install)
        self.install_button.pack(pady=20)

        # Progress Bar
        self.progress = ctk.CTkProgressBar(self.main_frame, width=400)
        self.progress.pack(pady=20)
        self.progress.set(0)

        # Status Label
        self.status_label = ctk.CTkLabel(self.main_frame, text="Status: Waiting for user action...",
                                         font=("Helvetica", 12))
        self.status_label.pack(pady=10)

    def download_and_install(self):
        # Step 1: Download ODT
        self.status_label.configure(text="Status: Downloading ODT...")
        self.update_idletasks()

        # Ensure the download folder exists
        if not os.path.exists(odt_folder_path):
            os.makedirs(odt_folder_path)

        response = requests.get(odt_download_url)
        with open(odt_exe_path, "wb") as file:
            file.write(response.content)

        # Check if download was successful
        if not os.path.exists(odt_exe_path):
            self.status_label.configure(text="Error: ODT download failed.")
            return

        # Step 2: Run ODT as Admin for Extraction
        self.status_label.configure(text="Status: Running ODT for extraction...")
        self.update_idletasks()

        # Run ODT and prompt user
        subprocess.run(f"powershell Start-Process '{odt_exe_path}' -Verb RunAs", shell=True)

        # Step 3: User Manual Intervention
        self.status_label.configure(text="Please accept the terms, click 'Continue', and extract to 'C:\\OfficeSetup'.")
        self.update_idletasks()

        # Ask user if extraction is successful
        if not self.confirm_extraction():
            self.status_label.configure(text="Extraction not confirmed. Exiting.")
            return

        # Verify required files exist
        if not os.path.exists(os.path.join(odt_folder_path, "setup.exe")):
            self.status_label.configure(text="Error: Required setup file 'setup.exe' not found after extraction.")
            return

        # Step 4: Write Configuration XML only after confirmation
        self.write_config_xml()

        # Step 5: Run the Installer with Admin Privileges
        self.status_label.configure(text="Status: Installing Office 2024 Professional Plus...")
        self.update_idletasks()
        command = f'powershell Start-Process "{os.path.join(odt_folder_path, "setup.exe")}" -ArgumentList "/configure, \\"{config_xml_path}\\"" -Verb RunAs'

        # Execute the installation command
        installation_process = subprocess.run(command, shell=True, capture_output=True, text=True)

        # Check installation result
        if installation_process.returncode != 0:
            self.status_label.configure(text=f"Error: Installation failed. {installation_process.stderr}")
            return

        # Step 6: Update UI when Installation is Complete
        self.status_label.configure(text="Status: Installation Complete!")
        self.progress.set(1)
        self.install_button.configure(state="disabled")

    def confirm_extraction(self):
        dialog = ctk.CTkInputDialog(text="Has the ODT extraction completed successfully? (yes/no)",
                                    title="Extraction Confirmation")
        response = dialog.get_input()
        if response.lower() == 'yes':
            self.status_label.configure(text="Extraction confirmed.")
            return True
        else:
            self.status_label.configure(text="Extraction not confirmed.")
            return False

    def write_config_xml(self):
        self.status_label.configure(text="Status: Writing configuration file...")
        self.update_idletasks()

        config_xml_content = """
<Configuration ID="15046a64-fe78-44e0-bb01-07521b0b21c1">
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume" PIDKEY="XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB">
      <Language ID="en-us" />
      <ExcludeApp ID="Lync" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <User Key="software\\microsoft\\office\\16.0\\excel\\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\\microsoft\\office\\16.0\\powerpoint\\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\\microsoft\\office\\16.0\\word\\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
  </AppSettings>
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"""
        with open(config_xml_path, "w") as config_file:
            config_file.write(config_xml_content)
            self.status_label.configure(text="Status: Configuration file written successfully.")


# Run the App
if __name__ == "__main__":
    app = OfficeInstallerApp()
    app.mainloop()