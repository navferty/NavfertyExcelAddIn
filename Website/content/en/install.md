{
  "title": "Installation and removal",
  "description": "How to install the Navferty Excel Add-In, what to do about the certificate warning, how to install offline and how to remove it.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[Home](/en/) › Installation and removal

# Installation and removal

<div class="installation-steps">

<div class="step">

<div class="step-number">

1

</div>

<div class="step-content">

#### Download and unpack the installer archive

Get the latest version from our server:

<a href="https://deploy.navferty.net/setup.zip" class="btn btn-primary btn-small">Download setup.zip</a>

</div>

</div>

<div class="step">

<div class="step-number">

2

</div>

<div class="step-content">

#### Run the installation

Double-click the downloaded `setup.exe` file and follow the installation wizard.

</div>

</div>

<div class="step">

<div class="step-number">

3

</div>

<div class="step-content">

#### Restart Excel

If Excel was open during installation, restart it. You'll see the new **Navferty's Tools** tab in the ribbon.

</div>

</div>

</div>

<span id="certificate"></span>

## Certificate installation (if needed)

<div class="warning-box">

If you encounter a security warning about the certificate:

1.  Right-click on `setup.exe` and select **Properties**
2.  Go to the **Digital Signatures** tab
3.  Select the signature and click **Details**
4.  Click **View Certificate** → **Install Certificate**
5.  Choose **Local Machine** and place it in **Trusted Root Certification Authorities**

<img src="/digital-certificate.jpg" loading="lazy" width="1518" height="822" alt="Certificate installation wizard" />

*You can remove the certificate after installation, but it will be required for future updates.*

</div>

<span id="offline"></span>

## Offline install

Every release is built with GitHub Actions, and the full set of installation files is attached to it:

1.  Open the <a href="https://github.com/navferty/NavfertyExcelAddIn/releases/latest" rel="noopener" target="_blank">latest release</a>
2.  Download the attached `NavfertyExcelAddIn-*.zip` archive
3.  Extract the files to a folder and run the `.vsto` file

Excel records the location of the `.vsto` file you run and uses it to check for updates. Create a dedicated folder for the add-in, keep the extracted installation files there, and use that same folder when installing future updates.

<img src="/images/offline-install-addin-location.png" loading="lazy" width="1082" height="484" alt="Excel COM Add-ins dialog showing the add-in&#39;s installation location" />

After installation, restart Excel. The new **Navferty's Tools** tab appears in the ribbon.

<span id="uninstall"></span>

## Removal

**Method 1: Windows Settings (recommended)**

1.  Open Windows Settings (<kbd>Win</kbd> + <kbd>I</kbd>)
2.  Go to **Apps** → **Installed apps** (**Apps & features** on Windows 10)
3.  Find **Navferty Excel Add-In** in the list
4.  Click **Uninstall** and follow the prompts

**Method 2: Control Panel**

1.  Open **Control Panel**
2.  Go to **Programs** → **Programs and Features**
3.  Right-click **Navferty Excel Add-In** and select **Uninstall**

Problems with installation? See [support and feedback](/en/support/) for how to report a bug and attach logs.
