```

__/\\\\\\\\\\\\___________/\\\________/\\\________/\\\\____________/\\\\________/\\\\\\\\\\\\\_________        
 _\/\\\////////\\\________\/\\\_______\/\\\_______\/\\\\\\________/\\\\\\_______\/\\\/////////\\\_______       
  _\/\\\______\//\\\_______\/\\\_______\/\\\_______\/\\\//\\\____/\\\//\\\_______\/\\\_______\/\\\_______      
   _\/\\\_______\/\\\_______\/\\\_______\/\\\_______\/\\\\///\\\/\\\/_\/\\\_______\/\\\\\\\\\\\\\\________     
    _\/\\\_______\/\\\_______\/\\\_______\/\\\_______\/\\\__\///\\\/___\/\\\_______\/\\\/////////\\\_______    
     _\/\\\_______\/\\\_______\/\\\_______\/\\\_______\/\\\____\///_____\/\\\_______\/\\\_______\/\\\_______   
      _\/\\\_______/\\\________\//\\\______/\\\________\/\\\_____________\/\\\_______\/\\\_______\/\\\_______  
       _\/\\\\\\\\\\\\/____/\\\__\///\\\\\\\\\/____/\\\_\/\\\_____________\/\\\__/\\\_\/\\\\\\\\\\\\\/___/\\\_ 
        _\////////////_____\///_____\/////////_____\///__\///______________\///__\///__\/////////////____\///__
                                                           
```

## Overview
**D.U.M.B - Data Unifier & Management Bot** is a tool designed to streamline the process of generating **Bills of Materials (BoMs)** for Odoo's manufacturing module, particularly for businesses dealing with **highly customizable products** that have numerous **variants with subtle changes**.

This tool eliminates the hassle of manually preparing BoMs by **automating file generation** and ensuring that all data is formatted correctly for direct import into Odoo, requiring **no manual adjustments**.

## Features
✅ **Automates BoM Creation** – Quickly generates BoMs ready for import into Odoo’s manufacturing module.
✅ **Handles Complex Variants** – Processes products with multiple variant options and structures them properly.
✅ **Removes Manual Adjustments** – Ensures data is correctly formatted, reducing errors and saving time.
✅ **Streamlines Workflow** – Designed for businesses that use Odoo's manufacturing module for highly customizable products.
✅ **Easy Data Mapping** – Assigns raw materials to specific attribute values with a user-friendly interface.

## Installation
No installation is needed, it's packaged to run on its own. go to [Releases](https://github.com/E1310/Data-Unifier-and-Management-Bot/releases/tag/v1.0.0) to donwload. 

### **Dependencies**
Ensure you have the following installed:
- Python 3.x
- Pandas
- PyQt6
- OpenPyXL
- 

Install dependencies with:
```sh
pip install -r requirements.txt
```

## Usage Guide
### **1️⃣ Export Data from Odoo**
#### **Step 1: Export Product Variants**
- **Go to the Product Variant view in Odoo** (not the Product view)
- **Filter:** Ensure **Variant Values** is set
- **Export:** Select *I want to update data*
- **Required Fields:** `Product Template ID`, `Name`, `Variant Values`

#### **Step 2: Export Products (Raw Materials & Components)**
- **Go to the Product Variant view in Odoo** (not the Product view)
- **Filter:** Remove finished products, keep only raw materials
- **Export:** Select *I want to update data*
- **Required Fields:** `Name`, `Unit of Measure`

### **2️⃣ Assign Components to Product Variants**
- You can only process **one product template at a time**
- Select the **product template** to assign components to
- For each **attribute value**, assign one or more **components** (raw materials) for the BoM

### **3️⃣ Generate the BoM CSV File**
- Select **BoM type** (Standard or Kit)
- Set the **quantity of products** to be produced
- Export the **BoM file** for **direct import into Odoo**

## Packaging & Distribution *(To be completed)*
- Convert the script into an **executable file** (Windows `.exe`, macOS/Linux binaries)
- Package dependencies so users **don’t need to install Python manually**
- Create an **installer** for easy setup
- Add **versioning & updates**

## Tech Stack
🚀 **Built with:**
- **Python** – Core scripting language
- **Pandas** – Data processing & manipulation
- **PyQt6** – Graphical User Interface
- **OpenPyXL** – Excel file handling

## License
📜 **Open Source** – Free to use, modify, and distribute.

## Author
👤 **Daher Zaidan**  
🔗 [GitHub](https://github.com/E1310)

