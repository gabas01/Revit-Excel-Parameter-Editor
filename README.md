# Revit Excel Parameter Editor

![GitHub repo size](https://img.shields.io/github/repo-size/YOUR-USERNAME/Revit-Excel-Parameter-Editor?style=for-the-badge)
![GitHub language count](https://img.shields.io/github/languages/count/YOUR-USERNAME/Revit-Excel-Parameter-Editor?style=for-the-badge)
![GitHub top language](https://img.shields.io/github/languages/top/YOUR-USERNAME/Revit-Excel-Parameter-Editor?style=for-the-badge)
![GitHub last commit](https://img.shields.io/github/last-commit/YOUR-USERNAME/Revit-Excel-Parameter-Editor?style=for-the-badge)

An Autodesk Revit plugin that enables bulk editing of element parameters by exporting and re-importing schedules with Microsoft Excel.

![Plugin Demo](https://i.imgur.com/your-image-url.gif)  
*(Suggestion: Record a short screen GIF showing the plugin in action and replace the link above)*

---

## 🚀 Features

* **Intuitive UI:** A simple and straightforward window inside Revit.
* **Schedule Selection:** Automatically loads all project schedules into a dropdown list.
* **Safe Export:** Exports schedule data to an `.xlsx` file, ensuring each object's `ElementID` is included for error-free re-importing.
* **Smart Import:** Reads the Excel file, identifies elements by their `ElementID`, and updates only the modified parameters.
* **Safe Transactions:** All modifications are wrapped in a single Revit transaction, allowing the entire operation to be undone with a single click.

## 📦 Installation (For Users)

1.  Go to the [**Releases**](https://github.com/YOUR-USERNAME/Revit-Excel-Parameter-Editor/releases) page of this repository.
2.  Download the `.zip` file from the latest version.
3.  Unzip the file. You will find the main items inside:
    * The `RevitParameterEditor.bundle` folder (or the `RevitParameterEditor.dll` and `RevitParameterEditor.addin` files).
4.  Copy the `RevitParameterEditor.bundle` folder to the Revit Add-ins folder. The path is typically:
    * `C:\Users\YOUR-USERNAME\AppData\Roaming\Autodesk\Revit\Addins\[REVIT-VERSION]\`
5.  Start Autodesk Revit. The plugin will be available in the **Add-Ins** tab.

## 📋 How to Use

1.  With your project open in Revit, go to the **Add-Ins** tab.
2.  Click the **"Parameter Editor"** button.
3.  In the plugin window, select the schedule you want to edit from the list.
4.  Click **"Export to Excel"** and save the `.xlsx` file.
5.  Open the file in Excel. **Do not modify the `ElementID` column!** Edit the values in the other parameter columns as needed.
6.  Save your changes in Excel.
7.  Go back to the plugin window in Revit and click **"Import from Excel & Update"**.
8.  Select the Excel file you just modified.
9.  A success message will appear, telling you how many elements were updated. Your changes are now live in the Revit model!

## ⚙️ For Developers

This project was built with C# using .NET Framework 4.8 and the Revit 2025 API.

1.  Clone the repository:
    ```bash
    git clone [https://github.com/YOUR-USERNAME/Revit-Excel-Parameter-Editor.git](https://github.com/YOUR-USERNAME/Revit-Excel-Parameter-Editor.git)
    ```
2.  Open the solution file (`.sln`) in Visual Studio.
3.  The `RevitAPI.dll` and `RevitAPIUI.dll` references might need to be re-pathed. Point them to the corresponding files in your Revit installation directory.
4.  The `EPPlus` library dependency can be restored via NuGet Package Manager.
5.  Build the project.

## 📄 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
