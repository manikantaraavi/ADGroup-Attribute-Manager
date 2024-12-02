
# ADGroup Attribute Manager

A robust PowerShell script for managing and updating the `extensionAttribute6` field in Active Directory (AD) groups. This script allows you to easily remove a specified user from this attribute across multiple groups with logging and manual or file-based group selection.

---

## Features
- **Interactive User Input**: Specify groups manually or via a `groups.txt` file.
- **Smart Attribute Handling**: Automatically clears the attribute if empty after removing the user.
- **Automatic Formatting**: Handles multiple entries in `extensionAttribute6` values by ensuring proper formatting (e.g., removing the specified user and rejoining remaining users).
- **Detailed Logging**: Logs all operations to a timestamped log file for auditing.
- **Error Handling**: Reports issues for non-existent or inaccessible groups.
- **Preview Before Changes**: View proposed updates and confirm before execution.

---

## Prerequisites
1. **Active Directory PowerShell Module**: Ensure the module is installed and available.
   ```powershell
   Import-Module ActiveDirectory
   ```
2. **Permissions**: The script requires permissions to:
   - Read group attributes.
   - Modify `extensionAttribute6` for the specified groups.
3. **PowerShell Environment**: Run the script in a Windows environment with PowerShell 5.1+.

---

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/ADGroup-Attribute-Manager.git
   cd ADGroup-Attribute-Manager
   ```
2. Place a `groups.txt` file in the same directory if using the file-based group input method. Each group name should be on a new line.

---

## Usage

1. Open a PowerShell terminal and navigate to the script's directory.

2. Run the script:
   ```powershell
   .\ADGroup-Attribute-Manager.ps1
   ```

3. Follow the prompts:
   - Enter the username to remove.
   - Choose the group input method:
     - **Manual Input**: Type group names directly into the console.
     - **Import from File**: Use the `groups.txt` file.

4. Review the preview of changes and confirm if you want to proceed.

5. Logs of the operation will be saved in the same directory with a timestamp.

---

## Example `groups.txt`
```
GroupName1
GroupName2
GroupName3
```

---

## Output
### Log File Example:
- A detailed log will be created in the script directory.
- File name format: `attribute_changes_YYYYMMDD_HHMMSS.log`.

Sample log entry:
```
2024-12-02 12:45:03 - Starting script execution
2024-12-02 12:45:10 - Found 3 groups containing user 'JohnDoe' in extensionAttribute6
2024-12-02 12:47:15 - Successfully updated group 'GroupName1' - New value: UserA, UserB
```

---

## Troubleshooting
- **Module Error**: Ensure the Active Directory module is installed and loaded:
  ```powershell
  Install-WindowsFeature RSAT-AD-PowerShell
  ```
- **Permission Issues**: Verify you have sufficient rights to modify AD groups.

---

## Contributing
Contributions are welcome! Feel free to:
- Fork the repository.
- Create a pull request for bug fixes or new features.
- Open issues for feature suggestions or bugs.

---

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## Support
For any issues or questions, feel free to open an [issue](https://github.com/yourusername/ADGroup-Attribute-Manager/issues) or create a pull request. 

Happy managing! 🚀
