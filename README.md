# A GUI for Active Directory with PowerShell
A powershell script using a GUI in XAML to work with AD users and computers faster than RSAT AD tool
### Notes
- Gets all computers and users under the domain the user that ran the script 
- The computers are listed by description and the selected value is the computer name
- The users are listed by displayname and the selected value is the distinguishedName
- For users includes:
    - Get groups a user is member of (output in a GridView table)
    - Unlock a user that has blocked the account
    - Reset the password of a user, asking for a new password that the user will need to change on next logon
- For computers includes:
    - Ping the computer
    - Open remote C:\ drive
    - Get groups the computer is member of
- Other options:
    - Update the computer and user list

### Examples
- Image
- Image
- Image
### Prerequisites
```
Powershell 3.0 or higher

```

## Built With

* [VSCommunity](https://visualstudio.microsoft.com/es/vs/community/) - Framework to create the GUI in XAML
* [Powershell](https://docs.microsoft.com/es-es/powershell/) - Main code
