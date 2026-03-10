# Prompt password (recommended)
.\ocs_export.ps1 `
  -BaseUrl "https://your-server/ocsreports" `
  -Username "your_user" `
  -PromptForPassword `
  -OutputDirectory "C:\Exports\OCS"

# Password via parameter (less secure, but possible)
.\ocs_export.ps1 `
  -BaseUrl "https://your-server/ocsreports" `
  -Username "your_user" `
  -Password "your_password"