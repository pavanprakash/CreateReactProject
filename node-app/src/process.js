const shell = require("shelljs");
shell.exec(`powershell ../ps/checkPowershellRunStatus.ps1 -folder monthly`);
