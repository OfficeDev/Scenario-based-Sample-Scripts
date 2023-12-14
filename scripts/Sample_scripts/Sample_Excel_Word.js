"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const shell = require("shelljs");
const cli_spinner_1 = require("cli-spinner");
const child_process_1 = require("child_process");
const readline = require("readline");
const usageData = require("office-addin-usage-data");
const os = require("os");
const fs = require("fs");
shell.config.silent = true;
const usageDataOptions = {
    groupName: usageData.groupName,
    projectName: "Sample_scripts",
    raisePrompt: false,
    instrumentationKey: usageData.instrumentationKeyForOfficeAddinCLITools,
    usageDataLevel: usageData.UsageDataLevel.off,
    method: usageData.UsageDataReportingMethod.applicationInsights,
    isForTesting: false
};
let usageDataObject = new usageData.OfficeAddinUsageData(usageDataOptions);
function exec_script_Excel_Mail() {
    return __awaiter(this, void 0, void 0, function* () {
        // shell.exec('code .');
        console.log('Explore the Excel Mail Merge Add-in project and dive into this Office add-in sample for an immersive experience!');
        console.log('--------------------------------------------------------------------------------------------------------');
        return new Promise((resolve, reject) => {
            let is_vscode_installed = false;
            // Step 1: Get sample code
            console.log('Step [1/3]: Getting sample code...');
            let spinner = new cli_spinner_1.Spinner('Processing.. %s');
            spinner.setSpinnerString('|/-\\');
            spinner.start();
            shell.exec('git clone https://github.com/OfficeDev/Excel-Scenario-based-Add-in-Samples.git', { async: true }, (code, stdout, stderr) => {
                replaceUrl('https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-add-in-mail-merge-localhost', 'https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-add-in-mail-merge-script', './Excel-Scenario-based-Add-in-Samples/Mail-Merge-Sample-Add-in/src/taskpane/taskpane.html')
                    .then(() => {
                    shell.cd('./Excel-Scenario-based-Add-in-Samples/Mail-Merge-Sample-Add-in');
                })
                    .catch((err) => {
                    console.error(err);
                });
                // shell.exec('git sparse-checkout set Mail-Merge-Sample-Add-in/', {async:true}, (code, stdout, stderr) => {
                spinner.stop(true);
                readline.clearLine(process.stdout, 0);
                readline.cursorTo(process.stdout, 0);
                console.log('Step [1/3] completed!');
                console.log('--------------------------------------------------------------------------------------------------------');
                // Ask user if sample Add-in automatic launch is needed
                let rl = readline.createInterface({
                    input: process.stdin,
                    output: process.stdout
                });
                let auto_launch_answer = false;
                rl.question('Proceed to launch Office with the sample add-in? (Y/N)\n', (answer) => {
                    if (answer.trim().toLowerCase() == 'y') {
                        auto_launch_answer = true;
                    }
                    rl.close();
                    // Step 2: Check if VSCode is installed
                    console.log('Step [2/3]: Checking if Visual Studio Code is installed...');
                    if (shell.which('code')) {
                        console.log('Visual Studio Code is installed on your machine. Ready to launch for code exploration.');
                        is_vscode_installed = true;
                        shell.exec('code -n . ./README.md');
                    }
                    else {
                        console.log('Visual Studio Code is not installed on your machine.');
                        if (os.platform() == 'darwin') {
                            shell.exec('open Mail-Merge-Sample-Add-in');
                        }
                        else if (os.platform() == 'win32') {
                            shell.exec('start Mail-Merge-Sample-Add-in');
                        }
                    }
                    console.log('Step [2/3] completed!');
                    console.log('--------------------------------------------------------------------------------------------------------');
                    reportUsageData('Excel_Mail', auto_launch_answer, is_vscode_installed);
                    if (auto_launch_answer) {
                        // Continue with the operations
                        // Step 3: Provide user the command to side-load add-in directly 
                        console.log('Step [3/3]: Automatically launch add-in with Excel...');
                        console.log('The process is expected to finish shortly, thank you for your patience...');
                        spinner.text = 'Processing... (installation of all dependencies may take a few minutes)';
                        spinner.start();
                        shell.cd('./Mail-Merge-Sample-Add-in');
                        let command_npm_install = 'npm install';
                        shell.exec('npm install', { async: true }, (code, stdout, stderr) => {
                            shell.exec('npm run start', { async: true }, (code, stdout, stderr) => {
                                spinner.stop(true);
                                readline.clearLine(process.stdout, 0);
                                readline.cursorTo(process.stdout, 0);
                                console.log('Step [3/3] completed!');
                                console.log('--------------------------------------------------------------------------------------------------------');
                                FreePortAlert();
                                console.log('Finished!');
                                console.log('--------------------------------------------------------------------------------------------------------');
                                // console.log('Hint: To try out the full functionality, please follow the instruction in the opening web page: Register a web application with the Azure Active Directory admin center.');
                                // open('https://github.com/OfficeDev/Excel-Scenario-based-Add-in-Samples/tree/main/Mail-Merge-Sample-Add-in');
                                resolve(is_vscode_installed);
                            });
                        });
                    }
                    else {
                        // Don't continue with the operations
                        console.log('Step [3/3] skipped. Auto-launch for the sample has been excluded based on your choice.');
                        console.log('And you can initiate the sample add-in by executing the following commands:');
                        console.log('--------------------------------------------------------------------------------------------------------');
                        console.log('npm install');
                        console.log('npm run start');
                        console.log('--------------------------------------------------------------------------------------------------------');
                        FreePortAlert();
                        console.log('Finished!');
                        console.log('--------------------------------------------------------------------------------------------------------');
                        // console.log('Hint: To try out the full functionality, please follow the instruction in the opening web page: Register a web application with the Azure Active Directory admin center.');
                        // open('https://github.com/OfficeDev/Excel-Scenario-based-Add-in-Samples/tree/main/Mail-Merge-Sample-Add-in');
                        resolve(is_vscode_installed);
                    }
                });
            });
        });
    });
}
function exec_script_Word_AIGC() {
    return __awaiter(this, void 0, void 0, function* () {
        // shell.exec('code .');
        console.log('Explore the Word AIGC Add-in project and dive into this Office add-in sample for an immersive experience!');
        console.log('--------------------------------------------------------------------------------------------------------');
        return new Promise((resolve, reject) => {
            let is_vscode_installed = false;
            // Step 1: Get sample code
            console.log('Step [1/3]: Getting sample code...');
            let spinner = new cli_spinner_1.Spinner('Processing.. %s');
            spinner.setSpinnerString('|/-\\');
            spinner.start();
            shell.exec('git clone https://github.com/OfficeDev/Word-Scenario-based-Add-in-Samples.git', { async: true }, (code, stdout, stderr) => {
                replaceUrl('https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-add-in-aigc-localhost', 'https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-add-in-aigc-script', './Word-Scenario-based-Add-in-Samples/Word-Add-in-AIGC/src/taskpane/taskpane.html')
                    .then(() => {
                    shell.cd('./Word-Scenario-based-Add-in-Samples/Word-Add-in-AIGC');
                })
                    .catch((err) => {
                    console.error(err);
                });
                // shell.cd('./Word-Scenario-based-Add-in-Samples/Word-Add-in-AIGC');
                // shell.exec('git sparse-checkout set Mail-Merge-Sample-Add-in/', {async:true}, (code, stdout, stderr) => {
                spinner.stop(true);
                readline.clearLine(process.stdout, 0);
                readline.cursorTo(process.stdout, 0);
                // Step 2: Check if VSCode is installed
                console.log('Step [1/3] completed!');
                console.log('--------------------------------------------------------------------------------------------------------');
                // Ask user if sample Add-in automatic launch is needed
                let rl = readline.createInterface({
                    input: process.stdin,
                    output: process.stdout
                });
                let auto_launch_answer = false;
                rl.question('Proceed to launch Office with the sample add-in? (Y/N)\n', (answer) => {
                    if (answer.trim().toLowerCase() == 'y') {
                        auto_launch_answer = true;
                    }
                    rl.close();
                    // Step 2: Check if VSCode is installed
                    console.log('Step [2/3]: Checking if Visual Studio Code is installed...');
                    if (shell.which('code')) {
                        console.log('Visual Studio Code is installed on your machine. Ready to launch for code exploration.');
                        is_vscode_installed = true;
                        shell.exec('code -n . ./README.md');
                    }
                    else {
                        console.log('Visual Studio Code is not installed on your machine.');
                        if (os.platform() == 'darwin') {
                            shell.exec('open Word-Add-in-AIGC');
                        }
                        else if (os.platform() == 'win32') {
                            shell.exec('start Word-Add-in-AIGC');
                        }
                    }
                    console.log('Step [2/3] completed!');
                    console.log('--------------------------------------------------------------------------------------------------------');
                    reportUsageData('Word_AIGC', auto_launch_answer, is_vscode_installed);
                    if (auto_launch_answer) {
                        // Continue with the operations
                        // Step 3: Provide user the command to side-load add-in directly 
                        console.log('Step [3/3]: Automatically launch add-in with Word...');
                        console.log('The process is expected to finish shortly, thank you for your patience...');
                        // spinner.text = 'Processing...(installation of all dependencies may take a few minutes)';
                        // spinner.start();
                        shell.cd('./Word-Add-in-AIGC');
                        shell.exec('npm set progress always');
                        const install = (0, child_process_1.spawn)('cmd.exe', ['/c', 'npm install --loglevel verbose']);
                        install.stdout.on('data', (data) => {
                            process.stdout.write(data);
                        });
                        install.stderr.on('data', (data) => {
                            process.stderr.write(data);
                        });
                        install.on('close', (code) => {
                            if (code !== 0) {
                                console.log(`Err: npm install process exited with code ${code}`);
                            }
                            const start = (0, child_process_1.spawn)('cmd.exe', ['/c', 'npm run start']);
                            start.stdout.on('data', (data) => {
                                console.log(`${data}`);
                            });
                            start.stderr.on('data', (data) => {
                                console.error(`stderr: ${data}`);
                            });
                            start.on('close', (code) => {
                                if (code !== 0) {
                                    console.log(`npm run start process exited with code ${code}`);
                                }
                                spinner.stop(true);
                                readline.clearLine(process.stdout, 0);
                                readline.cursorTo(process.stdout, 0);
                                console.log('Step [3/3] completed!');
                                console.log('--------------------------------------------------------------------------------------------------------');
                                FreePortAlert();
                                console.log('Finished!');
                                console.log('--------------------------------------------------------------------------------------------------------');
                                const rl = readline.createInterface({
                                    input: process.stdin,
                                    output: process.stdout
                                });
                                rl.question('Press any key to exit...', (answer) => {
                                    rl.close();
                                    resolve(is_vscode_installed);
                                });
                            });
                        });
                    }
                    else {
                        // Don't continue with the operations
                        console.log('Step [3/3] skipped. Auto-launch for the sample has been excluded based on your choice.');
                        console.log('And you can initiate the sample add-in by executing the following commands:');
                        console.log('--------------------------------------------------------------------------------------------------------');
                        console.log('npm install');
                        console.log('npm run start');
                        console.log('--------------------------------------------------------------------------------------------------------');
                        FreePortAlert();
                        console.log('Finished!');
                        console.log('--------------------------------------------------------------------------------------------------------');
                        resolve(is_vscode_installed);
                    }
                });
            });
        });
    });
}
function exec_script_Excel_Hello_World() {
    return __awaiter(this, void 0, void 0, function* () {
        // shell.exec('code .');
        console.log('Explore the Excel Hello World Add-in project and dive into this Office add-in sample for an immersive experience!');
        console.log('--------------------------------------------------------------------------------------------------------');
        return new Promise((resolve, reject) => {
            let is_vscode_installed = false;
            // Step 1: Get sample code
            console.log('Step [1/3]: Getting sample code...');
            let spinner = new cli_spinner_1.Spinner('Processing.. %s');
            spinner.setSpinnerString('|/-\\');
            spinner.start();
            shell.exec('git clone https://github.com/OfficeDev/Office-Addin-TaskPane-React.git', { async: true }, (code, stdout, stderr) => {
                shell.cd('./Office-Addin-TaskPane-React');
                shell.exec('npm run convert-to-single-host --if-present -- excel', { async: true }, (code, stdout, stderr) => {
                    spinner.stop(true);
                    readline.clearLine(process.stdout, 0);
                    readline.cursorTo(process.stdout, 0);
                    console.log('Step [1/3] completed!');
                    console.log('--------------------------------------------------------------------------------------------------------');
                    // Ask user if sample Add-in automatic launch is needed
                    let rl = readline.createInterface({
                        input: process.stdin,
                        output: process.stdout
                    });
                    let auto_launch_answer = false;
                    rl.on('error', (err) => {
                        console.error(`An error occurred: ${err.message}`);
                    });
                    rl.question('Proceed to launch Office with the sample add-in? (Y/N)\n', (answer) => {
                        if (answer.trim().toLowerCase() == 'y') {
                            auto_launch_answer = true;
                        }
                        rl.close();
                        // Step 2: Check if VSCode is installed
                        console.log('Step [2/3]: Checking if Visual Studio Code is installed...');
                        if (shell.which('code')) {
                            console.log('Visual Studio Code is installed on your machine. Ready to launch for code exploration.');
                            is_vscode_installed = true;
                            shell.exec('code -n . ./src/taskpane/office-document.ts');
                        }
                        else {
                            console.log('Visual Studio Code is not installed on your machine.');
                            if (os.platform() == 'darwin') {
                                shell.exec('open .');
                            }
                            else if (os.platform() == 'win32') {
                                shell.exec('start .');
                            }
                        }
                        console.log('Step [2/3] completed!');
                        console.log('--------------------------------------------------------------------------------------------------------');
                        reportUsageData('Excel_Hello_World', auto_launch_answer, is_vscode_installed);
                        if (auto_launch_answer) {
                            // Continue with the operations
                            // Step 3: Provide user the command to side-load add-in directly 
                            console.log('Step [3/3]: Automatically launch add-in with Excel...');
                            console.log('The process is expected to finish shortly, thank you for your patience...');
                            spinner.text = 'Processing... (installation of all dependencies may take a few minutes)';
                            spinner.start();
                            // shell.cd('./Mail-Merge-Sample-Add-in');
                            console.log(`Current path is: ${process.cwd()}`);
                            shell.exec('npm install', { async: true }, (code, stdout, stderr) => {
                                shell.exec('npm run start', { async: true }, (code, stdout, stderr) => {
                                    spinner.stop(true);
                                    readline.clearLine(process.stdout, 0);
                                    readline.cursorTo(process.stdout, 0);
                                    console.log('Step [3/3] completed!');
                                    console.log('--------------------------------------------------------------------------------------------------------');
                                    FreePortAlert();
                                    console.log('Finished!');
                                    console.log('--------------------------------------------------------------------------------------------------------');
                                    // console.log('Hint: To try out the full functionality, please follow the instruction in the opening web page: Register a web application with the Azure Active Directory admin center.');
                                    resolve(is_vscode_installed);
                                });
                            });
                        }
                        else {
                            // Don't continue with the operations
                            console.log('Step [3/3] skipped. Auto-launch for the sample has been excluded based on your choice.');
                            console.log('And you can initiate the sample add-in by executing the following commands:');
                            console.log('--------------------------------------------------------------------------------------------------------');
                            console.log('npm install');
                            console.log('npm run start');
                            console.log('--------------------------------------------------------------------------------------------------------');
                            console.log('Finished!');
                            console.log('--------------------------------------------------------------------------------------------------------');
                            // console.log('Hint: To try out the full functionality, please follow the instruction in the opening web page: Register a web application with the Azure Active Directory admin center.');
                            resolve(is_vscode_installed);
                        }
                    });
                });
            });
        });
    });
}
function exec_script_Word_Hello_World() {
    return __awaiter(this, void 0, void 0, function* () {
        // shell.exec('code .');
        console.log('Explore the Word Hello World Add-in project and dive into this Office add-in sample for an immersive experience!');
        console.log('--------------------------------------------------------------------------------------------------------');
        return new Promise((resolve, reject) => {
            let is_vscode_installed = false;
            // Step 1: Get sample code
            console.log('Step [1/3]: Getting sample code...');
            let spinner = new cli_spinner_1.Spinner('Processing.. %s');
            spinner.setSpinnerString('|/-\\');
            spinner.start();
            shell.exec('git clone https://github.com/OfficeDev/Office-Addin-TaskPane-React.git', { async: true }, (code, stdout, stderr) => {
                shell.cd('./Office-Addin-TaskPane-React');
                shell.exec('npm run convert-to-single-host --if-present -- word', { async: true }, (code, stdout, stderr) => {
                    spinner.stop(true);
                    readline.clearLine(process.stdout, 0);
                    readline.cursorTo(process.stdout, 0);
                    // Step 2: Check if VSCode is installed
                    console.log('Step [1/3] completed!');
                    console.log('--------------------------------------------------------------------------------------------------------');
                    // Ask user if sample Add-in automatic launch is needed
                    let rl = readline.createInterface({
                        input: process.stdin,
                        output: process.stdout
                    });
                    let auto_launch_answer = false;
                    rl.question('Proceed to launch Office with the sample add-in? (Y/N)\n', (answer) => {
                        if (answer.trim().toLowerCase() == 'y') {
                            auto_launch_answer = true;
                        }
                        rl.close();
                        // Step 2: Check if VSCode is installed
                        console.log('Step [2/3]: Checking if Visual Studio Code is installed...');
                        if (shell.which('code')) {
                            console.log('Visual Studio Code is installed on your machine. Ready to launch for code exploration.');
                            is_vscode_installed = true;
                            shell.exec('code -n . ./src/taskpane/office-document.ts');
                        }
                        else {
                            console.log('Visual Studio Code is not installed on your machine.');
                            if (os.platform() == 'darwin') {
                                shell.exec('open .');
                            }
                            else if (os.platform() == 'win32') {
                                shell.exec('start .');
                            }
                        }
                        console.log('Step [2/3] completed!');
                        console.log('--------------------------------------------------------------------------------------------------------');
                        reportUsageData('Word_Hello_World', auto_launch_answer, is_vscode_installed);
                        if (auto_launch_answer) {
                            // Continue with the operations
                            // Step 3: Provide user the command to side-load add-in directly 
                            console.log('Step [3/3]: Automatically launch add-in with Word...');
                            console.log('The process is expected to finish shortly, thank you for your patience...');
                            spinner.text = 'Processing... (installation of all dependencies may take a few minutes)';
                            spinner.start();
                            // shell.cd('./Word-Add-in-AIGC');
                            shell.exec('npm install', { async: true }, (code, stdout, stderr) => {
                                shell.exec('npm run start', { async: true }, (code, stdout, stderr) => {
                                    spinner.stop(true);
                                    readline.clearLine(process.stdout, 0);
                                    readline.cursorTo(process.stdout, 0);
                                    console.log('Step [3/3] completed!');
                                    console.log('--------------------------------------------------------------------------------------------------------');
                                    FreePortAlert();
                                    console.log('Finished!');
                                    console.log('--------------------------------------------------------------------------------------------------------');
                                    resolve(is_vscode_installed);
                                });
                            });
                        }
                        else {
                            // Don't continue with the operations
                            console.log('Step [3/3] skipped. Auto-launch for the sample has been excluded based on your choice.');
                            console.log('And you can initiate the sample add-in by executing the following commands:');
                            console.log('--------------------------------------------------------------------------------------------------------');
                            console.log('npm install');
                            console.log('npm run start');
                            console.log('--------------------------------------------------------------------------------------------------------');
                            FreePortAlert;
                            console.log('Finished!');
                            console.log('--------------------------------------------------------------------------------------------------------');
                            resolve(is_vscode_installed);
                        }
                    });
                });
            });
        });
    });
}
function replaceUrl(url, newUrl, filePath) {
    return new Promise((resolve, reject) => {
        fs.readFile(filePath, 'utf8', (err, data) => {
            if (err) {
                console.error(err);
                return;
            }
            const result = data.replace(url, newUrl);
            fs.writeFile(filePath, result, 'utf8', (err) => {
                if (err) {
                    console.error(err);
                    return;
                }
                resolve(true);
            });
        });
    });
}
function reportUsageData(scriptName, isAutomaticallyLaunch, isVscodeInstalled) {
    const projectInfo = {
        ScriptName: [scriptName],
        ScriptType: ["TypeScript"],
        isAutomaticallyLaunch: [isAutomaticallyLaunch],
        isVscodeInstalled: [isVscodeInstalled]
    };
    usageDataObject.reportEvent("sample_scripts", projectInfo);
}
function FreePortAlert() {
    if (os.platform() == 'darwin') {
        console.log('Hint: For mac users, if the add-in cannot be loaded correctly because of the add-in you loaded before, please try to run the commands below:');
        console.log('lsof -i:3000');
        console.log('kill -9 <PID>');
        console.log('--------------------------------------------------------------------------------------------------------');
        console.log('Then try to run the add-in again:');
        console.log('npm run start');
        console.log('--------------------------------------------------------------------------------------------------------');
    }
    else if (os.platform() == 'win32') {
        console.log('Hint: For windows users, if the add-in cannot be loaded correctly because of the add-in you loaded before, please try to run the commands below:');
        console.log('netstat -ano | findstr :3000');
        console.log('taskkill /PID <PID> /F');
        console.log('--------------------------------------------------------------------------------------------------------');
        console.log('Then try to run the add-in again:');
        console.log('npm run start');
        console.log('--------------------------------------------------------------------------------------------------------');
    }
}
function installDependencies() {
    const npmInstall = (0, child_process_1.spawn)('npm', ['install']);
    npmInstall.stdout.on('data', (data) => {
        process.stdout.write(data);
    });
    npmInstall.stderr.on('data', (data) => {
        process.stderr.write(data);
    });
    npmInstall.on('close', (code) => {
        if (code !== 0) {
            console.log(`npm install process exited with code ${code}`);
        }
        else {
            console.log('Dependencies installed successfully');
        }
    });
}
function launchAddinAndExit() {
}
module.exports = { exec_script_Excel_Mail, exec_script_Word_AIGC, exec_script_Excel_Hello_World, exec_script_Word_Hello_World };
//# sourceMappingURL=Sample_Excel_Word.js.map