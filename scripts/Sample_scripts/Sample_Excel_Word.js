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
const readline = require("readline");
const usageData = require("office-addin-usage-data");
const os = require("os");
const fs = require("fs");
// shell.config.silent = true;
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
        console.log('Welcome to experience this Office add-in sample: Excel Mail merge Add-in project!');
        return new Promise((resolve, reject) => {
            let is_vscode_installed = false;
            // Step 1: Get sample code
            console.log('Step [1/3]: Getting sample code...');
            let spinner = new cli_spinner_1.Spinner('Processing.. %s');
            spinner.setSpinnerString('|/-\\');
            spinner.start();
            shell.exec('git clone https://github.com/OfficeDev/Excel-Scenario-based-Add-in-Samples.git', { async: true }, (code, stdout, stderr) => {
                shell.cd('./Excel-Scenario-based-Add-in-Samples/Mail-Merge-Sample-Add-in');
                // shell.exec('git sparse-checkout set Mail-Merge-Sample-Add-in/', {async:true}, (code, stdout, stderr) => {
                spinner.stop(true);
                readline.clearLine(process.stdout, 0);
                readline.cursorTo(process.stdout, 0);
                replaceUrl('https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-add-in-mail-merge-localhost', 'https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-add-in-mail-merge-script', './src/taskpane/taskpane.html');
                console.log('Step [1/3] completed!');
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
                        console.log('Visual Studio Code is installed on your machine. Would open in VSCode for exploring the code.');
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
                    reportUsageData('Excel_Mail', auto_launch_answer);
                    if (auto_launch_answer) {
                        // Continue with the operations
                        // Step 3: Provide user the command to side-load add-in directly 
                        console.log('Step [3/3]: Automatically side-load add-in directly...');
                        console.log('It may take longer time to complete the process. Please wait patiently...');
                        spinner.start();
                        shell.cd('./Mail-Merge-Sample-Add-in');
                        shell.exec('npm install', { async: true }, (code, stdout, stderr) => {
                            shell.exec('npm run start', { async: true }, (code, stdout, stderr) => {
                                spinner.stop(true);
                                readline.clearLine(process.stdout, 0);
                                readline.cursorTo(process.stdout, 0);
                                console.log('Step [3/3] completed!');
                                console.log('Finished!');
                                // console.log('Hint: To try out the full functionality, please follow the instruction in the opening web page: Register a web application with the Azure Active Directory admin center.');
                                // open('https://github.com/OfficeDev/Excel-Scenario-based-Add-in-Samples/tree/main/Mail-Merge-Sample-Add-in');
                                resolve(is_vscode_installed);
                            });
                        });
                    }
                    else {
                        // Don't continue with the operations
                        console.log('Step [3/3] skipped. You decided not to auto-launch the sample.');
                        console.log('No problem. You can always launch the sample add-in by running the following commands:');
                        console.log('--------------------------------------------');
                        console.log('npm install');
                        console.log('npm run start');
                        console.log('--------------------------------------------');
                        console.log('Finished!');
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
        console.log('Welcome to experience this Office add-in sample: Word AIGC Add-in project!');
        return new Promise((resolve, reject) => {
            let is_vscode_installed = false;
            // Step 1: Get sample code
            console.log('Step [1/3]: Getting sample code...');
            let spinner = new cli_spinner_1.Spinner('Processing.. %s');
            spinner.setSpinnerString('|/-\\');
            spinner.start();
            shell.exec('git clone https://github.com/OfficeDev/Word-Scenario-based-Add-in-Samples.git', { async: true }, (code, stdout, stderr) => {
                shell.cd('./Word-Scenario-based-Add-in-Samples/Word-Add-in-AIGC');
                // shell.exec('git sparse-checkout set Mail-Merge-Sample-Add-in/', {async:true}, (code, stdout, stderr) => {
                replaceUrl('https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-add-in-aigc-localhost', 'https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-add-in-aigc-script', './src/taskpane/taskpane.html');
                spinner.stop(true);
                readline.clearLine(process.stdout, 0);
                readline.cursorTo(process.stdout, 0);
                // Step 2: Check if VSCode is installed
                console.log('Step [1/3] completed!');
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
                        console.log('Visual Studio Code is installed on your machine. Would open in VSCode for exploring the code.');
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
                    reportUsageData('Word_AIGC', auto_launch_answer);
                    if (auto_launch_answer) {
                        // Continue with the operations
                        // Step 3: Provide user the command to side-load add-in directly 
                        console.log('Step [3/3]: Automatically side-load add-in directly...');
                        spinner.start();
                        shell.cd('./Word-Add-in-AIGC');
                        // shell.config.silent = false;
                        shell.exec('npm install --loglevel verbose', { async: true }, (code, stdout, stderr) => {
                            console.log(stdout);
                            shell.exec('npm run start', { async: true }, (code, stdout, stderr) => {
                                spinner.stop(true);
                                readline.clearLine(process.stdout, 0);
                                readline.cursorTo(process.stdout, 0);
                                console.log('Step [3/3] completed!');
                                console.log('Finished!');
                                resolve(is_vscode_installed);
                            });
                        });
                    }
                    else {
                        // Don't continue with the operations
                        console.log('Step [3/3] skipped. You decided not to auto-launch the sample.');
                        console.log('No problem. You can always launch the sample add-in by running the following commands:');
                        console.log('--------------------------------------------');
                        console.log('npm install');
                        console.log('npm run start');
                        console.log('--------------------------------------------');
                        console.log('Finished!');
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
        console.log('Welcome to experience this Office add-in sample: Excel Hello World Add-in!');
        return new Promise((resolve, reject) => {
            let is_vscode_installed = false;
            console.log('Welcome to experience this Office add-in sample!');
            // Step 1: Get sample code
            console.log('Step [1/3]: Getting sample code...');
            let spinner = new cli_spinner_1.Spinner('Processing.. %s');
            spinner.setSpinnerString('|/-\\');
            spinner.start();
            shell.exec('git clone https://github.com/OfficeDev/Office-Addin-TaskPane-React.git', { async: true }, (code, stdout, stderr) => {
                shell.cd('./Office-Addin-TaskPane-React');
                shell.exec('npm run convert-to-single-host --if-present --Excel', { async: true }, (code, stdout, stderr) => {
                    spinner.stop(true);
                    readline.clearLine(process.stdout, 0);
                    readline.cursorTo(process.stdout, 0);
                    console.log('Step [1/3] completed!');
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
                            console.log('Visual Studio Code is installed on your machine. Would open in VSCode for exploring the code.');
                            is_vscode_installed = true;
                            shell.exec('code -n . ./README.md');
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
                        reportUsageData('Excel_Hello_World', auto_launch_answer);
                        if (auto_launch_answer) {
                            // Continue with the operations
                            // Step 3: Provide user the command to side-load add-in directly 
                            console.log('Step [3/3]: Automatically side-load add-in directly...');
                            console.log('It may take longer time to complete the process. Please wait patiently...');
                            spinner.start();
                            // shell.cd('./Mail-Merge-Sample-Add-in');
                            console.log(`Current path is: ${process.cwd()}`);
                            shell.exec('npm install', { async: true }, (code, stdout, stderr) => {
                                shell.exec('npm run start', { async: true }, (code, stdout, stderr) => {
                                    spinner.stop(true);
                                    readline.clearLine(process.stdout, 0);
                                    readline.cursorTo(process.stdout, 0);
                                    console.log('Step [3/3] completed!');
                                    console.log('Finished!');
                                    // console.log('Hint: To try out the full functionality, please follow the instruction in the opening web page: Register a web application with the Azure Active Directory admin center.');
                                    resolve(is_vscode_installed);
                                });
                            });
                        }
                        else {
                            // Don't continue with the operations
                            console.log('Step [3/3] skipped. You decided not to auto-launch the sample.');
                            console.log('No problem. You can always launch the sample add-in by running the following commands:');
                            console.log('--------------------------------------------');
                            console.log('npm install');
                            console.log('npm run start');
                            console.log('--------------------------------------------');
                            console.log('Finished!');
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
        console.log('Welcome to experience this Office add-in sample: Word Hello World Add-in project!');
        return new Promise((resolve, reject) => {
            let is_vscode_installed = false;
            console.log('Welcome to experience this Office add-in sample!');
            // Step 1: Get sample code
            console.log('Step [1/3]: Getting sample code...');
            let spinner = new cli_spinner_1.Spinner('Processing.. %s');
            spinner.setSpinnerString('|/-\\');
            spinner.start();
            shell.exec('git clone git clone https://github.com/OfficeDev/Office-Addin-TaskPane-React.git', { async: true }, (code, stdout, stderr) => {
                shell.cd('./Office-Addin-TaskPane-React');
                // shell.exec('git sparse-checkout set Mail-Merge-Sample-Add-in/', {async:true}, (code, stdout, stderr) => {
                spinner.stop(true);
                readline.clearLine(process.stdout, 0);
                readline.cursorTo(process.stdout, 0);
                // Step 2: Check if VSCode is installed
                console.log('Step [1/3] completed!');
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
                        console.log('Visual Studio Code is installed on your machine. Would open in VSCode for exploring the code.');
                        is_vscode_installed = true;
                        shell.exec('code -n . ./README.md');
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
                    reportUsageData('Word_Hello_World', auto_launch_answer);
                    if (auto_launch_answer) {
                        // Continue with the operations
                        // Step 3: Provide user the command to side-load add-in directly 
                        console.log('Step [3/3]: Automatically side-load add-in directly...');
                        spinner.start();
                        // shell.cd('./Word-Add-in-AIGC');
                        shell.exec('npm install', { async: true }, (code, stdout, stderr) => {
                            shell.exec('npm run start', { async: true }, (code, stdout, stderr) => {
                                spinner.stop(true);
                                readline.clearLine(process.stdout, 0);
                                readline.cursorTo(process.stdout, 0);
                                console.log('Step [3/3] completed!');
                                console.log('Finished!');
                                resolve(is_vscode_installed);
                            });
                        });
                    }
                    else {
                        // Don't continue with the operations
                        console.log('Step [3/3] skipped. You decided not to auto-launch the sample.');
                        console.log('No problem. You can always launch the sample add-in by running the following commands:');
                        console.log('--------------------------------------------');
                        console.log('npm install');
                        console.log('npm run start');
                        console.log('--------------------------------------------');
                        console.log('Finished!');
                        resolve(is_vscode_installed);
                    }
                });
            });
        });
    });
}
function replaceUrl(url, newUrl, filePath) {
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
        });
    });
}
function reportUsageData(scriptName, isAutomaticallyLaunch) {
    const projectInfo = {
        ScriptName: [scriptName],
        ScriptType: ["TypeScript"],
        isAutomaticallyLaunch: [isAutomaticallyLaunch]
    };
    usageDataObject.reportEvent("sample_scripts", projectInfo);
}
module.exports = { exec_script_Excel_Mail, exec_script_Word_AIGC, exec_script_Excel_Hello_World, exec_script_Word_Hello_World };
//# sourceMappingURL=Sample_Excel_Word.js.map