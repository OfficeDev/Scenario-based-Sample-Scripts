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
const shelljs_1 = require("shelljs");
const cli_spinner_1 = require("cli-spinner");
const readline = require("readline");
shelljs_1.default.config.silent = true;
function exec_script() {
    return __awaiter(this, void 0, void 0, function* () {
        // shell.exec('code .');
        console.log('Welcome to experience this Office add-in sample!');
        return new Promise((resolve, reject) => {
            let is_vscode_installed = false;
            console.log('Welcome to experience this Office add-in sample!');
            // Step 1: Get sample code
            console.log('Step [1/3]: Getting sample code...');
            let spinner = new cli_spinner_1.Spinner('Processing.. %s');
            spinner.setSpinnerString('|/-\\');
            spinner.start();
            shelljs_1.default.exec('git clone --depth 1 --filter=blob:none --sparse https://github.com/OfficeDev/Office-Add-in-samples.git ./Office_add_in_sample', { async: true }, (code, stdout, stderr) => {
                shelljs_1.default.cd('./Office_add_in_sample');
                shelljs_1.default.exec('git sparse-checkout set Samples/Excel.OfflineStorageAddin/', { async: true }, (code, stdout, stderr) => {
                    spinner.stop(true);
                    readline.clearLine(process.stdout, 0);
                    readline.cursorTo(process.stdout, 0);
                    // Step 2: Check if VSCode is installed
                    console.log('Step [1/3] completed!');
                    console.log('Step [2/3]: Checking if Visual Studio Code is installed...');
                    if (shelljs_1.default.which('code')) {
                        console.log('Visual Studio Code is installed on your machine. Would open in VSCode for exploring the code.');
                        is_vscode_installed = true;
                        shelljs_1.default.exec('code ./Samples/Excel.OfflineStorageAddin README.md');
                    }
                    else {
                        console.log('Visual Studio Code is not installed on your machine.');
                        shelljs_1.default.exec('start Samples\\Excel.OfflineStorageAddin');
                    }
                    console.log('Step [2/3] completed!');
                    // Ask user if sample Add-in automatic launch is needed
                    let rl = readline.createInterface({
                        input: process.stdin,
                        output: process.stdout
                    });
                    let auto_launch_answer = false;
                    rl.question('Do you want to continue with some operations? (Y/N)\n', (answer) => {
                        console.log(`Your input was: ${answer}`);
                        if (answer.trim().toLowerCase() == 'y') {
                            // Continue with the operations
                            // Step 3: Provide user the command to side-load add-in directly 
                            console.log('Step [3/3]: Automatically side-load add-in directly...');
                            spinner.start();
                            shelljs_1.default.cd('./Samples/Excel.OfflineStorageAddin');
                            shelljs_1.default.exec('npm install', { async: true }, (code, stdout, stderr) => {
                                shelljs_1.default.exec('npm run start', { async: true }, (code, stdout, stderr) => {
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
                            console.log('No problem. You can always launch the sample add-in by running the following commands:');
                            resolve(is_vscode_installed);
                        }
                        rl.close();
                    });
                    // resolve(is_vscode_installed);
                    // if (!auto_launch_answer) {
                    //     resolve(is_vscode_installed);
                    // }
                });
            });
        });
    });
}
// exec_script();
module.exports = { exec_script };
//# sourceMappingURL=Sample_script.js.map