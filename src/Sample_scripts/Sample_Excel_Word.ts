import * as shell from 'shelljs';
import ProgressBar from 'progress';
import log from 'single-line-log';
import { Spinner } from 'cli-spinner';
import { spawn, exec, execSync } from 'child_process';
import * as readline from 'readline';
import open from 'open';
import * as usageData from 'office-addin-usage-data';
import { Writable } from 'stream';

import * as os from 'os';
import * as fs from 'fs';
shell.config.silent = true;

const usageDataOptions: usageData.IUsageDataOptions = {
    groupName: usageData.groupName,
    projectName: "Sample_scripts",
    raisePrompt: false,
    instrumentationKey: usageData.instrumentationKeyForOfficeAddinCLITools,
    usageDataLevel: usageData.UsageDataLevel.off,
    method: usageData.UsageDataReportingMethod.applicationInsights,
    isForTesting: false
  }
let usageDataObject: usageData.OfficeAddinUsageData = new usageData.OfficeAddinUsageData(usageDataOptions);

async function exec_script_Excel_Mail(){
  // shell.exec('code .');
  console.log('Explore the Excel Mail Merge Add-in project and dive into this Office add-in sample for an immersive experience!');
  console.log('--------------------------------------------------------------------------------------------------------');

  return new Promise<boolean>((resolve, reject) => {
    
        let is_vscode_installed = false;

        // Step 1: Get sample code
        console.log('Step [1/3]: Getting sample code...');
        let spinner = new Spinner('Processing.. %s');
        spinner.setSpinnerString('|/-\\');
        spinner.start();

        shell.exec('git clone https://github.com/OfficeDev/Excel-Scenario-based-Add-in-Samples.git', {async:true}, (code, stdout, stderr) => {
            
            replaceUrl('https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-add-in-mail-merge-localhost', 'https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-add-in-mail-merge-script', './Excel-Scenario-based-Add-in-Samples/Mail-Merge-Sample-Add-in/src/taskpane/taskpane.html')
                .then(() => {
                    shell.cd('./Excel-Scenario-based-Add-in-Samples/Mail-Merge-Sample-Add-in');
                    spinner.stop(true);

                    //Stop the spinner and clear the console
                    readline.clearLine(process.stdout, 0);
                    readline.cursorTo(process.stdout, 0);

                    console.log('Step [1/3] completed!');
                    console.log('--------------------------------------------------------------------------------------------------------');
            
                    // Ask user if sample Add-in automatic launch is needed
                    let originalWrite = process.stdout.write;
                    let silentStream = new Writable({
                        write(chunk, encoding, callback) {
                            callback();
                        }
                    });

                    let rl = readline.createInterface({
                        input: process.stdin,
                        output: silentStream
                    });

                    let auto_launch_answer = false;

                    console.log('Proceed to launch Office with the sample add-in? (Y/N)');
                    rl.question('', (answer) => {

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
                        } else {
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
                        reportUsageData('Excel_Mail', auto_launch_answer, is_vscode_installed);

                        if (auto_launch_answer) {
                            // Step 3: Provide user the command to side-load add-in directly 
                            console.log('Step [3/3]: Automatically launch add-in with Excel...');
                            console.log('The process is expected to finish shortly, thank you for your patience...');

                            shell.cd('./Mail-Merge-Sample-Add-in');
                            shell.exec('npm set progress always');

                            let env = 'cmd.exe';
                            let para = '/c';
                            if (os.platform() == 'darwin') {
                                env = 'sh';
                                para = '-c';
                            }

                            const install = spawn(env, [para, 'npm install --loglevel verbose']);
                            install.stdout.on('data', (data) => {
                                process.stdout.write(data);
                            });
                        
                            install.stderr.on('data', (data) => {
                                process.stderr.write(data);
                            });
                            install.on('close', (code) => {
                                if (code !== 0) {
                                    console.log(`Err: npm install process exited with code ${code}`);

                                    // Error handling on mac
                                    if (os.platform() == 'darwin') {
                                        // if npm install failed because of access issue on mac
                                        if (code == 243) {
                                            console.log('Mac access issue detected. Trying to automatically fix the issue...');

                                            // Get the current user's UID and GID
                                            const uid = process.getuid();
                                            const gid = process.getgid();

                                            shell.exec(`sudo chown -R ${uid}:${gid} ~/.npm`, {async:true}, (code, stdout, stderr) => {
                                                if (code !== 0) {
                                                    console.log(`Err: sudo chown process exited with code ${code}`);
                                                    console.error(`stderr: ${stderr}`);
                                                    console.log('Automatically fix the issue failed. Please try to run the following commands manually:');
                                                    console.log('sudo chown -R $uid:$gid ~/.npm');
                                                    console.log('where $uid and $gid are the current user\'s UID and GID, which can be get by running the following commands:');
                                                    console.log('id -u');
                                                    console.log('id -g');
                                                }
                                                else{
                                                    console.log('Issue fixed. Please try to run the sample command again.');
                                                    console.log('--------------------------------------------------------------------------------------------------------');
                                                    console.log('Hint: If the issue persists, please try to run the following commands manually:');
                                                    console.log('sudo chown -R $uid:$gid ~/.npm');
                                                    console.log('where $uid and $gid are the current user\'s UID and GID, which can be get by running the following commands:');
                                                    console.log('id -u');
                                                    console.log('id -g');
                                                }

                                                const rl = readline.createInterface({
                                                    input: process.stdin,
                                                    output: process.stdout
                                                });
                                            
                                                rl.question('Press Enter to exit...', (answer) => {
                                                    rl.close();
                                                    resolve(is_vscode_installed);
                                                });
                                            });
                                        }
                                    }
                                }
                                else {
                                    const start = spawn(env, [para, 'npm run start']);

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
                                    
                                        rl.question('Press Enter to exit...', (answer) => {
                                            rl.close();
                                            resolve(is_vscode_installed);
                                        });
                                    });

                                    // Make sure npm run start process will not be blocked by the prompt
                                    start.stdin.write('n\n');
                                    }
                            });
                        }
                        else{
                            // Don't continue with the operations
                            console.log('Step [3/3] skipped. Auto-launch for the sample has been excluded based on your choice.')
                            console.log('And you can initiate the sample add-in by executing the following commands:');
                            console.log('--------------------------------------------------------------------------------------------------------');
                            console.log('npm install');
                            console.log('npm run start');
                            console.log('--------------------------------------------------------------------------------------------------------');
                            FreePortAlert();     
                            console.log('Finished!');     
                            console.log('--------------------------------------------------------------------------------------------------------');        
                            let rl = readline.createInterface({
                                input: process.stdin,
                                output: process.stdout
                            });
                            rl.question('Press Enter to exit...', (answer) => {
                                rl.close();
                                resolve(is_vscode_installed);
                            });
                        }
                    });
                })
                .catch((err) => {
                    //Stop the spinner and clear the console
                    spinner.stop(true);
                    readline.clearLine(process.stdout, 0);
                    readline.cursorTo(process.stdout, 0);

                    console.error(err);
                    console.log('Error occurred when downloading the code. This may be caused by the network issue. Please rerun the command to try again.')
                    console.log('--------------------------------------------------------------------------------------------------------');

                    let rl = readline.createInterface({
                        input: process.stdin,
                        output: process.stdout
                    });
                
                    rl.question('Press Enter to exit...', (answer) => {
                        rl.close();
                        reject(err);
                    });
                });
        });
    });
}

async function exec_script_Word_AIGC(){
    // shell.exec('code .');
    console.log('Explore the Word AIGC Add-in project and dive into this Office add-in sample for an immersive experience!');
    console.log('--------------------------------------------------------------------------------------------------------');
  
    return new Promise<boolean>((resolve, reject) => {
      
          let is_vscode_installed = false;
  
          // Step 1: Get sample code
          console.log('Step [1/3]: Getting sample code...');
          let spinner = new Spinner('Processing.. %s');
          spinner.setSpinnerString('|/-\\');
          spinner.start();
  
          shell.exec('git clone https://github.com/OfficeDev/Word-Scenario-based-Add-in-Samples.git', {async:true}, (code, stdout, stderr) => {

            replaceUrl('https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-add-in-aigc-localhost', 'https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-add-in-aigc-script', './Word-Scenario-based-Add-in-Samples/Word-Add-in-AIGC/src/taskpane/taskpane.html')
                .then(() => {
                    shell.cd('./Word-Scenario-based-Add-in-Samples/Word-Add-in-AIGC');
                    spinner.stop(true);

                    //Stop the spinner and clear the console
                    readline.clearLine(process.stdout, 0);
                    readline.cursorTo(process.stdout, 0);

                    // Step 2: Check if VSCode is installed
                    console.log('Step [1/3] completed!');
                    console.log('--------------------------------------------------------------------------------------------------------');

                    let originalWrite = process.stdout.write;
                    let silentStream = new Writable({
                        write(chunk, encoding, callback) {
                            callback();
                        }
                    });

                    let rl = readline.createInterface({
                        input: process.stdin,
                        output: silentStream
                    });

                    let auto_launch_answer = false;

                    console.log('Proceed to launch Office with the sample add-in? (Y/N)');
                    rl.question('', (answer) => {
                        
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
                        } else {
                            console.log('Visual Studio Code is not installed on your machine.');
                            if (os.platform() == 'darwin') {
                                console.log(process.cwd());
                                shell.exec('open .');
                            }
                            else if (os.platform() == 'win32') {
                                shell.exec('start .');
                            }
                        }

                        console.log('Step [2/3] completed!');
                        console.log('--------------------------------------------------------------------------------------------------------');
                        reportUsageData('Word_AIGC', auto_launch_answer, is_vscode_installed);

                        if (auto_launch_answer) {
                            // Step 3: Provide user the command to side-load add-in directly 
                            console.log('Step [3/3]: Automatically launch add-in with Word...');
                            console.log('The process is expected to finish shortly, thank you for your patience...');

                            shell.cd('./Word-Add-in-AIGC');
                            shell.exec('npm set progress always');

                            let env = 'cmd.exe';
                            let para = '/c';
                            if (os.platform() == 'darwin') {
                                env = 'sh';
                                para = '-c';
                            }

                            const install = spawn(env, [para, 'npm install --loglevel verbose']);
                            install.stdout.on('data', (data) => {
                                process.stdout.write(data);
                            });
                        
                            install.stderr.on('data', (data) => {
                                process.stderr.write(data);
                            });
                            install.on('close', (code) => {
                                if (code !== 0) {
                                    console.log(`Err: npm install process exited with code ${code}`);

                                    // Error handling on mac
                                    if (os.platform() == 'darwin') {
                                        // if npm install failed because of access issue on mac
                                        if (code == 243) {
                                            console.log('Mac access issue detected. Trying to automatically fix the issue...');

                                            // Get the current user's UID and GID
                                            const uid = process.getuid();
                                            const gid = process.getgid();

                                            shell.exec(`sudo chown -R ${uid}:${gid} ~/.npm`, {async:true}, (code, stdout, stderr) => {
                                                if (code !== 0) {
                                                    console.log(`Err: sudo chown process exited with code ${code}`);
                                                    console.error(`stderr: ${stderr}`);
                                                    console.log('Automatically fix the issue failed. Please try to run the following commands manually:');
                                                    console.log('sudo chown -R $uid:$gid ~/.npm');
                                                    console.log('where $uid and $gid are the current user\'s UID and GID, which can be get by running the following commands:');
                                                    console.log('id -u');
                                                    console.log('id -g');
                                                }
                                                else{
                                                    console.log('Issue fixed. Please try to run the sample command again.');
                                                    console.log('--------------------------------------------------------------------------------------------------------');
                                                    console.log('Hint: If the issue persists, please try to run the following commands manually:');
                                                    console.log('sudo chown -R $uid:$gid ~/.npm');
                                                    console.log('where $uid and $gid are the current user\'s UID and GID, which can be get by running the following commands:');
                                                    console.log('id -u');
                                                    console.log('id -g');
                                                }

                                                const rl = readline.createInterface({
                                                    input: process.stdin,
                                                    output: process.stdout
                                                });
                                            
                                                rl.question('Press Enter to exit...', (answer) => {
                                                    rl.close();
                                                    resolve(is_vscode_installed);
                                                });
                                            });
                                        }
                                    }
                                }
                                else {
                                    // if npm install succeeded
                                    const start = spawn(env, [para, 'npm run start']);

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
                                    
                                        rl.question('Press Enter to exit...', (answer) => {
                                            rl.close();
                                            resolve(is_vscode_installed);
                                        });
                                    });

                                    // Make sure npm run start process will not be blocked by the prompt
                                    start.stdin.write('n\n');
                                    }
                            });
                        }
                        else{
                            // Don't continue with the operations
                            console.log('Step [3/3] skipped. Auto-launch for the sample has been excluded based on your choice.')
                            console.log('And you can initiate the sample add-in by executing the following commands:');
                            console.log('--------------------------------------------------------------------------------------------------------');
                            console.log('npm install');
                            console.log('npm run start');
                            console.log('--------------------------------------------------------------------------------------------------------');
                            FreePortAlert();     
                            console.log('Finished!');     
                            console.log('--------------------------------------------------------------------------------------------------------');        
                            let rl = readline.createInterface({
                                input: process.stdin,
                                output: process.stdout
                            });
                            rl.question('Press Enter to exit...', (answer) => {
                                rl.close();
                                resolve(is_vscode_installed);
                            });
                        }
                    });
                })
                .catch((err) => {
                    //Stop the spinner and clear the console
                    spinner.stop(true);
                    readline.clearLine(process.stdout, 0);
                    readline.cursorTo(process.stdout, 0);

                    console.error(err);
                    console.log('Error occurred when downloading the code. This may be caused by the network issue. Please rerun the command to try again.')
                    console.log('--------------------------------------------------------------------------------------------------------');

                    let rl = readline.createInterface({
                        input: process.stdin,
                        output: process.stdout
                    });
                
                    rl.question('Press Enter to exit...', (answer) => {
                        rl.close();
                        reject(err);
                    });
                });
          });
      });
}

async function exec_script_Excel_Hello_World(){
    // shell.exec('code .');
    console.log('Explore the Excel Hello World Add-in project and dive into this Office add-in sample for an immersive experience!');
    console.log('--------------------------------------------------------------------------------------------------------');
  
    return new Promise<boolean>((resolve, reject) => {
      
        let is_vscode_installed = false;

        // Step 1: Get sample code
        console.log('Step [1/3]: Getting sample code...');
        let spinner = new Spinner('Processing.. %s');
        spinner.setSpinnerString('|/-\\');
        spinner.start();

        shell.exec('git clone https://github.com/OfficeDev/Office-Addin-TaskPane-React.git', {async:true}, (code, stdout, stderr) => {

            if (code != 0) {
                //Stop the spinner and clear the console
                spinner.stop(true);
                readline.clearLine(process.stdout, 0);
                readline.cursorTo(process.stdout, 0);

                console.error(stderr);
                console.log('Error occurred when downloading the code. This may be caused by the network issue. Please rerun the command to try again.')
                console.log('--------------------------------------------------------------------------------------------------------');

                let rl = readline.createInterface({
                    input: process.stdin,
                    output: process.stdout
                });
            
                rl.question('Press Enter to exit...', (answer) => {
                    rl.close();
                    reject(stderr);
                });
            }
            else{
                //download succeeded
                shell.cd('./Office-Addin-TaskPane-React');
                shell.exec('npm run convert-to-single-host --if-present -- excel', {async:true}, (code, stdout, stderr) => {
                    spinner.stop(true);

                    //Stop the spinner and clear the console
                    readline.clearLine(process.stdout, 0);
                    readline.cursorTo(process.stdout, 0);

                    console.log('Step [1/3] completed!');
                    console.log('--------------------------------------------------------------------------------------------------------');
            
                    // Ask user if sample Add-in automatic launch is needed
                    let silentStream = new Writable({
                        write(chunk, encoding, callback) {
                            callback();
                        }
                    });

                    let rl = readline.createInterface({
                        input: process.stdin,
                        output: silentStream
                    });

                    let auto_launch_answer = false;

                    console.log('Proceed to launch Office with the sample add-in? (Y/N)');
                    rl.question('', (answer) => {

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
                        } else {
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
                            // Step 3: Provide user the command to side-load add-in directly 
                            console.log('Step [3/3]: Automatically launch add-in with Excel...');
                            console.log('The process is expected to finish shortly, thank you for your patience...');

                            shell.exec('npm set progress always');

                            let env = 'cmd.exe';
                            let para = '/c';
                            if (os.platform() == 'darwin') {
                                env = 'sh';
                                para = '-c';
                            }

                            const install = spawn(env, [para, 'npm install --loglevel verbose']);
                            install.stdout.on('data', (data) => {
                                process.stdout.write(data);
                            });
                        
                            install.stderr.on('data', (data) => {
                                process.stderr.write(data);
                            });
                            install.on('close', (code) => {
                                if (code !== 0) {
                                    //If npm install failed
                                    console.log(`Err: npm install process exited with code ${code}`);

                                    // Error handling on mac
                                    if (os.platform() == 'darwin') {
                                        // if npm install failed because of access issue on mac
                                        if (code == 243) {
                                            console.log('Mac access issue detected. Trying to automatically fix the issue...');

                                            // Get the current user's UID and GID
                                            const uid = process.getuid();
                                            const gid = process.getgid();

                                            shell.exec(`sudo chown -R ${uid}:${gid} ~/.npm`, {async:true}, (code, stdout, stderr) => {
                                                if (code !== 0) {
                                                    console.log(`Err: sudo chown process exited with code ${code}`);
                                                    console.error(`stderr: ${stderr}`);
                                                    console.log('Automatically fix the issue failed. Please try to run the following commands manually:');
                                                    console.log('sudo chown -R $uid:$gid ~/.npm');
                                                    console.log('where $uid and $gid are the current user\'s UID and GID, which can be get by running the following commands:');
                                                    console.log('id -u');
                                                    console.log('id -g');
                                                }
                                                else{
                                                    console.log('Issue fixed. Please try to run the sample command again.');
                                                    console.log('--------------------------------------------------------------------------------------------------------');
                                                    console.log('Hint: If the issue persists, please try to run the following commands manually:');
                                                    console.log('sudo chown -R $uid:$gid ~/.npm');
                                                    console.log('where $uid and $gid are the current user\'s UID and GID, which can be get by running the following commands:');
                                                    console.log('id -u');
                                                    console.log('id -g');
                                                }

                                                const rl = readline.createInterface({
                                                    input: process.stdin,
                                                    output: process.stdout
                                                });
                                            
                                                rl.question('Press Enter to exit...', (answer) => {
                                                    rl.close();
                                                    resolve(is_vscode_installed);
                                                });
                                            });
                                        }
                                    }
                                }
                                else {
                                    //If npm install succeeded
                                    const start = spawn(env, [para, 'npm run start']);

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
                                    
                                        rl.question('Press Enter to exit...', (answer) => {
                                            rl.close();
                                            resolve(is_vscode_installed);
                                        });
                                    });

                                    // Make sure npm run start process will not be blocked by the prompt
                                    start.stdin.write('n\n');
                                    }
                            });
                        }
                        else{
                            // Don't continue with the operations
                            console.log('Step [3/3] skipped. Auto-launch for the sample has been excluded based on your choice.')
                            console.log('And you can initiate the sample add-in by executing the following commands:');
                            console.log('--------------------------------------------------------------------------------------------------------');
                            console.log('npm install');
                            console.log('npm run start');
                            console.log('--------------------------------------------------------------------------------------------------------');
                            FreePortAlert();     
                            console.log('Finished!');     
                            console.log('--------------------------------------------------------------------------------------------------------');        
                            let rl = readline.createInterface({
                                input: process.stdin,
                                output: process.stdout
                            });
                            rl.question('Press Enter to exit...', (answer) => {
                                rl.close();
                                resolve(is_vscode_installed);
                            });
                        }
                    });
                });
            }     
        });
    });
}
  
async function exec_script_Word_Hello_World(){
    // shell.exec('code .');
    console.log('Explore the Excel Hello World Add-in project and dive into this Office add-in sample for an immersive experience!');
    console.log('--------------------------------------------------------------------------------------------------------');
  
    return new Promise<boolean>((resolve, reject) => {
      
        let is_vscode_installed = false;

        // Step 1: Get sample code
        console.log('Step [1/3]: Getting sample code...');
        let spinner = new Spinner('Processing.. %s');
        spinner.setSpinnerString('|/-\\');
        spinner.start();

        shell.exec('git clone https://github.com/OfficeDev/Office-Addin-TaskPane-React.git', {async:true}, (code, stdout, stderr) => {

            if (code != 0) {
                //Stop the spinner and clear the console
                spinner.stop(true);
                readline.clearLine(process.stdout, 0);
                readline.cursorTo(process.stdout, 0);

                console.error(stderr);
                console.log('Error occurred when downloading the code. This may be caused by the network issue. Please rerun the command to try again.')
                console.log('--------------------------------------------------------------------------------------------------------');

                let rl = readline.createInterface({
                    input: process.stdin,
                    output: process.stdout
                });
            
                rl.question('Press Enter to exit...', (answer) => {
                    rl.close();
                    reject(stderr);
                });
            }
            else{
                //download succeeded
                shell.cd('./Office-Addin-TaskPane-React');
                shell.exec('npm run convert-to-single-host --if-present -- word', {async:true}, (code, stdout, stderr) => {
                    spinner.stop(true);

                    //Stop the spinner and clear the console
                    readline.clearLine(process.stdout, 0);
                    readline.cursorTo(process.stdout, 0);

                    console.log('Step [1/3] completed!');
                    console.log('--------------------------------------------------------------------------------------------------------');
            
                    // Ask user if sample Add-in automatic launch is needed
                    let silentStream = new Writable({
                        write(chunk, encoding, callback) {
                            callback();
                        }
                    });

                    let rl = readline.createInterface({
                        input: process.stdin,
                        output: silentStream
                    });

                    let auto_launch_answer = false;

                    console.log('Proceed to launch Office with the sample add-in? (Y/N)');
                    rl.question('', (answer) => {

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
                        } else {
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
                            // Step 3: Provide user the command to side-load add-in directly 
                            console.log('Step [3/3]: Automatically launch add-in with Word...');
                            console.log('The process is expected to finish shortly, thank you for your patience...');

                            shell.exec('npm set progress always');

                            let env = 'cmd.exe';
                            let para = '/c';
                            if (os.platform() == 'darwin') {
                                env = 'sh';
                                para = '-c';
                            }

                            const install = spawn(env, [para, 'npm install --loglevel verbose']);
                            install.stdout.on('data', (data) => {
                                process.stdout.write(data);
                            });
                        
                            install.stderr.on('data', (data) => {
                                process.stderr.write(data);
                            });
                            install.on('close', (code) => {
                                if (code !== 0) {
                                    //If npm install failed
                                    console.log(`Err: npm install process exited with code ${code}`);

                                    // Error handling on mac
                                    if (os.platform() == 'darwin') {
                                        // if npm install failed because of access issue on mac
                                        if (code == 243) {
                                            console.log('Mac access issue detected. Trying to automatically fix the issue...');

                                            // Get the current user's UID and GID
                                            const uid = process.getuid();
                                            const gid = process.getgid();

                                            shell.exec(`sudo chown -R ${uid}:${gid} ~/.npm`, {async:true}, (code, stdout, stderr) => {
                                                if (code !== 0) {
                                                    console.log(`Err: sudo chown process exited with code ${code}`);
                                                    console.error(`stderr: ${stderr}`);
                                                    console.log('Automatically fix the issue failed. Please try to run the following commands manually:');
                                                    console.log('sudo chown -R $uid:$gid ~/.npm');
                                                    console.log('where $uid and $gid are the current user\'s UID and GID, which can be get by running the following commands:');
                                                    console.log('id -u');
                                                    console.log('id -g');
                                                }
                                                else{
                                                    console.log('Issue fixed. Please try to run the sample command again.');
                                                    console.log('--------------------------------------------------------------------------------------------------------');
                                                    console.log('Hint: If the issue persists, please try to run the following commands manually:');
                                                    console.log('sudo chown -R $uid:$gid ~/.npm');
                                                    console.log('where $uid and $gid are the current user\'s UID and GID, which can be get by running the following commands:');
                                                    console.log('id -u');
                                                    console.log('id -g');
                                                }

                                                const rl = readline.createInterface({
                                                    input: process.stdin,
                                                    output: process.stdout
                                                });
                                            
                                                rl.question('Press Enter to exit...', (answer) => {
                                                    rl.close();
                                                    resolve(is_vscode_installed);
                                                });
                                            });
                                        }
                                    }
                                }
                                else {
                                    //If npm install succeeded
                                    const start = spawn(env, [para, 'npm run start']);

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
                                    
                                        rl.question('Press Enter to exit...', (answer) => {
                                            rl.close();
                                            resolve(is_vscode_installed);
                                        });
                                    });

                                    // Make sure npm run start process will not be blocked by the prompt
                                    start.stdin.write('n\n');
                                    }
                            });
                        }
                        else{
                            // Don't continue with the operations
                            console.log('Step [3/3] skipped. Auto-launch for the sample has been excluded based on your choice.')
                            console.log('And you can initiate the sample add-in by executing the following commands:');
                            console.log('--------------------------------------------------------------------------------------------------------');
                            console.log('npm install');
                            console.log('npm run start');
                            console.log('--------------------------------------------------------------------------------------------------------');
                            FreePortAlert();     
                            console.log('Finished!');     
                            console.log('--------------------------------------------------------------------------------------------------------');        
                            let rl = readline.createInterface({
                                input: process.stdin,
                                output: process.stdout
                            });
                            rl.question('Press Enter to exit...', (answer) => {
                                rl.close();
                                resolve(is_vscode_installed);
                            });
                        }
                    });
                });
            }     
        });
    });
}

function replaceUrl(url: string, newUrl: string, filePath: string) {
    return new Promise((resolve, reject) => {
        fs.readFile(filePath, 'utf8', (err, data) => {
            if (err) {
                console.error(err);
                reject(err);
                return;
            }

            const result = data.replace(url, newUrl);

            fs.writeFile(filePath, result, 'utf8', (err) => {
                if (err) {
                    console.error(err);
                    reject(err);
                    return;
                }
                resolve(true);
            });
        });
    });
}

function reportUsageData(scriptName: string, isAutomaticallyLaunch: boolean, isVscodeInstalled: boolean) {
    const projectInfo = {
        ScriptName: [scriptName],
        ScriptType: ["TypeScript"],
        isAutomaticallyLaunch: [isAutomaticallyLaunch],
        isVscodeInstalled: [isVscodeInstalled]
      };

    usageDataObject.reportEvent("sample_scripts", projectInfo);
    
}

function FreePortAlert() {
    if(os.platform() == 'darwin'){
        console.log('Hint: For mac users, if the add-in cannot be loaded correctly because of the add-in you loaded before, please try to run the commands below:');
        console.log('lsof -i:3000');
        console.log('kill -9 <PID>');
        console.log('--------------------------------------------------------------------------------------------------------');
        console.log('Then try to run the add-in again:');
        console.log('npm run start');
        console.log('--------------------------------------------------------------------------------------------------------');
    }
    else if(os.platform() == 'win32'){
        console.log('Hint: For windows users, if the add-in cannot be loaded correctly because of the add-in you loaded before, please try to run the commands below:');
        console.log('netstat -ano | findstr :3000');
        console.log('taskkill /PID <PID> /F');
        console.log('--------------------------------------------------------------------------------------------------------');
        console.log('Then try to run the add-in again:');
        console.log('npm run start');
        console.log('--------------------------------------------------------------------------------------------------------');
    }
}

module.exports = { exec_script_Excel_Mail, exec_script_Word_AIGC, exec_script_Excel_Hello_World, exec_script_Word_Hello_World };