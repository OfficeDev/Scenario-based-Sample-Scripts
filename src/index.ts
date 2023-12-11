#!/usr/bin/env node

import * as yargs from 'yargs';
import * as fs from 'fs';
const { exec_script_Excel_Mail, exec_script_Word_AIGC, exec_script_Excel_Hello_World, exec_script_Word_Hello_World } = require('./Sample_scripts/Sample_Excel_Word');

function exec_script(){
    yargs
    .command(
        'launch <sampleType> <sampleFolder>',
        'Launch the sample choosed',
        (yargs) => {
            yargs.positional('sampleType', {
                type: 'string',
                describe: 'The sample type decide to choose'
            })
            .positional('sampleFolder', {
                type: 'string',
                describe: 'The folder decide to store the sample'
            })
        },
        (argv) => {
            console.log(`Launching ${argv.sampleType} for you ...`)
            const currentWorkingDirectory = process.cwd();

            // Check if folder under the sample path is exist
            if (!fs.existsSync(argv.sampleFolder)) {
                console.log('The sample path is valid. Creating folder for you ...');
                fs.mkdirSync(argv.sampleFolder);

                //Check if the folder is created successfully
                if (!fs.existsSync(argv.sampleFolder)) {
                    console.log('Failed to create the folder. Please try a new path.');
                    return;
                }

                //Log the current working directory
                console.log("Create sample folder successfully. The current working directory is: " + currentWorkingDirectory + "\\" + argv.sampleFolder);

                //switch to the sample folder
                process.chdir(argv.sampleFolder);
            }
            else{
                console.log('The sample path is exist. Please try a new path.');
                return;
            }

            if (argv.sampleType == 'excel_hello_world') {
                exec_script_Excel_Hello_World();
            }
            else if (argv.sampleType == 'word_hello_world') {
                exec_script_Word_Hello_World();
            }
            else if (argv.sampleType == 'excel_mail') {
                exec_script_Excel_Mail();
            }
            else if (argv.sampleType == 'word_aigc') {
                exec_script_Word_AIGC();
            }
            else {
                console.log('Please enter the correct sample type.');
            }
        }
    )
    .help()
    .argv;
}

exec_script();



