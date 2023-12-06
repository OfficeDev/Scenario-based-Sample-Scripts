import yargs from 'yargs';
const { exec_script_Excel_Mail, exec_script_Word_AIGC } = require('./Sample_scripts/Sample_Excel_Word');
const { exec_script_Excel_Hello_World, exec_script_Word_Hello_World } = require('./Sample_scripts/Sample_Hello_world_script');

yargs
    .command(
        'launch <sampleType> <sampleName>',
        'Launch the sample choosed',
        (yargs) => {
            yargs.positional('sampleType', {
                type: 'string',
                describe: 'The sample type decide to choose'
            })
        },
        (argv) => {
            console.log(`Launching ${argv.sampleType} for you ...`)
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

