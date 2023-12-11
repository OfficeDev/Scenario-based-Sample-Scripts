#!/bin/bash
 
# Check if Homebrew is installed
if ! command -v brew &> /dev/null
then
    echo "Homebrew is not installed, installing now..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
else
    echo "Homebrew is already installed!"
fi
 
# Check if git is installed
if ! command -v git &> /dev/null
then
    echo "Git is not installed, installing now..."
    brew install git
else
    echo "Git is already installed!"
fi
 
# Check if Node.js is installed
if ! command -v node &> /dev/null
then
    echo "Node.js is not installed, installing now..."
    brew install node@18
    brew link --overwrite --force node@18
else
    echo "Node.js is already installed!"
    #check the version of Node.js
    NODE_VERSION=$(node -v)
    if [[ "$NODE_VERSION" != "v16"*  && "$NODE_VERSION" != "v18"* ]]
    then
        echo "The current version of Node.js is not 16 or 18, installing Node.js 18 now..."
        brew install node@18
        brew link --overwrite --force node@18
    fi
fi
 
# Check the version of Node.js
echo "The current version of Node.js is: $(node -v)"
 
# check if typescript & tsc have been installed
if ! command -v tsc &> /dev/null
then
    echo "TypeScript is not installed, installing now..."
    npm install -g typescript
else
    echo "TypeScript is already installed!"
fi
 
# Check the version of npm
echo "The current version of npm is: $(npm -v)"
 
# Check if office_addin_sample_scripts are installed
if ! npm list -g --depth=0 | grep office_addin_sample_scripts &> /dev/null
then
    echo "office_addin_sample_scripts is not installed, installing now..."
    npm install -g office_addin_sample_scripts
else
    echo "office_addin_sample_scripts has already been installed."
fi
 
# Now Office add-in sample scripts have been installed. Create a sample project.
foldername="Office_sample_Excel_Mail"
counter=0
 
while [ -d "$foldername" ]
do
    counter=$((counter + 1))
    foldername="Office_sample_Excel_Mail_$counter"
done

#Automatically clear port 3000:
pid=$(lsof -t -i:3000)
if [ -n "$pid" ]; then
    echo "Port 3000 is in use by PID $pid. Killing..."
    kill -9 $pid
else
    echo "Port 3000 is not in use."
fi
 
office_addin_sample_scripts launch excel_mail $foldername