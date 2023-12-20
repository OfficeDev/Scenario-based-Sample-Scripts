#!/bin/bash

# echo "Checking if Xcode Command Line Tools are installed..."
# if ! xcrun --version > /dev/null 2>&1; then
#     echo "Xcode Command Line Tools are not installed or configured correctly."
#     echo "Installing Xcode Command Line Tools..."
#     xcode-select --install
#     echo "Please follow the prompts to install Xcode Command Line Tools."
# else
#     echo "Xcode Command Line Tools are installed."
# fi

# Check if git is installed
if ! git --version > /dev/null 2>&1; then
    echo "Git is not installed, please install git first."
    exit 1
else
    echo "Git is already installed!"
fi
 
# Check if Node.js is installed. If Node is not installed, check homebrew/install homebrew and Node.js 18.
if ! command -v node &> /dev/null
then
    # echo "Node.js is not installed, installing Node.js using homebrew now..."
    echo "Node.js is not installed, please install Node.js first."
    exit

    # # Check if Homebrew is installed
    # if ! command -v brew &> /dev/null
    # then
    #     echo "Homebrew is not installed, installing now..."
        
    #     /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    #     exit_status=$?
    #     echo "Exit status of Installing homebrew: $exit_status"

    #     if [ $exit_status -ne 0 ]; then
    #         echo "An error occurred while running the command. Trying to fix it..."
    #         sudo chown -R $(whoami) /usr/local/share/zsh /usr/local/share/zsh/site-functions
    #     fi
    # else
    #     echo "Homebrew is already installed!"
    # fi

    # brew install node@16
    # brew link --overwrite --force node@16
else
    echo "Node.js is already installed!"
    # check the version of Node.js
    NODE_VERSION=$(node -v)
    # if [[ "$NODE_VERSION" != "v16"* ]]
    # then
    #     echo "The current version of Node.js is not 16 or 18, installing Node.js 18 now..."
    #     brew install node@16
    #     brew link --overwrite --force node@16
    #     exit_status=$?
    #     if [ $exit_status -ne 0 ]; then
    #         echo "An error occurred while linking node. Trying to fix it..."
    #         sudo chown -R $(whoami) /usr/local
    #     fi
    # fi
fi


# Check the version of Node.js
echo "The current version of Node.js is: $(node -v)"
 
# # check if typescript & tsc have been installed
# if ! command -v tsc &> /dev/null
# then
#     echo "TypeScript is not installed, installing now..."
#     sudo npm install -g typescript
# else
#     echo "TypeScript is already installed!"
# fi
 
# # Check the version of npm
# NPM_VERSION=$(npm -v)
# echo "The current version of npm is: $NPM_VERSION"

# # Check if npm version is >=7 and <10
# if (( $(echo "$NPM_VERSION >= 7" | bc -l) )) && (( $(echo "$NPM_VERSION < 10" | bc -l) )); then
#     echo "npm version is in the correct range."
# else
#     echo "npm version is not in the correct range, reinstalling npm to version 9..."
#     npm install -g npm@9
# fi

# Check if office_addin_sample_scripts are installed
if ! npm list -g --depth=0 | grep office_addin_sample_scripts &> /dev/null
then
    echo "office_addin_sample_scripts is not installed, installing now..."
    sudo npm install -g office_addin_sample_scripts
else
    echo "office_addin_sample_scripts has already been installed. Updating to the latest version..."
    sudo npm update -g office_addin_sample_scripts
    echo "office_addin_sample_scripts has been updated to the latest version!"
fi

#Check if Excel is installed
echo "Checking if Excel is installed..."
if [ "$1" != "bypass" ]; then
    if ! find /Applications -name "Microsoft Excel.app" 2>/dev/null | grep -q "Microsoft Excel.app"
    then
        echo "Microsoft Excel is not installed. Please install Microsoft Excel and then rerun the script."
        echo "If you make sure the application is installed, please run the script with "bypass":"
        echo "bash <(curl -L -s aka.ms/exceladdin/mail_mac) bypass"
        exit 1
    fi
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

# sudo chown -R $(whoami) ~/.npm

office_addin_sample_scripts launch excel_mail $foldername