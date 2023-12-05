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
    brew install node
else
    echo "Node.js is already installed!"
fi

# Check if Yeoman and generator-office are installed
if ! npm list -g --depth=0 | grep generator-office &> /dev/null
then
    echo "Yeoman Office is not installed, installing now..."
    npm install -g yo generator-office
else
    echo "Yeoman Office has already been installed."
fi

# Now Yeoman Office has been installed. Create a sample project.
foldername="Office_sample_Excel_Hello_World"
counter=0

while [ -d "$foldername" ]
do
    counter=$((counter + 1))
    foldername="Office_sample_Excel_Hello_World_$counter"
done

yo office --output $foldername --projectType excel_hello_world --no-insight