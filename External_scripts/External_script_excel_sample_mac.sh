#!/bin/bash

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
    brew link --overwrite node
else
    echo "Node.js is already installed!"
fi

# Check if Yeoman and generator-office are installed
if ! npm list -g --depth=0 | grep generator-office &> /dev/null
then
    echo "Yeoman Office is not installed, installing now..."
    sudo npm install -g yo generator-office
else
    echo "Yeoman Office has already been installed."
fi

# Now Yeoman Office has been installed. Create a sample project.
foldername="Office_sample_Excel_Mail"
counter=0

while [ -d "$foldername" ]
do
    counter=$((counter + 1))
    foldername="Office_sample_Excel_Mail_$counter"
done

yo office --output $foldername --projectType excel_sample --no-insight