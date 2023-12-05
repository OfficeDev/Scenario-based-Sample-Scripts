#!/bin/bash



# Check if git is installed
if ! command -v git &> /dev/null
then
    echo "Git is not installed, installing now..."
    # Manual installation of Git
    cd ~
    curl -O https://github.com/git/git/archive/v2.31.1.tar.gz
    tar -zxf v2.31.1.tar.gz
    cd git-2.31.1
    make prefix=/usr/local all
    sudo make prefix=/usr/local install
else
    echo "Git is already installed!"
fi

# Check if Node.js is installed
if ! command -v node &> /dev/null
then
    echo "Node.js is not installed, installing now..."
    # Manual installation of Node.js
    cd ~
    curl -O https://nodejs.org/dist/v20.10.0/node-v20.10.0.tar.gz
    tar -zxf node-v20.10.0.tar.gz
    cd node-v20.10.0
    ./configure
    make -j4
    sudo make install
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
foldername="Office_sample_Excel_Hello_World"
counter=0

while [ -d "$foldername" ]
do
    counter=$((counter + 1))
    foldername="Office_sample_Excel_Hello_World_$counter"
done

yo office --output $foldername --projectType excel_hello_world --no-insight