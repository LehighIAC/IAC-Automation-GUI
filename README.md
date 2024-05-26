# IAC-Automation
Automated Python script for Lehigh University Industrial Assessment Center

This tutorial is designed for Mechanical Engineering people with zero programming knowledge.
## Required Software
Of course, you need to install Microsoft Office. https://confluence.cc.lehigh.edu/display/LKB/Windows+or+macOS%3A++Download+and+Install+Office+365
### Windows
Install [VS code](https://code.visualstudio.com/download) and [Anaconda](https://www.anaconda.com/download).
### macOS
Install [homebrew](https://brew.sh) then```brew install --cask visual-studio-code anaconda```
### Linux
LibreOffice compatibility is not guaranteed.

## Setting up Python environment
### Open Anaconda Prompt
```
conda create -n iac python=3.8 
conda activate iac 
```
### Install the Following Packages
```
conda install json5 numpy pandas openpyxl
conda install -c conda-forge python-docx docxcompose easydict
pip install python-docx-replace tkcalendar tktooltip
```
`conda` always has the highest priority. If not available, install packages from `conda-forge`. Don't install from `pip` unless you have to, otherwise there might be dependency issue.
### Configure VS Code
Go to `Extensions`(Ctrl+Shift+X), install `Python` (from Microsoft) and `JSON5 syntax` (from mrmlnc).
Press Ctrl+Shift+P, search `Python: Select Interpreter` and select the `iac` environment you just created.
### NOTE: IF YOU WISH TO REMOVE THIS ENVIRONMENT
```
conda remove --name iac --all
```