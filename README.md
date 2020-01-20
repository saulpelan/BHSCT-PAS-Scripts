<img src="images/BHSCT%20Logo%20in%20Colour%20Cropped.jpg" height="50px">

# Introduction
This repo is a portfolio of the scripts I have developed for the Belfast Health and Social Care Trust PAS (**P**atient **A**dministration **S**ystem) Support Team to aid the maintenance of the three Patient Administration Systems.

## CliniCom Patient Administration System
The three Belfast Trust hospitals each run on a PAS system called CliniCom, a mainframe system developed roughly in the late 1980s by Shared Medical Systems Limited (SMS UK) and now supported by [DXC Technology](https://dxc.technology). Clinicom was originally designed to be accessed by "dumb" terminals but is now accessed via terminal emulation software on PCs.

![CliniCom AMS Function Set Menu](images/CliniCom%20AMS%20Function%20Set.PNG)

## CRT 
The terminal emulation software the BHSCT uses to access PAS is [CRT by VanDyke Software, Inc](https://www.vandyke.com/download/crt/index.html). CRT allows the user to run scripts in various scripting languages as long as a script engine is installed for a particular scripting language. The only script engine available on BHSCT PCs is Microsoft's VBScript.

# Scripts
Certain maintenance tasks on CliniCom are done on a very regular basis by the PAS Team for example:
 * Batch recording deaths on the system, using information from alert emails
 * Validating phone numbers held on digital patient records
 * Setting up and maintaining clinic sessions/timeslots
 * User account maintenance
 
Due to CliniCom being a 24 row x 80 column terminal system, performing these tasks can require a lot of navigation through menus and functions. The purpose of the scripts in **CRT-PAS-Scripts** is to:
 * Perform full tasks at the single touch of a key
 * Extract information from the screen and export it to more user friendly format, such as Excel spreadsheets
 * Assist with navigation
