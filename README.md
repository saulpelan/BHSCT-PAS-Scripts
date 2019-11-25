# CRT-PAS-Scripts
## Background
The BHSCT* has a dedicated team that maintains its PAS*. The following tasks are completed on a regular basis:
1. Setting up and maintaining user accounts
1. Setting up and maintaining clinic sessions (ongoing and ad hoc)
1. Setting up and maintaining virtual printers and print queues
1. Maintaining quality of data (validating information held on records)

After a short time working in the BHSCT PAS team, I came across the feature in CRT that allows users to run scripts. I instantly realised the advantages of this and started creating scripts to automate certain processes. 
## * Terminology
### PAS
Every hospital trust in the UK relies on a PAS (Patient Administration System) to manage its day to day operation. The BHSCT among other trusts in Northern Ireland currently use a PAS called CLINiCOM, a system developed by Shared Medical Systems Limited (SMS UK) in the 1980s. It is a unix mainframe video terminal (VT220) system, originally accessed by terminals but today is accessed by PCs using terminal emulation software.
### CRT
The terminal emulation software used by BHSCT to access its PAS is a program called CRT developed by VanDyke Software, Inc., and it provides the ability to run scripts written in various different programming/scripting languages such as Python, VBScript, Ruby, and possibly others. 
### Scripts
With the use of scripts, certain maintenance tasks can be automated (such as recording deaths based on information from emails from the NI Health & Care Index, set up new users, tidy up/validate recorded patient phone numbers, etc).
CRT relies on a script engine to run scripts written in a particular language. BHSCT doesn't provide its users with script engines of any sort for Ruby, Python, etc - however it does provide the Microsoft Office suite which includes support for VBScript. Therefore all scripts listed are written in VBScript.
### BHSCT
The Belfast Health & Social Care Trust

*Last updated 22.10.19 12:38 Saul Pelan*
